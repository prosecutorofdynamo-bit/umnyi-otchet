import pandas as pd
import io
import unicodedata
import re

# === Константы ===
OUTSIDE = "шлюз"
INSIDE_HINT = "офис"
DEDUP_WINDOW_MIN = 3        # слипание дублей (минуты)
CORE_START_H, CORE_START_M = 9, 0
CORE_END_H,   CORE_END_M   = 18, 0
DAY_CORE_MIN = 8 * 60       # 8 часов ядра
LATE_H, LATE_M = 9, 1       # опоздание с 09:01


def fmt_hm(m) -> str:
    """минуты -> 'Xч Yмин' (0 -> '0ч 0мин', пустое если NaN)."""
    if m is None or pd.isna(m):
        return ""
    try:
        m = int(m)
    except Exception:
        return ""
    if m < 0:
        m = 0
    h, mm = divmod(m, 60)
    return f"{h}ч {mm}мин"


def fio_norm(s: str) -> str:
    s = "" if pd.isna(s) else str(s)
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("ё", "е").replace("Ё", "Е")
    s = " ".join(s.strip().split()).lower()
    return s


def work_day(ts):
    """Рабочие сутки 06:00–06:00."""
    ts = pd.to_datetime(ts)
    return (ts - pd.Timedelta(days=1)).date() if ts.hour < 6 else ts.date()


def norm(s):
    s = "" if pd.isna(s) else str(s)
    return unicodedata.normalize("NFKC", s).strip().casefold()


def inside_minutes_between(
    grp: pd.DataFrame,
    right_col: str,
    a: pd.Timestamp,
    b: pd.Timestamp,
) -> int:
    """
    Сколько минут сотрудник был ВНУТРИ офиса в окне [a, b].
    Основано на направлениях (офис/шлюз).
    """
    if grp is None or grp.empty or a >= b:
        return 0

    g = grp.sort_values("Дата события")[[ "Дата события", right_col ]].copy()
    g["dest_n"] = g[right_col].map(norm)

    start_look = a - pd.Timedelta(hours=6)
    sec = g[(g["Дата события"] >= start_look) & (g["Дата события"] <= b)].copy()
    sec["label"] = sec["dest_n"].apply(
        lambda s: "in" if INSIDE_HINT in s else ("out" if OUTSIDE in s else None)
    )
    sec = sec.dropna(subset=["label"]).reset_index(drop=True)

    # состояние на момент a
    hist = g[g["Дата события"] <= a]
    last_dest = hist.iloc[-1]["dest_n"] if len(hist) else ""
    inside = OUTSIDE not in last_dest

    # дедуп одинаковых подряд меток
    ded = []
    for _, row in sec.iterrows():
        t, lab = row["Дата события"], row["label"]
        if ded:
            t_prev, lab_prev = ded[-1]
            if lab == lab_prev and (t - t_prev).total_seconds() / 60.0 <= DEDUP_WINDOW_MIN:
                continue
        ded.append((t, lab))

    mins = 0.0
    last_t = a
    for t, lab in ded:
        t_clamp = min(max(t, a), b)
        if inside:
            mins += max(0.0, (t_clamp - last_t).total_seconds() / 60.0)
        inside = lab == "in"
        last_t = t_clamp
        if last_t >= b:
            break

    if last_t < b and inside:
        mins += (b - last_t).total_seconds() / 60.0

    return int(round(mins))


def longest_outside_gap_between(
    grp: pd.DataFrame,
    right_col: str,
    a: pd.Timestamp,
    b: pd.Timestamp,
):
    """
    Самый длинный непрерывный интервал 'вне офиса' в окне [a,b].
    Возвращает (gap_min, t_from, t_to).
    """
    if grp is None or grp.empty or a >= b:
        return 0, None, None

    g = grp.sort_values("Дата события")[[ "Дата события", right_col ]].copy()
    g["dest_n"] = g[right_col].map(norm)

    start_look = a - pd.Timedelta(hours=6)
    sec = g[(g["Дата события"] >= start_look) & (g["Дата события"] <= b)].copy()
    sec["label"] = sec["dest_n"].apply(
        lambda s: "in" if INSIDE_HINT in s else ("out" if OUTSIDE in s else None)
    )
    sec = sec.dropna(subset=["label"]).reset_index(drop=True)

    # состояние на момент a
    hist = g[g["Дата события"] <= a]
    last_dest = hist.iloc[-1]["dest_n"] if len(hist) else ""
    outside = OUTSIDE in last_dest

    ded = []
    for _, row in sec.iterrows():
        t, lab = row["Дата события"], row["label"]
        if ded:
            t_prev, lab_prev = ded[-1]
            if lab == lab_prev and (t - t_prev).total_seconds() / 60.0 <= DEDUP_WINDOW_MIN:
                continue
        ded.append((t, lab))

    best = 0.0
    best_a = None
    best_b = None
    last_t = a

    for t, lab in ded:
        t_clamp = min(max(t, a), b)
        if outside:
            gap = max(0.0, (t_clamp - last_t).total_seconds() / 60.0)
            if gap > best:
                best, best_a, best_b = gap, last_t, t_clamp
        outside = lab == "out"
        last_t = t_clamp
        if last_t >= b:
            break

    if last_t < b and outside:
        gap = (b - last_t).total_seconds() / 60.0
        if gap > best:
            best, best_a, best_b = gap, last_t, b

    return int(round(best)), best_a, best_b


def compute_outside_table(df: pd.DataFrame, right_col: str) -> pd.DataFrame:
    """
    Таблица «Вне офиса» по каждому (ФИО, Рабочий_день).
    right_col = 'Вход' или 'Выход' — по какой колонке считать направления.
    """
    rows = []

    for (fio, day), grp in df.groupby(["ФИО", "Рабочий_день"], sort=False):
        base = pd.Timestamp(day).normalize()
        start0600 = base + pd.Timedelta(hours=6)
        end0600 = start0600 + pd.Timedelta(days=1)

        first = grp["Дата события"].min()
        last = grp["Дата события"].max()

        # окно ядра 09–18
        a = base + pd.Timedelta(hours=CORE_START_H, minutes=CORE_START_M)
        b = base + pd.Timedelta(hours=CORE_END_H,   minutes=CORE_END_M)
        a = max(a, start0600)
        b = min(b, end0600)

        if a >= b:
            out_core_min = 0
            gap_period = ""
        else:
            total_core = (b - a).total_seconds() / 60.0
            in_core = inside_minutes_between(grp, right_col, a, b)
            out_core_min = max(0.0, total_core - in_core)
            gap_min, g_a, g_b = longest_outside_gap_between(grp, right_col, a, b)
            gap_period = (
                f"{g_a:%H:%M}–{g_b:%H:%M}"
                if (gap_min and gap_min >= 120 and g_a and g_b)
                else ""
            )

        rows.append(
            {
                "ФИО": fio,
                "Дата": base.date(),
                "Время прихода": first.strftime("%H:%M") if pd.notna(first) else "",
                "Время ухода": last.strftime("%H:%M") if pd.notna(last) else "",
                "Вне_ядра_мин": int(round(out_core_min)),
                "Отсутствие более 2 часов подряд": gap_period,
            }
        )

    res = pd.DataFrame(rows).sort_values(["ФИО", "Дата"])
    res["Вне офиса"] = res["Вне_ядра_мин"].apply(lambda m: f"{m // 60}ч {m % 60}мин")
    return res[
        [
            "ФИО",
            "Дата",
            "Время прихода",
            "Время ухода",
            "Вне офиса",
            "Отсутствие более 2 часов подряд",
            "Вне_ядра_мин",
        ]
    ]


# --- Фильтрация «не людей» (карты, клининг и т.п.) ---
NONPERSON_TOKENS = [
    "студент",
    "клининг",
    "уборщ",
    "водител",
    "охран",
    "технич",
    "персонал",
    "инженер без",
    "без фио",
    "безфио",
    "аэростар",
    "aerostar",
    "техносервис",
    "техно-сервис",
    "техносерв",
    "отель",
    "гостиниц",
    "стажер",
    "стажёр",
    "практикант",
    "интерн",
    "ассистент",
    "ученик",
]
WHOLE_WORD_TOKENS = ["ооо", "оао", "пао", "зао", "ип"]
EXCLUDE_NAME_ALIASES = {"пелешок", "пешелка"}


def is_nonperson(fio: str) -> bool:
    s = "" if fio is None else str(fio)
    s = unicodedata.normalize("NFKC", s).strip().casefold()
    if not s:
        return True
    if any(alias in s for alias in EXCLUDE_NAME_ALIASES):
        return True
    if any(tok in s for tok in NONPERSON_TOKENS):
        return True
    if re.search(r"\b(?:" + "|".join(map(re.escape, WHOLE_WORD_TOKENS)) + r")\b", s):
        return True
    if any(ch.isdigit() for ch in s):
        return True
    return False


def read_journal(file_obj) -> pd.DataFrame:
    """
    Читаем журнал проходов из Excel.
    Ожидаем колонки:
    ['Событие','Дата события','Фамилия','Имя','Отчество','Вход','Выход']
    """
    need = ["Событие", "Дата события", "Фамилия", "Имя", "Отчество", "Вход", "Выход"]

    content = file_obj.read()
    df_raw = None
    for skip in (3, 0, 1, 2):
        try:
            _tmp = pd.read_excel(io.BytesIO(content), engine="openpyxl", skiprows=skip)
            if set(need).issubset(_tmp.columns):
                df_raw = _tmp
                break
        except Exception:
            continue

    if df_raw is None:
        raise RuntimeError(
            "Не удалось прочитать журнал: не найдены нужные колонки "
            f"(ожидались: {need}). Проверьте формат файла."
        )

    df = df_raw[need].copy()
    df["Событие_n"] = df["Событие"].apply(norm)
    df = df[
        df["Событие_n"].str.contains("проход по идентификатору", na=False)
    ].copy()

    for c in ["Фамилия", "Имя", "Отчество"]:
        df[c] = df[c].where(df[c].notna(), "").astype(str).str.strip()

    def _join_fio(row):
        parts = [row["Фамилия"], row["Имя"], row["Отчество"]]
        return " ".join(p for p in parts if p)

    df["ФИО"] = (
        df.apply(_join_fio, axis=1)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )

    df["Дата события"] = pd.to_datetime(df["Дата события"], errors="coerce")
    df = df.dropna(subset=["Дата события"]).sort_values("Дата события")
    df["Рабочий_день"] = df["Дата события"].apply(work_day)

    df["Вход_n"] = df["Вход"].apply(norm)
    df["Выход_n"] = df["Выход"].apply(norm)
    ok = ~(
        df["Вход_n"].str.contains("неконтролируем", na=False)
        | df["Выход_n"].str.contains("неконтролируем", na=False)
    )
    df = df[ok].copy()

    df = df[~df["ФИО"].apply(is_nonperson)].copy()

    return df


def read_kadry(file_obj) -> pd.DataFrame:
    """
    Читаем кадровый файл и разворачиваем интервалы в посуточный список.
    Ожидаем колонки: 'Сотрудник', 'Вид отсутствия', 'с', 'до'.
    """
    kadry = pd.read_excel(file_obj, header=None)

    # ищем строку, где в любой колонке есть 'Сотрудник'
    def _is_sotr_cell(x):
        s = "" if pd.isna(x) else str(x)
        return s.strip().casefold() == "сотрудник"

    mask_rows = kadry.apply(lambda row: row.map(_is_sotr_cell).any(), axis=1)
    idxs = kadry.index[mask_rows]
    if len(idxs) == 0:
        raise RuntimeError(
            "Не удалось найти строку с заголовком 'Сотрудник' в кадровом файле."
        )

    hdr_row = idxs[0]
    kadry.columns = kadry.iloc[hdr_row]
    kadry = kadry.iloc[hdr_row + 1 :]

    kadry = kadry.rename(
        columns={
            "Сотрудник": "ФИО",
            "Вид отсутствия": "Тип",
            "с": "Дата_с",
            "до": "Дата_по",
        }
    )
    kadry = kadry[["ФИО", "Тип", "Дата_с", "Дата_по"]].copy()
    kadry = kadry.dropna(subset=["ФИО", "Тип"], how="any")

    for col in ["Дата_с", "Дата_по"]:
        kadry[col] = kadry[col].astype(str).str.extract(
            r"(\d{2}\.\d{2}\.\d{4})", expand=False
        )
        kadry[col] = pd.to_datetime(kadry[col], dayfirst=True, errors="coerce")

    kadry["Дата_по"] = kadry["Дата_по"].fillna(kadry["Дата_с"])

    rows = []
    for _, r in kadry.iterrows():
        d1, d2 = r["Дата_с"], r["Дата_по"]
        if pd.isna(d1) or pd.isna(d2):
            continue
        for d in pd.date_range(d1, d2, freq="D"):
            rows.append({"ФИО": r["ФИО"], "Дата": d.date(), "Тип": r["Тип"]})

    kadry_dates = pd.DataFrame(rows)

    # замена «гос. обязанности» -> «Сдача крови»
    kadry_dates["Тип"] = kadry_dates["Тип"].replace(
        to_replace=r"(?i).*гос.*обязан.*", value="Сдача крови", regex=True
    )

    return kadry_dates


# --- Доп. логика: опоздания, выходы, флаг "возможен проход вне терминала" ---


def _core_window_for_day(day):
    base = pd.Timestamp(day).normalize()
    a = base + pd.Timedelta(hours=CORE_START_H, minutes=CORE_START_M)
    b = base + pd.Timedelta(hours=CORE_END_H, minutes=CORE_END_M)
    return a, b


def _calc_group_stats(df: pd.DataFrame):
    """
    Для каждого (ФИО, Рабочий_день):
      - первый/последний проход
      - длительность
      - опоздание / вовремя
    """
    rows = []
    for (fio, day), grp in df.groupby(["ФИО", "Рабочий_день"], sort=False):
        grp = grp.sort_values("Дата события")
        first_ts = grp["Дата события"].iloc[0]
        last_ts = grp["Дата события"].iloc[-1]
        dur_min = int(
            (last_ts - first_ts).total_seconds() / 60.0
        ) if pd.notna(last_ts) and pd.notna(first_ts) else 0

        # порог опоздания 09:01
        plan_start = pd.Timestamp(day) + pd.Timedelta(hours=LATE_H, minutes=LATE_M)
        late_min = (
            first_ts - plan_start
        ).total_seconds() / 60.0 if pd.notna(first_ts) else 0
        status = "опоздание" if late_min > 0 else "вовремя"

        rows.append(
            {
                "ФИО": fio,
                "Дата": pd.to_datetime(day).date(),
                "first_ts": first_ts,
                "last_ts": last_ts,
                "Продолжительность_мин": max(dur_min, 0),
                "Опоздание": status,
            }
        )

    st = pd.DataFrame(rows)
    st["Общее время"] = st["Продолжительность_мин"].apply(fmt_hm)
    return st


def _calc_exits_and_suspect(df: pd.DataFrame, right_col: str):
    """
    Для каждого дня:
      - количество выходов в ядре (09–18)
      - флаг 'suspect' (возможен проход вне терминала)
    """
    rows = []

    for (fio, day), grp in df.groupby(["ФИО", "Рабочий_день"], sort=False):
        base = pd.Timestamp(day).normalize()
        a_core, b_core = _core_window_for_day(base)

        g = grp.sort_values("Дата события")[[ "Дата события", right_col ]].copy()
        g["dest_n"] = g[right_col].map(norm)

        # только события в ядре
        g = g[(g["Дата события"] >= a_core) & (g["Дата события"] <= b_core)]

        labels = []
        times = []
        for _, r in g.iterrows():
            s = r["dest_n"]
            lab = "in" if INSIDE_HINT in s else ("out" if OUTSIDE in s else None)
            if lab is None:
                continue
            t = r["Дата события"]
            # дедуп дрожания
            if labels:
                t_prev = times[-1]
                lab_prev = labels[-1]
                if (
                    lab == lab_prev
                    and (t - t_prev).total_seconds() / 60.0 <= DEDUP_WINDOW_MIN
                ):
                    continue
            labels.append(lab)
            times.append(t)

        # выходы: переход in -> out
        exits = 0
        for i in range(1, len(labels)):
            if labels[i - 1] == "in" and labels[i] == "out":
                exits += 1

        # suspect: два одинаковых подряд события с разрывом > DEDUP_WINDOW_MIN
        suspect = False
        for i in range(1, len(labels)):
            if labels[i] == labels[i - 1]:
                gap = (times[i] - times[i - 1]).total_seconds() / 60.0
                if gap > DEDUP_WINDOW_MIN:
                    suspect = True
                    break

        rows.append(
            {
                "ФИО": fio,
                "Дата": base.date(),
                "Выходы": exits,
                "suspect": suspect,
            }
        )

    return pd.DataFrame(rows)


def build_report(journal_file, kadry_file) -> pd.DataFrame:
    """
    Главная функция: получает два файла (журнал и кадры) и
    возвращает готовый pandas.DataFrame для выгрузки в Excel.
    """
    # 1) читаем журнал
    df = read_journal(journal_file)

    # 2) выбираем нужную колонку для направлений ('Вход' или 'Выход')
    def _total_outside(col):
        t = compute_outside_table(df, col)
        return pd.to_numeric(t["Вне_ядра_мин"], errors="coerce").fillna(0).sum()

    sum_exit = _total_outside("Выход")
    sum_entry = _total_outside("Вход")
    right_col = "Вход" if sum_entry <= sum_exit else "Выход"

    # 3) таблица "вне офиса"
    out_df = compute_outside_table(df, right_col)

    # 4) опоздания и общее время
    stats_df = _calc_group_stats(df)
    out_df = out_df.merge(
        stats_df[["ФИО", "Дата", "Продолжительность_мин", "Общее время", "Опоздание"]],
        on=["ФИО", "Дата"],
        how="left",
    )

    # 5) выходы и флаг "возможен проход вне терминала"
    ex_df = _calc_exits_and_suspect(df, right_col)
    out_df = out_df.merge(ex_df, on=["ФИО", "Дата"], how="left")
    out_df["Выходы"] = out_df["Выходы"].fillna(0).astype(int)

    # добавляем надпись к "Вне офиса"
    note = "возм. проход вне терминала"
    out_df["Вне офиса"] = out_df.apply(
        lambda r: f"{r['Вне офиса']}\n{note}" if bool(r.get("suspect", False)) else r["Вне офиса"],
        axis=1,
    )

    # 6) дневной итог и недоработки
    dur = pd.to_numeric(out_df["Продолжительность_мин"], errors="coerce").fillna(0)

    # обед: если смена >= 60 мин — 60 минут, иначе 0
    lunch = dur.apply(lambda x: 60 if x >= 60 else 0)

    # штраф за "вне ядра": буфер 60 мин
    out_core = pd.to_numeric(out_df["Вне_ядра_мин"], errors="coerce").fillna(0)
    penalty = (out_core - 60).clip(lower=0)
    penalty = penalty.where(dur >= 60, 0)  # если смена < 60 мин, не штрафуем

    # эффективное время за день
    eff_day = (dur - lunch - penalty).clip(lower=0)
    out_df["Итого_дня_мин"] = eff_day.astype(int)
    out_df["Итого за день"] = out_df["Итого_дня_мин"].apply(fmt_hm)

    # недоработки (по ядру)
    out_df["Недоработки_мин"] = penalty.astype(int)
    out_df["Недоработки"] = out_df["Недоработки_мин"].apply(
        lambda m: fmt_hm(m) if m > 0 else ""
    )

    # 7) недельный итог (по фактически присутствующим дням)
    out_df["Дата_dt"] = pd.to_datetime(out_df["Дата"])
    out_df["week_monday"] = out_df["Дата_dt"] - out_df["Дата_dt"].dt.weekday * pd.Timedelta(
        days=1
    )

    week_sums = (
        out_df.groupby(["ФИО", "week_monday"])["Итого_дня_мин"].sum().reset_index()
    )
    week_sums.rename(columns={"Итого_дня_мин": "Итого_нед_мин"}, inplace=True)

    out_df = out_df.merge(
        week_sums, on=["ФИО", "week_monday"], how="left"
    )

    # Для читабельности: заполняем "Итого за неделю" только на последнем дне недели сотрудника
    out_df.sort_values(["ФИО", "Дата_dt"], inplace=True)
    out_df["Итого за неделю"] = ""

    for (fio, w), sub in out_df.groupby(["ФИО", "week_monday"], sort=False):
        if sub.empty:
            continue
        idx_last = sub.index[-1]
        val = sub["Итого_нед_мин"].iloc[0]
        out_df.at[idx_last, "Итого за неделю"] = fmt_hm(val)

        # 8) подмешиваем кадровые отсутствия
    kadry_dates = read_kadry(kadry_file)

    # ❗ Берём из кадров только те даты, которые есть в отчёте по проходам
    valid_dates = out_df["Дата_dt"].dt.date.unique()
    kadry_dates = kadry_dates[kadry_dates["Дата"].isin(valid_dates)].copy()

    # ключи для склейки
    out_df["Дата_key"] = out_df["Дата_dt"].dt.date
    kadry_dates["Дата_key"] = kadry_dates["Дата"]

    out_df["ФИО_key"] = out_df["ФИО"].astype(str).str.strip().str.lower()
    kadry_dates["ФИО_key"] = kadry_dates["ФИО"].astype(str).str.strip().str.lower()

    # добавим в кадры исходное ФИО, чтобы подтянуть его, если проходов не было
    kadry_merge = kadry_dates[["ФИО_key", "Дата_key", "Тип", "ФИО"]].rename(
        columns={"ФИО": "ФИО_кадры"}
    )

    # ВАЖНО: объединяем "снаружи", чтобы дни только из кадров тоже попали
    final = out_df.merge(
        kadry_merge,
        on=["ФИО_key", "Дата_key"],
        how="outer",
    )

    # восстанавливаем ФИО и дату там, где проходов не было
    final["ФИО"] = final["ФИО"].fillna(final["ФИО_кадры"])
    final["Дата_dt"] = final["Дата_dt"].fillna(
        pd.to_datetime(final["Дата_key"])
    )

    # причина отсутствия (может быть пустой, если просто обычный рабочий день)
    final["Причина отсутствия"] = final["Тип"]

    # 9) финальная косметика
    final["Дата"] = final["Дата_dt"].dt.strftime("%d-%m-%Y")

    # чистим NaN в текстовых колонках, чтобы в Excel не было "nan"
    for col in [
        "Опоздание",
        "Общее время",
        "Вне офиса",
        "Отсутствие более 2 часов подряд",
        "Итого за день",
        "Итого за неделю",
        "Недоработки",
        "Причина отсутствия",
    ]:
        if col in final.columns:
            final[col] = final[col].fillna("")

    # числа без проходов = 0
    if "Выходы" in final.columns:
        final["Выходы"] = final["Выходы"].fillna(0).astype(int)
    if "Вне_ядра_мин" in final.columns:
        final["Вне_ядра_мин"] = final["Вне_ядра_мин"].fillna(0).astype(int)

    cols_order = [
        "ФИО",
        "Дата",
        "Время прихода",
        "Время ухода",
        "Опоздание",
        "Общее время",
        "Вне офиса",
        "Выходы",
        "Отсутствие более 2 часов подряд",
        "Итого за день",
        "Итого за неделю",
        "Недоработки",
        "Причина отсутствия",
        "Вне_ядра_мин",
    ]
    for c in cols_order:
        if c not in final.columns:
            final[c] = ""

    final = final[cols_order].copy()
    
    # 🔒 Маскировка ФИО для демонстрации (Сотрудник 001, Сотрудник 002, ...)
    unique_fios = final['ФИО'].unique()
    fio_map = {fio: f"Сотрудник {i+1:03d}" for i, fio in enumerate(unique_fios)}
    final['ФИО'] = final['ФИО'].map(fio_map)

    # Убираем служебные столбцы, которых не должно быть в финальном отчёте
    final = final.drop(columns=['Вне_ядра_мин'], errors='ignore')
    
    return final
