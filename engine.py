import pandas as pd
import io
import unicodedata
import re
from datetime import datetime, date

# === Умный парсер дат ===
def smart_parse_date(x):
    """
    Аккуратный разбор дат:
    - если уже Timestamp/датa → просто приводим к pandas
    - если Excel-число → пытаемся трактовать как серию Excel
    - если строка → пробуем несколько форматов и общий to_datetime(dayfirst=True)
    - при неуспехе → NaT
    """
    # Уже Timestamp или date
    if isinstance(x, (pd.Timestamp, datetime, date)):
        return pd.to_datetime(x, errors="coerce")

    # Пустое / NaN / None
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return pd.NaT

    # Excel-сериал (целое или float)
    if isinstance(x, (int, float)):
        try:
            # стандартный Excel origin
            return pd.to_datetime(x, origin="1899-12-30", unit="D")
        except Exception:
            return pd.NaT

    # Остальное считаем строкой
    s = str(x).strip()
    if not s or s.lower() in {"nan", "none", "nat"}:
        return pd.NaT

    # немного чистим: 21-11-24 → 21.11.24; 2024/11/21 → 2024.11.21
    s_clean = re.sub(r"[-/]", ".", s)

    # Несколько популярных форматов
    for fmt in ("%d.%m.%Y", "%d.%m.%y", "%Y.%m.%d"):
        try:
            return datetime.strptime(s_clean, fmt)
        except ValueError:
            pass

    # Общий резервный вариант: пусть pandas попробует
    try:
        return pd.to_datetime(s, dayfirst=True, errors="coerce")
    except Exception:
        return pd.NaT


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

    g = grp.sort_values("Дата события")[["Дата события", right_col]].copy()
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

    g = grp.sort_values("Дата события")[["Дата события", right_col]].copy()
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


# ===================== ЧТЕНИЕ ЖУРНАЛА =====================

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

    # умный разбор даты
    df["Дата события"] = df["Дата события"].apply(smart_parse_date)
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


# ===================== ЧТЕНИЕ КАДРОВОГО ФАЙЛА =====================

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
    kadry = kadry["ФИО Тип Дата_с Дата_по".split()].copy()
    kadry = kadry.dropna(subset=["ФИО", "Тип"], how="any")

    # умный разбор дат
    for col in ["Дата_с", "Дата_по"]:
        kadry[col] = kadry[col].apply(smart_parse_date)

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


# --- Доп. логика: опоздания, выходы ---

def _core_window_for_day(day):
    base = pd.Timestamp(day).normalize()
    a = base + pd.Timedelta(hours=CORE_START_H, minutes=CORE_START_M)
    b = base + pd.Timedelta(hours=CORE_END_H, minutes=CORE_END_M)
    return a, b


def _calc_group_stats(df: pd.DataFrame):
    """
    Для каждого (ФИО, Рабочий_день):
      - длительность
      - опоздание / вовремя
    """
    rows = []
    for (fio, day), grp in df.groupby(["ФИО", "Рабочий_день"], sort=False):
        grp = grp.sort_values("Дата события")
        first_ts = grp["Дата события"].iloc[0]
        last_ts = grp["Дата события"].iloc[-1]
        if pd.notna(last_ts) and pd.notna(first_ts):
            dur_min = int((last_ts - first_ts).total_seconds() / 60.0)
        else:
            dur_min = 0

        # порог опоздания 09:01
        plan_start = pd.Timestamp(day) + pd.Timedelta(hours=LATE_H, minutes=LATE_M)
        if pd.notna(first_ts):
            late_min = (first_ts - plan_start).total_seconds() / 60.0
        else:
            late_min = 0
        status = "опоздание" if late_min > 0 else "вовремя"

        rows.append(
            {
                "ФИО": fio,
                "Дата": pd.to_datetime(day).date(),
                "Продолжительность_мин": max(dur_min, 0),
                "Опоздание": status,
            }
        )

    st_df = pd.DataFrame(rows)
    st_df["Общее время"] = st_df["Продолжительность_мин"].apply(fmt_hm)
    return st_df


def _calc_exits(df: pd.DataFrame, right_col: str):
    """
    Для каждого дня:
      - количество выходов в ядре (09–18)
    """
    rows = []

    for (fio, day), grp in df.groupby(["ФИО", "Рабочий_день"], sort=False):
        base = pd.Timestamp(day).normalize()
        a_core, b_core = _core_window_for_day(base)

        g = grp.sort_values("Дата события")[["Дата события", right_col]].copy()
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

        rows.append(
            {
                "ФИО": fio,
                "Дата": base.date(),
                "Выходы": exits,
            }
        )

    return pd.DataFrame(rows)


# === Ключ для сопоставления ФИО (журнал ↔ кадры) ===
def fio_match_key(s):
    s = "" if pd.isna(s) else str(s)
    s = unicodedata.normalize("NFKC", s)         # нормализуем символы и пробелы
    s = s.replace("ё", "е").replace("Ё", "Е")   # убираем различие Ё/Е
    s = re.sub(r"\s+", " ", s)                  # множественные пробелы → один
    return s.strip().lower()                    # обрезаем края, в нижний регистр


# ===================== ГЛАВНАЯ ФУНКЦИЯ ОТЧЁТА =====================

def build_report(journal_file, kadry_file=None) -> pd.DataFrame:
    """
    Главная функция: получает файл журнала и, при наличии, кадровый файл.
    Возвращает готовый pandas.DataFrame для выгрузки в Excel.
    Если kadry_file = None, колонка «Причина отсутствия» остаётся пустой.
    """
    # читаем журнал
    df = read_journal(journal_file)

    # автоматически выбираем колонку для направлений ('Вход' или 'Выход')
    def _total_outside(col):
        t = compute_outside_table(df, col)
        return pd.to_numeric(t["Вне_ядра_мин"], errors="coerce").fillna(0).sum()

    sum_exit = _total_outside("Выход")
    sum_entry = _total_outside("Вход")
    right_col = "Вход" if sum_entry <= sum_exit else "Выход"

    # пересчитываем окончательную таблицу «Вне офиса» по выбранной колонке
    out_df = compute_outside_table(df, right_col)

    # доп. статистика: опоздания, общее время
    stats_df = _calc_group_stats(df)
    exits_df = _calc_exits(df, right_col)

    # объединяем всё
    final = out_df.merge(
        stats_df[["ФИО", "Дата", "Опоздание", "Продолжительность_мин", "Общее время"]],
        on=["ФИО", "Дата"],
        how="left",
    )
    final = final.merge(
        exits_df[["ФИО", "Дата", "Выходы"]],
        on=["ФИО", "Дата"],
        how="left",
    )

    # --- ИТОГО ЗА ДЕНЬ / НЕДОРАБОТКИ ---
    final["Вне_ядра_мин"] = (
        pd.to_numeric(final["Вне_ядра_мин"], errors="coerce")
        .fillna(0)
        .astype(int)
    )
    final["Итого_дня_мин"] = (DAY_CORE_MIN - final["Вне_ядра_мин"]).clip(lower=0)
    final["Итого за день"] = final["Итого_дня_мин"].apply(fmt_hm)

    deficit = (DAY_CORE_MIN - final["Итого_дня_мин"]).clip(lower=0)
    final["Недоработки"] = deficit.apply(fmt_hm)

    # --- ИТОГО ЗА НЕДЕЛЮ ---
    final["Дата_dt"] = pd.to_datetime(final["Дата"], errors="coerce")
    final["week_key"] = final["Дата_dt"].dt.to_period("W").astype(str)

    final["Итого за неделю"] = ""

    for (fio, week), grp in final.groupby(["ФИО", "week_key"]):
        idxs = grp.sort_values("Дата_dt").index
        if len(idxs) == 0:
            continue
        total_week_min = int(grp["Итого_дня_мин"].sum())
        last_idx = idxs[-1]
        final.at[last_idx, "Итого за неделю"] = fmt_hm(total_week_min)

    # ----- ПРИЧИНА ОТСУТСТВИЯ (кадры) -----
    if kadry_file is not None:
        kadry_dates = read_kadry(kadry_file)

        final["ФИО_key"] = final["ФИО"].apply(fio_match_key)
        kadry_dates["ФИО_key"] = kadry_dates["ФИО"].apply(fio_match_key)

        final["Дата_key"] = final["Дата_dt"].dt.date
        kadry_dates["Дата_key"] = pd.to_datetime(
            kadry_dates["Дата"], errors="coerce"
        ).dt.date

        final = final.merge(
            kadry_dates[["ФИО_key", "Дата_key", "Тип"]],
            on=["ФИО_key", "Дата_key"],
            how="left",
        )
        final["Причина отсутствия"] = final["Тип"]
        final = final.drop(columns=["Тип", "ФИО_key", "Дата_key"], errors="ignore")
    else:
        final["Причина отсутствия"] = ""

    # форматируем дату дд-мм-гггг
    final["Дата"] = final["Дата_dt"].dt.strftime("%d-%m-%Y")

    # порядок колонок
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

    final = final[cols_order]

    return final
