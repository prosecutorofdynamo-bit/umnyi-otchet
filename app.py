import streamlit as st
import pandas as pd
import io  # для формирования файла Excel в памяти
import unicodedata
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from engine import build_report  # берём функцию из engine.py


def fio_norm(s: str) -> str:
    s = "" if pd.isna(s) else str(s)
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("ё", "е").replace("Ё", "Е")
    s = " ".join(s.strip().split()).lower()
    return s


def build_kadry_dates_from_df(kadry: pd.DataFrame) -> pd.DataFrame:
    """
    Принимает сырую таблицу кадров (как из Excel),
    возвращает kadry_dates: ФИО, Дата (date), Тип (строка).
    Логика идентична Колабу.
    """

    # 1) Ищем строку, где в ЛЮБОЙ колонке есть "Сотрудник"
    def _is_sotr_cell(x):
        s = "" if pd.isna(x) else str(x)
        return s.strip().casefold() == "сотрудник"

    mask_rows = kadry.apply(lambda row: row.map(_is_sotr_cell).any(), axis=1)
    idxs = kadry.index[mask_rows]

    if len(idxs) == 0:
        raise RuntimeError(
            "Не удалось найти строку с заголовком 'Сотрудник' "
            "(поиск по всем колонкам, без учёта регистра). "
            "Проверьте, что в кадровом отчёте есть колонка 'Сотрудник'."
        )

    hdr_row = idxs[0]

    # 2) Берём эту строку как шапку
    kadry = kadry.copy()
    kadry.columns = kadry.iloc[hdr_row]
    kadry = kadry.iloc[hdr_row + 1:]  # ниже заголовка

    # 3) Оставляем нужные колонки
    kadry = kadry.rename(
        columns={
            "Сотрудник": "ФИО",
            "Вид отсутствия": "Тип",
            "с": "Дата_с",
            "до": "Дата_по",
        }
    )

    need_cols = ["ФИО", "Тип", "Дата_с", "Дата_по"]
    kadry = kadry[need_cols].copy()
    kadry = kadry.dropna(subset=["ФИО", "Тип"], how="any")

    # 4) Вытаскиваем даты формата дд.мм.гггг
    for col in ["Дата_с", "Дата_по"]:
        kadry[col] = kadry[col].astype(str).str.extract(
            r"(\d{2}\.\d{2}\.\d{4})", expand=False
        )
        kadry[col] = pd.to_datetime(kadry[col], dayfirst=True, errors="coerce")

    # Если "до" пусто — отсутствие один день
    kadry["Дата_по"] = kadry["Дата_по"].fillna(kadry["Дата_с"])

    # 5) Разворачиваем интервалы в посуточный список
    rows = []
    for _, r in kadry.iterrows():
        d1, d2 = r["Дата_с"], r["Дата_по"]
        if pd.isna(d1) or pd.isna(d2):
            continue
        for d in pd.date_range(d1, d2, freq="D"):
            rows.append({"ФИО": r["ФИО"], "Дата": d.date(), "Тип": r["Тип"]})

    kadry_dates = pd.DataFrame(rows)

    # 6) Замена типа отсутствия (как в Колабе)
    kadry_dates["Тип"] = kadry_dates["Тип"].replace(
        to_replace=r"(?i).*гос.*обязан.*", value="Сдача крови", regex=True
    )

    return kadry_dates


# ---------------- НАСТРОЙКИ СТРАНИЦЫ ----------------
st.set_page_config(
    page_title="Умный отчет",
    page_icon="📊",
    layout="wide",
)

# ---------------- ГЛОБАЛЬНЫЙ СТИЛЬ (CSS) ----------------
st.markdown(
    """
    <style>
    .stApp {
        background: linear-gradient(135deg, #e4f0ff 0%, #ffffff 55%) !important;
        color: #102A43 !important;
        font-size: 16px !important;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif !important;
    }

    .block-container {
        padding-top: 1rem;
        padding-bottom: 2rem;
    }

    /* === ЗАГРУЗЧИК ФАЙЛОВ === */

    /* Светлый фон всей плашки загрузчика */
    [data-testid="stFileUploader"] section {
        background-color: #f5f7fb !important;
        border: 1px solid #d0d7ea !important;
        border-radius: 8px !important;
        color: #102A43 !important;
    }

    /* Скрываем английские подсказки */
    [data-testid="stFileDropzone"] span,
    [data-testid="stFileUploaderInstructions"] {
        display: none !important;
    }

    /* Кнопка "Browse files" — светлая */
    [data-testid="stFileUploader"] button {
        background-color: #eef3ff !important;
        color: #003366 !important;
        border: 1px solid #d0d7ea !important;
        border-radius: 6px !important;
        padding: 6px 14px !important;
        font-weight: 600 !important;
        box-shadow: none !important;
    }
    [data-testid="stFileUploader"] button:hover {
        background-color: #d6e4ff !important;
    }

    /* Убираем тёмный блок вокруг кнопки */
    [data-testid="stFileDropzone"] {
        background-color: transparent !important;
        border: none !important;
    }

    /* Название файла и размер — читаемые, с отдельным фоном */
    [data-testid="stFileUploaderFileName"] {
        color: #003366 !important;
        background-color: #ffffff !important;
        padding: 4px 8px !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        display: inline-block !important;
    }
    [data-testid="stFileUploaderSize"] {
        color: #4a637e !important;
        background-color: #ffffff !important;
        padding: 2px 6px !important;
        border-radius: 4px !important;
        margin-left: 4px !important;
        font-size: 13px !important;
    }

    /* Кнопки (Обработать данные, Скачать отчёт) */
    .stButton > button, .stDownloadButton > button {
        background-color: #1E88E5 !important;
        color: white !important;
        border-radius: 8px !important;
        padding: 10px 22px !important;
        font-size: 16px !important;
        border: none !important;
        font-weight: 600 !important;
        transition: 0.3s ease-in-out;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        background-color: #1565C0 !important;
        transform: translateY(-1px);
    }

    /* Заголовки */
    h1, h2, h3, h4 {
        color: #102A43 !important;
        font-weight: 700 !important;
    }

    /* Текст "Журнал: файл.xlsx" и "Кадровый файл: ..." */
    .file-label {
        padding: 4px 10px;
        margin: 4px 0;
        border-radius: 6px;
        background-color: #eef3ff;
        color: #003366;
        font-weight: 600;
        display: inline-block;
    }

    /* Таблица предпросмотра (если появится) — белый фон */
    [data-testid="stDataFrame"] div[role="grid"] {
        background-color: #ffffff !important;
        color: #102A43 !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# ---------------- ГЛАВНЫЙ ЗАГОЛОВОК ----------------
st.markdown(
    """
    <div style="text-align: center; padding: 20px; background-color: #F0F4FF;
                border-radius: 10px; margin-bottom: 1.5rem;">
        <h2 style="color: #003366; margin-bottom: 0.5rem;">
            📊 Умный контроль рабочего времени
        </h2>
        <p style="color: #003366; font-size:16px; margin: 0;">
            Загрузите журнал проходов и (по желанию) файл кадров — система автоматически сформирует табель,
            рассчитает недоработки, выходы, длительные отсутствия и причины прогула.
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

# --- Шаг 1. Загрузка файлов ---
st.header("Шаг 1. Загрузка файлов")

col_left, col_right = st.columns([2, 1])

with col_left:
    # -------- ЖУРНАЛ ПРОХОДОВ --------
    st.subheader("📘 Журнал проходов")

    st.markdown(
        """
        <div style="
            padding: 10px; 
            background-color: #eef3ff; 
            border-radius: 6px; 
            border: 1px solid #d0d7ea; 
            margin-bottom: 8px; 
            color:#003366;
        ">
            📤 <b>Загрузите файл журнала проходов</b><br>
            <span style="font-size: 14px;">
                Формат: XLS или XLSX, размер до 200 МБ.
            </span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    file_journal = st.file_uploader(
        "Журнал проходов",
        type=["xls", "xlsx"],
        label_visibility="collapsed",
        help="Файл журнала из системы проходов (XLS/XLSX)."
    )

    st.markdown("---")

    # -------- ФАЙЛ КАДРОВ --------
    st.subheader("📗 Сведения из кадров (по желанию)")

    st.markdown(
        """
        <div style="
            padding: 10px; 
            background-color: #eef3ff; 
            border-radius: 6px; 
            border: 1px solid #d0d7ea; 
            margin-bottom: 8px; 
            color:#003366;
        ">
            📤 <b>Загрузите файл с отсутствиями (кадровый отчёт)</b><br>
            <span style="font-size: 14px;">
                Отпуска, больничные, командировки и др. причины отсутствий. 
                Можно не загружать — тогда колонка «Причина отсутствия» останется пустой.
            </span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    kadry_file = st.file_uploader(
        "Загрузите файл кадров (.xls / .xlsx)",
        type=["xls", "xlsx"]
    )

kadry_dates = None
if kadry_file is not None:
    raw_kadry_df = pd.read_excel(kadry_file, header=None)
    kadry_dates = build_kadry_dates_from_df(raw_kadry_df)
    st.write(f"Загружено записей по отсутствиям: {len(kadry_dates)}")

with col_right:
    st.markdown(
        """
        **Подсказки:**
        - Журнал — стандартная выгрузка из системы проходов.
        - Кадровый файл — со столбцами:
          *«Сотрудник», «Вид отсутствия», «с», «до»*.
        - Можно загрузить только журнал —
          тогда «Причина отсутствия» останется пустой.
        """,
        unsafe_allow_html=False,
    )

# Если журнал не загружен — дальше не идём
if file_journal is None:
    st.warning("⬆ Сначала загрузите файл журнала проходов.")
    st.stop()

st.caption("Перетащите файл сюда или нажмите «Browse files» для выбора файла журнала.")

# Пояснение по кадровому файлу
if kadry_file is None:
    st.info(
        "Кадровый файл *не обязателен*. "
        "Можете загрузить его для указания причин отсутствия "
        "или сразу перейти к обработке."
    )
else:
    st.success("✅ Оба файла загружены!")

# Красивый вывод названий файлов
st.markdown(
    f"<div class='file-label'>📘 Журнал: {file_journal.name}</div>",
    unsafe_allow_html=True,
)

if kadry_file is not None:
    st.markdown(
        f"<div class='file-label'>📗 Кадровый файл: {kadry_file.name}</div>",
        unsafe_allow_html=True,
    )
else:
    st.markdown(
        "<div class='file-label' style='background-color:#f5f5f5; color:#555;'>"
        "📗 Кадровый файл: не загружен"
        "</div>",
        unsafe_allow_html=True,
    )

# ---------------- ШАГ 2. ОБРАБОТКА ДАННЫХ ----------------
st.header("Шаг 2. Обработка данных")

final_df = None

if st.button("🚀 Обработать данные"):
    try:
        # kadry_file может быть None — это нормально
        final_df = build_report(file_journal, kadry_file)
    except Exception as e:
        st.error(f"❌ Ошибка при обработке данных: {e}")
    else:
        st.success("✅ Обработка завершена.")

        # --- ДОБАВЛЯЕМ ПРИЧИНЫ ОТСУТСТВИЯ ИЗ КАДРОВОГО ОТЧЁТА (как в Колабе) ---
        if kadry_dates is not None and not kadry_dates.empty:
            tmp = final_df.copy()

            # 1) Ключи по дате
            tmp["Дата_key"] = pd.to_datetime(
                tmp["Дата"], dayfirst=True, errors="coerce"
            ).dt.date
            kd = kadry_dates.copy()
            kd["Дата_key"] = kd["Дата"]          # там уже date

            # 2) Ключи по ФИО (нижний регистр, без лишних пробелов)
            tmp["ФИО_key"] = tmp["ФИО"].astype(str).str.strip().str.lower()
            kd["ФИО_key"] = kd["ФИО"].astype(str).str.strip().str.lower()

            # 3) Соединяем
            tmp = tmp.merge(
                kd[["ФИО_key", "Дата_key", "Тип"]],
                on=["ФИО_key", "Дата_key"],
                how="left",
            )

            # 4) Переносим в человекочитаемую колонку
            tmp["Причина отсутствия"] = tmp["Тип"]

            # 5) Убираем служебные колонки
            tmp = tmp.drop(columns=["Тип", "ФИО_key", "Дата_key"], errors="ignore")

            final_df = tmp

# Если ещё не нажали кнопку или произошла ошибка — дальше не идём
if final_df is None:
    st.stop()

# ---------------- ШАГ 3. ПРЕДПРОСМОТР И ВЫГРУЗКА ----------------
st.header("Шаг 3. Выгрузка отчёта")

# Определяем, есть ли смысл показывать «Причина отсутствия»
show_reason = False
if "Причина отсутствия" in final_df.columns:
    non_empty = final_df["Причина отсутствия"].astype(str).str.strip().ne("")
    show_reason = non_empty.any()

# Базовый набор колонок
visible_cols = [
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
]

# Добавляем «Причина отсутствия» только если она непустая
if show_reason:
    visible_cols.append("Причина отсутствия")

# Оставляем только существующие колонки
visible_cols = [c for c in visible_cols if c in final_df.columns]

if not visible_cols:
    st.warning("В итоговом отчёте нет ожидаемых колонок для отображения.")
    final_view = final_df.copy()
else:
    final_view = final_df[visible_cols].copy()

# Сортировка по ФИО и дате
if "ФИО" in final_view.columns and "Дата" in final_view.columns:
    final_view = final_view.sort_values(["ФИО", "Дата"])

# ---------------- ФОРМИРОВАНИЕ И СКАЧИВАНИЕ EXCEL ----------------
buffer = io.BytesIO()
with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
    sheet_name = "Журнал"

    # пишем таблицу с отступом (чтобы сверху уместить заголовок)
    final_view.to_excel(writer, index=False, sheet_name=sheet_name, startrow=3)

    wb = writer.book
    ws = writer.sheets[sheet_name]

    max_col = ws.max_column
    last_col_letter = get_column_letter(max_col)

    # --- Большой заголовок ---
    title_cell = ws["A1"]
    title_cell.value = "ОТЧЁТ ЗА НЕДЕЛЮ"
    title_cell.font = Font(name="Times New Roman", size=14, bold=True)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(f"A1:{last_col_letter}1")

    # --- Шапка таблицы (строка 4) ---
    header_row = 4
    header_fill = PatternFill("solid", fgColor="DCE6F1")  # нежно-голубой фон
    header_font = Font(name="Times New Roman", size=11, bold=True)

    # заголовки
    for col in range(1, max_col + 1):
        cell = ws.cell(row=header_row, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )

    # --- Выравнивание данных и ширина столбцов ---
    col_names = [cell.value for cell in ws[header_row]]

    # выравниваем все ячейки в таблице по центру
    for col_idx, name in enumerate(col_names, start=1):
        align = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )
        for row in range(header_row + 1, ws.max_row + 1):
            ws.cell(row=row, column=col_idx).alignment = align

    # Ширины столбцов
    width_map = {
        "ФИО": 30,
        "Дата": 12,
        "Время прихода": 15,
        "Время ухода": 15,
        "Опоздание": 14,
        "Вне офиса": 16,
        "Выходы": 12,
        "Отсутствие более 2 часов подряд": 28,
        "Итого за день": 14,
        "Итого за неделю": 16,
        "Недоработки": 16,
        "Причина отсутствия": 28,
    }

    for col_idx, name in enumerate(col_names, start=1):
        if name in width_map:
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = width_map[name]

    # Общий шрифт Times New Roman 11 для всех непустых ячеек
    base_font = Font(name="Times New Roman", size=11)
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                cell.font = base_font

    # Заморозить строки до данных (курсор сразу под шапкой)
    ws.freeze_panes = "A5"

buffer.seek(0)

st.download_button(
    label="💾 Скачать итоговый отчёт (Excel)",
    data=buffer,
    file_name="умный_табель.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)



