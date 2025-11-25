import streamlit as st
import pandas as pd
import io  # для формирования файла Excel в памяти
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from engine import build_report  # берём функцию из engine.py

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

    /* === Загрузчик файлов: светлый фон, без чёрных полос === */
    .stFileUploader > div:nth-child(1),
    [data-testid="stFileDropzone"] {
        background-color: #f5f7fb !important;
        border-radius: 8px !important;
        border: 1px solid #d0d7ea !important;
    }

    .stFileUploader label {
        font-weight: 600 !important;
        color: #102A43 !important;
    }

    .stFileUploader div[role="button"] {
        background-color: #ffffff !important;
        border: 1px solid #d0d7ea !important;
        color: #102A43 !important;
    }

    /* Название загруженного файла и размер — читабельные */
    [data-testid="stFileUploaderFileName"] {
        color: #102A43 !important;
        font-weight: 600 !important;
    }
    [data-testid="stFileUploaderSize"] {
        color: #4a637e !important;
        font-weight: 500 !important;
    }

    /* Прячем английский текст внутри дропзоны, чтобы не мозолил глаза */
    [data-testid="stFileDropzone"] span {
        color: transparent !important;
    }

    /* === Кнопки (Обработать данные, Скачать отчёт) === */
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

    /* === Таблица предпросмотра — белый фон === */
    .stDataFrame {
        background-color: #ffffff !important;
        border-radius: 6px !important;
        padding: 0.4rem !important;
    }
    .stDataFrame div[role="grid"] {
        background-color: #ffffff !important;
        color: #102A43 !important;
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

# ---------------- ШАГ 1. ЗАГРУЗКА ФАЙЛОВ ----------------
st.header("Шаг 1. Загрузка файлов")

col_left, col_right = st.columns([2, 1])

with col_left:
    st.subheader("📘 Журнал проходов")
    st.markdown("Загрузите файл с проходами через турникеты.")

    file_journal = st.file_uploader(
        "Файл журнала проходов",
        type=["xls", "xlsx"],
        help="Формат: .xls или .xlsx. Убедитесь, что файл содержит таблицу проходов."
    )

    # Наш русский текст вместо скрытого английского
    st.markdown(
        '<div class="upload-hint">'
        'Перетащите файл журнала сюда или нажмите кнопку «Browse files».'
        '</div>',
        unsafe_allow_html=True,
    )

    st.markdown("---")

    st.subheader("📗 Сведения из кадров (по желанию)")
    st.markdown(
        "Загрузите кадровый файл с отпусками, больничными, "
        "командировками и другими причинами отсутствий."
    )

    file_kadry = st.file_uploader(
        "Файл кадров / отсутствий",
        type=["xls", "xlsx"],
        help="Формат: .xls или .xlsx. Используется для учёта отпусков, больничных и т.д."
    )

st.caption("Перетащите файл сюда или нажмите «Browse files» для выбора кадрового файла.")

st.markdown(
        '<div class="upload-hint">'
        'Перетащите кадровый файл сюда или нажмите кнопку «Browse files».'
        '</div>',
        unsafe_allow_html=True,
    )

with col_right:
    st.markdown(
        """
        **Подсказки:**
        - Журнал — стандартная выгрузка из системы проходов.
        - Кадровый файл — со столбцами:
          *«Сотрудник», «Вид отсутствия», «с», «до»*.
        - Можно загрузить только журнал —
          тогда «Причина отсутствия» останется пустой.
        """
    )

# Если журнал не загружен — дальше не идём
if file_journal is None:
    st.warning("⬆ Сначала загрузите файл журнала проходов.")
    st.stop()

st.caption("Перетащите файл сюда или нажмите «Browse files» для выбора файла журнала.")

# Пояснение по кадровому файлу
if file_kadry is None:
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
if file_kadry is not None:
    st.markdown(
        f"<div class='file-label'>📗 Кадровый файл: {file_kadry.name}</div>",
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
        # file_kadry может быть None — это нормально
        final_df = build_report(file_journal, file_kadry)
    except Exception as e:
        st.error(f"❌ Ошибка при обработке данных: {e}")
    else:
        st.success("✅ Обработка завершена.")

# Если ещё не нажали кнопку или произошла ошибка — дальше не идём
if final_df is None:
    st.stop()

# ---------------- ШАГ 3. ПРЕДПРОСМОТР И ВЫГРУЗКА ----------------
st.header("Шаг 3. Предпросмотр и выгрузка отчёта")

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









