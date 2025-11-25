import streamlit as st
import pandas as pd
import io  # для формирования файла Excel в памяти
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from engine import build_report  # берём функцию из engine.py

# Настройки страницы
st.set_page_config(
    page_title="Умный отчет",
    page_icon="📊",
    layout="wide",
)

# 🎨 Глобальное оформление (принудительно светлое)
st.markdown("""
<style>
:root {
    color-scheme: light;  /* просим браузер/Streamlit вести себя как в светлой теме */
}

/* Главный фон приложения */
html, body, [data-testid="stAppViewContainer"], .stApp {
    background: linear-gradient(135deg, #e8efff 0%, #ffffff 60%) !important;
    color: #102A43 !important;
    font-size: 16px;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
}

/* Верхняя панель Streamlit (убираем темный фон) */
[data-testid="stHeader"] {
    background: rgba(255, 255, 255, 0.0) !important;
}

/* Сайдбар, если он появится */
[data-testid="stSidebar"] {
    background-color: #f3f5ff !important;
    color: #102A43 !important;
}

/* Контейнер с основным контентом */
.block-container {
    font-size: 16px;
}

/* Заголовки */
h1, h2, h3, h4 {
    color: #102A43 !important;
    font-weight: 700 !important;
}

/* === Загрузчик файлов (dropzone) === */
[data-testid="stFileUploadDropzone"] {
    background-color: #f5f7fb !important;       /* светлый голубоватый фон */
    border-radius: 8px !important;
    border: 1px dashed #d0d7ea !important;
    color: #102A43 !important;
}

/* Текст "Drag and drop..." и подсказки */
.stFileUploader label {
    font-weight: 600 !important;
    color: #102A43 !important;
}

/* Название загруженного файла — обычный тёмный текст, без чёрного прямоугольника */
[data-testid="stFileUploaderFileName"] {
    color: #102A43 !important;
    background: transparent !important;
}

/* Кнопка "Browse files" / "Выбрать файл" */
.stFileUploader div[role="button"] {
    background-color: #ffffff !important;
    border: 1px solid #d0d7ea !important;
    color: #102A43 !important;
    border-radius: 6px !important;
}

/* === Обычные кнопки (в том числе "Обработать данные") === */
.stButton > button {
    background-color: #1E88E5 !important;   /* приятный синий */
    color: white !important;
    border-radius: 8px !important;
    padding: 8px 20px !important;
    font-size: 16px !important;
    border: none !important;
}
.stButton > button:hover {
    background-color: #1565C0 !important;
}

/* === Кнопка скачивания === */
.stDownloadButton button {
    background-color: #1E88E5 !important;
    color: white !important;
    border-radius: 8px !important;
    padding: 10px 22px !important;
    font-size: 16px !important;
    border: none !important;
}
.stDownloadButton button:hover {
    background-color: #1565C0 !important;
}

/* === Предпросмотр таблицы (st.dataframe) — белый фон, читаемый текст === */
[data-testid="stDataFrame"] {
    background-color: #ffffff !important;
    border-radius: 8px !important;
    padding: 0.25rem !important;
}

/* Внутри самого грида тоже принудительно светлый фон и тёмный текст */
[data-testid="stDataFrame"] div[role="grid"] {
    background-color: #ffffff !important;
    color: #102A43 !important;
}

/* Чуть уменьшенный шрифт в таблице для компактности */
[data-testid="stDataFrame"] table {
    font-size: 14px;
}
</style>
""", unsafe_allow_html=True)

# ================= ГЛАВНЫЙ ЗАГОЛОВОК =================
st.markdown(
    """
    <div style="text-align: center; padding: 20px; background-color: #F0F4FF; border-radius: 10px; margin-bottom: 1.5rem;">
        <h2 style="color: #003366; margin-bottom: 0.5rem;">
            📊 Умный контроль рабочего времени
        </h2>
        <p style="color: #003366; font-size:16px; margin: 0;">
            Загрузите журнал проходов и (по желанию) кадровый файл — система автоматически сформирует табель,
            рассчитает недоработки, выходы, длительные отсутствия и причины отсутствия.
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

# --- Шаг 1: Загрузка файлов ---
st.header("Шаг 1. Загрузка файлов")

col_left, col_right = st.columns([2, 1])

with col_left:
    st.subheader("📘 Журнал проходов")
    file_journal = st.file_uploader(
        "Файл журнала проходов",
        type=["xls", "xlsx"],
        help="Формат: .xls или .xlsx"
    )

    st.markdown("---")

    st.subheader("📗 Сведения из кадров (по желанию)")
    file_kadry = st.file_uploader(
        "Файл кадров / отсутствий",
        type=["xls", "xlsx"],
        help="Можно не загружать — тогда столбец 'Причина отсутствия' будет пустым"
    )

with col_right:
    st.markdown(
        """
        **Подсказки:**
        - Журнал — выгрузка из системы проходов.
        - Кадровый файл — со столбцами:
          *«Сотрудник», «Вид отсутствия», «с», «до»*.
        - Можно загрузить только журнал — тогда причины отсутствия не будут указаны.
        """
    )

# --- Запуск обработки ---
if file_journal is None:
    st.warning("⬆ Сначала загрузите файл журнала проходов.")
else:
    st.markdown(f"📘 **Файл журнала:** {file_journal.name}")
    if file_kadry:
        st.markdown(f"📗 **Файл кадров:** {file_kadry.name}")
    else:
        st.markdown("📗 **Файл кадров не загружен**")

    st.header("Шаг 2. Обработка данных")

    if st.button("🚀 Обработать данные"):
        try:
            final_df = build_report(file_journal, file_kadry)
        except Exception as e:
            st.error(f"❌ Ошибка при обработке: {e}")
        else:
            st.success("✅ Обработка завершена.")

            # --- Шаг 3. Предпросмотр и выгрузка ---
            st.header("Шаг 3. Предпросмотр и выгрузка отчёта")

            # Показывать ли «Причина отсутствия»
            show_reason = (
                "Причина отсутствия" in final_df.columns
                and final_df["Причина отсутствия"].astype(str).str.strip().ne("").any()
            )

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
            if show_reason:
                visible_cols.append("Причина отсутствия")

            visible_cols = [c for c in visible_cols if c in final_df.columns]
            final_view = final_df[visible_cols].copy()

            if "ФИО" in final_view.columns and "Дата" in final_view.columns:
                final_view = final_view.sort_values(["ФИО", "Дата"])

            st.write(f"Строк в итоговой таблице: **{len(final_view)}**")
            st.dataframe(final_view.head(200))

            # 📥 Подготовка Excel
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                sheet_name = "Журнал"
                final_view.to_excel(writer, index=False, sheet_name=sheet_name, startrow=3)

                ws = writer.sheets[sheet_name]
                max_col = ws.max_column
                last_col_letter = get_column_letter(max_col)

                # Заголовок
                title_cell = ws["A1"]
                title_cell.value = "ОТЧЁТ ЗА НЕДЕЛЮ"
                title_cell.font = Font(name="Times New Roman", size=14, bold=True)
                title_cell.alignment = Alignment(horizontal="center", vertical="center")
                ws.merge_cells(f"A1:{last_col_letter}1")

                # Шапка таблицы
                header_row = 4
                header_fill = PatternFill("solid", fgColor="DCE6F1")
                header_font = Font(name="Times New Roman", size=11, bold=True)

                for col in range(1, max_col + 1):
                    cell = ws.cell(row=header_row, column=col)
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(
                        horizontal="center",
                        vertical="center",
                        wrap_text=True,
                    )

                col_names = [cell.value for cell in ws[header_row]]

                # Центровка данных
                for col_idx, name in enumerate(col_names, start=1):
                    align = Alignment(
                        horizontal="center",
                        vertical="center",
                        wrap_text=True,
                    )
                    for row in range(header_row + 1, ws.max_row + 1):
                        ws.cell(row=row, column=col_idx).alignment = align

                # Ширины колонок
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
                        ws.column_dimensions[get_column_letter(col_idx)].width = width_map[name]

                # Общий шрифт
                base_font = Font(name="Times New Roman", size=11)
                for row in ws.iter_rows():
                    for cell in row:
                        if cell.value is not None:
                            cell.font = base_font

                ws.freeze_panes = "A5"

            buffer.seek(0)

            # Кнопка скачивания
            st.download_button(
                label="💾 Скачать итоговый отчёт (Excel)",
                data=buffer,
                file_name="умный_табель.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
