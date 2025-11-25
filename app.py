import streamlit as st
import pandas as pd
import io
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from engine import build_report

# Настройки страницы
st.set_page_config(
    page_title="Умный отчет",
    page_icon="📊",
    layout="wide",
)


# ========= 🎨 ГЛАВНОЕ ОФОРМЛЕНИЕ =========
st.markdown(
    """
<style>
/* Общий фон, шрифт, отступы */
html, body, .stApp {
    background: linear-gradient(135deg, #eef4ff 0%, #ffffff 60%);
    color: #102A43 !important;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    font-size: 16px;
}

/* Блок с контентом — делаем читабельнее */
.block-container {
    padding-top: 1rem !important;
    padding-bottom: 2rem !important;
}

/* 🔹 Кнопка "Обработать данные" и другие кнопки */
.stButton > button {
    background-color: #1E88E5 !important;
    color: white !important;
    border-radius: 8px !important;
    border: none !important;
    padding: 10px 20px !important;
    font-size: 15px !important;
    font-weight: 500 !important;
    cursor: pointer;
    transition: 0.2s ease-in-out !important;
}
.stButton > button:hover {
    background-color: #1565C0 !important;
}

/* 🔹 Кнопка скачивания */
.stDownloadButton > button {
    background-color: #2E7D32 !important;
    color: white !important;
    border-radius: 8px !important;
    padding: 10px 22px !important;
    font-size: 16px !important;
    border: none !important;
}
.stDownloadButton > button:hover {
    background-color: #1B5E20 !important;
}

/* 🔹 Файловые загрузчики — нормальные, светлые */
.stFileUploader > div:nth-child(1) {
    background-color: #f7f9fd !important;
    border-radius: 10px !important;
    border: 1px solid #c8d6ff !important;
    padding: 10px;
}
.stFileUploader label {
    font-weight: 600 !important;
    color: #003366 !important;
}

/* 🔹 Фон таблицы предварительного просмотра */
[data-testid="stDataFrame"] {
    background-color: white !important;
    color: #102A43 !important;
}

/* Заголовки — жирные и читабельные */
h1, h2, h3, h4 {
    color: #003366 !important;
    font-weight: 700 !important;
}
</style>
""",
    unsafe_allow_html=True,
)

# ========= 🏷 ГЛАВНЫЙ ЗАГОЛОВОК =========
st.markdown(
    """
    <div style="text-align: center; padding: 20px; background-color: #F0F4FF;
                border-radius: 10px; margin-bottom: 1.5rem;">
        <h2 style="color: #003366; margin-bottom: 0.5rem;">
            📊 Умный контроль рабочего времени
        </h2>
        <p style="color: #003366; font-size:16px; margin: 0;">
            Загрузите журнал проходов и (по желанию) файл кадров — система автоматически сформирует табель,
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
    st.markdown("Загрузите Excel-файл с проходами через турникеты.")
    file_journal = st.file_uploader(
        "Файл журнала проходов",
        type=["xls", "xlsx"],
        help="Формат: .xls или .xlsx"
    )

    st.markdown("---")

    st.subheader("📗 Сведения из кадров (необязательно)")
    file_kadry = st.file_uploader(
        "Файл кадров / отсутствий",
        type=["xls", "xlsx"],
        help="Можно не загружать. Тогда причина отсутствия будет пустой."
    )

with col_right:
    st.markdown(
        """
        **Как использовать:**
        1️⃣ Сначала загрузите **журнал проходов**  
        2️⃣ При наличии — загрузите **кадровый файл**  
        3️⃣ Нажмите кнопку **Обработать данные**  

        📌 Кадровый файл должен содержать колонки:
        *ФИО*, *Вид отсутствия*, *с*, *до*
        """
    )

# --- Шаг 2: Обработка ---
if file_journal is None:
    st.warning("⬆ Сначала загрузите файл журнала проходов.")
else:
    if file_kadry is None:
        st.info("Кадровый файл не загружен — причина отсутствия останется пустой.")

    st.markdown(f"<div style='color:#003366; font-weight:600;'>📘 Журнал: {file_journal.name}</div>", unsafe_allow_html=True)

if file_kadry is not None:
    st.markdown(f"<div style='color:#003366; font-weight:600;'>📗 Кадровый файл: {file_kadry.name}</div>", unsafe_allow_html=True)
else:
    st.markdown("<div style='color:#555;'>📗 Кадровый файл: не загружен</div>", unsafe_allow_html=True)

    st.header("Шаг 2. Обработка данных")

    if st.button("🚀 Обработать данные"):
        try:
            final_df = build_report(file_journal, file_kadry)
        except Exception as e:
            st.error(f"❌ Ошибка при обработке: {e}")
        else:
            st.success("✅ Обработка завершена!")

            # === Шаг 3. Предпросмотр ===
            st.header("Шаг 3. Предпросмотр и выгрузка")

            show_reason = False
            if "Причина отсутствия" in final_df.columns:
                non_empty = final_df["Причина отсутствия"].astype(str).str.strip().ne("")
                show_reason = non_empty.any()

            visible_cols = [
                "ФИО", "Дата", "Время прихода", "Время ухода", "Опоздание",
                "Общее время", "Вне офиса", "Выходы",
                "Отсутствие более 2 часов подряд",
                "Итого за день", "Итого за неделю", "Недоработки",
            ]
            if show_reason:
                visible_cols.append("Причина отсутствия")

            visible_cols = [c for c in visible_cols if c in final_df.columns]
            final_view = final_df[visible_cols].copy()
            final_view = final_view.sort_values(["ФИО", "Дата"])

            st.write(f"📄 Строк в таблице: **{len(final_view)}**")
            st.dataframe(final_view.head(200), use_container_width=True)

            # === Скачивание Excel ===
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                sheet_name = "Журнал"
                final_view.to_excel(writer, sheet_name=sheet_name, index=False, startrow=3)

                ws = writer.book[sheet_name]
                last_col_letter = get_column_letter(ws.max_column)

                # Заголовок
                cell = ws["A1"]
                cell.value = "ОТЧЁТ ЗА НЕДЕЛЮ"
                cell.font = Font(name="Times New Roman", size=14, bold=True)
                cell.alignment = Alignment(horizontal="center")
                ws.merge_cells(f"A1:{last_col_letter}1")

                ws.freeze_panes = "A5"

            buffer.seek(0)
            st.download_button(
                label="💾 Скачать итоговый отчёт (Excel)",
                data=buffer,
                file_name="умный_табель.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

