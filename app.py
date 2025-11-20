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

# Главный заголовок
st.title("📊 Умный отчет")
st.caption("Создание отчёта по проходам за пару кликов")

# --- Шаг 1: Загрузка файлов ---
st.header("Шаг 1. Загрузка файлов")

# Блок журнала проходов
st.subheader("📘 Журнал проходов")
st.markdown("Загрузите файл с проходами через турникеты.")

file_journal = st.file_uploader(
    label="",
    type=["xls", "xlsx"],
    label_visibility="collapsed",
    help="Формат: .xls или .xlsx. Убедитесь, что файл содержит таблицу проходов.",
    key="journal",
)

st.markdown("---")  # разделитель

# Блок кадров / отсутствий
st.subheader("📗 Сведения из кадров")
st.markdown(
    "Загрузите кадровый файл с отпусками, больничными, "
    "командировками и другими причинами отсутствий."
)

file_kadry = st.file_uploader(
    label="",
    type=["xls", "xlsx"],
    label_visibility="collapsed",
    help="Формат: .xls или .xlsx. Используется для учёта отпусков, больничных и т.д.",
    key="kadry",
)

# --- Подсказки пользователю ---
if file_journal is None and file_kadry is None:
    st.warning("⬆ Сначала загрузите файл журнала проходов.")
elif file_journal is not None and file_kadry is None:
    st.info("Теперь загрузите файл кадров (отсутствий).")
elif file_journal is None and file_kadry is not None:
    st.warning("Не найден файл журнала проходов — загрузите его.")
else:
    st.success("✅ Оба файла загружены!")

    st.markdown(f"**📘 Журнал проходов:** `{file_journal.name}`")
    st.markdown(f"**📗 Файл кадров:** `{file_kadry.name}`")

    # --- Шаг 2: Обработка данных ---
    st.header("Шаг 2. Обработка данных")

    if st.button("🚀 Обработать данные"):
        try:
            # вызываем наш основной движок
            final_df = build_report(file_journal, file_kadry)
        except Exception as e:
            st.error(f"❌ Ошибка при обработке данных: {e}")
        else:
            st.success("✅ Обработка завершена.")

            # --- Шаг 3: Предпросмотр результата ---
            st.header("Шаг 3. Предпросмотр и выгрузка отчёта")

            # Оставляем только понятные пользователю колонки (если они есть)
            visible_cols = [
                "ФИО",
                "Дата",
                "Время прихода",
                "Время ухода",
                "Опоздание",
                "Вне офиса",
                "Выходы",
                "Отсутствие более 2 часов подряд",
                "Итого за день",
                "Итого за неделю",
                "Недоработки",
                "Причина отсутствия",
            ]
            visible_cols = [c for c in visible_cols if c in final_df.columns]

            if not visible_cols:
                st.warning("В итоговом отчёте нет ожидаемых колонок для отображения.")
                final_view = final_df.copy()
            else:
                final_view = final_df[visible_cols].copy()

            # Красивее сортируем: сначала по ФИО, потом по дате (если есть)
            if "ФИО" in final_view.columns and "Дата" in final_view.columns:
                final_view = final_view.sort_values(["ФИО", "Дата"])

            st.write(f"Строк в итоговой таблице: **{len(final_view)}**")
            st.dataframe(final_view.head(200))

            # --- Подготовка Excel-файла с оформлением ---
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

            # --- Кнопка скачивания ---
            st.download_button(
                label="💾 Скачать итоговый отчёт (Excel)",
                data=buffer,
                file_name="умный_табель.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
