import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
import tempfile

# Настройки страницы
st.set_page_config(page_title="Отчет по ресурсам", layout="centered")

st.title("📊 Формирование отчета по ресурсам")

# Загрузка файлов
uploaded_files = st.file_uploader(
    "Загрузите файлы проектов (Excel)",
    type=["xlsx"],
    accept_multiple_files=True
)

# Выбор периода
col1, col2 = st.columns(2)
with col1:
    date_from = st.date_input("Начало")
with col2:
    date_to = st.date_input("Окончание")

generate = st.button("🚀 Сформировать отчет")

# Функция чтения
def read_data(file):
    df = pd.read_excel(file, sheet_name="Data")
    return df


if generate:

    # Проверки
    if not uploaded_files:
        st.error("Загрузите хотя бы один файл")
        st.stop()

    if date_from > date_to:
        st.error("Неверный период")
        st.stop()

    # Читаем все файлы
    dfs = []
    for file in uploaded_files:
        try:
            df = read_data(file)
            dfs.append(df)
        except Exception as e:
            st.error(f"Ошибка чтения файла: {file.name}")
            st.stop()

    # Объединяем
    data = pd.concat(dfs, ignore_index=True)

    st.subheader("Выберите колонку с датой")

    columns = data.columns.tolist()

    date_col = st.selectbox(
        "Колонка с датой:",
        columns
    )

    # Приводим к дате
    data[date_col] = pd.to_datetime(data[date_col], errors="coerce")

    # Фильтрация
    mask = (
        (data[date_col] >= pd.to_datetime(date_from)) &
        (data[date_col] <= pd.to_datetime(date_to))
    )

    filtered_df = data[mask]

    st.success(f"Найдено строк: {len(filtered_df)}")

    # Работа с шаблоном
    wb = load_workbook("template.xlsx")
    ws = wb["Data"]

    # Очищаем старые строки
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row)

    # Записываем новые данные
    for _, row in filtered_df.iterrows():
        ws.append(list(row))

    # Сохраняем временный файл
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(tmp.name)

    with open(tmp.name, "rb") as f:
        st.success("✅ Отчет готов!")
        st.download_button(
            "⬇ Скачать отчет",
            f,
            file_name="Отчет_по_ресурсам.xlsx"
        )
