import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
import tempfile

st.set_page_config(page_title="Отчет по ресурсам", layout="centered")

st.title("📊 Формирование отчета по ресурсам")

uploaded_files = st.file_uploader(
    "Загрузите файлы проектов (Excel)",
    type=["xlsx"],
    accept_multiple_files=True
)

col1, col2 = st.columns(2)
with col1:
    date_from = st.date_input("Начало")
with col2:
    date_to = st.date_input("Окончание")

generate = st.button("🚀 Сформировать отчет")

def read_data(file):
    df = pd.read_excel(file, sheet_name="Data")
    return df

if generate:

    if not uploaded_files:
        st.error("Загрузите хотя бы один файл")
        st.stop()

    if date_from > date_to:
        st.error("Неверный период")
        st.stop()

    dfs = []
    for file in uploaded_files:
        try:
            df = read_data(file)
            dfs.append(df)
        except:
            st.error(f"Ошибка чтения файла: {file.name}")
            st.stop()

    data = pd.concat(dfs, ignore_index=True)

    # Преобразуем дату
    date_columns = [c for c in data.columns if "дата" in c.lower()]
    for col in date_columns:
        data[col] = pd.to_datetime(data[col], errors="coerce")

    # Фильтр по периоду
    main_date_col = date_columns[0]
    mask = (data[main_date_col] >= pd.to_datetime(date_from)) & \
           (data[main_date_col] <= pd.to_datetime(date_to))
    filtered = data[mask]

    # Работаем с шаблоном
    wb = load_workbook("template.xlsx")
    ws = wb["Data"]

    ws.delete_rows(2, ws.max_row)

    for i, row in filtered.iterrows():
        ws.append(list(row))

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(tmp.name)

    with open(tmp.name, "rb") as f:
        st.success("Отчет готов!")
        st.download_button(
            "⬇ Скачать отчет",
            f,
            file_name="Отчет_по_ресурсам.xlsx"
        )

