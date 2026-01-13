import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import tempfile
import os

st.set_page_config(page_title="Отчет по ресурсам", layout="centered")
st.title("📊 Формирование отчета по ресурсам")

# ---------- SESSION ----------
if "files" not in st.session_state:
    st.session_state.files = None

# ---------- ЗАГРУЗКА ----------
uploaded_files = st.file_uploader(
    "Загрузите файлы проектов (Excel)",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    st.session_state.files = uploaded_files

# ---------- ПЕРИОД ----------
col1, col2 = st.columns(2)
with col1:
    date_from = st.date_input("Начало периода")
with col2:
    date_to = st.date_input("Окончание периода")

generate = st.button("🚀 Сформировать отчет")

# ---------- ЧТЕНИЕ ----------
def read_data(file):
    return pd.read_excel(file, sheet_name="Data")

# ---------- ОСНОВНАЯ ЛОГИКА ----------
if generate:

    if not st.session_state.files:
        st.error("Загрузите файлы")
        st.stop()

    dfs = []
    for f in st.session_state.files:
        df = read_data(f)
        dfs.append(df)

    data = pd.concat(dfs, ignore_index=True)

    st.subheader("Выберите колонки дат")

    cols = data.columns.tolist()
    col_start = st.selectbox("Колонка НАЧАЛА", cols)
    col_end   = st.selectbox("Колонка ОКОНЧАНИЯ", cols)

    data[col_start] = pd.to_datetime(data[col_start], errors="coerce")
    data[col_end]   = pd.to_datetime(data[col_end], errors="coerce")

    mask = (
        (data[col_start] <= pd.to_datetime(date_to)) &
        (data[col_end]   >= pd.to_datetime(date_from))
    )

    filtered = data[mask]

    st.success(f"Найдено строк: {len(filtered)}")

    # ---------- ЗАПИСЬ В ШАБЛОН ----------
    if not os.path.exists("template.xlsx"):
        st.error("Файл template.xlsx не найден рядом с app.py")
        st.stop()

    wb = load_workbook("template.xlsx")
    ws = wb["Data"]   # ВАЖНО: имя листа в шаблоне

    # очистка старых данных
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row)

    # запись
    for _, row in filtered.iterrows():
        ws.append(list(row))

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(tmp.name)

    with open(tmp.name, "rb") as f:
        st.download_button(
            "⬇ Скачать отчет",
            f,
            file_name="Отчет_по_ресурсам.xlsx"
        )
