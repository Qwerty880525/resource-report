import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import tempfile

st.set_page_config(page_title="Отчет по ресурсам", layout="centered")

st.title("📊 Формирование отчета по ресурсам")

# --- SESSION ---
if "files" not in st.session_state:
    st.session_state.files = None

# --- Upload ---
uploaded_files = st.file_uploader(
    "Загрузите файлы проектов (Excel)",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    st.session_state.files = uploaded_files

# --- Dates ---
col1, col2 = st.columns(2)
with col1:
    date_from = st.date_input("Начало")
with col2:
    date_to = st.date_input("Окончание")

# --- Button ---
generate = st.button("🚀 Сформировать отчет")

# --- Read ---
def read_data(file):
    return pd.read_excel(file, sheet_name="Data")


if generate:

    files = st.session_state.files

    if not files:
        st.error("Загрузите файл")
        st.stop()

    if date_from > date_to:
        st.error("Неверный период")
        st.stop()

    dfs = []
    for f in files:
        dfs.append(read_data(f))

    data = pd.concat(dfs, ignore_index=True)

    st.subheader("Выберите колонки с датами")

    cols = data.columns.tolist()

    start_col = st.selectbox("Колонка НАЧАЛА:", cols)
    end_col = st.selectbox("Колонка ОКОНЧАНИЯ:", cols)

    # convert
    data[start_col] = pd.to_datetime(data[start_col], errors="coerce")
    data[end_col] = pd.to_datetime(data[end_col], errors="coerce")

    # logic: пересечение периодов
    mask = (
        (data[start_col] <= pd.to_datetime(date_to)) &
        (data[end_col] >= pd.to_datetime(date_from))
    )

    filtered = data[mask]

    st.success(f"Найдено строк: {len(filtered)}")

    # ---- Save to template ----
    wb = load_workbook("template.xlsx")
    ws = wb["Data"]

    ws.delete_rows(2, ws.max_row)

    for _, r in filtered.iterrows():
        ws.append(list(r))

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(tmp.name)

    with open(tmp.name, "rb") as f:
        st.download_button(
            "⬇ Скачать отчет",
            f,
            file_name="Отчет_по_ресурсам.xlsx"
        )
