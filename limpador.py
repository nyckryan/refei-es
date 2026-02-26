import streamlit as st
import pandas as pd
import re
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

st.set_page_config(
    page_title="Cafeteria Access Report - PDC Campinas",
    layout="wide"
)

st.title("Cafeteria Access Report – PDC Campinas")

EXPECTED_COLUMNS = [
    "Date",
    "Time",
    "Event",
    "Reader",
    "SAPID",
    "LASTNAME",
    "FIRSTNAME",
    "MIDNAME",
    "BADGE",
    "Badge Type",
    "Company",
    "Unit Code",
    "Location"
]

SAPDC_KEYWORDS = [
    "SAPDC CAMPINAS",
    "SAPDC - CAMPINAS",
    "SAPDC-CAMPINAS",
    "PDC CAMPINAS",
    "SA PARTS DISTRIBUTION CENTER"
]

FILTER_COLUMNS = [
    "Reader",
    "Company",
    "Location",
    "Unit Code",
    "FIRSTNAME"
]

MONTH_MAP = {
    "JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4,
    "MAI": 5, "JUN": 6, "JUL": 7, "AGO": 8,
    "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12
}

def load_excel_with_real_header(file):
    raw = pd.read_excel(file, header=None)
    header_row = None
    for i, row in raw.iterrows():
        if "DATE" in row.astype(str).str.upper().values:
            header_row = i
            break
    if header_row is None:
        raise ValueError("Cabeçalho não encontrado")
    return pd.read_excel(file, header=header_row)

def normalize_text(df):
    df = df.copy()
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].astype(str).str.upper().str.strip()
    return df

def is_sapdc_campinas(row):
    for col in FILTER_COLUMNS:
        if col in row and any(k in str(row[col]) for k in SAPDC_KEYWORDS):
            return True
    return False

def enforce_columns(df):
    for col in EXPECTED_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df[EXPECTED_COLUMNS]

def extract_dates_from_filename(filename):
    match = re.search(r"(\d{2})A(\d{2})([A-Z]{3})(\d{4})", filename.upper())
    if not match:
        return None, None, None
    d_start, d_end, mon, year = match.groups()
    month = MONTH_MAP[mon]
    start = f"{month}/{int(d_start)}/{year} 12:00:00 AM"
    end = f"{month}/{int(d_end)}/{year} 11:59:59 PM"
    label = f"{int(d_start):02d}.{month:02d} a {int(d_end):02d}.{month:02d}"
    return start, end, label

def build_query(start, end):
    return (
        f"QUERY:  START DATE:  {start};   "
        f"END DATE:  {end};   "
        "READERS:  BR-GS-CATALAO CAFE 1 ENTRY, "
        "BR-GS-CATALAO CAFE 2 ENTRY, "
        "BR-SP-INTBA C&F 1 OFF CAFE TS #1 ENTRY, "
        "BR-SP-INTBA C&F 1 OFF CAFE TS #2 ENTRY, "
        "BR-SP-INTBA C&F 2 OFF CAFE TS..."
    )

uploaded_file = st.file_uploader(
    "Upload do arquivo Excel (.xlsx)",
    type=["xlsx"]
)

if uploaded_file:
    filename = uploaded_file.name

    df_raw = load_excel_with_real_header(uploaded_file)
    df_norm = normalize_text(df_raw)

    mask = df_norm.apply(is_sapdc_campinas, axis=1)
    df_filtered = df_raw.loc[mask].copy()

    if df_filtered.empty:
        st.error("Nenhum registro SAPDC Campinas encontrado.")
        st.stop()

    df_final = enforce_columns(df_filtered)

    st.metric("Total de itens no relatório", len(df_final))
    st.dataframe(df_final, use_container_width=True)

    start_date, end_date, label = extract_dates_from_filename(filename)
    if not start_date:
        st.error("Nome do arquivo fora do padrão.")
        st.stop()

    query = build_query(start_date, end_date)

    output_name = f"Relatorio de Controle do Restaurante {label}.xlsx"

    df_final.to_excel(output_name, index=False, startrow=1)

    wb = load_workbook(output_name)
    ws = wb.active

    ws["A1"] = "Cafeteria Access Report"
    ws["B1"] = query

    last_row = ws.max_row
    last_col = ws.max_column
    table_ref = f"A2:{chr(64 + last_col)}{last_row}"

    table = Table(
        displayName="CafeteriaAccessReport",
        ref=table_ref
    )

    style = TableStyleInfo(
        name="TableStyleLight1",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=False,
        showColumnStripes=False
    )

    table.tableStyleInfo = style
    ws.add_table(table)

    wb.save(output_name)

    with open(output_name, "rb") as f:
        st.download_button(
            "Baixar relatório final",
            f,
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
