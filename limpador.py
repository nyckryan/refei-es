import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


st.set_page_config(page_title="Cafeteria Access Report - PDC Campinas", layout="wide")
st.title("Cafeteria Access Report – PDC Campinas")


EXPECTED_COLUMNS = [
    "Date", "Time", "Event", "Reader", "SAPID", "LASTNAME", "FIRSTNAME",
    "MIDNAME", "BADGE", "Badge Type", "Company", "Unit Code", "Location"
]

SAPDC_KEYWORDS = [
    "SAPDC CAMPINAS",
    "SAPDC - CAMPINAS",
    "SAPDC-CAMPINAS",
    "PDC CAMPINAS",
    "SA PARTS DISTRIBUTION CENTER"
]

FILTER_COLUMNS = ["Reader", "Company", "Location", "Unit Code", "FIRSTNAME"]

MONTH_MAP = {
    "JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4,
    "MAI": 5, "JUN": 6, "JUL": 7, "AGO": 8,
    "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12
}


# ----------------------------
# Helpers
# ----------------------------
def normalize_text(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].fillna("").astype(str).str.upper().str.strip()
    return df


def enforce_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in EXPECTED_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df[EXPECTED_COLUMNS]


def is_sapdc_campinas(row) -> bool:
    # row é uma Series
    for col in FILTER_COLUMNS:
        if col in row.index:
            val = str(row[col])
            if any(k in val for k in SAPDC_KEYWORDS):
                return True
    return False


def detect_header_row(raw: pd.DataFrame) -> int | None:
    # robusto: procura linha que contenha "DATE" e "TIME" (não precisa ser exato)
    for i, row in raw.iterrows():
        cells = row.fillna("").astype(str).str.upper().str.strip().tolist()
        has_date = any("DATE" == c or "DATE" in c for c in cells)
        has_time = any("TIME" == c or "TIME" in c for c in cells)
        if has_date and has_time:
            return i
    # fallback: só DATE
    for i, row in raw.iterrows():
        cells = row.fillna("").astype(str).str.upper().str.strip().tolist()
        if any("DATE" == c or "DATE" in c for c in cells):
            return i
    return None


def try_read_as_already_standard(file_bytes: bytes) -> pd.DataFrame | None:
    """
    Tenta ler supondo que o arquivo já tem header real na linha 2 (A2),
    ou que a primeira linha do excel já é header.
    Se encontrar pelo menos metade das EXPECTED_COLUMNS, considera como padrão.
    """
    bio = BytesIO(file_bytes)
    # tentativa 1: header na linha 2 (0-based = 1)
    try:
        df1 = pd.read_excel(bio, header=1, engine="openpyxl")
        cols = set(map(str, df1.columns))
        hit = sum([1 for c in EXPECTED_COLUMNS if c in cols])
        if hit >= 7:
            return df1
    except Exception:
        pass

    # tentativa 2: header na primeira linha
    bio = BytesIO(file_bytes)
    try:
        df2 = pd.read_excel(bio, header=0, engine="openpyxl")
        cols = set(map(str, df2.columns))
        hit = sum([1 for c in EXPECTED_COLUMNS if c in cols])
        if hit >= 7:
            return df2
    except Exception:
        pass

    return None


def read_raw_with_detected_header(file_bytes: bytes) -> pd.DataFrame:
    # lê sem header e descobre linha do header
    bio = BytesIO(file_bytes)
    raw = pd.read_excel(bio, header=None, engine="openpyxl")
    header_row = detect_header_row(raw)
    if header_row is None:
        raise ValueError("Cabeçalho não encontrado (não achei DATE/TIME).")

    bio = BytesIO(file_bytes)
    return pd.read_excel(bio, header=header_row, engine="openpyxl")


def extract_dates_from_filename(filename: str, year_default: int) -> tuple[str | None, str | None, str | None]:
    """
    Suporta:
    1) 16A22FEV2026  -> (start,end,label)
    2) ... 24.11 a 30.11  -> usa year_default (não tem ano no nome)
    3) ... 24.11 a 30.11.2026  -> usa ano informado no nome
    """
    name = filename.upper()

    # padrão 1: 16A22FEV2026
    m = re.search(r"(\d{2})A(\d{2})([A-Z]{3})(\d{4})", name)
    if m:
        d_start, d_end, mon, year = m.groups()
        month = MONTH_MAP.get(mon)
        if not month:
            return None, None, None
        start = f"{month}/{int(d_start)}/{year} 12:00:00 AM"
        end   = f"{month}/{int(d_end)}/{year} 11:59:59 PM"
        label = f"{int(d_start):02d}.{month:02d} a {int(d_end):02d}.{month:02d}"
        return start, end, label

    # padrão 2: 24.11 a 30.11 (com ou sem ano)
    m = re.search(r"(\d{2})\.(\d{2})\s*A\s*(\d{2})\.(\d{2})(?:\.(\d{4}))?", name)
    if m:
        d1, m1, d2, m2, y = m.groups()
        year = int(y) if y else year_default
        start = f"{int(m1)}/{int(d1)}/{year} 12:00:00 AM"
        end   = f"{int(m2)}/{int(d2)}/{year} 11:59:59 PM"
        label = f"{int(d1):02d}.{int(m1):02d} a {int(d2):02d}.{int(m2):02d}"
        return start, end, label

    return None, None, None


def build_query(start: str, end: str) -> str:
    # troquei &amp; por &
    return (
        f"QUERY:  START DATE:  {start};   "
        f"END DATE:  {end};   "
        "READERS:  BR-GS-CATALAO CAFE 1 ENTRY, "
        "BR-GS-CATALAO CAFE 2 ENTRY, "
        "BR-SP-INTBA C&F 1 OFF CAFE TS #1 ENTRY, "
        "BR-SP-INTBA C&F 1 OFF CAFE TS #2 ENTRY, "
        "BR-SP-INTBA C&F 2 OFF CAFE TS..."
    )


def write_final_excel(df_final: pd.DataFrame, query: str) -> bytes:
    """
    Gera o Excel final em memória (bytes), com:
    - A1 = Cafeteria Access Report
    - B1 = query
    - Tabela a partir A2 (header) e dados A3...
    - Excel Table aplicada
    """
    out = BytesIO()

    # escreve DF a partir da linha 2 (startrow=1)
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df_final.to_excel(writer, index=False, startrow=1, sheet_name="Report")

    out.seek(0)
    wb = load_workbook(out)
    ws = wb.active

    ws["A1"] = "Cafeteria Access Report"
    ws["B1"] = query

    last_row = ws.max_row
    last_col = ws.max_column
    table_ref = f"A2:{get_column_letter(last_col)}{last_row}"

    # nome de tabela único (excel não aceita repetido)
    table = Table(displayName="CafeteriaAccessReport", ref=table_ref)

    style = TableStyleInfo(
        name="TableStyleLight1",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=False,
        showColumnStripes=False
    )
    table.tableStyleInfo = style
    ws.add_table(table)

    # opcional: freeze panes pra ficar top
    ws.freeze_panes = "A3"

    out2 = BytesIO()
    wb.save(out2)
    return out2.getvalue()


# ----------------------------
# UI
# ----------------------------
apply_filter = st.checkbox("Aplicar filtro SAPDC Campinas", value=True)
year_default = st.number_input("Ano (usado se o nome do arquivo não tiver ano)", value=datetime.now().year, step=1)

uploaded_file = st.file_uploader("Upload do arquivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    file_bytes = uploaded_file.getvalue()
    filename = uploaded_file.name

    # 1) tenta ler como já padronizado
    df_guess = try_read_as_already_standard(file_bytes)

    try:
        if df_guess is not None:
            df_raw = df_guess
            st.info("Arquivo parece já estar padronizado. Vou apenas repadronizar (colunas/ordem/tabela/query).")
        else:
            df_raw = read_raw_with_detected_header(file_bytes)
            st.info("Arquivo parece ser bruto. Vou detectar cabeçalho e padronizar.")

        df_norm = normalize_text(df_raw)

        # 2) filtro opcional
        if apply_filter:
            mask = df_norm.apply(is_sapdc_campinas, axis=1)
            df_filtered = df_raw.loc[mask].copy()
        else:
            df_filtered = df_raw.copy()

        if df_filtered.empty:
            st.error("Nenhum registro encontrado (após filtro). Desmarque o filtro para testar.")
            st.stop()

        # 3) enforce columns
        df_final = enforce_columns(df_filtered)

        st.metric("Total de itens no relatório", len(df_final))
        st.dataframe(df_final, use_container_width=True)

        # 4) datas e query via nome do arquivo
        start_date, end_date, label = extract_dates_from_filename(filename, year_default=year_default)
        if not start_date:
            st.warning("Não consegui extrair datas do nome do arquivo. Vou gerar um nome genérico e query vazia.")
            label = "Padronizado"
            query = ""
        else:
            query = build_query(start_date, end_date)

        output_name = f"Relatorio de Controle do Restaurante {label}.xlsx"

        # 5) escreve excel final (bytes)
        final_bytes = write_final_excel(df_final, query)

        st.download_button(
            "Baixar relatório final",
            data=final_bytes,
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Erro ao padronizar: {e}")
