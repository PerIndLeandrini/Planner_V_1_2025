import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from datetime import datetime

st.set_page_config(page_title="Orderbook Generator", layout="wide")
st.title("📘 Orderbook – Generatore da CSV")

st.caption(
    "Carica il CSV (compilato) + il template Orderbook vuoto (solo intestazioni). "
    "Ottieni l'Excel compilato e una preview online."
)

# -----------------------------
# Helpers
# -----------------------------
def read_csv_flexible(uploaded_file) -> pd.DataFrame:
    """
    Legge CSV in modo tollerante:
    - prova con ; poi con ,
    - forza dtype=str
    """
    raw = uploaded_file.getvalue()

    # prova separatore ;
    try:
        df = pd.read_csv(BytesIO(raw), sep=";", dtype=str).fillna("")
        if df.shape[1] > 1:
            return df
    except Exception:
        pass

    # fallback separatore ,
    df = pd.read_csv(BytesIO(raw), sep=",", dtype=str).fillna("")
    return df


def fill_template_from_df(template_bytes: bytes, df: pd.DataFrame, sheet_index: int = 0) -> bytes:
    """
    Riempie il template xlsx (vuoto) usando le intestazioni del template:
    - legge la riga 1 come header del template
    - scrive i dati da riga 2 in poi, allineando per nome colonna
    """
    wb = load_workbook(BytesIO(template_bytes))
    ws = wb.worksheets[sheet_index]

    # header template = riga 1
    template_headers = []
    col = 1
    while True:
        v = ws.cell(row=1, column=col).value
        if v is None or str(v).strip() == "":
            break
        template_headers.append(str(v).strip())
        col += 1

    if not template_headers:
        raise ValueError("Il template non ha intestazioni in riga 1 (o sono vuote).")

    # pulizia righe sotto header (se il template non è proprio vuoto)
    if ws.max_row >= 2:
        ws.delete_rows(2, ws.max_row - 1)

    # normalizza nomi colonna DF (strip)
    df2 = df.copy()
    df2.columns = [str(c).strip() for c in df2.columns]

    # scrittura dati: da riga 2
    start_row = 2
    n = len(df2)

    for i in range(n):
        r_excel = start_row + i
        for j, h in enumerate(template_headers, start=1):
            val = df2.at[i, h] if h in df2.columns else ""
            ws.cell(row=r_excel, column=j, value="" if pd.isna(val) else str(val))

    out = BytesIO()
    wb.save(out)
    return out.getvalue()


# -----------------------------
# UI
# -----------------------------
c1, c2 = st.columns(2)
with c1:
    up_csv = st.file_uploader("📤 Carica CSV orderbook (compilato)", type=["csv"], key="csv")
with c2:
    up_tpl = st.file_uploader("📤 Carica template Orderbook vuoto (.xlsx)", type=["xlsx"], key="tpl")

if not up_csv or not up_tpl:
    st.info("Carica entrambi i file per procedere.")
    st.stop()

# Leggi CSV
try:
    df = read_csv_flexible(up_csv)
except Exception as e:
    st.error(f"Errore lettura CSV: {e}")
    st.stop()

st.subheader("👀 Preview CSV")
st.dataframe(df, use_container_width=True, hide_index=True)

# KPI veloci (facoltativi)
k1, k2 = st.columns(2)
k1.metric("Righe CSV", len(df))
k2.metric("Colonne CSV", df.shape[1])

st.divider()

if st.button("🚀 Genera Orderbook Excel compilato", use_container_width=True):
    try:
        out_bytes = fill_template_from_df(up_tpl.getvalue(), df, sheet_index=0)
    except Exception as e:
        st.error(f"Errore generazione Excel: {e}")
        st.stop()

    fname = f"orderbook_compilato_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

    st.success("✅ Excel generato.")
    st.download_button(
        "⬇️ Scarica Orderbook compilato",
        data=out_bytes,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
