import streamlit as st
import pandas as pd
import hashlib
import random
from io import BytesIO
from datetime import datetime, timedelta
from openpyxl import load_workbook

# -----------------------------
# CONFIG TEMPLATE / COLONNE
# -----------------------------
TEMPLATE_PATH = "template_orderbook.xlsx"   # nel repo
SHEET_TEMPLATE_INDEX = 0                   # primo foglio del template

# Colonne di scrittura nel template (Excel letters)
OB_COL_DATA_OUT = "X"   # data stimata
OB_COL_STATO_OUT = "AF" # stato

# Colonne che il template usa per matchare (dove sono in template)
# Se nel template queste colonne sono diverse, dimmelo e le cambiamo.
OB_COL_MATERIALE = "A"
OB_COL_ODA = "E"
OB_COL_POS = "F"

# Colonne richieste nel file upload (ORDINI_export)
REQ_UPLOAD_COLS = ["CODICE","ODA","POS","STATO","DATA_PASSAGGIO_PRD"]

# -----------------------------
# AUTH
# -----------------------------
def check_login() -> bool:
    st.sidebar.subheader("🔐 Login")
    u = st.sidebar.text_input("Username")
    p = st.sidebar.text_input("Password", type="password")

    users = st.secrets.get("auth", {}).get("users", [])
    pwds  = st.secrets.get("auth", {}).get("passwords", [])

    if st.sidebar.button("Entra", use_container_width=True):
        if u in users:
            idx = users.index(u)
            if idx < len(pwds) and p == pwds[idx]:
                st.session_state["auth_ok"] = True
            else:
                st.sidebar.error("Password errata.")
        else:
            st.sidebar.error("Utente non valido.")
    return st.session_state.get("auth_ok", False)

# -----------------------------
# UTILS
# -----------------------------
def stable_randint(seed_key: str, a: int, b: int) -> int:
    h = hashlib.sha256(seed_key.encode("utf-8")).hexdigest()
    seed = int(h[:8], 16)
    rnd = random.Random(seed)
    return rnd.randint(a, b)

def parse_dt(x):
    s = str(x).strip() if x is not None else ""
    if not s:
        return None
    dt = pd.to_datetime(s, errors="coerce", dayfirst=True)
    if pd.isna(dt):
        return None
    return dt.to_pydatetime()

def calc_eta_date(base_dt: datetime, stato: str, key: str) -> datetime:
    s = (stato or "").strip().upper()

    if s == "SALA METROLOGICA":
        days = 7
    elif s == "OUTSOURCING":
        days = stable_randint(key, 10, 14)
    elif s == "SCAFFALE":
        days = stable_randint(key, 7, 15)
    else:
        days = stable_randint(key, 18, 30)

    return base_dt + timedelta(days=int(days))

def xl_cell(row_idx: int, col_letter: str) -> str:
    return f"{col_letter}{row_idx}"

def build_planner_map(df_ord: pd.DataFrame) -> dict:
    df = df_ord.copy()

    # normalizza
    for c in ["CODICE","ODA","POS","STATO","DATA_PASSAGGIO_PRD"]:
        df[c] = df[c].astype(str).fillna("").str.strip()

    df["_BASE_DT"] = df["DATA_PASSAGGIO_PRD"].apply(parse_dt)

    planner_map = {}
    for _, r in df.iterrows():
        cod, oda, pos = r["CODICE"], r["ODA"], r["POS"]
        if not (cod and oda and pos):
            continue
        stato = r["STATO"]
        base_dt = r["_BASE_DT"] or datetime.now()
        k = f"{cod}|{oda}|{pos}"
        planner_map[k] = (stato, base_dt)

    return planner_map

def redigi_orderbook(template_bytes: bytes, planner_map: dict) -> bytes:
    wb = load_workbook(BytesIO(template_bytes))
    ws = wb.worksheets[SHEET_TEMPLATE_INDEX]

    updated = 0
    no_match = 0
    skipped = 0

    max_row = ws.max_row

    for r in range(2, max_row + 1):
        cod = str(ws[xl_cell(r, OB_COL_MATERIALE)].value or "").strip()
        oda = str(ws[xl_cell(r, OB_COL_ODA)].value or "").strip()
        pos = str(ws[xl_cell(r, OB_COL_POS)].value or "").strip()

        if not (cod and oda and pos):
            skipped += 1
            continue

        k = f"{cod}|{oda}|{pos}"
        if k not in planner_map:
            no_match += 1
            continue

        stato, base_dt = planner_map[k]

        seed_key = f"{k}|{stato}|{base_dt.strftime('%Y-%m-%d')}"
        target = calc_eta_date(base_dt, stato, seed_key)

        # scrivi su template: X = data, AF = stato
        ws[xl_cell(r, OB_COL_DATA_OUT)].value = target.strftime("%d/%m/%Y")
        ws[xl_cell(r, OB_COL_STATO_OUT)].value = stato

        updated += 1

    out = BytesIO()
    wb.save(out)
    return out.getvalue(), {"updated": updated, "no_match": no_match, "skipped": skipped, "rows": max_row-1}

# -----------------------------
# UI
# -----------------------------
st.set_page_config(page_title="Orderbook Viewer", layout="wide")
st.title("📘 Orderbook Viewer (Clienti)")

if not check_login():
    st.info("Effettua il login dalla sidebar.")
    st.stop()

st.success("✅ Accesso consentito.")

uploaded = st.file_uploader("📤 Carica il file ORDINI_export.xlsx", type=["xlsx"])

if not uploaded:
    st.stop()

try:
    df = pd.read_excel(uploaded, sheet_name="ORDINI", dtype=str).fillna("")
except Exception:
    # fallback: primo foglio
    df = pd.read_excel(uploaded, sheet_name=0, dtype=str).fillna("")

missing = [c for c in REQ_UPLOAD_COLS if c not in df.columns]
if missing:
    st.error(f"File non valido: mancano colonne {missing}")
    st.stop()

st.subheader("🔎 Anteprima dati caricati")
st.dataframe(df.head(30), use_container_width=True)

planner_map = build_planner_map(df)

st.divider()
st.subheader("🧾 Redigi orderbook")

with open(TEMPLATE_PATH, "rb") as f:
    template_bytes = f.read()

if st.button("🚀 Genera Orderbook compilato", use_container_width=True):
    out_bytes, stats = redigi_orderbook(template_bytes, planner_map)

    c1, c2, c3 = st.columns(3)
    c1.metric("Righe aggiornate", stats["updated"])
    c2.metric("Righe senza match", stats["no_match"])
    c3.metric("Righe incomplete template", stats["skipped"])

    fname = f"orderbook_compilato_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    st.download_button(
        "⬇️ Scarica Orderbook compilato",
        data=out_bytes,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
