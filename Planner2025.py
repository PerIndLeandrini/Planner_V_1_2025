import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta
import plotly.express as px
import numpy as np

# --- Config ---
st.set_page_config(page_title="Pianificazione Produzione", layout="wide")
st.title("📦 Gestione Pianificazione Produzione")

# Stati attività (unica fonte di verità)
STATI_ATTIVITA = ["Programmato", "In Produzione", "Controllo qualità", "Da definire", "Completato", "In Corso"]

# ---------- Helper ----------
def parse_eur_to_float(x):
    """Converte stringhe tipo '€ 1.234,56' o numeri in float; altrimenti NaN."""
    if pd.isna(x):
        return np.nan
    if isinstance(x, (int, float, np.integer, np.floating)):
        return float(x)
    s = str(x).strip()
    s = s.replace("€", "").replace("EUR", "").replace("\xa0", "").replace(" ", "")
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return np.nan

def fmt_eur(x):
    """Ritorna stringa in formato € 1.234,56 (italiano)."""
    if pd.isna(x):
        return ""
    try:
        val = float(x)
        s = f"{val:,.2f}"
        s = s.replace(",", "X").replace(".", ",").replace("X", ".")
        return f"€ {s}"
    except Exception:
        return str(x)

def excel_to_datetime(series):
    """Converte seriali Excel, stringhe varie (IT/EN), datetime -> datetime64[ns]."""
    it2en = {"gen":"Jan","feb":"Feb","mar":"Mar","apr":"Apr","mag":"May","giu":"Jun",
             "lug":"Jul","ago":"Aug","set":"Sep","ott":"Oct","nov":"Nov","dic":"Dec"}
    s = series.copy()
    is_dt = s.apply(lambda x: isinstance(x, (pd.Timestamp, datetime)))
    s_num = pd.to_numeric(s, errors="coerce")
    is_num = s_num.notna()

    def _norm_str(v):
        if pd.isna(v): return v
        t = str(v).strip()
        if not t: return t
        tl = t.lower()
        for k,v_ in it2en.items():
            tl = tl.replace(f"-{k}-", f"-{v_}-").replace(f" {k} ", f" {v_} ")
        return tl

    s_str = s.where(~(is_dt | is_num)).apply(_norm_str)
    out = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")

    out.loc[is_num] = pd.to_datetime(s_num[is_num], unit="D", origin="1899-12-30", errors="coerce")
    out.loc[is_dt]  = pd.to_datetime(s[is_dt], errors="coerce")
    mask_str = s_str.notna() & (s_str.astype(str).str.len() > 0)
    out.loc[mask_str] = pd.to_datetime(s_str[mask_str], dayfirst=True, errors="coerce")
    return out

def dt_to_it_str(series):
    return series.dt.strftime("%d/%m/%Y").fillna("")

def to_str_noint_if_int(x):
    """Per ODA/Posizione: niente .0; mantieni stringa pulita."""
    if pd.isna(x):
        return ""
    if isinstance(x, (int, np.integer)):
        return str(int(x))
    if isinstance(x, (float, np.floating)):
        return str(int(x)) if float(x).is_integer() else str(x)
    return str(x).strip()

def add_days(d0, giorni, only_business=False):
    """Somma giorni calendario o lavorativi (lun-ven)."""
    if not only_business:
        return d0 + timedelta(days=giorni)
    d = d0
    added = 0
    while added < giorni:
        d += timedelta(days=1)
        if d.weekday() < 5:  # 0-4 lun-ven
            added += 1
    return d

def pick_color_column(df, candidates=("Operatore","Centro di lavoro","Macchina","Attività","Stato attività")):
    """Restituisce la prima colonna disponibile tra i candidati; None se nessuna presente."""
    for c in candidates:
        if c in df.columns:
            return c
    return None

def normalize_state(s: str) -> str:
    """Normalizza etichette stato alle categorie usate nei Gantt aggregati."""
    if not isinstance(s, str):
        return ""
    t = s.strip().lower()
    mapping = {
        "in produzione": "In Produzione",
        "in lavorazione": "In Produzione",   # per robustezza
        "in corso": "In Corso",
        "programmato": "Programmato",
        "programmata": "Programmato",
        "completato": "Completato",
        "completata": "Completato",
        "controllo qualità": "Controllo Qualità",
        "controllo qualita": "Controllo Qualità",
        "da definire": "Da definire",
    }
    return mapping.get(t, s.strip())

def draw_gantt(df_src, titolo, y_col, color_candidates=("Operatore","Centro di lavoro","Attività","Macchina","Stato attività")):
    """Disegna un Gantt robusto su df_src con colonne Data inizio/Data fine."""
    if df_src is None or df_src.empty:
        st.info(f"Nessuna riga per {titolo}.")
        return
    dfp = df_src.copy()
    dfp["Data inizio"] = pd.to_datetime(dfp["Data inizio"], errors="coerce")
    dfp["Data fine"]   = pd.to_datetime(dfp["Data fine"],   errors="coerce")
    dfp = dfp.dropna(subset=["Data inizio","Data fine"])
    if dfp.empty:
        st.info(f"Nessuna data valida per {titolo}.")
        return

    if "Completamento" in dfp.columns:
        dfp["LabelAvanzamento"] = dfp["Completamento"].fillna(0).astype(int).astype(str) + "%"
    else:
        dfp["LabelAvanzamento"] = ""

    order = pd.Index(dfp.sort_values(["Data inizio","Data fine", y_col])[y_col]).unique().tolist()

    color_col = pick_color_column(dfp, color_candidates)
    kwargs = {}
    if color_col:
        kwargs["color"] = color_col

    fig = px.timeline(
        dfp,
        x_start="Data inizio",
        x_end="Data fine",
        y=y_col,
        text="LabelAvanzamento",
        **kwargs
    )
    fig.update_traces(textposition="inside", insidetextanchor="middle")
    fig.update_yaxes(categoryorder="array", categoryarray=order, autorange="reversed")
    fig.update_layout(title=titolo, xaxis_title="Data", yaxis_title=y_col)
    st.plotly_chart(fig, use_container_width=True)


# =========================
# 1) Caricamento file Excel ORDINI per pianificare (flusso classico)
# =========================
uploaded_file = st.file_uploader("Carica file Excel ordini", type=[".xlsx"])

if uploaded_file:
    # Intestazione reale alla riga 2 (header=1)
    df_raw = pd.read_excel(uploaded_file, header=1)

    # ===== MAPPING ESATTO =====
    # A=0, B=1, C=2, E=4, F=5, J=9, V=21, N=13, W=22, AE=30
    cols = [0, 1, 2, 4, 5, 9, 21, 13, 22, 30]
    df = df_raw.iloc[:, cols].copy()
    df.columns = [
        "MATERIALE", "Revisione", "Descrizione", "ODA", "Posizione",
        "Quantità", "Valore", "Data consegna originale",
        "Data consegna ritrattata", "Note"
    ]

    # --- Conversioni robuste ---
    df["Valore_num"] = df["Valore"].apply(parse_eur_to_float)
    df["Data consegna originale"] = excel_to_datetime(df["Data consegna originale"])
    df["Data consegna ritrattata"] = excel_to_datetime(df["Data consegna ritrattata"])
    df["ODA"] = df["ODA"].apply(to_str_noint_if_int)
    df["Posizione"] = df["Posizione"].apply(to_str_noint_if_int)

    df_display = df.copy()
    df_display["Valore"] = df_display["Valore_num"].apply(fmt_eur)
    df_display["Data consegna originale"] = dt_to_it_str(df_display["Data consegna originale"])
    df_display["Data consegna ritrattata"] = dt_to_it_str(df_display["Data consegna ritrattata"])

    with st.expander("📋 Anteprima Dati Filtrati", expanded=False):
        st.dataframe(df_display.drop(columns=["Valore_num"]), use_container_width=True)

    codice_sel = st.selectbox(
        "Seleziona codice materiale da pianificare",
        df["MATERIALE"].astype(str).str.strip().unique()
    )

    if codice_sel:
        df["MATERIALE"] = df["MATERIALE"].astype(str).str.strip()
        codice_sel = str(codice_sel).strip()

        riga_match = df[df["MATERIALE"] == codice_sel].copy()
        if not riga_match.empty:
            riga_match_display = riga_match.copy()
            riga_match_display["Valore"] = riga_match_display["Valore_num"].apply(fmt_eur)
            riga_match_display["Data consegna originale"] = dt_to_it_str(riga_match_display["Data consegna originale"])
            riga_match_display["Data consegna ritrattata"] = dt_to_it_str(riga_match_display["Data consegna ritrattata"])

            with st.expander("📄 Tutte le righe trovate per questo materiale", expanded=False):
                st.dataframe(riga_match_display.drop(columns=["Valore_num"]), use_container_width=True)

            def _fmt_opt(i):
                oda = riga_match.loc[i, 'ODA']
                pos = riga_match.loc[i, 'Posizione']
                dco = riga_match.loc[i, 'Data consegna originale']
                dco_s = dco.strftime("%d/%m/%Y") if pd.notna(dco) else ""
                return f"ODA: {oda} | Posizione: {pos} | Consegna: {dco_s}"

            index_riga = st.selectbox(
                "Seleziona la riga da pianificare",
                riga_match.index,
                format_func=_fmt_opt
            )

            riga = riga_match.loc[index_riga]
            st.markdown("### 📄 Riga selezionata per la pianificazione")
            dettaglio = riga.to_dict()
            dettaglio["Valore"] = fmt_eur(dettaglio.get("Valore_num"))
            for k in ["Data consegna originale", "Data consegna ritrattata"]:
                if isinstance(dettaglio.get(k), pd.Timestamp):
                    dettaglio[k] = dettaglio[k].strftime("%d/%m/%Y") if pd.notna(dettaglio[k]) else ""
            st.write(dettaglio)

            st.markdown("### 🛠️ Pianificazione Produzione")
            data_inizio = st.date_input("Data inizio produzione", datetime.today(), format="DD/MM/YYYY")

            attivita_possibili = [
                "Progettazione", "Preparazione materiale", "F1", "F2", "F3", "F4", "F5", "F6",
                "Trattamento", "Lav. Esterna", "Verniciatura", "Marcatura", "Controllo Qualità", "Imballaggio"
            ]
            operatori = ["PAOLO", "TONINO", "MICHELE", "ALESSANDRO", "VALERIO", "LUCA", "MARCO", "TOMMI",
                         "IACOPO", "ALESSIO", "ALEANDRO", "DANIELE", "SOUKAINA", "SIMONE", "MICHEL", "ELENA"]
            centri = ["Hurco", "Mazak 5assi", "Mazak 4assi", "Mazak HCN", "Mazak 3assi", "Hyundai", "Macchine //", "DMG Mori", "Takisawa", "SALA METROLOGICA", "TRONCATRICE"]
            fornitori = {
                "Trattamento": ["MOCHEM", "SAMACROMO", "ART.ING.", "F.LLI BUGLI", "AVIORUBBER"],
                "Verniciatura": ["Verniciatura industriale", "Birindelli"],
                "Lav. Esterna": ["Galli & Sesti", "Pazzaglia", "Donatello"],
                "Marcatura": ["Pazzaglia", "INTERNA LUPPICHINI"]
            }

            n_attivita = st.number_input("Quante attività vuoi pianificare?", min_value=1, max_value=15, value=3)
            pianificazione = []

            for i in range(n_attivita):
                st.markdown(f"#### Attività {i+1}")
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    attivita = st.selectbox(f"Attività {i+1}", attivita_possibili, key=f"att_{i}")
                with col2:
                    operatore = st.selectbox(f"Operatore", operatori, key=f"op_{i}")
                with col3:
                    centro = st.selectbox(f"Centro di lavoro", centri, key=f"cl_{i}")
                with col4:
                    stato = st.selectbox(f"Stato attività", STATI_ATTIVITA, key=f"st_{i}")

                data_ini = st.date_input(f"Data inizio attività {i+1}", datetime.today(), format="DD/MM/YYYY", key=f"dini_{i}")
                data_fine = st.date_input(f"Data fine attività {i+1}", datetime.today(), format="DD/MM/YYYY", key=f"dfine_{i}")
                completamento = st.slider(f"Completamento attività {i+1} (%)", 0, 100, 0, step=5, key=f"comp_{i}")

                # Auto-stato: se completamento 100 -> Completato
                if completamento == 100 and stato != "Completato":
                    stato = "Completato"
                    st.session_state[f"st_{i}"] = "Completato"

                fornitore = ""
                if attivita in fornitori and fornitori[attivita]:
                    fornitore = st.selectbox(f"Fornitore per {attivita}", fornitori[attivita], key=f"forn_{i}")

                pianificazione.append({
                    "Attività": attivita,
                    "Operatore": operatore,
                    "Centro di lavoro": centro,
                    "Data inizio": data_ini.strftime("%d/%m/%Y"),
                    "Data fine": data_fine.strftime("%d/%m/%Y"),
                    "Stato attività": stato,
                    "Completamento": completamento,
                    "Fornitore": fornitore
                })

            st.markdown("### 💾 Salvataggio Pianificazione")
            if st.button("Salva pianificazione su CSV"):
                output_data = []
                for att in pianificazione:
                    output_data.append({
                        "MATERIALE": riga["MATERIALE"],
                        "Revisione": riga["Revisione"],
                        "Descrizione": riga["Descrizione"],
                        "ODA": riga["ODA"],
                        "Posizione": riga["Posizione"],
                        "Quantità": riga["Quantità"],
                        "Valore": riga["Valore_num"],
                        "Valore (visuale)": fmt_eur(riga["Valore_num"]),
                        "Data consegna originale": riga["Data consegna originale"].strftime("%d/%m/%Y") if pd.notna(riga["Data consegna originale"]) else "",
                        "Data consegna ritrattata": riga["Data consegna ritrattata"].strftime("%d/%m/%Y") if pd.notna(riga["Data consegna ritrattata"]) else "",
                        "Note": riga["Note"],
                        **att
                    })
                df_out = pd.DataFrame(output_data)
                os.makedirs("dati_pianificati", exist_ok=True)
                nome_file = f"dati_pianificati/pianificazione_{codice_sel}.csv"
                df_out.to_csv(nome_file, index=False, encoding="utf-8-sig")
                st.success(f"Dati salvati in {nome_file}")

            st.markdown("### 📆 Gantt delle Attività")
            df_gantt = pd.DataFrame(pianificazione)
            if not df_gantt.empty:
                df_gantt["Inizio"] = pd.to_datetime(df_gantt["Data inizio"], format="%d/%m/%Y", errors="coerce")
                df_gantt["Fine"]   = pd.to_datetime(df_gantt["Data fine"],   format="%d/%m/%Y", errors="coerce")
                df_gantt = df_gantt.dropna(subset=["Inizio","Fine"])
                if not df_gantt.empty:
                    df_gantt["LabelAvanzamento"] = df_gantt["Completamento"].fillna(0).astype(int).astype(str) + "%"

                    order = pd.Index(
                        df_gantt.sort_values(["Inizio", "Fine", "Attività"])["Attività"]
                    ).unique().tolist()

                    # hardening: scegli color disponibile
                    color_col = pick_color_column(df_gantt, ("Operatore","Centro di lavoro","Attività","Stato attività"))
                    kwargs = {}
                    if color_col:
                        kwargs["color"] = color_col

                    fig = px.timeline(
                        df_gantt,
                        x_start="Inizio",
                        x_end="Fine",
                        y="Attività",
                        text="LabelAvanzamento",
                        **kwargs
                    )
                    fig.update_traces(textposition="inside", insidetextanchor="middle")
                    fig.update_yaxes(categoryorder="array", categoryarray=order, autorange="reversed")
                    fig.update_layout(
                        xaxis_title="Data",
                        yaxis_title="Attività pianificata",
                        title=f"Gantt - {codice_sel}"
                    )
                    st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("⚠️ Nessuna riga trovata per il materiale selezionato.")
else:
    st.info("📂 Carica un file Excel per iniziare (oppure vai sotto e carica un CSV esistente).")

# =========================
# 2) Carica Pianificazione ESISTENTE (CSV) e mostra TUTTO come in pianificazione
# =========================
st.markdown("---")
st.header("📂 Carica Pianificazione Esistente")

csv_file = st.file_uploader("Carica un file CSV di pianificazione esistente", type=[".csv"], key="csv")

if csv_file:
    df_pianif = pd.read_csv(csv_file)

    # Normalizza date per la vista e per il Gantt
    for col in ["Data inizio", "Data fine"]:
        if col in df_pianif.columns:
            dt = pd.to_datetime(df_pianif[col], dayfirst=True, errors="coerce")
            df_pianif[col] = dt.dt.strftime("%d/%m/%Y").fillna("")

    if "Valore" in df_pianif.columns and pd.api.types.is_numeric_dtype(df_pianif["Valore"]):
        df_pianif["Valore (visuale)"] = df_pianif["Valore"].apply(fmt_eur)

    st.subheader("📋 Pianificazione caricata (tutte le righe)")
    st.dataframe(df_pianif, use_container_width=True, height=380)
    # === EXPANDER: Gantt per macchina (dal CSV) — per stato ===
    with st.expander("📦 Gantt per macchina (dal CSV) — per stato", expanded=False):
        dfn = df_pianif.copy()
    
        # normalizza etichette stato (usa la tua helper normalize_state)
        if "Stato attività" in dfn.columns:
            dfn["Stato attività"] = dfn["Stato attività"].apply(normalize_state)
        else:
            dfn["Stato attività"] = ""
    
        # parse date (il CSV qui è formattato DD/MM/YYYY ma i grafici vogliono datetime)
        for c in ["Data inizio","Data fine"]:
            if c in dfn.columns:
                dfn[c] = pd.to_datetime(dfn[c], dayfirst=True, errors="coerce")
    
        # bucket: In lavorazione = In Produzione ∪ In Corso ; Programmato = Programmato
        mask_lav  = dfn["Stato attività"].isin(["In Produzione","In Corso"])
        mask_prog = dfn["Stato attività"].eq("Programmato")
    
        if "Macchina" not in dfn.columns:
            st.info("Nel CSV non c'è la colonna 'Macchina': impossibile raggruppare per macchina.")
        else:
            macchine = (
                dfn["Macchina"]
                .dropna()
                .astype(str).str.strip()
                .unique()
                .tolist()
            )
            if not macchine:
                st.info("Nessuna macchina trovata nel CSV.")
            else:
                for mac in macchine:
                    st.markdown(f"#### 🏭 {mac}")
    
                    sub = dfn[dfn["Macchina"] == mac]
                    sub_lav  = sub[mask_lav  & (sub["Macchina"] == mac)]
                    sub_prog = sub[mask_prog & (sub["Macchina"] == mac)]
    
                    # y = MATERIALE per vedere a colpo d'occhio i pezzi in coda su quella macchina
                    draw_gantt(sub_lav,  f"Gantt — {mac} — In lavorazione", "MATERIALE")
                    draw_gantt(sub_prog, f"Gantt — {mac} — Programmato",   "MATERIALE")
    
                    st.markdown("---")
    # === FINE EXPANDER ===



    # Selezione per MATERIALE (come nel flusso di pianificazione)
    materiali = df_pianif["MATERIALE"].astype(str).str.strip().unique().tolist()
    mat_sel = st.selectbox("Seleziona materiale per dettaglio e Gantt", materiali)

    df_mat = df_pianif[df_pianif["MATERIALE"].astype(str).str.strip() == str(mat_sel).strip()].copy()

    # Gantt per il materiale selezionato
    st.markdown("### 📆 Gantt del materiale selezionato")
    if not df_mat.empty and {"Data inizio","Data fine","Attività"}.issubset(df_mat.columns):
        gantt = df_mat.copy()
        gantt["Inizio"] = pd.to_datetime(gantt["Data inizio"], dayfirst=True, errors="coerce")
        gantt["Fine"]   = pd.to_datetime(gantt["Data fine"],   dayfirst=True, errors="coerce")
        gantt = gantt.dropna(subset=["Inizio","Fine"])
        if not gantt.empty:
            gantt["LabelAvanzamento"] = gantt["Completamento"].fillna(0).astype(int).astype(str) + "%"

            # ordina per data d'inizio (più vecchia in alto)
            order = pd.Index(
                gantt.sort_values(["Inizio", "Fine", "Attività"])["Attività"]
            ).unique().tolist()

            # hardening: scegli color disponibile
            color_col = pick_color_column(gantt, ("Operatore","Centro di lavoro","Macchina","Attività","Stato attività"))
            kwargs = {}
            if color_col:
                kwargs["color"] = color_col

            fig = px.timeline(
                gantt,
                x_start="Inizio",
                x_end="Fine",
                y="Attività",
                text="LabelAvanzamento",
                **kwargs
            )
            fig.update_traces(textposition="inside", insidetextanchor="middle")
            fig.update_yaxes(categoryorder="array", categoryarray=order, autorange="reversed")
            fig.update_layout(
                xaxis_title="Data",
                yaxis_title="Attività",
                title=f"Gantt - {mat_sel}"
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Nessuna data valida per disegnare il Gantt.")

    # --- Modifica SEQUENZIALE di tutte le attività del materiale selezionato ---
    st.markdown("### ✏️ Modifica SEQUENZIALE di tutte le attività del materiale selezionato")

    if not df_mat.empty:
        # Liste di riferimento
        operatori = ["PAOLO", "TONINO", "MICHELE", "ALESSANDRO", "VALERIO", "LUCA", "MARCO", "TOMMI",
                     "IACOPO", "ALESSIO", "ALEANDRO", "DANIELE", "SOUKAINA", "SIMONE", "MICHEL", "ELENA"]
        centri = ["Hurco", "Mazak 5assi", "Mazak 4assi", "Mazak HCN", "Mazak 3assi", "Hyundai", "Macchine //",
                  "DMG Mori", "Takisawa", "SALA METROLOGICA", "TRONCATRICE"]
        fornitori = {
            "Trattamento": ["MOCHEM", "SAMACROMO", "ART.ING.", "F.LLI BUGLI", "AVIORUBBER"],
            "Verniciatura": ["Verniciatura industriale", "Birindelli"],
            "Lav. Esterna": ["Galli & Sesti", "Pazzaglia", "Donatello"],
            "Marcatura": ["Pazzaglia", "INTERNA LUPPICHINI"]
        }

        edited_rows = []

        st.caption("🔎 Stai modificando TUTTE le fasi del materiale selezionato. Le modifiche verranno salvate insieme.")
        for i, (idx, r) in enumerate(df_mat.iterrows(), start=1):
            st.markdown(f"#### Attività {i} — indice riga CSV: `{idx}`")
            st.write(f"**Attività:** {r.get('Attività','')}  |  **Operatore attuale:** {r.get('Operatore','')}  |  **Centro:** {r.get('Centro di lavoro','')}")

            # Date (robuste al formato IT)
            di = pd.to_datetime(r.get("Data inizio",""), dayfirst=True, errors="coerce")
            dfine = pd.to_datetime(r.get("Data fine",""), dayfirst=True, errors="coerce")
            if pd.isna(di): di = datetime.today()
            if pd.isna(dfine): dfine = datetime.today()

            col1, col2, col3 = st.columns(3)
            with col1:
                new_di = st.date_input("Data inizio", di, format="DD/MM/YYYY", key=f"bulk_di_{idx}")
            with col2:
                new_df = st.date_input("Data fine", dfine, format="DD/MM/YYYY", key=f"bulk_df_{idx}")
            with col3:
                comp_default_raw = r.get("Completamento", 0)
                try:
                    comp_default = int(comp_default_raw) if pd.notna(comp_default_raw) else 0
                except Exception:
                    comp_default = 0
                new_comp = st.slider("Completamento (%)", 0, 100, comp_default, step=5, key=f"bulk_comp_{idx}")

            col4, col5, col6 = st.columns(3)
            with col4:
                stato_opts = STATI_ATTIVITA
                stato_att = r.get("Stato attività")
                default_idx = stato_opts.index(stato_att) if stato_att in stato_opts else 0
                new_stato = st.selectbox("Stato attività", stato_opts, index=default_idx, key=f"bulk_stato_{idx}")
            with col5:
                op_default = r.get("Operatore") if pd.notna(r.get("Operatore","")) else operatori[0]
                idx_op = operatori.index(op_default) if op_default in operatori else 0
                new_op = st.selectbox("Operatore", operatori, index=idx_op, key=f"bulk_op_{idx}")
            with col6:
                cl_default = r.get("Centro di lavoro") if pd.notna(r.get("Centro di lavoro","")) else centri[0]
                idx_cl = centri.index(cl_default) if cl_default in centri else 0
                new_cl = st.selectbox("Centro di lavoro", centri, index=idx_cl, key=f"bulk_cl_{idx}")

            # Fornitore dinamico se la tipologia attività lo prevede
            att = r.get("Attività")
            forn_list = fornitori.get(att, [])
            new_forn = r.get("Fornitore","")
            if forn_list:
                idx_f = forn_list.index(new_forn) if new_forn in forn_list else 0
                new_forn = st.selectbox(f"Fornitore per {att}", forn_list, index=idx_f, key=f"bulk_forn_{idx}")

            # Coerenza stato ↔ completamento
            if new_comp == 100 and new_stato != "Completato":
                new_stato = "Completato"
                st.session_state[f"bulk_stato_{idx}"] = "Completato"
                st.caption("🔁 Stato impostato automaticamente a **Completato** (100%).")
            if new_stato == "Completato" and new_comp < 100:
                new_comp = 100
                st.session_state[f"bulk_comp_{idx}"] = 100

            edited_rows.append({
                "idx": idx,
                "Data inizio": new_di.strftime("%d/%m/%Y"),
                "Data fine": new_df.strftime("%d/%m/%Y"),
                "Completamento": new_comp,
                "Stato attività": new_stato,
                "Operatore": new_op,
                "Centro di lavoro": new_cl,
                "Fornitore": new_forn if forn_list else r.get("Fornitore","")
            })

            st.markdown("---")

        # Salvataggio in blocco
        if st.button("💾 Salva TUTTE le modifiche del materiale"):
            os.makedirs("dati_pianificati", exist_ok=True)
            for upd in edited_rows:
                ridx = upd.pop("idx")
                for k, v in upd.items():
                    df_pianif.at[ridx, k] = v

            out_path = f"dati_pianificati/{csv_file.name}"
            st.success(f"Aggiornato e salvato in {out_path}")

            # assicura formato ITA prima del salvataggio
            df_save = df_pianif.copy()
            for c in ["Data inizio","Data fine"]:
                df_save[c] = pd.to_datetime(df_save[c], dayfirst=True, errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
            df_save.to_csv(out_path, index=False, encoding="utf-8-sig")

            # Refresh preview
            st.dataframe(df_pianif, use_container_width=True, height=300)

# =========================
# 3) Nuovo flusso: Assegna Macchina e Modalità, poi genera Pianificazione
# =========================

st.markdown("---")
st.header("🧭 Assegnazione Macchina & Modalità")

MACCHINE = [
    "DMG MORI","HURCO","HYUNDAI","MAZAK 3 ASSI","MAZAK 4 ASSI","MAZAK 5 ASSI","MAZAK HCN",
    "SALA SMN","TAKISAWA","TORNIO PAOLO","TORNIO TONINO"
]
MODALITA = ["", "In lavorazione", "Programmazione"]

# (Opzionale) Pianificazione esistente per concatenare correttamente le code per macchina
st.caption("🔗 (Opzionale) Carica una pianificazione esistente per concatenare correttamente le code per macchina.")
csv_esistente = st.file_uploader("Carica CSV pianificazione esistente (opzionale)", type=[".csv"], key="csv_esistente")
df_exist = None
if csv_esistente:
    df_exist = pd.read_csv(csv_esistente)
    # normalizza date e macchina
    for col in ["Data inizio","Data fine"]:
        if col in df_exist.columns:
            dtx = pd.to_datetime(df_exist[col], dayfirst=True, errors="coerce")
            df_exist[col] = dtx
    if "Macchina" in df_exist.columns:
        df_exist["Macchina"] = df_exist["Macchina"].astype(str).str.upper().str.strip()

# Prepara una vista sintetica per l’assegnazione
if uploaded_file:
    base_cols = [
        "MATERIALE","Revisione","Descrizione","ODA","Posizione",
        "Quantità","Valore_num","Data consegna originale","Data consegna ritrattata","Note"
    ]
    ordini = df[base_cols].copy()
    ordini["Assegna Macchina"] = ""
    ordini["Modalità"] = ""  # "In lavorazione" o "Programmazione"

    st.subheader("📋 Seleziona Macchina e Modalità riga per riga")
    st.caption("Imposta **Assegna Macchina** e **Modalità** (vuoto = ignorata). Le righe con Modalità impostata verranno elaborate sotto.")

    # Editor tabellare per assegnazioni
    ordini_edit = st.data_editor(
        ordini,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Valore_num": st.column_config.NumberColumn("Valore (€)", help="Valore ordine (numero).", format="%.2f"),
            "Assegna Macchina": st.column_config.SelectboxColumn(options=MACCHINE, required=False),
            "Modalità": st.column_config.SelectboxColumn(options=MODALITA, required=False),
            "Data consegna originale": st.column_config.DateColumn(format="DD/MM/YYYY"),
            "Data consegna ritrattata": st.column_config.DateColumn(format="DD/MM/YYYY"),
        }
    )

    st.markdown("### ⚙️ Configurazione semplificata per Programmazione")
    st.caption("Per i materiali in **Programmazione** imposta quanti step simulare. Durate fisse: Lavoro=3g; Trattamenti/Lav. Esterne=10g; CQ=2g; Imballaggio=1g.")
    colA, colB, colC = st.columns(3)
    with colA:
        usa_giorni_lavorativi = st.toggle("Usa giorni lavorativi (lun–ven)", value=True)
    with colB:
        unico_csv = st.toggle("Salva in **un solo CSV** consolidato", value=True)
    with colC:
        mostra_gantt_macchine = st.toggle("Mostra Gantt per macchina", value=True)

    # Filtro righe selezionate (robusto)
    tmp = ordini_edit.copy()
    tmp["Assegna Macchina"] = tmp["Assegna Macchina"].fillna("").astype(str).str.strip()
    tmp["Modalità"] = tmp["Modalità"].fillna("").astype(str).str.strip()
    sel = tmp[(tmp["Modalità"].isin(["In lavorazione", "Programmazione"])) & (tmp["Assegna Macchina"] != "")]
    if sel.empty:
        st.info("Seleziona almeno una riga impostando **Assegna Macchina** e **Modalità**.")
    else:
        st.subheader("🧾 Configurazione fasi per **In lavorazione** (manuale) e **Programmazione** (automatica)")

        # 1) Editor MANUALE per "In lavorazione"
        df_lav = sel[sel["Modalità"] == "In lavorazione"].copy()
        out_manual = []
        if not df_lav.empty:
            st.markdown("#### ✍️ In lavorazione — inserisci manualmente fasi, date, stato, completamento")
            for i, r in df_lav.reset_index(drop=True).iterrows():
                with st.container():
                    st.markdown("---")
                    st.write(f"**{r['MATERIALE']}** — {r['Descrizione']} | **Macchina:** {r['Assegna Macchina']}")
                    n_fasi = st.number_input(f"Quante fasi vuoi definire per {r['MATERIALE']} (manuale)?",
                                             min_value=1, max_value=20, value=3, key=f"m_n_{i}")
                    for k in range(n_fasi):
                        st.markdown(f"**Fase {k+1}**")
                        c1, c2, c3, c4 = st.columns(4)
                        with c1:
                            nome_fase = st.text_input("Nome fase", f"F{k+1}", key=f"m_nome_{i}_{k}")
                        with c2:
                            d_ini = st.date_input("Data inizio", datetime.today(), format="DD/MM/YYYY", key=f"m_di_{i}_{k}")
                        with c3:
                            d_fine = st.date_input("Data fine", datetime.today(), format="DD/MM/YYYY", key=f"m_df_{i}_{k}")
                        with c4:
                            stato = st.selectbox("Stato", STATI_ATTIVITA, key=f"m_st_{i}_{k}")
                        comp = st.slider("Completamento (%)", 0, 100, 0, 5, key=f"m_cp_{i}_{k}")

                        # Coerenza stato ↔ completamento
                        if comp == 100 and stato != "Completato":
                            stato = "Completato"
                            st.session_state[f"m_st_{i}_{k}"] = "Completato"
                        if st.session_state.get(f"m_st_{i}_{k}") == "Completato" and comp < 100:
                            comp = 100
                            st.session_state[f"m_cp_{i}_{k}"] = 100

                        out_manual.append({
                            "MATERIALE": r["MATERIALE"],
                            "Revisione": r["Revisione"],
                            "Descrizione": r["Descrizione"],
                            "ODA": r["ODA"],
                            "Posizione": r["Posizione"],
                            "Quantità": r["Quantità"],
                            "Valore": r["Valore_num"],
                            "Macchina": r["Assegna Macchina"],
                            "Attività": nome_fase,
                            "Data inizio": pd.to_datetime(d_ini, dayfirst=True, errors="coerce"),
                            "Data fine": pd.to_datetime(d_fine, dayfirst=True, errors="coerce"),
                            "Stato attività": stato,
                            "Completamento": comp,
                            "Fornitore": ""
                        })

        # 2) CONFIG per “Programmazione” (solo numeri di fasi; le date le calcolo io)
        df_prog = sel[sel["Modalità"] == "Programmazione"].copy()
        cfg_prog = {}
        if not df_prog.empty:
            st.markdown("#### 🧪 Programmazione — definisci solo quanti step simulare per ciascun materiale")
            st.caption("Imposterò le date in **coda** alla macchina: Lavoro=3g cadauna; Tratt./Esterna=10g cadauna; CQ 2g; Imballaggio 1g.")
            for i, r in df_prog.reset_index(drop=True).iterrows():
                with st.container():
                    st.markdown("---")
                    st.write(f"**{r['MATERIALE']}** — {r['Descrizione']} | **Macchina:** {r['Assegna Macchina']}")
                    c1, c2, c3, c4, c5 = st.columns(5)
                    with c1:
                        n_fasi_lav = st.number_input("Fasi di lavoro (×3g)", min_value=0, max_value=20, value=3, key=f"p_lav_{i}")
                    with c2:
                        n_tratt = st.number_input("Trattamenti (×10g)", min_value=0, max_value=10, value=1, key=f"p_tr_{i}")  # default 1 per ~3 settimane
                    with c3:
                        n_ext = st.number_input("Lav. esterne (×10g)", min_value=0, max_value=10, value=0, key=f"p_ext_{i}")
                    with c4:
                        add_cq = st.checkbox("Controllo Qualità (2g)", value=True, key=f"p_cq_{i}")
                    with c5:
                        add_imball = st.checkbox("Imballaggio (1g)", value=True, key=f"p_pack_{i}")

                    cfg_prog[i] = {
                        "mat": r["MATERIALE"],
                        "desc": r["Descrizione"],
                        "rev": r["Revisione"],
                        "oda": r["ODA"],
                        "pos": r["Posizione"],
                        "qta": r["Quantità"],
                        "val": r["Valore_num"],
                        "mac": r["Assegna Macchina"],
                        "n_lav": int(n_fasi_lav),
                        "n_tr": int(n_tratt),
                        "n_ex": int(n_ext),
                        "cq": bool(add_cq),
                        "pack": bool(add_imball)
                    }

        # 3) COSTRUISCO LE CODE PER MACCHINA (ultima data come base)
        def last_end_for_machine(machine, df_exist, manual_rows):
            """Trova l'ultima Data fine per la macchina da:
               - CSV esistente (se caricato)
               - righe manuali inserite in questa sessione
            """
            dates = []
            mch = str(machine).upper().strip()
            if df_exist is not None and not df_exist.empty and "Macchina" in df_exist.columns and "Data fine" in df_exist.columns:
                sub = df_exist[df_exist["Macchina"] == mch].copy()
                if not sub.empty:
                    dtf = pd.to_datetime(sub["Data fine"], errors="coerce")
                    dtf = dtf.dropna()
                    if not dtf.empty:
                        dates.append(dtf.max())
            if manual_rows:
                m = [r["Data fine"] for r in manual_rows if str(r.get("Macchina","")).upper().strip() == mch]
                m = [pd.to_datetime(x, errors="coerce") for x in m]
                m = [x for x in m if pd.notna(x)]
                if m:
                    dates.append(max(m))
            if not dates:
                return pd.Timestamp(datetime.today().date())
            return max(dates)

        # 4) GENERO LA SIMULAZIONE IN CODA
        DUR = {"lavoro": 3, "trattamento": 10, "esterna": 10, "cq": 2, "imballaggio": 1}
        out_auto = []
        if cfg_prog:
            # raggruppa per macchina rispettando l'ordine di apparizione
            by_mac = {}
            for i, cfg in cfg_prog.items():
                by_mac.setdefault(cfg["mac"], []).append(cfg)

            for mac, items in by_mac.items():
                cursor = last_end_for_machine(mac, df_exist, out_manual)
                cursor = add_days(cursor, 1, usa_giorni_lavorativi)  # inizio giorno successivo

                for cfg in items:
                    # fasi di lavoro
                    for j in range(cfg["n_lav"]):
                        di = cursor
                        dfine = add_days(di, DUR["lavoro"], usa_giorni_lavorativi)
                        out_auto.append({
                            "MATERIALE": cfg["mat"], "Revisione": cfg["rev"], "Descrizione": cfg["desc"],
                            "ODA": cfg["oda"], "Posizione": cfg["pos"], "Quantità": cfg["qta"], "Valore": cfg["val"],
                            "Macchina": mac, "Attività": f"Lavoro {j+1}",
                            "Data inizio": di, "Data fine": dfine,
                            "Stato attività": "Programmato", "Completamento": 0, "Fornitore": ""
                        })
                        cursor = add_days(dfine, 1, usa_giorni_lavorativi)

                    # trattamenti
                    for j in range(cfg["n_tr"]):
                        di = cursor
                        dfine = add_days(di, DUR["trattamento"], usa_giorni_lavorativi)
                        out_auto.append({
                            "MATERIALE": cfg["mat"], "Revisione": cfg["rev"], "Descrizione": cfg["desc"],
                            "ODA": cfg["oda"], "Posizione": cfg["pos"], "Quantità": cfg["qta"], "Valore": cfg["val"],
                            "Macchina": mac, "Attività": f"Trattamento {j+1}",
                            "Data inizio": di, "Data fine": dfine,
                            "Stato attività": "Programmato", "Completamento": 0, "Fornitore": ""
                        })
                        cursor = add_days(dfine, 1, usa_giorni_lavorativi)

                    # lavorazioni esterne
                    for j in range(cfg["n_ex"]):
                        di = cursor
                        dfine = add_days(di, DUR["esterna"], usa_giorni_lavorativi)
                        out_auto.append({
                            "MATERIALE": cfg["mat"], "Revisione": cfg["rev"], "Descrizione": cfg["desc"],
                            "ODA": cfg["oda"], "Posizione": cfg["pos"], "Quantità": cfg["qta"], "Valore": cfg["val"],
                            "Macchina": mac, "Attività": f"Lavorazione Esterna {j+1}",
                            "Data inizio": di, "Data fine": dfine,
                            "Stato attività": "Programmato", "Completamento": 0, "Fornitore": ""
                        })
                        cursor = add_days(dfine, 1, usa_giorni_lavorativi)

                    # CQ
                    if cfg["cq"]:
                        di = cursor
                        dfine = add_days(di, DUR["cq"], usa_giorni_lavorativi)
                        out_auto.append({
                            "MATERIALE": cfg["mat"], "Revisione": cfg["rev"], "Descrizione": cfg["desc"],
                            "ODA": cfg["oda"], "Posizione": cfg["pos"], "Quantità": cfg["qta"], "Valore": cfg["val"],
                            "Macchina": mac, "Attività": "Controllo Qualità",
                            "Data inizio": di, "Data fine": dfine,
                            "Stato attività": "Programmato", "Completamento": 0, "Fornitore": ""
                        })
                        cursor = add_days(dfine, 1, usa_giorni_lavorativi)

                    # Imballaggio
                    if cfg["pack"]:
                        di = cursor
                        dfine = add_days(di, DUR["imballaggio"], usa_giorni_lavorativi)
                        out_auto.append({
                            "MATERIALE": cfg["mat"], "Revisione": cfg["rev"], "Descrizione": cfg["desc"],
                            "ODA": cfg["oda"], "Posizione": cfg["pos"], "Quantità": cfg["qta"], "Valore": cfg["val"],
                            "Macchina": mac, "Attività": "Imballaggio",
                            "Data inizio": di, "Data fine": dfine,
                            "Stato attività": "Programmato", "Completamento": 0, "Fornitore": ""
                        })
                        cursor = add_days(dfine, 1, usa_giorni_lavorativi)

        # 5) OUTPUT CONSOLIDATO + (opzionale) Gantt per macchina
        st.markdown("### 🧾 Anteprima pianificazione generata")
        df_out = pd.DataFrame((out_manual or []) + (out_auto or []))
        if not df_out.empty:
            if "Data inizio" in df_out.columns:
                df_out["Data inizio"] = pd.to_datetime(df_out["Data inizio"], errors="coerce")
            if "Data fine" in df_out.columns:
                df_out["Data fine"] = pd.to_datetime(df_out["Data fine"], errors="coerce")
            df_out = df_out.sort_values(["Macchina","Data inizio","MATERIALE","Attività"], na_position="last")

            df_show = df_out.copy()
            for c in ["Data inizio","Data fine"]:
                df_show[c] = df_show[c].dt.strftime("%d/%m/%Y").fillna("")
            if "Valore" in df_show.columns:
                df_show["Valore (visuale)"] = df_show["Valore"].apply(fmt_eur)

            st.dataframe(df_show, use_container_width=True, height=420)
            with st.expander("📦 Gantt per macchina — per stato", expanded=False):
               if df_out is None or df_out.empty:
                   st.info("Nessuna pianificazione disponibile.")
               else:
                   dfn = df_out.copy()

                   # normalizza stato
                   if "Stato attività" in dfn.columns:
                       dfn["Stato attività"] = dfn["Stato attività"].apply(normalize_state)
                   else:
                       dfn["Stato attività"] = ""

                   # assicura datetime per il grafico
                   for c in ["Data inizio","Data fine"]:
                       if c in dfn.columns:
                           dfn[c] = pd.to_datetime(dfn[c], errors="coerce")

                   # bucket stati
                   mask_lav  = dfn["Stato attività"].isin(["In Produzione","In Corso"])
                   mask_prog = dfn["Stato attività"].eq("Programmato")

                   if "Macchina" not in dfn.columns:
                       st.info("Nessuna colonna 'Macchina' trovata.")
                   else:
                       macchine = dfn["Macchina"].dropna().astype(str).str.strip().unique().tolist()
                       if not macchine:
                           st.info("Nessuna macchina trovata.")
                       else:
                           for mac in macchine:
                               st.markdown(f"#### 🏭 {mac}")

                               sub = dfn[dfn["Macchina"] == mac]
                               sub_lav  = sub[mask_lav  & (sub["Macchina"] == mac)]
                               sub_prog = sub[mask_prog & (sub["Macchina"] == mac)]

                               draw_gantt(sub_lav,  f"Gantt — {mac} — In lavorazione", "MATERIALE")
                               draw_gantt(sub_prog, f"Gantt — {mac} — Programmato",   "MATERIALE")

                               st.markdown("---")


            if mostra_gantt_macchine:
                st.markdown("### 📆 Gantt per Macchina")
                for mac in df_out["Macchina"].dropna().unique().tolist():
                    sub = df_out[df_out["Macchina"] == mac].copy()
                    if sub.empty: 
                        continue
                    sub = sub.dropna(subset=["Data inizio","Data fine"])
                    if sub.empty:
                        continue
                    sub["LabelAvanzamento"] = sub["Completamento"].fillna(0).astype(int).astype(str) + "%"
                    order = pd.Index(
                        sub.sort_values(["Data inizio","Data fine","MATERIALE","Attività"])["MATERIALE"]
                    ).unique().tolist()

                    # hardening: scegli color disponibile
                    color_col = pick_color_column(sub, ("Operatore","Centro di lavoro","Attività","Macchina","Stato attività"))
                    kwargs = {}
                    if color_col:
                        kwargs["color"] = color_col

                    fig = px.timeline(
                        sub,
                        x_start="Data inizio",
                        x_end="Data fine",
                        y="MATERIALE",
                        text="LabelAvanzamento",
                        **kwargs
                    )
                    fig.update_traces(textposition="inside", insidetextanchor="middle")
                    fig.update_yaxes(categoryorder="array", categoryarray=order, autorange="reversed")
                    fig.update_layout(title=f"Gantt - {mac}", xaxis_title="Data", yaxis_title="Materiale")
                    st.plotly_chart(fig, use_container_width=True)

            # Salvataggio
            if st.button("💾 Salva pianificazione"):
                os.makedirs("dati_pianificati", exist_ok=True)
                if unico_csv:
                    nome = "dati_pianificati/pianificazione_consolidata.csv"
                    df_save = df_out.copy()
                    for c in ["Data inizio","Data fine"]:
                        df_save[c] = pd.to_datetime(df_save[c], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
                    df_save.to_csv(nome, index=False, encoding="utf-8-sig")
                    st.success(f"Pianificazione salvata in {nome}")
                else:
                    mac_list = df_out["Macchina"].dropna().unique().tolist()
                    for mac in mac_list:
                        sub = df_out[df_out["Macchina"] == mac].copy()
                        for c in ["Data inizio","Data fine"]:
                            sub[c] = pd.to_datetime(sub[c], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
                        safe_mac = mac.lower().replace(" ", "_")
                        nome = f"dati_pianificati/pianificazione_{safe_mac}.csv"
                        sub.to_csv(nome, index=False, encoding="utf-8-sig")
                    st.success("Pianificazioni salvate per macchina in cartella dati_pianificati/")
        else:
            st.warning("Nessuna pianificazione generata: imposta almeno una riga su 'In lavorazione' o 'Programmazione'.")
