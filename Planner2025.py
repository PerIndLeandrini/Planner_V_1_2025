import streamlit as st
import pandas as pd
import os
from datetime import datetime
import plotly.express as px
import numpy as np

# --- Config ---
st.set_page_config(page_title="Pianificazione Produzione", layout="wide")
st.title("📦 Gestione Pianificazione Produzione")

# Stati attività (unica fonte di verità)
STATI_ATTIVITA = ["Programmato", "In Produzione", "Controllo qualità", "Da definire", "Completato"]

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
        s = f"{val:,.2f}"          # 1,234.56
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
            centri = ["Hurco", "Mazak 5assi", "Mazak 4assi", "Mazak HCN", "Mazak 3assi", "Hyandai", "Macchine //", "SALA METROLOGICA", "SEGA"]
            fornitori = {
                "Trattamento": ["Mochem", "Samacromo", "Art Ing", "Bugli", "Aviorubber"],
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
                if completamento == 100:
                    stato = "Completato"

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
                for attivita in pianificazione:
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
                        **attivita
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
                    # etichetta avanzamento
                    df_gantt["LabelAvanzamento"] = df_gantt["Completamento"].fillna(0).astype(int).astype(str) + "%"

                    # ordinamento per data d'inizio (più vecchia in alto)
                    order = pd.Index(
                        df_gantt.sort_values(["Inizio", "Fine", "Attività"])["Attività"]
                    ).unique().tolist()

                    fig = px.timeline(
                        df_gantt,
                        x_start="Inizio",
                        x_end="Fine",
                        y="Attività",
                        color="Operatore",
                        text="LabelAvanzamento"
                    )
                    fig.update_traces(textposition="inside", insidetextanchor="middle")

                    # applica l'ordine al verticale (y)
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
            # mantieni sia stringa formattata per tabella, sia colonne datetime per gantt
            dt = pd.to_datetime(df_pianif[col], dayfirst=True, errors="coerce")
            df_pianif[col] = dt.dt.strftime("%d/%m/%Y").fillna("")

    if "Valore" in df_pianif.columns and pd.api.types.is_numeric_dtype(df_pianif["Valore"]):
        df_pianif["Valore (visuale)"] = df_pianif["Valore"].apply(fmt_eur)

    st.subheader("📋 Pianificazione caricata (tutte le righe)")
    st.dataframe(df_pianif, use_container_width=True, height=380)

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

            # ordinamento per data d'inizio (più vecchia in alto)
            order = pd.Index(
                gantt.sort_values(["Inizio", "Fine", "Attività"])["Attività"]
            ).unique().tolist()

            fig = px.timeline(
                gantt,
                x_start="Inizio",
                x_end="Fine",
                y="Attività",
                color="Operatore",
                text="LabelAvanzamento"
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

    # Maschera di modifica identica (per singola attività)
    st.markdown("### ✏️ Modifica attività selezionata")
    if not df_mat.empty:
        # Select attività (indice reale per poter salvare sulla riga corretta)
        idx = st.selectbox(
            "Seleziona attività da modificare",
            df_mat.index,
            format_func=lambda i: f"{df_mat.loc[i, 'Attività']} | {df_mat.loc[i, 'MATERIALE']}" if "Attività" in df_mat.columns else str(i)
        )

        row = df_pianif.loc[idx]

        # Pre-parsing date
        di = pd.to_datetime(row.get("Data inizio",""), dayfirst=True, errors="coerce")
        dfine = pd.to_datetime(row.get("Data fine",""), dayfirst=True, errors="coerce")
        if pd.isna(di): di = datetime.today()
        if pd.isna(dfine): dfine = datetime.today()

        new_di = st.date_input("Nuova data inizio", di)
        new_df = st.date_input("Nuova data fine", dfine)

        # Stato attività con lista unica
        stato_opts = STATI_ATTIVITA
        idx_stato = stato_opts.index(row["Stato attività"]) if row.get("Stato attività") in stato_opts else 0
        new_stato = st.selectbox("Stato attività", stato_opts, index=idx_stato)

        # Completamento con auto-link allo stato
        new_comp = st.slider("Completamento (%)", 0, 100, int(row.get("Completamento",0)), step=5)
        if new_comp == 100:
            new_stato = "Completato"
            st.caption("🔁 Stato impostato automaticamente a **Completato** (100%).")

        if st.button("💾 Salva modifica"):
            os.makedirs("dati_pianificati", exist_ok=True)
            df_pianif.at[idx, "Data inizio"] = new_di.strftime("%d/%m/%Y")
            df_pianif.at[idx, "Data fine"] = new_df.strftime("%d/%m/%Y")
            df_pianif.at[idx, "Stato attività"] = new_stato
            df_pianif.at[idx, "Completamento"] = new_comp

            nome_file_csv = csv_file.name
            out_path = f"dati_pianificati/{nome_file_csv}"
            df_pianif.to_csv(out_path, index=False, encoding="utf-8-sig")
            st.success(f"Aggiornamento salvato in {out_path}")

            # Refresh preview
            st.dataframe(df_pianif, use_container_width=True, height=300)
