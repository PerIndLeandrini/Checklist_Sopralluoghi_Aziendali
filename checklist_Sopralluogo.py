# checklist_Sopralluogo.py
import streamlit as st
import pandas as pd
from PIL import Image
from io import BytesIO
from datetime import date
from typing import Optional
from calendar import monthrange

# ReportLab
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    BaseDocTemplate, PageTemplate, Frame, NextPageTemplate,
    Paragraph, Spacer, Table, TableStyle,
    PageBreak, Image as RLImage
)
from reportlab.lib.utils import ImageReader

# ---------- Utility ----------
def fmt_date(d: Optional[date]) -> str:
    return d.strftime("%d/%m/%Y") if d else ""

def filename_date(d: Optional[date]) -> str:
    return fmt_date(d).replace("/", "-") if d else "data"

def date_input_eu(label: str, key: str, value: Optional[date] = None, allow_empty: bool = False,
                  min_year: int = 2000, max_year: int = 2100) -> Optional[date]:
    """Date picker europeo (GG/MM/AAAA) con 3 selectbox. Se allow_empty=True, può restare vuoto."""
    if allow_empty:
        enabled = st.checkbox(f"{label} — inserisci data", value=bool(value), key=f"{key}_enable")
        if not enabled:
            st.caption(f"{label}: —")
            return None
    else:
        st.caption(label)

    _val = value or date.today()
    years = list(range(min_year, max_year + 1))

    c1, c2, c3 = st.columns([1.3, 1.3, 2.2])
    with c1:
        gg = st.selectbox("GG", list(range(1, 32)), index=_val.day - 1, key=f"{key}_d")
    with c2:
        mm = st.selectbox("MM", list(range(1, 13)), index=_val.month - 1, key=f"{key}_m")
    with c3:
        try:
            y_index = years.index(_val.year)
        except ValueError:
            y_index = years.index(date.today().year)
        aa = st.selectbox("AAAA", years, index=y_index, key=f"{key}_y")

    gg_max = monthrange(aa, mm)[1]
    if gg > gg_max:
        gg = gg_max

    return date(aa, mm, gg)

st.set_page_config(page_title="Audit Fornitore — D.Lgs. 81/08", page_icon="✅", layout="wide")

# ---------- CSS ----------
st.markdown("""
<style>
:root { --brand:#0f766e; --muted:#6b7280; --soft:#e5e7eb; --danger:#b91c1c; --danger-bg:#fee2e2; }
.block-container { padding-top: 1rem; }
h1,h2,h3 { letter-spacing: .2px; }
div[role="radiogroup"] > label {
  padding: 6px 10px; border: 1px solid var(--soft); border-radius: 8px; margin-right: 6px; margin-bottom: 6px;
}
[data-testid="stMetricValue"] { color: var(--brand); }
.badge-nc { background:#b91c1c; color:#fff; padding:4px 8px; border-radius:999px; font-size:12px; font-weight:700; }
.ref { color:#666; font-size:12px; }
/* Sidebar: allarga i select */
[data-testid="stSidebar"] div[data-baseweb="select"] { min-width: 88px; }
[data-testid="stSidebar"] .stSelectbox label { font-size: 12px; }
/* Brand in sidebar */
[data-testid="stSidebar"] .brand-box {
  position: sticky; top: 0;
  background: inherit;
  padding: 12px 8px 10px 8px;
  margin: -8px -8px 8px -8px; /* allineati alla sidebar */
  border-bottom: 1px solid var(--soft);
  text-align: center;
  z-index: 100;
}
[data-testid="stSidebar"] .brand-box small {
  color: var(--muted);
}
</style>
""", unsafe_allow_html=True)

st.title("Checklist Audit Fornitore — D.Lgs. 81/08 & SMEI")



# ---------- Sidebar ----------
with st.sidebar:
    # --- BRANDING SIDEBAR (solo UI, non PDF) ---
    BRAND_LOCK = True  # lascia True: così forzi il tuo logo per tutti
    # se vuoi sbloccarlo per te su Streamlit Cloud: metti BRAND_LOCK = st.secrets.get("BRAND_LOCK", True)

    brand_img = None
    try:
        brand_img = Image.open("assets/sidebar_brand.png")
    except Exception:
        pass

    st.markdown('<div class="brand-box">', unsafe_allow_html=True)
    if brand_img is not None:
        st.image(brand_img, use_container_width=True)
    else:
        st.write("**Audit Tool**")
    st.write("<small>© 2025 Simone Leandrini</small>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)
    st.header("Dati Audit")
    fornitore = st.text_input("Fornitore")
    data_audit = date_input_eu("Data audit", key="data_audit", value=date.today(), allow_empty=False)
    auditor = st.text_input("Auditor")

    st.divider()
    st.caption("Logo per il PDF (opzionale)")
    logo_up = st.file_uploader("Carica logo (JPG/PNG)", type=["jpg","jpeg","png"], key="logo_upl")

    st.divider()
    st.caption("Filtri (solo vista)")
    if st.button("🔄 Reset filtri"):
        st.session_state["filtro_nc"] = False
        st.session_state["filtro_testo"] = ""
    filtro_nc = st.checkbox("Mostra solo Non conformi", value=st.session_state.get("filtro_nc", False), key="filtro_nc")
    filtro_testo = st.text_input("Cerca (testo in requisito/note)", value=st.session_state.get("filtro_testo", ""), key="filtro_testo")

st.write(f"**Fornitore:** {fornitore or '—'}  |  **Data:** {fmt_date(data_audit)}  |  **Auditor:** {auditor or '—'}")

# ---------- Catalogo requisiti ----------
CATALOGO = {
    "Documentazione e organizzazione": [
        ("DVR presente, firmato, data ≤ 12 mesi o aggiornato a variazioni", "Art. 17, 28-29 D.Lgs. 81/08"),
        ("Nomina RSPP disponibile e coerente con macrosettore ATECO", "Art. 17, 31-33 D.Lgs. 81/08"),
        ("Nomina ASPP (se presente) e evidenze formazione modulo A-B-C", "Art. 32 Acc. Stato-Regioni 2011"),
        ("Designazione Addetti Primo Soccorso (elenchi e turnazioni)", "Art. 18, 45 D.Lgs. 81/08; DM 388/03"),
        ("Designazione Addetti Antincendio e registro prove", "Art. 18 D.Lgs. 81/08; DM 02/09/21"),
        ("Informazione e consultazione RLS/RLST documentata", "Art. 47-50 D.Lgs. 81/08"),
        ("Gestione appalti: idoneità tecnico-professionale fornitori", "Art. 26 D.Lgs. 81/08"),
        ("D.U.V.R.I. emesso ove necessario (rischi interferenziali)", "Art. 26 c.3 D.Lgs. 81/08"),
        ("Piano di emergenza con planimetrie aggiornate/esposte", "Art. 43-46 D.Lgs. 81/08; DM 02/09/21"),
    ],
    "Impianti elettrici e verifiche": [
        ("Verbali verifica messa a terra e differenziali (DPR 462/01)", "DPR 462/01; Art. 86 D.Lgs. 81/08"),
        ("Quadri elettrici: chiusura, targhette, schemi, IP adeguato", "CEI 64-8; Art. 80-87 D.Lgs. 81/08"),
        ("Prese e cavi: integrità, assenza di giunte volanti, protezioni", "CEI 64-8; Art. 80-87 D.Lgs. 81/08"),
        ("Illuminazione di emergenza funzionante e manutenzionata", "UNI EN 1838; Art. 63-64 D.Lgs. 81/08"),
    ],
    "Macchine e attrezzature": [
        ("Marcatura CE e Dichiarazione CE conformità disponibili", "Dir. 2006/42/CE; Art. 70-71 D.Lgs. 81/08"),
        ("Manuale d’uso/manutenzione disponibile in lingua italiana", "Dir. 2006/42/CE; All. V D.Lgs. 81/08"),
        ("Ripari fissi/mobili e microinterruttori funzionanti", "Allegato V e VI D.Lgs. 81/08"),
        ("Dispositivi arresto di emergenza accessibili e testati", "EN ISO 13850; Art. 71 D.Lgs. 81/08"),
        ("Check-list manutenzione periodica e registri compilati", "Art. 71 c.8-9 D.Lgs. 81/08"),
        ("Verifiche periodiche attrezzature (carrelli, sollev.)", "Art. 71 c.11; DM 11/04/2011"),
    ],
    "Attrezzature in pressione / gas": [
        ("PED/recipienti in pressione: collaudi/verifiche in validità", "D.Lgs. 81/08 Art. 71; PED 2014/68/UE"),
        ("Bombole gas: fissaggio, cappellotti, etichette, stoccaggio", "Linee guida INAIL; CLP"),
    ],
    "Sostanze chimiche / CLP / REACH": [
        ("Inventario sostanze aggiornato con codici e quantità", "Titolo IX Capo I D.Lgs. 81/08; REACH"),
        ("Schede di sicurezza (SDS) 16 sezioni, ≤ 5 anni, in ITA", "REACH; Art. 223 D.Lgs. 81/08"),
        ("Etichettatura CLP corretta su imballaggi e contenitori secondari", "Reg. CLP (CE) n.1272/2008"),
        ("Stoccaggio per compatibilità e bacini di contenimento", "Titolo IX D.Lgs. 81/08"),
        ("Procedure sversamenti e kit assorbenti disponibili", "Titolo IX D.Lgs. 81/08"),
        ("Valutazioni specifiche: cancerogeni/mutageni dove presenti", "Titolo IX Capo II D.Lgs. 81/08"),
    ],
    "Movimentazione merci / carrelli": [
        ("Abilitazione carrellisti (Accordo CSR 2012) in corso di validità", "Art. 73 D.Lgs. 81/08; CSR 22/02/2012"),
        ("Check giornaliero carrelli (freni, forche, luci, allarmi)", "Buone pratiche; Art. 71 D.Lgs. 81/08"),
        ("Viabilità interna segnalata (corsie, limiti, specchi)", "Allegato IV D.Lgs. 81/08"),
        ("Zone carico/scarico: protezioni bordi, STOP, paraurti", "Allegato IV D.Lgs. 81/08"),
        ("Mezzi di sollevamento: brache/ganci con certificazione e stato", "Art. 71 D.Lgs. 81/08"),
    ],
    "Ambienti di lavoro / antincendio": [
        ("Estintori adeguati e manutenzione UNI 9994-1 aggiornata", "DM 02/09/21; UNI 9994-1"),
        ("Idranti/naspi UNI 10779: ispezioni e prova pressione", "UNI 10779; DM 02/09/21"),
        ("Uscite di emergenza libere e segnaletica UNI EN ISO 7010", "Allegato IV D.Lgs. 81/08"),
        ("Ordine e pulizia (5S) nelle aree operative e stoccaggi", "Art. 64 D.Lgs. 81/08"),
        ("Rumore: valutazione e misure (cuffie disponibili dove ≥85 dB)", "Titolo VIII Capo II D.Lgs. 81/08"),
        ("Vibrazioni: valutazione e controllo esposizioni", "Titolo VIII Capo III D.Lgs. 81/08"),
    ],
    "Sorveglianza sanitaria": [
        ("Nomina Medico Competente ove dovuta", "Art. 18, 25, 41 D.Lgs. 81/08"),
        ("Protocolli sanitari coerenti con rischi valutati", "Art. 25, 41 D.Lgs. 81/08"),
        ("Giudizi di idoneità disponibili e comunicati ai preposti", "Art. 41 D.Lgs. 81/08"),
        ("Gestione idoneità con prescrizioni e follow-up attivi", "Art. 18, 41 D.Lgs. 81/08"),
    ],
    "DPI (scelta, consegna, uso)": [
        ("Valutazione scelta DPI per ciascun rischio", "Art. 76-77 D.Lgs. 81/08"),
        ("Registro consegna DPI firmato dai lavoratori", "Art. 77 D.Lgs. 81/08"),
        ("Addestramento DPI di III categoria documentato", "Art. 77 c.5 D.Lgs. 81/08"),
        ("Sostituzione DPI usurati e stoccaggio corretto", "Art. 77 D.Lgs. 81/08"),
    ],
    "Aspetti formativi (dettaglio)": [
        ("Formazione generale lavoratori (≥4h) erogata a tutti i neoassunti", "Accordi Stato-Regioni 2011/2016"),
        ("Formazione specifica (4/8/12h secondo rischio) completata", "Accordi SR 2011/2016"),
        ("Aggiornamento lavoratori (≥6h/5 anni) tracciato", "Accordi SR 2011/2016"),
        ("Formazione Preposti (moduli aggiuntivi) + aggiornamento", "Accordo SR 2021 (Preposti)"),
        ("Formazione Dirigenti (≥16h) + aggiornamento quinquennale", "Accordi SR 2011/2016"),
        ("Antincendio (Liv. 1-2-3) con addestramento pratico", "DM 02/09/21"),
        ("Primo soccorso (A/B/C) + aggiornamento triennale", "DM 388/03"),
        ("VDT (ove previsto): informazione/formazione addetti", "Art. 173-176 D.Lgs. 81/08"),
        ("Attrezzature particolari (carrelli, PLE, gru): abilitazioni", "Accordo CSR 22/02/2012"),
        ("RSPP/ASPP: moduli A-B-C e aggiornamenti quinquennali", "Art. 32 D.Lgs. 81/08; Accordi SR"),
    ],
}

RESPONSABILI = ["Datore di Lavoro", "Dirigente", "Preposto", "RSPP", "Medico Competente", "Addetto Sicurezza", "Altro"]

def render_requisito(idx, sezione, requisito, riferimento):
    with st.container(border=True):
        applicabile = st.toggle("Applicabile?", value=False, key=f"{sezione}_{idx}_appl")

        ctitle, cbadge = st.columns([10,1])
        with ctitle:
            st.markdown(f"**{idx}. {requisito}**", unsafe_allow_html=True)
            st.markdown(f"<span class='ref'>Rif.: {riferimento}</span>", unsafe_allow_html=True)

        stato = "Non applicabile"
        livello = cause = trattamento = periodo = responsabile = ""
        data_tratt = data_verifica = None
        note = ""
        files = []

        if applicabile:
            stato = st.radio("Stato", ["Conforme", "Non conforme"], key=f"{sezione}_{idx}_st", horizontal=True)
            with cbadge:
                if stato == "Non conforme":
                    st.markdown("<div style='text-align:right;'><span class='badge-nc'>NC</span></div>", unsafe_allow_html=True)
        else:
            with cbadge:
                st.write("")

        if applicabile and stato == "Non conforme":
            livello = st.selectbox("Classificazione NC", ["Livello 1", "Livello 2"], key=f"{sezione}_{idx}_lvl")
            st.markdown("**Gestione Non Conformità**")
            cause = st.text_area("Analisi delle cause", key=f"{sezione}_{idx}_cause",
                                 placeholder="Es. mancata pianificazione aggiornamenti; assenza procedura…")
            c1, c2 = st.columns(2)
            with c1:
                trattamento = st.text_area("Trattamento / Azione correttiva", key=f"{sezione}_{idx}_tratt",
                                           placeholder="Descrivi l'azione correttiva o migliorativa da attuare")
                periodo = st.selectbox("Periodo di gestione",
                                       ["BREVE (≤ 1 mese)", "MEDIO (≤ 6 mesi)", "LUNGO (≤ 12 mesi)"],
                                       key=f"{sezione}_{idx}_periodo")
            with c2:
                data_tratt = date_input_eu("Data prevista completamento trattamento",
                                           key=f"{sezione}_{idx}_dtratt", value=None, allow_empty=True)
                data_verifica = date_input_eu("Verifica prevista per",
                                              key=f"{sezione}_{idx}_dver", value=None, allow_empty=True)

            resp = st.selectbox("Responsabile", RESPONSABILI, key=f"{sezione}_{idx}_resp")
            responsabile = st.text_input("Specifica altro responsabile", key=f"{sezione}_{idx}_resp_alt") if resp == "Altro" else resp

            note = st.text_area("Note / Evidenze testuali", key=f"{sezione}_{idx}_note",
                                placeholder="Annota evidenze, riferimenti documentali, ubicazione file…")
            files = st.file_uploader("Allega foto/documenti (jpg, png, pdf)", accept_multiple_files=True,
                                     type=["jpg","jpeg","png","pdf"], key=f"{sezione}_{idx}_files")

        elif applicabile and stato == "Conforme":
            note = st.text_area("Note / Evidenze testuali", key=f"{sezione}_{idx}_note",
                                placeholder="Annota evidenze, riferimenti documentali, ubicazione file…")
            files = st.file_uploader("Allega foto/documenti (jpg, png, pdf)", accept_multiple_files=True,
                                     type=["jpg","jpeg","png","pdf"], key=f"{sezione}_{idx}_files")

        allegati = [f.name for f in files] if files else []

        # punteggi
        score_simple = 1 if (applicabile and stato == "Conforme") else (0 if (applicabile and stato == "Non conforme") else None)
        score_weight = (1.0 if (applicabile and stato == "Conforme") else (0.5 if livello == "Livello 1" and applicabile else (0.0 if applicabile and stato == "Non conforme" else None)))

        return {
            "Applicabile": "Sì" if applicabile else "No",
            "N": idx,
            "Sezione": sezione,
            "Requisito": requisito,
            "Riferimento": riferimento,
            "Stato": stato,
            "NC Livello": livello,
            "Note": note,
            "Allegati": ", ".join(allegati),
            "Punteggio": score_simple,
            "Punteggio ponderato": score_weight,
            "Cause": cause,
            "Trattamento": trattamento,
            "Periodo": periodo,
            "Data trattamento": fmt_date(data_tratt),
            "Data verifica": fmt_date(data_verifica),
            "Responsabile": responsabile,
            "_files": files
        }

# ---------- Raccolta risultati ----------
records_all, records_visible = [], []
for sezione, requisiti in CATALOGO.items():
    st.header(sezione)
    for i, (req, rif) in enumerate(requisiti, start=1):
        rec = render_requisito(i, sezione, req, rif)
        records_all.append(rec)

        vis_ok = True
        if st.session_state.get("filtro_nc") and rec["Stato"] != "Non conforme":
            vis_ok = False
        if st.session_state.get("filtro_testo"):
            text_blob = f"{rec['Requisito']} {rec['Note']} {rec['Sezione']}".lower()
            if st.session_state["filtro_testo"].lower() not in text_blob:
                vis_ok = False
        if vis_ok:
            records_visible.append(rec)

st.divider()

# ---------- Statistiche ----------
st.subheader("Valutazione complessiva")
df_all = pd.DataFrame(records_all)
df_vis = pd.DataFrame(records_visible)

def percentuali(df_section, col="Punteggio"):
    if df_section.empty or col not in df_section: return (None, 0, 0, 0)
    validi = df_section[col].dropna()
    if len(validi) == 0: return (None, 0, 0, 0)
    if col == "Punteggio":
        conformi = (validi == 1).sum()
        nonconf = (validi == 0).sum()
        perc = round((conformi / len(validi)) * 100, 1)
        return perc, conformi, nonconf, len(validi)
    else:
        media = round(validi.mean() * 100, 1)
        conformi = (df_section["Stato"] == "Conforme").sum()
        nonconf = (df_section["Stato"] == "Non conforme").sum()
        return media, conformi, nonconf, len(validi)

def stats_per_sezione(df, col="Punteggio"):
    if df.empty: return pd.DataFrame()
    out = []
    for sezione, dfg in df.groupby("Sezione"):
        perc, conf, nc, tot = percentuali(dfg, col=col)
        out.append({
            "Sezione": sezione,
            ("% Conformità" if col=="Punteggio" else "% Conformità ponderata"): perc if perc is not None else "—",
            "Conformi": conf, "Non conformi": nc, "Requisiti valutati": tot
        })
    sdf = pd.DataFrame(out)
    metric_col = "% Conformità" if col=="Punteggio" else "% Conformità ponderata"
    sdf["_ord"] = pd.to_numeric(sdf[metric_col], errors="coerce")
    sdf = sdf.sort_values("_ord", ascending=False, na_position="last").drop(columns=["_ord"])
    return sdf

if not df_all.empty:
    colL, colR = st.columns([2,1])
    with colL:
        st.markdown("**Per sezione (semplice)**")
        st.dataframe(stats_per_sezione(df_all, col="Punteggio"), use_container_width=True)
        st.markdown("**Per sezione (ponderata L1=0.5 / L2=0)**")
        st.dataframe(stats_per_sezione(df_all, col="Punteggio ponderato"), use_container_width=True)
    with colR:
        perc_tot_s, conf_tot_s, nc_tot_s, tot_tot_s = percentuali(df_all, col="Punteggio")
        st.metric("Conformità totale (semplice)", f"{perc_tot_s if perc_tot_s is not None else '—'}%")
        perc_tot_w, _, _, _ = percentuali(df_all, col="Punteggio ponderato")
        st.metric("Conformità totale (ponderata)", f"{perc_tot_w if perc_tot_w is not None else '—'}%")
        st.write(f"**Conformi:** {conf_tot_s}  \n**Non conformi:** {nc_tot_s}  \n**Requisiti valutati:** {tot_tot_s}")
else:
    st.info("Compila almeno un requisito per vedere le statistiche.")

st.divider()

# ---------- Riepilogo & Export (sul visibile) ----------
st.subheader("Riepilogo requisiti (filtri applicati)")
if not df_vis.empty:
    st.dataframe(df_vis.drop(columns=["_files"]), use_container_width=True)
    data_str = filename_date(data_audit)
    st.download_button(
        "⬇️ Scarica CSV",
        data=df_vis.drop(columns=["_files"]).to_csv(index=False).encode("utf-8"),
        file_name=f"audit_{(fornitore or 'fornitore')}_{data_str}.csv",
        mime="text/csv"
    )
    excel_buffer = BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
        df_vis.drop(columns=["_files"]).to_excel(writer, index=False, sheet_name="Audit")
    st.download_button(
        "⬇️ Scarica Excel",
        data=excel_buffer.getvalue(),
        file_name=f"audit_{(fornitore or 'fornitore')}_{data_str}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Nessun dato da esportare (verifica i filtri).")

# ---------- PDF ----------
st.subheader("Report PDF")
st.caption("Sezioni in verticale; riepilogo di tutte le Non Conformità in orizzontale; immagini ridimensionate.")

def build_pdf(df_all: pd.DataFrame, logo_file) -> bytes:
    buf = BytesIO()

    # Margini
    lm, rm, tm, bm = 2*cm, 2*cm, 1.5*cm, 1.5*cm
    doc = BaseDocTemplate(buf, pagesize=A4,
                          leftMargin=lm, rightMargin=rm, topMargin=tm, bottomMargin=bm)

    # Frame Portrait
    frame_p = Frame(lm, bm, doc.width, doc.height, id='portrait_frame')
    # Frame LANDSCAPE con margini ridotti a 0.2 cm (≈ 2 mm)
    W, H = landscape(A4)
    lm_l = rm_l = tm_l = bm_l = 0.2*cm  # margini ridotti
    frame_l = Frame(lm_l, bm_l, W - lm_l - rm_l, H - tm_l - bm_l, id='landscape_frame')

    pt_portrait  = PageTemplate(id='Portrait',  frames=[frame_p], pagesize=A4)
    pt_landscape = PageTemplate(id='Landscape', frames=[frame_l], pagesize=landscape(A4))
    doc.addPageTemplates([pt_portrait, pt_landscape])

    story = []

    styles = getSampleStyleSheet()
    H1 = ParagraphStyle("H1", parent=styles["Heading1"], fontSize=18, spaceAfter=8, leading=22, textColor=colors.HexColor("#0f172a"))
    H2 = ParagraphStyle("H2", parent=styles["Heading2"], fontSize=14, spaceBefore=8, spaceAfter=4, textColor=colors.HexColor("#111827"))
    P  = ParagraphStyle("P",  parent=styles["BodyText"], fontSize=10, leading=13)
    Small = ParagraphStyle("Small", parent=styles["BodyText"], fontSize=9, textColor=colors.grey)

    # helper Paragraph con wrap
    def par(txt, style):
        if txt is None: txt = ""
        txt = str(txt).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        return Paragraph(txt, style)
    
    Cell      = ParagraphStyle("Cell",      parent=P, fontSize=8.55, leading=11,  wordWrap='LTR')
    CellSmall = ParagraphStyle("CellSmall", parent=P, fontSize=8.2, leading=10.5, wordWrap='LTR')
    HeaderCell= ParagraphStyle("HeaderCell",parent=styles["Heading4"], fontSize=8.0, leading=11)


    # Logo
    if logo_file is not None:
        try:
            logo_file.seek(0)
            story.append(RLImage(logo_file, width=3*cm))
            story.append(Spacer(1, 6))
        except Exception:
            pass

    # Intestazione
    story.append(Paragraph("Report Audit Fornitore — D.Lgs. 81/08 & SMEI", H1))
    story.append(Paragraph(
        f"Fornitore: <b>{fornitore or '—'}</b> &nbsp;&nbsp; Data: <b>{fmt_date(data_audit)}</b> &nbsp;&nbsp; Auditor: <b>{auditor or '—'}</b>",
        P
    ))
    story.append(Spacer(1, 8))

    # Sintesi
    def percentuali_local(df_section, col="Punteggio"):
        if df_section.empty or col not in df_section: return (None, 0, 0, 0)
        validi = df_section[col].dropna()
        if len(validi) == 0: return (None, 0, 0, 0)
        if col == "Punteggio":
            conformi = (validi == 1).sum()
            nonconf = (validi == 0).sum()
            perc = round((conformi / len(validi)) * 100, 1)
            return perc, conformi, nonconf, len(validi)
        else:
            media = round(validi.mean() * 100, 1)
            conformi = (df_section["Stato"] == "Conforme").sum()
            nonconf = (df_section["Stato"] == "Non conforme").sum()
            return media, conformi, nonconf, len(validi)

    perc_tot_s, conf_tot_s, nc_tot_s, tot_tot_s = percentuali_local(df_all, col="Punteggio")
    perc_tot_w, _, _, _ = percentuali_local(df_all, col="Punteggio ponderato")
    sintesi_data = [
        ["Conformità totale (semplice)", f"{perc_tot_s if perc_tot_s is not None else '—'}%"],
        ["Conformità totale (ponderata)", f"{perc_tot_w if perc_tot_w is not None else '—'}%"],
        ["Requisiti conformi", str(conf_tot_s)],
        ["Requisiti non conformi", str(nc_tot_s)],
        ["Requisiti valutati", str(tot_tot_s)],
    ]
    t = Table(sintesi_data, colWidths=[7*cm, 7*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#f3f4f6")),
        ("BOX", (0,0), (-1,-1), 0.5, colors.grey),
        ("INNERGRID", (0,0), (-1,-1), 0.25, colors.grey),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold")
    ]))
    story.append(t)
    story.append(Spacer(1, 12))

    # Tabelle per sezione (portrait)
    for sezione, dfg in df_all.groupby("Sezione"):
        dfg = dfg.sort_values("N")
        story.append(Paragraph(sezione, H2))

        rows = [[par("#", HeaderCell), par("Requisito", HeaderCell), par("Stato", HeaderCell),
                 par("Riferimento", HeaderCell), par("Note", HeaderCell)]]
        for _, r in dfg.iterrows():
            rows.append([
                par(str(r.get("N","")), Cell),
                par(r["Requisito"], Cell),
                par(r["Stato"], Cell),
                par(r["Riferimento"], CellSmall),
                par(r.get("Note",""), CellSmall)
            ])

        tbl = Table(rows, colWidths=[1.0*cm, 6.2*cm, 2.4*cm, 4.1*cm, 3.3*cm], repeatRows=1, splitByRow=1)
        tbl_style = TableStyle([
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#e5e7eb")),
            ("BOX", (0,0), (-1,-1), 0.5, colors.grey),
            ("INNERGRID", (0,0), (-1,-1), 0.25, colors.grey),
            ("VALIGN", (0,0), (-1,-1), "TOP"),
            ("LEFTPADDING", (0,0), (-1,-1), 4),
            ("RIGHTPADDING", (0,0), (-1,-1), 4),
            ("TOPPADDING", (0,0), (-1,-1), 3),
            ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ])
        # righe NC in rosso
        for ridx in range(1, len(rows)):
            if dfg.iloc[ridx-1]["Stato"] == "Non conforme":
                tbl_style.add("TEXTCOLOR", (0, ridx), (-1, ridx), colors.HexColor("#b91c1c"))
        tbl.setStyle(tbl_style)
        story.append(tbl)
        story.append(Spacer(1, 10))

    # ---- Appendice NC in Landscape (tutte insieme) ----
    df_nc_all = df_all[df_all["Stato"] == "Non conforme"].copy()
    if not df_nc_all.empty:
        story.append(NextPageTemplate('Landscape'))
        story.append(PageBreak())

        story.append(Paragraph("Non Conformità — Riepilogo complessivo", H1))
        story.append(Paragraph("Elenco sintetico di tutte le NC rilevate, con campi principali.", Small))
        story.append(Spacer(1, 6))

        df_nc_all = df_nc_all.sort_values(["Sezione", "N"], kind="stable")
        nc_rows = [[
            par("#", HeaderCell), par("Sezione", HeaderCell), par("Requisito", HeaderCell),
            par("Livello", HeaderCell), par("Cause", HeaderCell), par("Trattamento", HeaderCell),
            par("Periodo", HeaderCell), par("Data tratt.", HeaderCell), par("Verifica", HeaderCell),
            par("Responsabile", HeaderCell)
        ]]
        for _, r in df_nc_all.iterrows():
            nc_rows.append([
                par(str(r.get("N","")), Cell),
                par(r.get("Sezione",""), CellSmall),
                par(r.get("Requisito",""), Cell),
                par(r.get("NC Livello",""), Cell),
                par(r.get("Cause",""), CellSmall),
                par(r.get("Trattamento",""), CellSmall),
                par(r.get("Periodo",""), Cell),
                par(r.get("Data trattamento",""), Cell),
                par(r.get("Data verifica",""), Cell),
                par(r.get("Responsabile",""), CellSmall),
            ])

        nc_tbl = Table(
            nc_rows,
            colWidths=[
                0.9*cm,  # #
                3.2*cm,  # Sezione
                8.7*cm,  # Requisito
                1.4*cm,  # Livello
                4.3*cm,  # Cause
                4.0*cm,  # Trattamento
                1.4*cm,  # Periodo
                1.7*cm,  # Data tratt.
                1.7*cm,  # Verifica
                2.0*cm   # Responsabile
            ],
            repeatRows=1, splitByRow=1
        )
        
        nc_tbl.setStyle(TableStyle([
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#fee2e2")),
            ("BOX", (0,0), (-1,-1), 0.5, colors.grey),
            ("INNERGRID", (0,0), (-1,-1), 0.25, colors.grey),
            ("VALIGN", (0,0), (-1,-1), "TOP"),
            ("LEFTPADDING",(0,0),(-1,-1),2),
            ("RIGHTPADDING",(0,0),(-1,-1),2),
            ("TOPPADDING",(0,0),(-1,-1),2),
            ("BOTTOMPADDING",(0,0),(-1,-1),2),
            ("TEXTCOLOR", (0,1), (-1,-1), colors.HexColor("#7f1d1d")),
        ]))
        story.append(nc_tbl)
        story.append(Spacer(1, 12))

        # torna al portrait per le sezioni successive (firme/foto)
        story.append(NextPageTemplate('Portrait'))
        story.append(PageBreak())

    # Firme
    story.append(Paragraph("Firme", H2))
    firm_tbl = Table([
        ["Auditor", "Rappresentante Fornitore"],
        ["\n\n__________________________", "\n\n__________________________"]
    ], colWidths=[8*cm, 8*cm])
    firm_tbl.setStyle(TableStyle([("ALIGN", (0,0), (-1,-1), "CENTER"), ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold")]))
    story.append(firm_tbl)
    story.append(Spacer(1, 8))

    # Appendice fotografica (ridimensionamento sicuro)
    imgs = []
    for _, r in df_all.iterrows():
        files = r.get("_files", None)
        if files:
            for f in files:
                try:
                    if f.type and f.type.startswith("image/"):
                        imgs.append(f)
                except Exception:
                    pass

    if imgs:
        story.append(PageBreak())
        story.append(Paragraph("Appendice fotografica", H1))
        story.append(Paragraph("Selezione immagini caricate a supporto dell’audit.", Small))
        story.append(Spacer(1, 6))

        max_w = doc.width
        max_h = doc.height * 0.65
        for f in imgs:
            try:
                f.seek(0)
                ir = ImageReader(f)
                iw, ih = ir.getSize()
                scale = min(max_w / iw, max_h / ih)
                w, h = iw * scale, ih * scale
                story.append(RLImage(f, width=w, height=h))
                story.append(Spacer(1, 6))
                story.append(Paragraph(f.name, Small))
                story.append(Spacer(1, 12))
            except Exception:
                continue

    doc.build(story)
    pdf_bytes = buf.getvalue()
    buf.close()
    return pdf_bytes

# ---- Azione: genera/scarica il PDF ----
if not df_all.empty:
    if st.button("🧾 Genera Report PDF", key="btn_pdf"):
        pdf_data = build_pdf(df_all.copy(), logo_up)
        st.download_button(
            "⬇️ Scarica Report PDF",
            data=pdf_data,
            file_name=f"Audit_{(fornitore or 'fornitore')}_{filename_date(data_audit)}.pdf",
            mime="application/pdf",
            use_container_width=True,
            key="dl_pdf"
        )
else:
    st.info("Compila almeno un requisito per generare il PDF.")
