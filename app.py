# app.py
import streamlit as st
import sqlite3
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from decimal import Decimal, ROUND_HALF_UP
from datetime import datetime, date
import logging

from config import DB_FILE, PRODUCTION_MODE
from core.pricing_engine import PricingEngine, PricingInput
from core.hierarchy_resolver import HierarchyResolver

logging.basicConfig(level=logging.WARNING)

st.set_page_config(page_title="Bunker Commerciale - Salov", layout="wide")

# --- CSS PASTELLO AD ALTO CONTRASTO (Safe per Icone e Mobile) ---
st.markdown("""
<style>
    :root { color-scheme: light !important; }
    .stApp { background-color: #FCFAF7 !important; }
    section[data-testid="stSidebar"] { background-color: #F5F3EE !important; border-right: 1px solid #D1C9BC !important; }
    h1, h2, h3, h4, h5, h6 { color: #1B5E20 !important; font-weight: bold !important; }
    div[data-testid="stMetricValue"] { color: #D32F2F !important; font-weight: bold !important; }
    div[data-testid="stExpander"] { background-color: #FFFFFF !important; border: 1px solid #D1C9BC !important; border-radius: 4px !important; }
    .warning-box { background-color: #FFF3E0 !important; border-left: 5px solid #FF9800 !important; padding: 12px; border-radius: 4px; margin-bottom: 15px; color: #B78103 !important; font-weight: bold; }
    .info-box { background-color: #E8F5E9 !important; border-left: 5px solid #2E7D32 !important; padding: 12px; border-radius: 4px; margin-bottom: 15px; color: #1B5E20 !important; font-weight: bold; }
    .stMarkdown p, .stMarkdown li, label { color: #1C1C1C !important; font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif !important; font-size: 1.02rem !important; }
</style>
""", unsafe_allow_html=True)

def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT attivo FROM clienti LIMIT 1")
        cursor.execute("SELECT sottogruppo FROM accordi_commerciali LIMIT 1")
        cursor.execute("SELECT min_net_net_g FROM prodotti LIMIT 1")
    except sqlite3.OperationalError:
        cursor.execute("DROP TABLE IF EXISTS accordi_commerciali")
        cursor.execute("DROP TABLE IF EXISTS clienti")
        cursor.execute("DROP TABLE IF EXISTS prodotti")
        conn.commit()
    
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS prodotti (
        ean TEXT PRIMARY KEY, codice_sap TEXT, tipo_olio TEXT,
        descrizione_sap TEXT, descrizione_commerciale TEXT, formato_lt REAL,
        min_net_net_g REAL DEFAULT 0.0, confezione TEXT
    )""")
    
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS clienti (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        gruppo_macro TEXT, sottogruppo TEXT, associato_insegna TEXT,
        attivo BOOLEAN DEFAULT 1, UNIQUE(gruppo_macro, sottogruppo, associato_insegna)
    )""")
    
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS accordi_commerciali (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        gruppo_macro TEXT, sottogruppo TEXT, associato_insegna TEXT, livello TEXT, chiave_livello TEXT,
        listino_r REAL, sconto_1 REAL, sconto_2 REAL, sconto_3 REAL, sconto_4 REAL, sconto_5 REAL,
        sconto_6 REAL, sconto_7 REAL, sconto_y REAL, sconto_carico REAL, sconto_pagamento REAL,
        voce_contratto_1 REAL, voce_contratto_2 REAL, voce_contratto_3 REAL, voce_contratto_4 REAL, voce_contratto_5 REAL,
        UNIQUE(gruppo_macro, sottogruppo, associato_insegna, livello, chiave_livello)
    )""")
    conn.commit()
    
    cursor.execute("SELECT COUNT(*) FROM prodotti")
    if cursor.fetchone()[0] == 0:
        seed_baseline_data(conn)
    else:
        conn.close()

def seed_baseline_data(conn):
    cursor = conn.cursor()
    cursor.execute("DELETE FROM accordi_commerciali")
    cursor.execute("DELETE FROM clienti")
    cursor.execute("DELETE FROM prodotti")
    
    prodotti_salov = [
        ("8002210111110", "10002713", "EXTRAVERGINE", "SAGRA EXV BOT W12x1L CLASS IT", "Ex.v. Sagra Classico lt.1", 1.0, 10.00, "Bott.Lt 1"),
        ("8002210133440", "10003255", "EXTRAVERGINE", "SAGRA EXV 100%R-PET V12x750ML IT", "Ex.v. Sagra lt.0,75 PET", 0.75, 7.50, "Pet.Lt 0,75"),
        ("8002210115088", "10002716", "EXTRAVERGINE", "SAGRA GRAND EXV BOT W12x1L", "Ex.v. Sagra Grandulivo lt.1", 1.0, 10.00, "Bott.Lt 1"),
        ("8002210127562", "10002719", "EXTRAVERGINE", "SAGRA T.VIVE EXV BOT W 12x1L", "Ex.v. Sagra Terre Vive lt.1", 1.0, 10.00, "Bott.Lt 1"),
        ("8002210119543", "10000536", "EXTRAVERGINE", "SAGRA PROF. EXV PET C2x5L IT", "Ex.v. Sagra Prof Lt.5", 5.0, 50.00, "Pet lt 5"),
        ("8002210112827", "10002714", "EXTRAVERGINE", "SAGRA EXV 100%I BSA BOT W12x1L IT", "Ex.v. Sagra Bassa Acidità 100% ITA lt.1", 1.0, 15.00, "Bott.Lt 1"),
        ("8002210127425", "10002715", "EXTRAVERGINE", "SAGRA EXV 100%I BOT W 12x1L", "Ex.v. Sagra 100% Italiano lt.1", 1.0, 15.00, "Bott.Lt 1"),
        ("8002210128286", "10002720", "EXTRAVERGINE", "SAGRA EXV 100%I BIO BOT V12x1L IT", "Ex.v. Sagra Biologico 100% ITA lt.1", 1.0, 15.00, "Vetro lt 1"),
        ("8002210128248", "10002747", "EXTRAVERGINE", "SAGRA EXV BOT W12x750ML CLASS IT", "Ex.v. Sagra Classico lt.0,75", 0.75, 7.50, "Bott.Lt 0,75"),
        ("8002210121997", "10003315", "EXTRAVERGINE", "SAGRA GRAND EXV BOT W12x750ML  IT", "Ex.v. Sagra Grandulivo 0,75", 0.75, 7.50, "Bott.Lt 0,75"),
        ("8002210127197", "10003316", "EXTRAVERGINE", "SAGRA EXV 100%I BSA BOT W12x 750ML IT", "Ex.v. Sagra Bassa Acidità 100% ITA 0,75", 0.75, 11.25, "Bott.Lt 0,75"),
        ("8002210133792", "10003317", "EXTRAVERGINE", "SAGRA EXV 100% I BOT W 12x750ML IT", "Ex.v. Sagra 100% Italiano 0,75", 0.75, 11.25, "Bott.Lt 0,75"),
        ("8002210131815", "10003319", "EXTRAVERGINE", "SAGRA EXV 100%I BIO BOT W12x750ML IT", "Ex.v. Sagra Biologico 100% ITA  0,75", 0.75, 11.25, "Bott.Lt 0,75"),
        ("8002210130814", "60000444", "EXTRAVERGINE", "SAGRA EXV SPRAY C6x200ML ALLUMINIO IT", "Ex.v. Sagra Spray ml.200", 0.2, 2.00, "Spray Lt 0,20"),
        ("8002210124387", "10003061", "EXTRAVERGINE", "SAGRA PROF EXV PET T6x2L IT", "Ex.v. Sagra Prof lt.2", 2.0, 20.00, "Pet.Lt 2"),
        ("8002210131620", "10002724", "EXTRAVERGINE", "FBERIO EXV BOT W12x1L CLASS IT", "Ex.v. Filippo Berio Classico lt.1", 1.0, 12.00, "Bott.Lt 1"),
        ("8002210131644", "10002725", "EXTRAVERGINE", "FBERIO EXV BOT W12x1L BSA IT", "Ex.v. Filippo Berio Bassa Acidità lt.1", 1.0, 17.00, "Bott.Lt 1"),
        ("8002210131705", "10002726", "EXTRAVERGINE", "FBERIO EXV 100%I BOT W12x1L IT", "Ex.v. Filippo Berio 100% Italiano lt.1", 1.0, 18.00, "Bott.Lt 1"),
        ("8002210131767", "10002765", "EXTRAVERGINE", "FBERIO EXV BOT W12x750ML CLASS IT", "Ex.v. Filippo Berio Classico lt.0,75", 0.75, 12.00, "Bott.Lt 0,75"),
        ("8002210131668", "10002746", "EXTRAVERGINE", "FBERIO EXV BSA BOT W12x750ML IT", "Ex.v. Filippo Berio Bassa Acidità lt.0,75", 0.75, 17.00, "Bott.Lt 0,75"),
        ("8002210131804", "10002768", "EXTRAVERGINE", "FBERIO EXV 100%I BOT W12x750ML IT", "Ex.v. Filippo Berio 100% Italiano lt.0,75", 0.75, 18.00, "Bott.Lt 0,75"),
        ("8002210133013", "10003200", "EXTRAVERGINE", "FB R.O. EXV BIO 100%IT MB BOT W12X750 IT", "Ex.v. Filippo Berio Riserva Oro lt.0,75", 0.75, 19.00, "Bott.Lt 0,75"),
        ("8002210121461", "60000544", "EXTRAVERGINE", "EX.V. BUSTINA 10mlx250 FILIPPO BERIO ITA", "Ex.v. Filippo Berio Bustina ml.10", 0.01, 0.12, "bust lt 0,01"),
        ("8002210126572", "10003240", "OLIVA", "SAGRA OOL PUR R-PET V12X750ML CLASS IT", "Oliva Sagra RPET lt.0,75 PET", 0.75, 8.00, "Pet.Lt 0,75"),
        ("8002210001305", "10002717", "OLIVA", "SAGRA OOL BOT W12x1L CLASS", "Oliva Sagra lt.1", 1.0, 8.00, "Bott.Lt 1"),
        ("8002210128453", "10002718", "OLIVA", "SAGRA GRAND OOL BOT W12x1L", "Oliva Sagra Grandulivo lt.1", 1.0, 8.00, "Bott.Lt 1"),
        ("8002210126176", "10003288", "OLIVA", "SAGRA OOL PUR R-PET T6X1.5L IT", "Oliva Sagra lt.1,5", 1.5, 12.00, "Pet.Lt 1,5"),
        ("8002210119567", "10000537", "OLIVA", "SAGRA PROF. OOL PUR PET C2x5L IT", "Oliva Sagra Prof Lt.5", 5.0, 40.00, "Pet.Lt 5"),
        ("8002210132436", "10002965", "OLIVA", "FBERIO OOL PUR BOT V6X500ML IT", "Oliva Filippo Berio lt.0,50", 0.5, 4.11, "Bott.Lt 0,5"),
        ("8002210131729", "10002727", "OLIVA", "FBERIO OOL PUR BOT W12x1L IT", "Oliva Filippo Berio lt.1", 1.0, 7.75, "Bott.Lt 1"),
        ("8002210131781", "10002766", "OLIVA", "FBERIO OOL PUR BOT W12x750ML IT", "Oliva Filippo Berio lt.0,75", 0.75, 5.97, "Bott.Lt 0,75"),
        ("8002210122307", "10000922", "OLIVA", "FBERIO OOL PUR LAT V8x1L IT", "Oliva Filippo Berio Latta lt.1", 1.0, 8.10, "Latta lt 1"),
        ("8002210111486", "10003307", "SEMI", "SAGRA SEM MAIS PET V12x1L IT", "Mais Sagra lt.1", 1.0, 2.00, "Pet.Lt 1"),
        ("8002210127067", "10003286", "SEMI", "SAGRA SEM MAIS PET T6x1.5L IT", "Mais Sagrì lt.1,5", 1.5, 3.00, "Pet.Lt 1,5"),
        ("8002210112889", "10003089", "SEMI", "SAGRA SEM MAIS PET T6x2L IT", "Mais Sagra lt.2", 2.0, 4.00, "Pet.Lt 2"),
        ("8002210000551", "10003311", "SEMI", "SAGRA SEM ARACHIDE PET V12x1L IT", "Arachide Sagra lt.1", 1.0, 3.00, "Pet.Lt 1"),
        ("8002210126916", "10003284", "SEMI", "SAGRI SEM ARACHIDE PET T6x1.5L IT", "Arachide Sagrì lt.1,5", 1.5, 4.50, "Pet.Lt 1,5"),
        ("8002210112865", "10003086", "SEMI", "SAGRA SEM ARACHIDE PET T6x2L IT", "Arachide Sagra lt.2", 2.0, 6.00, "Pet.Lt 2"),
        ("8002210116160", "10000326", "SEMI", "SAGRA PROF SEM ARACHIDE PET C2x5L IT", "Arachide Sagra Prof. Lt.5", 5.0, 15.00, "Pet lt 5"),
        ("8002210111905", "10003310", "SEMI", "SAGRA SEM GIRAS PET V12x1L IT", "Girasole Sagra lt.1", 1.0, 2.20, "Pet.Lt 1"),
        ("8002210126817", "10003287", "SEMI", "SAGRI SEM GIRAS PET T6x1.5L IT", "Girasole Sagrì lt.1,5", 1.5, 3.30, "Pet.Lt 1,5"),
        ("8002210113107", "10003087", "SEMI", "SAGRA SEM GIRAS PET T6x2L IT", "Girasole Sagra lt.2", 2.0, 4.40, "Pet.Lt 2"),
        ("8002210115453", "10003062", "SEMI", "SAGRA PROF SEM GIRAS PET C2x5L IT", "Girasole Sagra Prof Lt.5", 5.0, 11.00, "Pet lt 5"),
        ("8002210111295", "10002933", "SEMI", "SAGRA FRIMX SEM FRITT PET V12x1L NOP IT", "Frimax Sagra lt.1", 1.0, 2.25, "Pet Lt 1"),
        ("8002210126893", "10003285", "SEMI", "SAGRI SEM FRITT PET T6x1.5L IT", "Frimax Sagrì lt.1,5", 1.5, 3.38, "Pet.Lt 1,5"),
        ("8002210112940", "10003085", "SEMI", "SAGRA FRIMX SEM FRITT PET T6x2L NOP IT", "Frimax Sagra lt.2", 2.0, 4.50, "Pet Lt 2"),
        ("8002210115484", "10002644", "SEMI", "SAGRA FRIMX SEM FRITT PET C2x5L NOP IT", "Frimax Sagra lt.5", 5.0, 11.25, "Pet Lt 5"),
        ("8002210134140", "10003327", "SEMI", "GRAZIA SEM GIRAS LAT 1x20L IT", "Frimax Spray ml.200", 0.2, 0.45, "Spray Lt 0,20"),
        ("8002210127401", "10003309", "SEMI", "SAGRA SEM GIRAS AO PET V12x1L IT", "Girasole Alto Oleico Sagra lt.1", 1.0, 2.80, "Pet.Lt 1"),
        ("8002210126336", "10003063", "SEMI", "SAGRA PROF SEM GIRAS AO PET C2x5L IT", "Girasole Alto Oleico Sagra Prof lt.5", 5.0, 14.00, "Pet lt 5"),
        ("8002210129290", "10003312", "SEMI", "SAGRA SEM VINACC PET V12x1L IT", "Vinacciolo Sagra lt.1", 1.0, 5.00, "Pet.Lt 1"),
        ("8002210130289", "10003082", "EXTRAVERGINE", "FBERIO EXV CLASS MB BOT V6x250ML IT", "Ex.v. F.Berio Anti Rab Classico lt.0,25", 0.25, 2.50, "Vetro lt 0,25"),
        ("8002210130210", "10003081", "EXTRAVERGINE", "FBERIO EXV 100%I MB BOT V6x250ML IT", "Ex.v. F.Berio Anti Rab 100% ITA lt.0,25", 0.25, 3.00, "Vetro lt 0,25"),
        ("8002210130340", "10003091", "EXTRAVERGINE", "FBERIO EXV CLASS MB BOT V6x500ML IT", "Ex.v. F.Berio Anti Rab Classico lt.0,50", 0.5, 4.30, "Vetro lt 0,50"),
        ("8002210130302", "10003079", "EXTRAVERGINE", "FBERIO EXV 100%I MB BOT V6x500ML IT", "Ex.v. F.Berio Anti Rab 100% ITA lt.0,50", 0.5, 4.80, "Vetro lt 0,50"),
        ("8002210132573", "10003072", "EXTRAVERGINE", "FBERIO EXV BOT V6x500ML TOSC IT", "Ex.v. F.Berio Toscano lt.0,50", 0.5, 10.00, "Vetro lt 0,50"),
        ("8002210130234", "60000591", "EXTRAVERGINE", "FBERIO EXV DRES BOT V6x250ML PEP TE IT", "Ex.v. F.Berio Peperoncino lt.0,25", 0.25, 3.50, "Vetro lt 0,25"),
        ("8002210130791", "60000590", "ACETO", "FBERIO ACE BALS BOT V6x250ML IT", "Aceto Balsamico F.Berio lt.0,25", 0.25, 2.00, "Vetro lt 0,25"),
        ("8002210130197", "60000589", "ACETO", "FBERIO ACE BALS BOT V6x500ML IT", "Aceto Balsamico F.Berio lt.0,50", 0.5, 2.10, "Vetro lt 0,50")
    ]
    for p in prodotti_salov:
        cursor.execute("INSERT OR REPLACE INTO prodotti VALUES (?, ?, ?, ?, ?, ?, ?, ?)", p)
        
    clienti_demo = [
        ("COOP ITALIA", "COOP ITALIA SOTTOGRUPPO", "ALLEANZA 3.0"),
        ("COOP ITALIA", "COOP ITALIA SOTTOGRUPPO", "NORDOVEST"),
        ("COOP ITALIA", "COOP ITALIA SOTTOGRUPPO", "LIGURIA"),
        ("CONAD", "CONAD SOTTOGRUPPO", "CONAD ADRIATICO"),
        ("CONAD", "CONAD SOTTOGRUPPO", "CONAD CENTRO NORD"),
        ("CONAD", "CONAD SOTTOGRUPPO", "PAC 2000A"),
        ("ESSELUNGA GRUPPO", "ESSELUNGA SOTTOGRUPPO", "ESSELUNGA"),
        ("SELEX GRUPPO", "SELEX SOTTOGRUPPO", "MAXI DI"),
        ("SELEX GRUPPO", "SELEX SOTTOGRUPPO", "DIMAR"),
        ("SELEX GRUPPO", "SELEX SOTTOGRUPPO", "UNICOMM")
    ]
    for c in clienti_demo:
        cursor.execute("INSERT OR IGNORE INTO clienti (gruppo_macro, sottogruppo, associato_insegna) VALUES (?, ?, ?)", c)
        
    fallback_data = [
        ('COOP ITALIA', '', '', 'GRUPPO', '', None, 10.0, 5.0, None, None, None, None, None, None, 1.5, 1.0, 5.0, 2.0, None, None, None),
        ('COOP ITALIA', 'COOP ITALIA SOTTOGRUPPO', '', 'SOTTOGRUPPO', '', None, None, None, None, 2.0, None, None, None, None, None, None, None, None, None, None, None),
        ('COOP ITALIA', 'COOP ITALIA SOTTOGRUPPO', 'ALLEANZA 3.0', 'CATEGORIA', 'EXTRAVERGINE', None, None, None, None, None, None, 3.0, None, None, None, None, None, None, None, None, 1.0),
        ('COOP ITALIA', 'COOP ITALIA SOTTOGRUPPO', 'ALLEANZA 3.0', 'REFERENZA', '8002210131620', 66.00, None, None, None, None, None, None, 12.0, 5.0, None, None, None, None, None, None, None),
        ('COOP ITALIA', 'COOP ITALIA SOTTOGRUPPO', 'ALLEANZA 3.0', 'REFERENZA', '8002210111110', 60.80, None, None, None, None, None, None, 15.0, 0.0, None, None, None, None, None, None, None),
        ('COOP ITALIA', 'COOP ITALIA SOTTOGRUPPO', 'ALLEANZA 3.0', 'REFERENZA', '8002210001305', 43.20, None, None, None, None, None, None, 12.0, 0.0, None, None, None, None, None, None, None),
        ('CONAD', '', '', 'GRUPPO', '', None, 10.0, 5.0, None, None, None, None, None, None, 1.5, 1.0, 5.0, 2.0, None, None, None),
        ('CONAD', 'CONAD SOTTOGRUPPO', '', 'SOTTOGRUPPO', '', None, None, None, None, 2.0, None, None, None, None, None, None, None, None, None, None, None),
        ('CONAD', 'CONAD SOTTOGRUPPO', 'CONAD ADRIATICO', 'CATEGORIA', 'EXTRAVERGINE', None, None, None, None, None, None, 3.0, None, None, None, None, None, None, None, None, 1.0),
        ('CONAD', 'CONAD SOTTOGRUPPO', 'CONAD ADRIATICO', 'REFERENZA', '8002210131620', 52.80, None, None, None, None, None, None, 12.0, 0.0, None, None, None, None, None, None, None),
        ('CONAD', 'CONAD SOTTOGRUPPO', 'CONAD ADRIATICO', 'REFERENZA', '8002210111110', 48.64, None, None, None, None, None, None, 15.0, 0.0, None, None, None, None, None, None, None),
        ('CONAD', 'CONAD SOTTOGRUPPO', 'CONAD ADRIATICO', 'REFERENZA', '8002210001305', 34.56, None, None, None, None, None, None, 12.0, 0.0, None, None, None, None, None, None, None),
    ]
    cursor.executemany("""
    INSERT OR REPLACE INTO accordi_commerciali (
        gruppo_macro, sottogruppo, associato_insegna, livello, chiave_livello, listino_r,
        sconto_1, sconto_2, sconto_3, sconto_4, sconto_5,
        sconto_6, sconto_7, sconto_y, sconto_carico, sconto_pagamento,
        voce_contratto_1, voce_contratto_2, voce_contratto_3, voce_contratto_4, voce_contratto_5
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, fallback_data)
    conn.commit()
    conn.close()

init_db()

menu = st.sidebar.radio("SELEZIONA SCHEDA", ["Simulatore Offerte", "Back-Office (Gestione Dati)", "Report Sintetico", "Guida Operativa"])

# ==========================================
# SCHEDA 1: SIMULATORE 
# ==========================================
if menu == "Simulatore Offerte":
    st.title("Bunker Commerciale Salov - Simulatore")
    conn = sqlite3.connect(DB_FILE)
    
    st.sidebar.header("Parametri Negoziazione")
    
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT gruppo_macro FROM clienti WHERE attivo=1 ORDER BY gruppo_macro")
    gruppi = [r[0] for r in cursor.fetchall()]
    
    if not gruppi:
        st.warning("ATTENZIONE: Nessun cliente caricato. Sblocca il sistema caricando i dati dal Back-Office.")
    else:
        gruppo_sel = st.sidebar.selectbox("1. Gruppo GDO", gruppi, help="Seleziona la centrale d'acquisto.")
        
        cursor.execute("SELECT DISTINCT sottogruppo FROM clienti WHERE gruppo_macro=? AND attivo=1 ORDER BY sottogruppo", (gruppo_sel,))
        sottogruppi = [r[0] for r in cursor.fetchall()]
        sottogruppo_sel = st.sidebar.selectbox("2. Sottogruppo GDO", sottogruppi, help="Seleziona il sottogruppo di canale.")
        
        cursor.execute("SELECT DISTINCT associato_insegna FROM clienti WHERE gruppo_macro=? AND sottogruppo=? AND attivo=1 ORDER BY associato_insegna", (gruppo_sel, sottogruppo_sel))
        associati = [r[0] for r in cursor.fetchall()]
        associato_sel = st.sidebar.selectbox("3. Insegna Locale / Associato", associati, help="Seleziona l'associato locale.")

        cursor.execute("SELECT ean, descrizione_commerciale, tipo_olio, min_net_net_g, codice_sap, formato_lt FROM prodotti")
        prodotti = cursor.fetchall()
        prodotti_dict = {f"{p[1]} [EAN: {p[0]}]": (p[0], p[2], p[3], p[4], p[5]) for p in prodotti}
        prodotto_scelto = st.sidebar.selectbox("4. Referenza Salov", list(prodotti_dict.keys()), help="Seleziona la referenza.")
        ean, tipo_olio, min_net_net_g, codice_sap, formato_lt = prodotti_dict[prodotto_scelto]

        contract = HierarchyResolver.resolve(conn, gruppo_sel, sottogruppo_sel, associato_sel, ean, tipo_olio)

        st.sidebar.markdown("---")
        st.sidebar.subheader("Verità Contrattuale")
        if contract.listino_r is None:
            st.error("ATTENZIONE: PRODOTTO FUORI ASSORTIMENTO PER QUESTO CLIENTE")
            st.stop()
        else:
            st.sidebar.info(f"Listino Base (R): {contract.listino_r:.2f} Euro")
            st.sidebar.text(f"Livello Risolto: {contract.livello_risolto}")

        st.markdown("### Scegli Metodologia di Negoziazione")
        metodo_lavoro = st.radio(
            "Seleziona l'approccio negoziale:",
            ["A. Partenza da Prezzo Target (Calcolo automatico Sconto Promo Z)", "B. Tentativi Spot Manuali (Immissione Sconto Z libera)"],
            horizontal=True
        )

        st.markdown("---")
        
        if "A. Partenza" in metodo_lavoro:
            st.markdown("### Definizione Obiettivo Economico")
            target_container = st.container(border=True)
            with target_container:
                col_t1, col_t2 = st.columns([2, 1])
                with col_t1:
                    target_net_net = st.number_input(
                        "PREZZO TARGET NET NET DESIDERATO (Euro/Pz)", 
                        min_value=0.0, 
                        value=float(min_net_net_g), 
                        step=0.10,
                        help="Fissa il ricavo reale netto a bottiglia (AM) che desideri ottenere. Di default è la soglia minima di sicurezza (G)."
                    )
                with col_t2:
                    st.markdown("<br>", unsafe_allow_html=True)
                    st.info("Modo Target Attivo")
        else:
            target_net_net = 0.0

        st.markdown("---")
        st.subheader("Manovre e Leve Sconti")
        col_l1, col_l2 = st.columns(2)
        
        with col_l1:
            sconto_y = st.number_input(
                "Sconto Continuativo Y (%)", 
                min_value=0.0, max_value=100.0, 
                value=float(contract.sconto_y), step=0.5
            )
            st.markdown(
                f"<div class='warning-box'>ATTENZIONE - VALORE CONCORDATO A DATABASE: {contract.sconto_y:.2f}%<br>"
                f"<span style='font-size:0.8em; font-weight:normal;'>La modifica in corsa potrebbe violare gli accordi locali già depositati.</span></div>", 
                unsafe_allow_html=True
            )
            
            sconto_aa = st.number_input(
                "Sconto Unitario in fattura (Euro/Pz) [AA]", 
                min_value=0.0, value=0.0, step=0.05
            )

        with col_l2:
            st.markdown("**Analisi Limiti Promozionali**")
            col_z1, col_z2 = st.columns(2)
            
            with col_z1:
                if target_net_net > 0:
                    target_dec = Decimal(f"{target_net_net:.5f}")
                    temp_input = PricingInput(
                        listino_r=contract.listino_r,
                        sconto_1=contract.sconto_1, sconto_2=contract.sconto_2, sconto_3=contract.sconto_3,
                        sconto_4=contract.sconto_4, sconto_5=contract.sconto_5, sconto_6=contract.sconto_6, sconto_7=contract.sconto_7,
                        sconto_y=Decimal(f"{sconto_y:.5f}"), sconto_z=Decimal("0.00"), sconto_aa=Decimal(f"{sconto_aa:.5f}"),
                        sconto_carico=contract.sconto_carico, sconto_pagamento=contract.sconto_pagamento,
                        voce_i=contract.voce_i, voce_ii=contract.voce_ii, voce_iii=contract.voce_iii, voce_iv=contract.voce_iv, voce_v=contract.voce_v,
                        min_net_net_g=Decimal(str(min_net_net_g))
                    )
                    sconto_z_val = PricingEngine.calculate_inverse(target_dec, temp_input, "Z")
                    sconto_z = sconto_z_val
                    st.number_input("Sconto Promozionale (%) [Z]", value=float(sconto_z), disabled=True, format="%.2f", help="Calcolato automaticamente per raggiungere il Net Net target.")
                else:
                    sconto_z_input = st.number_input("Sconto Promozionale (%) [Z] (Manuale)", min_value=0.0, max_value=100.0, value=0.0, step=0.5)
                    sconto_z = Decimal(f"{sconto_z_input:.5f}")
            
            with col_z2:
                temp_input_max_z = PricingInput(
                    listino_r=contract.listino_r,
                    sconto_1=contract.sconto_1, sconto_2=contract.sconto_2, sconto_3=contract.sconto_3,
                    sconto_4=contract.sconto_4, sconto_5=contract.sconto_5, sconto_6=contract.sconto_6, sconto_7=contract.sconto_7,
                    sconto_y=Decimal(f"{sconto_y:.5f}"), sconto_z=Decimal("0.00"), sconto_aa=Decimal(f"{sconto_aa:.5f}"),
                    sconto_carico=contract.sconto_carico, sconto_pagamento=contract.sconto_pagamento,
                    voce_i=contract.voce_i, voce_ii=contract.voce_ii, voce_iii=contract.voce_iii, voce_iv=contract.voce_iv, voce_v=contract.voce_v,
                    min_net_net_g=Decimal(str(min_net_net_g))
                )
                z_max_consentito = PricingEngine.calculate_inverse(Decimal(str(min_net_net_g)), temp_input_max_z, "Z")
                
                temp_input_max_aa = PricingInput(
                    listino_r=contract.listino_r,
                    sconto_1=contract.sconto_1, sconto_2=contract.sconto_2, sconto_3=contract.sconto_3,
                    sconto_4=contract.sconto_4, sconto_5=contract.sconto_5, sconto_6=contract.sconto_6, sconto_7=contract.sconto_7,
                    sconto_y=Decimal(f"{sconto_y:.5f}"), sconto_z=Decimal(f"{sconto_z:.5f}"), sconto_aa=Decimal("0.00"),
                    sconto_carico=contract.sconto_carico, sconto_pagamento=contract.sconto_pagamento,
                    voce_i=contract.voce_i, voce_ii=contract.voce_ii, voce_iii=contract.voce_iii, voce_iv=contract.voce_iv, voce_v=contract.voce_v,
                    min_net_net_g=Decimal(str(min_net_net_g))
                )
                aa_max_consentito = PricingEngine.calculate_inverse(Decimal(str(min_net_net_g)), temp_input_max_aa, "AA")
                
                st.number_input("Sconto Promo MAX Consentito [Z]", value=float(z_max_consentito), disabled=True, format="%.2f", help="Il massimo Sconto Z che puoi inserire (a parità di AA) prima di andare in blocco.")
                if not "A. Partenza" in metodo_lavoro:
                    st.number_input("Sconto Unitario MAX Consentito [AA]", value=float(aa_max_consentito), disabled=True, format="%.2f", help="Il massimo Sconto AA in Euro che puoi inserire (a parità di Z) prima di andare in blocco.")

        engine_input = PricingInput(
            listino_r=contract.listino_r,
            sconto_1=contract.sconto_1, sconto_2=contract.sconto_2, sconto_3=contract.sconto_3,
            sconto_4=contract.sconto_4, sconto_5=contract.sconto_5, sconto_6=contract.sconto_6, sconto_7=contract.sconto_7,
            sconto_y=Decimal(f"{sconto_y:.5f}"), sconto_z=sconto_z, sconto_aa=Decimal(f"{sconto_aa:.5f}"),
            sconto_carico=contract.sconto_carico, sconto_pagamento=contract.sconto_pagamento,
            voce_i=contract.voce_i, voce_ii=contract.voce_ii, voce_iii=contract.voce_iii, voce_iv=contract.voce_iv, voce_v=contract.voce_v,
            min_net_net_g=Decimal(str(min_net_net_g))
        )
        result = PricingEngine.calculate(engine_input)

        st.markdown("---")
        col_c1, col_c2 = st.columns(2)
        
        with col_c1:
            with st.expander("Verifica Margine e Stato", expanded=True):
                st.metric("PREZZO NET NET RISULTANTE", f"{result.net_net_finale:.3f} Euro")
                st.metric("SOGLIA MINIMA NET NET (G)", f"{min_net_net_g:.3f} Euro")
                if result.guardrail_ok:
                    st.success(f"VERDE (APPROVATO) - Margine sicuro. Delta: {result.delta_vs_min:+.3f} Euro")
                else:
                    st.error(f"ROSSO (BLOCCATO) - Sotto soglia di {abs(result.delta_vs_min):.3f} Euro")
        
        with col_c2:
            with st.expander("Finestra Temporale Promo", expanded=True):
                col_d1, col_d2 = st.columns(2)
                with col_d1:
                    sell_in_dal = st.date_input("Inizio Sell-In", date.today(), key="si_dal")
                    sell_in_al = st.date_input("Fine Sell-In", date.today(), key="si_al")
                with col_d2:
                    sell_out_dal = st.date_input("Inizio Sell-Out", date.today(), key="so_dal")
                    sell_out_al = st.date_input("Fine Sell-Out", date.today(), key="so_al")

        st.markdown("---")
        st.subheader("Tabella Sequenziale Estesa della Struttura di Costo")
        df_waterfall = pd.DataFrame([
            {"Fase Pricing": step.fase, "Valore Unitario": step.valore, "Dettaglio Operazione": step.descrizione}
            for step in result.steps
        ])
        st.dataframe(df_waterfall, use_container_width=True, hide_index=True)

        st.markdown("---")
        
        def genera_scheda_negoziale():
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Proposta_Commerciale"
            ws.views.sheetView[0].showGridLines = True
            
            font_title = Font(name="Arial", size=15, bold=True, color="FFFFFF")
            font_section = Font(name="Arial", size=11, bold=True, color="000000")
            font_label = Font(name="Arial", size=10, bold=True)
            font_value = Font(name="Arial", size=10)
            fill_header = PatternFill(start_color="1B5E20", end_color="1B5E20", fill_type="solid")
            fill_sub = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")
            thin_border = Border(
                left=Side(style='thin', color='DDDDDD'),
                right=Side(style='thin', color='DDDDDD'),
                top=Side(style='thin', color='DDDDDD'),
                bottom=Side(style='thin', color='DDDDDD')
            )
            
            ws.merge_cells('A1:D1')
            ws['A1'] = "SALOV S.p.A. - SCHEDA PROPOSTA COMMERCIALE"
            ws['A1'].font = font_title
            ws['A1'].fill = fill_header
            ws['A1'].alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[1].height = 40
            
            ws['A3'] = "ANAGRAFICA GDO"
            ws['A3'].font = font_section
            ws['A3'].fill = fill_sub
            ws.merge_cells('A3:D3')
            
            ws['A4'] = "Gruppo Macro:"
            ws['B4'] = gruppo_sel
            ws['A5'] = "Sottogruppo:"
            ws['B5'] = sottogruppo_sel
            ws['A6'] = "Insegna Locale / Associato:"
            ws['B6'] = associato_sel
            
            for r in range(4, 7):
                ws[f'A{r}'].font = font_label
                ws[f'B{r}'].font = font_value
            
            ws['A8'] = "DETTAGLIO REFERENZA"
            ws['A8'].font = font_section
            ws['A8'].fill = fill_sub
            ws.merge_cells('A8:D8')
            
            ws['A9'] = "Descrizione Articolo:"
            ws['B9'] = prodotto_scelto.split(" [EAN:")[0]
            ws['A10'] = "EAN:"
            ws['B10'] = ean
            ws['A11'] = "Codice SAP:"
            ws['B11'] = codice_sap
            ws['A12'] = "Formato:"
            ws['B12'] = f"{formato_lt} Litri"
            
            for r in range(9, 13):
                ws[f'A{r}'].font = font_label
                ws[f'B{r}'].font = font_value
                
            ws['A14'] = "FINESTRE TEMPORALI PROMO"
            ws['A14'].font = font_section
            ws['A14'].fill = fill_sub
            ws.merge_cells('A14:D14')
            
            ws['A15'] = "Periodo Sell-In:"
            ws['B15'] = f"Dal {sell_in_dal.strftime('%d/%m/%Y')} al {sell_in_al.strftime('%d/%m/%Y')}"
            ws['A16'] = "Periodo Sell-Out:"
            ws['B16'] = f"Dal {sell_out_dal.strftime('%d/%m/%Y')} al {sell_out_al.strftime('%d/%m/%Y')}"
            
            for r in range(15, 17):
                ws[f'A{r}'].font = font_label
                ws[f'B{r}'].font = font_value
            
            ws['A18'] = "CASCATA DI PRICING NEGOZIALE"
            ws['A18'].font = font_section
            ws['A18'].fill = fill_sub
            ws.merge_cells('A18:D18')
            
            ws['A19'] = "Elemento di Costo"
            ws['B19'] = "Valore"
            ws['C19'] = "Tipologia Operazione"
            for col in ['A', 'B', 'C']:
                ws[f'{col}19'].font = font_label
                
            row_idx = 20
            for step in result.steps:
                ws.cell(row=row_idx, column=1, value=step.fase).font = font_value
                ws.cell(row=row_idx, column=2, value=float(step.valore)).font = font_value
                ws.cell(row=row_idx, column=2).number_format = '#,##0.000 €'
                ws.cell(row=row_idx, column=3, value=step.descrizione).font = font_value
                row_idx += 1
                
            ws.cell(row=row_idx+1, column=1, value="SOGLIA MINIMA AM (G):").font = font_label
            ws.cell(row=row_idx+1, column=2, value=float(min_net_net_g)).font = font_value
            ws.cell(row=row_idx+1, column=2).number_format = '#,##0.00 €'
            
            ws.cell(row=row_idx+2, column=1, value="DELTA DI MARGINE VS SOGLIA:").font = font_label
            ws.cell(row=row_idx+2, column=2, value=float(result.delta_vs_min)).font = font_value
            ws.cell(row=row_idx+2, column=2).number_format = '#,##0.00 €'
            
            ws.cell(row=row_idx+3, column=1, value="STATO DEL MARGINE:").font = font_label
            stato_txt = "VERDE (APPROVATO)" if result.guardrail_ok else "ROSSO (SOTTO SOGLIA)"
            ws.cell(row=row_idx+3, column=2, value=stato_txt).font = font_label
            
            ws.column_dimensions['A'].width = 32
            ws.column_dimensions['B'].width = 38
            ws.column_dimensions['C'].width = 45
            
            for row in ws.iter_rows(min_row=1, max_row=row_idx+4, min_col=1, max_col=3):
                for cell in row:
                    cell.border = thin_border
            
            buffer = io.BytesIO()
            wb.save(buffer)
            return buffer.getvalue()

        proposta_excel = genera_scheda_negoziale()
        
        st.download_button(
            label="SCARICA PROPOSTA COMMERCIALE PER IL CLIENTE",
            data=proposta_excel,
            file_name=f"Proposta_{associato_sel}_{codice_sap}_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    conn.close()

# ==========================================
# SCHEDA 2: BACK-OFFICE (GDO CON TRANSAZIONI ATOMICHE SICURE)
# ==========================================
elif menu == "Back-Office (Gestione Dati)":
    st.title("Back-Office - Gestione Dati e Backup")
    conn = sqlite3.connect(DB_FILE)
    
    st.subheader("Modifica Diretta dei Contratti in Database (A caldo)")
    st.markdown("Questa griglia ti permette di modificare direttamente i record esistenti nel database locale. Qualsiasi modifica salvata si rifletterà all'istante sul Simulatore.")
    
    df_database_editor = pd.read_sql_query("""
        SELECT id, gruppo_macro, sottogruppo, associato_insegna, livello, chiave_livello, listino_r,
               sconto_1, sconto_2, sconto_3, sconto_4, sconto_5, sconto_6, sconto_7, sconto_y,
               sconto_carico, sconto_pagamento, voce_contratto_1, voce_contratto_2, voce_contratto_3,
               voce_contratto_4, voce_contratto_5
        FROM accordi_commerciali
    """, conn)
    
    edited_df = st.data_editor(
        df_database_editor, 
        num_rows="dynamic", 
        use_container_width=True,
        hide_index=True,
        key="db_data_editor"
    )
    
    if st.button("SALVA MODIFICHE DIRETTE NEL DATABASE"):
        cursor = conn.cursor()
        try:
            with conn:
                cursor.execute("DELETE FROM accordi_commerciali")
                for _, r in edited_df.iterrows():
                    def check_nan(val):
                        return float(val) if (pd.notna(val) and str(val).strip() != "") else None
                    
                    cursor.execute("""
                    INSERT OR REPLACE INTO accordi_commerciali (
                        id, gruppo_macro, sottogruppo, associato_insegna, livello, chiave_livello, listino_r,
                        sconto_1, sconto_2, sconto_3, sconto_4, sconto_5, sconto_6, sconto_7, sconto_y,
                        sconto_carico, sconto_pagamento, voce_contratto_1, voce_contratto_2, voce_contratto_3,
                        voce_contratto_4, voce_contratto_5
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        check_nan(r.get("id")),
                        str(r.get("gruppo_macro")).upper().strip() if pd.notna(r.get("gruppo_macro")) else "",
                        str(r.get("sottogruppo")).upper().strip() if pd.notna(r.get("sottogruppo")) else "",
                        str(r.get("associato_insegna")).upper().strip() if pd.notna(r.get("associato_insegna")) else "",
                        str(r.get("livello")).upper().strip() if pd.notna(r.get("livello")) else "GRUPPO",
                        str(r.get("chiave_livello")).strip() if pd.notna(r.get("chiave_livello")) else "",
                        check_nan(r.get("listino_r")),
                        check_nan(r.get("sconto_1")), check_nan(r.get("sconto_2")), check_nan(r.get("sconto_3")),
                        check_nan(r.get("sconto_4")), check_nan(r.get("sconto_5")), check_nan(r.get("sconto_6")),
                        check_nan(r.get("sconto_7")), check_nan(r.get("sconto_y")), check_nan(r.get("sconto_carico")), check_nan(r.get("sconto_pagamento")),
                        check_nan(r.get("voce_contratto_1")), check_nan(r.get("voce_contratto_2")), check_nan(r.get("voce_contratto_3")),
                        check_nan(r.get("voce_contratto_4")), check_nan(r.get("voce_contratto_5"))
                    ))
            st.success("VERDE (APPROVATO) - Il database locale è stato aggiornato correttamente.")
            st.rerun()
        except Exception as e:
            st.error(f"ROSSO (BLOCCATO) - Errore durante l'elaborazione del file: {e}")

    st.markdown("---")
    col_b1, col_b2 = st.columns(2)
    
    with col_b1:
        st.subheader("Esportazione e Backup")
        st.markdown("Scarica il tracciato attuale degli accordi commerciali per lavorarlo localmente in Excel.")
        
        query_accordi = """
        SELECT a.gruppo_macro as GRUPPO_MACRO, a.sottogruppo as SOTTOGRUPPO, a.associato_insegna as ASSOCIATO_INSEGNA,
               a.livello as LIVELLO, a.chiave_livello as CHIAVE_LIVELLO,
               CASE 
                    WHEN a.livello = 'REFERENZA' THEN p.descrizione_commerciale 
                    WHEN a.livello = 'CATEGORIA' THEN 'Accordo di Categoria: ' || a.chiave_livello
                    ELSE 'Contratto Quadro'
               END as DESCRIZIONE_PRODOTTO,
               a.listino_r as LISTINO_BASE_R,
               a.sconto_1 as SCONTO_1, a.sconto_2 as SCONTO_2, a.sconto_3 as SCONTO_3, a.sconto_4 as SCONTO_4, a.sconto_5 as SCONTO_5,
               a.sconto_6 as SCONTO_LOCAL_6, a.sconto_7 as SCONTO_LOCAL_7, a.sconto_y as SCONTO_CONTINUATIVO_Y,
               a.sconto_carico as SCONTO_CARICO_LOGISTICA, a.sconto_pagamento as SCONTO_PAGAMENTO_AC,
               a.voce_contratto_1 as PFA_VOCE_I, a.voce_contratto_2 as PFA_VOCE_II,
               a.voce_contratto_3 as PFA_VOCE_III, a.voce_contratto_4 as PFA_VOCE_IV, a.voce_contratto_5 as PFA_VOCE_V
        FROM accordi_commerciali a
        LEFT JOIN prodotti p ON a.chiave_livello = p.ean AND a.livello = 'REFERENZA'
        """
        df_accordi = pd.read_sql_query(query_accordi, conn)
        
        colonne_ordinate = [
            "GRUPPO_MACRO", "SOTTOGRUPPO", "ASSOCIATO_INSEGNA", "LIVELLO", "CHIAVE_LIVELLO", "DESCRIZIONE_PRODOTTO",
            "LISTINO_BASE_R", "SCONTO_1", "SCONTO_2", "SCONTO_3", "SCONTO_4", "SCONTO_5",
            "SCONTO_LOCAL_6", "SCONTO_LOCAL_7", "SCONTO_CONTINUATIVO_Y", "SCONTO_CARICO_LOGISTICA", "SCONTO_PAGAMENTO_AC",
            "PFA_VOCE_I", "PFA_VOCE_II", "PFA_VOCE_III", "PFA_VOCE_IV", "PFA_VOCE_V"
        ]
        df_accordi = df_accordi[colonne_ordinate]
        
        buffer_export = io.BytesIO()
        with pd.ExcelWriter(buffer_export, engine='openpyxl') as writer:
            df_accordi.to_excel(writer, index=False, sheet_name="Accordi_GDO")
            
        st.download_button(
            label="Scarica Backup / Template Excel",
            data=buffer_export.getvalue(),
            file_name=f"Backup_Bunker_Commerciale_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with col_b2:
        st.subheader("Importazione Massiva Contratti")
        st.markdown("Carica il file Excel compilato per sovrascrivere o aggiungere condizioni.")
        uploaded_file = st.file_uploader("Trascina il file Excel (.xlsx)", type=["xlsx"])
        
        if uploaded_file is not None:
            if st.button("Conferma Scrittura nel Bunker"):
                try:
                    df_import = pd.read_excel(uploaded_file)
                    colonne_obbligatorie = ["GRUPPO_MACRO", "SOTTOGRUPPO", "ASSOCIATO_INSEGNA", "LIVELLO", "CHIAVE_LIVELLO"]
                    missing_cols = [c for c in colonne_obbligatorie if c not in df_import.columns]
                    
                    if missing_cols:
                        st.error(f"ROSSO (BLOCCATO) - Struttura Excel non valida. Colonne mancanti: {', '.join(missing_cols)}")
                    else:
                        cursor = conn.cursor()
                        righe_inserite = 0
                        
                        with conn:
                            for idx, row in df_import.iterrows():
                                gruppo = str(row["GRUPPO_MACRO"]).upper().strip()
                                sottogruppo = str(row["SOTTOGRUPPO"]).upper().strip() if (pd.notna(row.get("SOTTOGRUPPO")) and str(row.get("SOTTOGRUPPO")).strip() != "") else ""
                                insegna = str(row["ASSOCIATO_INSEGNA"]).upper().strip() if (pd.notna(row.get("ASSOCIATO_INSEGNA")) and str(row.get("ASSOCIATO_INSEGNA")).strip() != "") else ""
                                livello = str(row["LIVELLO"]).upper().strip()
                                chiave_livello = str(row["CHIAVE_LIVELLO"]).strip() if pd.notna(row["CHIAVE_LIVELLO"]) else ""
                                
                                if livello == "REFERENZA" and chiave_livello:
                                    chiave_livello = str(chiave_livello).split('.')[0].zfill(13)

                                cursor.execute("""
                                INSERT OR IGNORE INTO clienti (gruppo_macro, sottogruppo, associato_insegna)
                                VALUES (?, ?, ?)
                                """, (gruppo, sottogruppo, insegna))

                                def to_float_or_none(val):
                                    if pd.isna(val) or str(val).strip() == "":
                                        return None
                                    try: return float(val)
                                    except: return None

                                cursor.execute("""
                                INSERT OR REPLACE INTO accordi_commerciali (
                                    gruppo_macro, sottogruppo, associato_insegna, livello, chiave_livello, listino_r,
                                    sconto_1, sconto_2, sconto_3, sconto_4, sconto_5,
                                    sconto_6, sconto_7, sconto_y, sconto_carico, sconto_pagamento,
                                    voce_contratto_1, voce_contratto_2, voce_contratto_3, voce_contratto_4, voce_contratto_5
                                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                                """, (
                                    gruppo, sottogruppo, insegna, livello, chiave_livello, to_float_or_none(row.get("LISTINO_BASE_R")),
                                    to_float_or_none(row.get("SCONTO_1")), to_float_or_none(row.get("SCONTO_2")), to_float_or_none(row.get("SCONTO_3")),
                                    to_float_or_none(row.get("SCONTO_4")), to_float_or_none(row.get("SCONTO_5")), to_float_or_none(row.get("SCONTO_LOCAL_6")),
                                    to_float_or_none(row.get("SCONTO_LOCAL_7")), to_float_or_none(row.get("SCONTO_CONTINUATIVO_Y")),
                                    to_float_or_none(row.get("SCONTO_CARICO_LOGISTICA")),
                                    to_float_or_none(row.get("SCONTO_PAGAMENTO_AC")), to_float_or_none(row.get("PFA_VOCE_I")),
                                    to_float_or_none(row.get("PFA_VOCE_II")), to_float_or_none(row.get("PFA_VOCE_III")),
                                    to_float_or_none(row.get("PFA_VOCE_IV")), to_float_or_none(row.get("PFA_VOCE_V"))
                                ))
                                righe_inserite += 1
                        st.success(f"VERDE (APPROVATO) - Elaborate {righe_inserite} regole commerciali nel Bunker.")
                        st.rerun()
                except Exception as e:
                    st.error(f"ROSSO (BLOCCATO) - Errore durante l'elaborazione del file: {e}")

    st.markdown("---")
    st.markdown("<h3 style='color: #D32F2F;'>🚨 Sezione Pericolo (Danger Zone)</h3>", unsafe_allow_html=True)
    
    if PRODUCTION_MODE:
        st.info("Modalità Produzione: Ripristino demo disattivato.")
    else:
        st.warning("ATTENZIONE: Questa operazione ripristinerà il database allo stato iniziale.")
        pin_conferma = st.text_input("Per procedere digita la password di sicurezza 'RESET' in lettere maiuscole:")
        
        if st.button("ESEGUI HARD RESET DATABASE", disabled=(pin_conferma != "RESET")):
            try:
                seed_baseline_data(conn)
                st.success("VERDE (APPROVATO) - Database ripristinato allo stato iniziale.")
                st.rerun()
            except Exception as ex:
                st.error(f"ROSSO (BLOCCATO) - Errore durante il reset: {ex}")

    conn.close()

# ==========================================
# SCHEDA 3: REPORT SINTETICO
# ==========================================
elif menu == "Report Sintetico":
    st.title("Report Sintetico e Analisi Contratti")
    st.markdown("Analisi e raggruppamento delle metriche chiave degli accordi commerciali presenti in database.")
    st.markdown("---")
    
    conn = sqlite3.connect(DB_FILE)
    
    col_k1, col_k2, col_k3, col_k4 = st.columns(4)
    cursor = conn.cursor()
    
    cursor.execute("SELECT COUNT(*) FROM accordi_commerciali")
    tot_contratti = cursor.fetchone()[0]
    col_k1.metric("Totale Regole Attive", f"{tot_contratti}")
    
    cursor.execute("SELECT COUNT(*) FROM clienti WHERE attivo=1")
    tot_clienti = cursor.fetchone()[0]
    col_k2.metric("Insegne / Associati", f"{tot_clienti}")
    
    cursor.execute("SELECT AVG(listino_r) FROM accordi_commerciali WHERE listino_r IS NOT NULL")
    avg_listino = cursor.fetchone()[0] or 0.0
    col_k3.metric("Listino Medio R", f"{avg_listino:.2f} Euro")
    
    cursor.execute("""
        SELECT AVG(voce_contratto_1 + COALESCE(voce_contratto_2,0) + COALESCE(voce_contratto_3,0) + 
                   COALESCE(voce_contratto_4,0) + COALESCE(voce_contratto_5,0)) 
        FROM accordi_commerciali
    """)
    avg_pfa = cursor.fetchone()[0] or 0.0
    col_k4.metric("PFA Medio off-invoice", f"{avg_pfa:.2f} %")
    
    st.markdown("---")
    st.subheader("Sintesi Dinamica e Analisi per Canale GDO")
    
    query_sintesi = """
        SELECT gruppo_macro as [Gruppo Macro],
               COUNT(*) as [Totale Righe],
               ROUND(AVG(listino_r), 2) as [Listino Medio (Euro)],
               ROUND(AVG(sconto_1), 2) as [Sconto 1 Medio (%)],
               ROUND(AVG(sconto_2), 2) as [Sconto 2 Medio (%)],
               ROUND(AVG(sconto_carico), 2) as [Oneri Logistica (%)],
               ROUND(AVG(sconto_pagamento), 2) as [Oneri Pagamento (%)],
               ROUND(AVG(voce_contratto_1 + COALESCE(voce_contratto_2,0) + COALESCE(voce_contratto_3,0) + 
                         COALESCE(voce_contratto_4,0) + COALESCE(voce_contratto_5,0)), 2) as [PFA Totale Off-Invoice (%)]
        FROM accordi_commerciali
        GROUP BY gruppo_macro
        ORDER BY [Totale Righe] DESC
    """
    df_sintesi = pd.read_sql_query(query_sintesi, conn)
    st.dataframe(df_sintesi, use_container_width=True, hide_index=True)
    
    st.markdown("---")
    st.subheader("Generatore ed Esportazione Report Consolidato di Sintesi")
    st.markdown("Questo strumento permette di simulare ed esportare in blocco la verità contrattuale netta di tutti i 59 prodotti Salov contemporaneamente per un cliente specifico.")
    
    col_ex1, col_ex2 = st.columns(2)
    with col_ex1:
        cursor.execute("SELECT DISTINCT gruppo_macro FROM clienti WHERE attivo=1 ORDER BY gruppo_macro")
        gruppi_report = [r[0] for r in cursor.fetchall()]
        grp_rep_sel = st.selectbox("Seleziona Gruppo Macro per Esportazione", gruppi_report, key="rep_grp")
    with col_ex2:
        cursor.execute("SELECT DISTINCT associato_insegna FROM clienti WHERE gruppo_macro=? AND attivo=1 ORDER BY associato_insegna", (grp_rep_sel,))
        associati_report = [r[0] for r in cursor.fetchall()]
        ass_rep_sel = st.selectbox("Seleziona Insegna Locale per Esportazione", associati_report, key="rep_ass")
        
    if st.button("GENERA E COMPILA REPORT CONSOLIDATO EXCEL"):
        cursor.execute("SELECT sottogruppo FROM clienti WHERE gruppo_macro=? AND associato_insegna=? LIMIT 1", (grp_rep_sel, ass_rep_sel))
        res_sub = cursor.fetchone()
        sub_rep_sel = res_sub[0] if res_sub else ""
        
        cursor.execute("SELECT ean, descrizione_commerciale, tipo_olio, min_net_net_g, codice_sap, formato_lt, confezione FROM prodotti")
        all_prods = cursor.fetchall()
        
        rows_report = []
        for p in all_prods:
            p_ean, p_desc, p_tipo, p_min_g, p_sap, p_form, p_conf = p
            resolved = HierarchyResolver.resolve(conn, grp_rep_sel, sub_rep_sel, ass_rep_sel, p_ean, p_tipo)
            
            if resolved.listino_r is not None:
                input_calc = PricingInput(
                    listino_r=resolved.listino_r,
                    sconto_1=resolved.sconto_1, sconto_2=resolved.sconto_2, sconto_3=resolved.sconto_3,
                    sconto_4=resolved.sconto_4, sconto_5=resolved.sconto_5, sconto_6=resolved.sconto_6, sconto_7=resolved.sconto_7,
                    sconto_y=resolved.sconto_y, sconto_z=Decimal("0.00"), sconto_aa=Decimal("0.00"),
                    sconto_carico=resolved.sconto_carico, sconto_pagamento=resolved.sconto_pagamento,
                    voce_i=resolved.voce_i, voce_ii=resolved.voce_ii, voce_iii=resolved.voce_iii, voce_iv=resolved.voce_iv, voce_v=resolved.voce_v,
                    min_net_net_g=Decimal(str(p_min_g))
                )
                res_calc = PricingEngine.calculate(input_calc)
                
                rows_report.append({
                    "EAN": p_ean,
                    "Codice SAP": p_sap,
                    "Descrizione Commerciale": p_desc,
                    "Formato Lt": p_form,
                    "Confezione": p_conf,
                    "Listino Base R (Euro)": float(resolved.listino_r),
                    "Sconto 1 (%)": float(resolved.sconto_1),
                    "Sconto 2 (%)": float(resolved.sconto_2),
                    "Sconto Local 6 (%)": float(resolved.sconto_6),
                    "Oneri Logistica (%)": float(resolved.sconto_carico),
                    "Oneri Pagamento (%)": float(resolved.sconto_pagamento),
                    "Prezzo Netto AF (Euro)": float(res_calc.netto_in_fattura_2),
                    "Premi Off-Invoice AL (%)": float(res_calc.contratto_tot_pfa),
                    "Prezzo Net Net AM (Euro)": float(res_calc.net_net_finale),
                    "Soglia Sicurezza G (Euro)": float(p_min_g),
                    "Delta Margine (Euro)": float(res_calc.delta_vs_min),
                    "Stato Approvazione": "VERDE" if res_calc.guardrail_ok else "ROSSO"
                })
        
        if not rows_report:
            st.warning("ATTENZIONE: Nessuna referenza in assortimento trovata per questo cliente nel database.")
        else:
            df_rep_out = pd.DataFrame(rows_report)
            
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Consolidato_Marginalita"
            ws.views.sheetView[0].showGridLines = True
            
            font_title = Font(name="Arial", size=14, bold=True, color="FFFFFF")
            font_header = Font(name="Arial", size=10, bold=True, color="FFFFFF")
            font_data = Font(name="Arial", size=9)
            font_alert = Font(name="Arial", size=9, bold=True, color="9C0006")
            font_ok = Font(name="Arial", size=9, bold=True, color="006100")
            
            fill_title = PatternFill(start_color="1B5E20", end_color="1B5E20", fill_type="solid")
            fill_header = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
            fill_alert = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            fill_ok = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            
            thin_border = Border(
                left=Side(style='thin', color='E0E0E0'),
                right=Side(style='thin', color='E0E0E0'),
                top=Side(style='thin', color='E0E0E0'),
                bottom=Side(style='thin', color='E0E0E0')
            )
            
            ws.merge_cells('A1:Q1')
            ws['A1'] = f"SALOV S.p.A. - REPORT SINTETICO CONSOLIDATO: {ass_rep_sel} ({grp_rep_sel})"
            ws['A1'].font = font_title
            ws['A1'].fill = fill_title
            ws['A1'].alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[1].height = 35
            
            headers = list(df_rep_out.columns)
            for col_num, h_text in enumerate(headers, 1):
                cell = ws.cell(row=3, column=col_num, value=h_text)
                cell.font = font_header
                cell.fill = fill_header
                cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[3].height = 25
            
            for row_num, row_data in enumerate(df_rep_out.values, 4):
                ws.row_dimensions[row_num].height = 18
                for col_num, val in enumerate(row_data, 1):
                    cell = ws.cell(row=row_num, column=col_num, value=val)
                    cell.font = font_data
                    cell.border = thin_border
                    
                    if col_num in [1, 2]:
                        cell.alignment = Alignment(horizontal="center")
                        cell.number_format = '@'
                    elif col_num in [4, 6, 12, 14, 15, 16]:
                        cell.alignment = Alignment(horizontal="right")
                        cell.number_format = '#,##0.000 €'
                    elif col_num in [7, 8, 9, 10, 11, 13]:
                        cell.alignment = Alignment(horizontal="right")
                        cell.number_format = '0.00" %"'
                    
                    if col_num == 17:
                        cell.alignment = Alignment(horizontal="center")
                        if val == "VERDE":
                            cell.font = font_ok
                            cell.fill = fill_ok
                        else:
                            cell.font = font_alert
                            cell.fill = fill_alert
                            
            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = openpyxl.utils.get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 3, 11)
            
            buffer_rep = io.BytesIO()
            wb.save(buffer_rep)
            
            st.download_button(
                label=f"SCARICA REPORT EXCEL SINTESI CONTRATTO {ass_rep_sel}",
                data=buffer_rep.getvalue(),
                file_name=f"Sintesi_Contratto_{ass_rep_sel}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    conn.close()

# ==========================================
# SCHEDA 4: GUIDA OPERATIVA
# ==========================================
else:
    st.title("Guida Operativa - Bunker Commerciale Salov")
    st.markdown("Questa guida spiega in modo semplice, esaustivo e pratico come funziona il motore di calcolo, l'ereditarietà dei contratti e l'uso dell'applicazione.")
    st.markdown("---")
    
    with st.expander("1. Il Motore di Pricing (Cascata In Fattura e Fuori Fattura)", expanded=True):
        st.markdown("""
        Il simulatore applica gli sconti in modo **sequenziale geometrico (a cascata)** e non per somma algebrica, rispettando rigorosamente le prassi negoziali e contabili dell'ufficio commerciale Salov.
        
        **La Scomposizione Sequenziale dei Calcoli (Flusso On-Invoice ed Off-Invoice):**
        1. **Listino Base R (Euro/Pz):** Rappresenta il prezzo lordo base di partenza stabilito esclusivamente in fase di contrattazione con lo specifico cliente per la singola referenza. Non è un listino fisso o teorico, ma il vero punto di inizio del pricing.
        2. **Sconti Centrali / Canale Fisso (S1 - S5):** Sconti contrattuali strutturali concessi al cliente in base agli accordi annuali. Si applicano in cascata sequenziale sul prezzo ridotto precedente.
        3. **Sconti Territoriali Locali (S6 - S7):** Trattenute contrattuali locali destinate alla gestione periferica o territoriale degli associati.
        4. **Sconto Continuativo Y (%):** Leva commerciale continuativa definita dal venditore e approvata per coprire trattative trimestrali o semestrali di canale.
        5. **Sconto Promozionale Z (%):** Sconto percentuale temporaneo legato rigidamente alle finestre di Sell-In e Sell-Out promozionali (campagne volantino).
        6. **Sconto Taglio Prezzo Secco (AA):** Detrazione diretta espressa in Euro/Pezzo applicata subito dopo gli sconti percentuali.
        7. **Oneri di Rete Logistica (AB) & Pagamento (AC):** AB (Sconto Carico) è legato all'efficienza logistica dei volumi d'ordine (es. bilici o pallet completi). AC (Sconto Pagamento) è legato ai termini di pagamento concordati (es. cassa vista o anticipato). Generano il **Netto in Fattura 2 (AF)**.
        8. **Premi Fuori Fattura (AL):** Somma algebrica dei premi fine anno (PFA Voci I-V) applicata sul Netto in Fattura 2 per calcolare il **Prezzo Net Net Reale Aziendale (AM)**.
        """)
        
    with st.expander("2. La Gerarchia dei Contratti (Regole di Ereditarietà a 4 Livelli)", expanded=False):
        st.markdown("""
        Per evitare di dover compilare migliaia di righe per ogni cliente, l'applicazione implementa un sistema di **ereditarietà gerarchica a 4 livelli**:
        
        ```text
        [ LIVELLO 1: GRUPPO GDO (Macro) ] (es. COOP ITALIA)
                     │
                     ▼
        [ LIVELLO 2: SOTTOGRUPPO ] (es. COOP ITALIA SOTTOGRUPPO)
                     │
                     ▼
        [ LIVELLO 3: CATEGORIA MERCEOLOGICA ] (es. EXTRAVERGINE, OLIVA, SEMI, ACETO)
                     │
                     ▼
        [ LIVELLO 4: REFERENZA SPECIFICA (EAN) ] (es. Ex.v. Sagra Classico 1L)
        ```
        
        **Mappatura Aziendale degli Sconti per Livello:**
        Per mantenere ordine nel database, l'azienda utilizza questa convenzione rigorosa per i "cassetti" degli sconti:
        * **SCONTI DI GRUPPO:** Sconti 1, 2 e 3
        * **SCONTO DI SOTTOGRUPPO:** Sconti 4 e 5
        * **SCONTO DI CATEGORIA:** Sconto 6
        * **SCONTI DI REFERENZA:** Sconto 7

        **Regole Fondamentali di Ereditarietà:**
        * **Cella Vuota (Blank):** Nel database o nel file Excel di caricamento, se una cella di sconto è vuota, l'applicazione erediterà automaticamente il valore inserito a livello superiore. Se lo Sconto 1 è vuoto sulla singola Referenza ma è popolato al 10% sul Gruppo, l'app applicherà il 10%.
        * **Override Esplicito (Valore 0.0):** Se inserisci esplicitamente lo `0` (o `0.0`) a livello di Referenza o Categoria, l'app **annullerà e azzererà** lo sconto ereditato, consentendo di bloccare sconti centrali non dovuti su determinati prodotti.
        * **Filtro Assortimento:** Un prodotto viene considerato in assortimento ed è selezionabile solo se esiste un valore di **Listino R** configurato per quel cliente specifico a livello di Referenza (Livello 4). Se il listino è assente, l'app visualizzerà l'errore di blocco *"Prodotto fuori assortimento"*.
        """)

    with st.expander("3. Metodologie di Negoziazione (Target vs Spot)", expanded=False):
        st.markdown("""
        Il simulatore offre due modalità di lavoro per adattarsi a ogni fase della trattativa:
        
        **Metodo A: Partenza da Prezzo Target (Calcolo Inverso)**
        1. All'attivazione, l'app imposta come **Prezzo Target Net Net** la soglia minima di sicurezza **G** della referenza selezionata.
        2. Inserendo il prezzo desiderato dal buyer, il sistema calcola istantaneamente lo **Sconto Promozionale Z (%)** necessario a raggiungere esattamente quell'obiettivo.
        3. Se inserisci un valore nello **Sconto Unitario in fattura AA (Euro/Pz)**, il sistema adeguerà in tempo reale lo Sconto Promo Z (%) per mantenere il Net Net target stabile.
        
        **Metodo B: Tentativi Spot Manuali (Sconto Libero)**
        1. Sblocca l'inserimento manuale dello Sconto Promozionale Z.
        2. Permette di fare tentativi liberi per vedere dove atterra il Net Net.
        3. Mostra costantemente a fianco lo **Sconto Massimo Consentito (AV)**: la percentuale limite che puoi inserire prima che il semaforo diventi rosso.
        """)

    with st.expander("4. Uso del Back-Office ed Excel (Sincronizzazione)", expanded=False):
        st.markdown("""
        L'operatore di sede può gestire i listini e le condizioni in corsa in due modi:
        
        * **Modifica Diretta (A caldo):** Utilizza lo strumento di modifica diretta inserito nella scheda Back-Office. Fai doppio clic sulle celle per variare sconti o listini e primi il pulsante **Salva Modifiche** per sincronizzare istantaneamente l'app.
        * **Caricamento da file Excel:** Scarica il tracciato attuale degli accordi, compilalo localmente e trascinalo nel widget di importazione.
        
        **ATTENZIONE - REGOLA DI FERRO PER LA FORMATTAZIONE DEGLI EAN IN EXCEL:**
        Excel tende a corrompere e abbreviare i codici EAN a 13 cifre in notazione scientifica (es. `8.00E+12`). Prima di salvare il file `demo_seed_data.xlsx` o qualsiasi altro file per l'importazione, assicurati che la colonna `CHIAVE_LIVELLO` sia **esplicitamente impostata in formato Testo** all'interno di Excel.
        """)
