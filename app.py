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
    
    /* Titoli in Verde Sagra */
    h1, h2, h3, h4, h5, h6 { color: #1B5E20 !important; font-weight: bold !important; }
    
    /* Metriche in Rosso Sagra */
    div[data-testid="stMetricValue"] { color: #D32F2F !important; font-weight: bold !important; }
    
    /* Expander puliti */
    div[data-testid="stExpander"] { background-color: #FFFFFF !important; border: 1px solid #D1C9BC !important; border-radius: 4px !important; }
    
    /* Box di allerta */
    .warning-box { background-color: #FFF3E0 !important; border-left: 5px solid #FF9800 !important; padding: 12px; border-radius: 4px; margin-bottom: 15px; color: #B78103 !important; font-weight: bold; }
    .info-box { background-color: #E8F5E9 !important; border-left: 5px solid #2E7D32 !important; padding: 12px; border-radius: 4px; margin-bottom: 15px; color: #1B5E20 !important; font-weight: bold; }
    
    /* Testi generici forzati scuri per leggibilità, MA ESCLUDENDO i tag span per salvare le icone Streamlit */
    .stMarkdown p, .stMarkdown li, label { color: #1C1C1C !important; font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif !important; font-size: 1.02rem !important; }
</style>
""", unsafe_allow_html=True)

def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT attivo FROM clienti LIMIT 1")
        cursor.execute("SELECT sconto_y FROM accordi_commerciali LIMIT 1") # Test per la nuova colonna Y
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
    
    # Aggiunto sconto_y nello schema
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
        
    # --- NUOVA MAPPATURA SCONTI E LISTINI RICALCOLATI (-60% COOP, -20% CONAD) ---
    # S1-S3 = Gruppo | S4-S5 = Sottogruppo | S6 = Categoria | S7 = Referenza | SY = Referenza
    fallback_data = [
        # COOP ITALIA (ALLEANZA 3.0)
        ('COOP ITALIA', '', '', 'GRUPPO', '', None, 10.0, 5.0, None, None, None, None, None, None, 1.5, 1.0, 5.0, 2.0, None, None, None),
        ('COOP ITALIA', 'COOP ITALIA SOTTOGRUPPO', '', 'SOTTOGRUPPO', '', None, None, None, None, 2.0, None, None, None, None, None, None, None, None, None, None, None),
        ('COOP ITALIA', 'COOP ITALIA SOTTOGRUPPO', 'ALLEANZA 3.0', 'CATEGORIA', 'EXTRAVERGINE', None, None, None, None, None, None, 3.0, None, None, None, None, None, None, None, None, 1.0),
        ('COOP ITALIA', 'COOP ITALIA SOTTOGRUPPO', 'ALLEANZA 3.0', 'REFERENZA', '8002210131620', 66.00, None, None, None, None, None, None, 12.0, 5.0, None, None, None, None, None, None, None),
        ('COOP ITALIA', 'COOP ITALIA SOTTOGRUPPO', 'ALLEANZA 3.0', 'REFERENZA', '8002210111110', 60.80, None, None, None, None, None, None, 15.0, 0.0, None, None, None, None, None, None, None),
        ('COOP ITALIA', 'COOP ITALIA SOTTOGRUPPO', 'ALLEANZA 3.0', 'REFERENZA', '8002210001305', 43.20, None, None, None, None, None, None, 12.0, 0.0, None, None, None, None, None, None, None),
        
        # CONAD ADRIATICO (Listini ulteriormente deflazionati)
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

        cursor.execute("SELECT ean, descrizione_commercial
