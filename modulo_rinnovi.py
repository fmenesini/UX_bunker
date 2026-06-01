import streamlit as st
import sqlite3
import pandas as pd
from decimal import Decimal
from config import DB_FILE

# ==========================================
# 1. DATABASE ISOLATO (Nessun impatto sul DB esistente)
# ==========================================
def init_modulo_rinnovi():
    """Crea tabelle isolate per la gerarchia a 3 livelli e i Floor di simulazione"""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    
    # Tabella di mappatura gerarchica estesa e Net_Net_Floor
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS ext_gerarchia_prodotti (
        ean TEXT PRIMARY KEY,
        macro_categoria TEXT,
        categoria TEXT,
        sub_categoria TEXT,
        net_net_floor REAL DEFAULT 0.0
    )""")
    
    # Inserimento dati di base (Seed) se la tabella è vuota
    cursor.execute("SELECT COUNT(*) FROM ext_gerarchia_prodotti")
    if cursor.fetchone()[0] == 0:
        seed_data = [
            ("8002210127425", "Olio di Oliva", "Extravergine", "Extravergini Italiani", 6.50),
            ("8002210111110", "Olio di Oliva", "Extravergine", "Extravergini Comunitari", 4.80),
            ("8002210001305", "Olio di Oliva", "Olio di Oliva", "Oli di Oliva Raffinati", 4.00),
            ("8002210000551", "Oli di Semi", "Arachide", "Semi di Arachide", 2.50),
            ("8002210111486", "Oli di Semi", "Mais", "Semi di Mais", 1.80),
            ("8002210111905", "Oli di Semi", "Girasole", "Semi di Girasole", 1.50),
            ("8002210111295", "Oli di Semi", "Frittura", "Oli per Frittura Specifici", 1.90)
        ]
        cursor.executemany("""
        INSERT INTO ext_gerarchia_prodotti (ean, macro_categoria, categoria, sub_categoria, net_net_floor)
        VALUES (?, ?, ?, ?, ?)
        """, seed_data)
        conn.commit()
    conn.close()

# ==========================================
# 2. MOTORE DI CALCOLO PONDERATO (Pandas)
# ==========================================
class RenewalEngine:
    @staticmethod
    def calculate_rollups(df_input):
        """
        Calcola i Net-Net e le medie ponderate per Referenza, Sub-Categoria e Totale.
        df_input deve contenere: ean, volumi, listino, sconto_perc, net_net_floor, sub_categoria
        """
        df = df_input.copy()
        
        # 1. Calcoli a livello di singola Referenza (SKU)
        df['Net_Net_Simulato'] = df['listino'] * (1 - (df['sconto_perc'] / 100))
        df['Fatturato_Simulato'] = df['Net_Net_Simulato'] * df['volumi']
        df['Valore_Floor_Totale'] = df['net_net_floor'] * df['volumi']
        df['Allarme_SKU'] = df['Net_Net_Simulato'] < df['net_net_floor']
        
        # 2. Roll-up a livello di Sub-Categoria (Media Ponderata sui Volumi)
        # Se i volumi sono 0, evitiamo divisioni per zero
        df_subcat = df.groupby('sub_categoria').agg(
            Volumi_Totali=('volumi', 'sum'),
            Fatturato_Totale=('Fatturato_Simulato', 'sum'),
            Valore_Floor_Totale=('Valore_Floor_Totale', 'sum')
        ).reset_index()
        
        df_subcat['Net_Net_Ponderato'] = df_subcat.apply(
            lambda x: x['Fatturato_Totale'] / x['Volumi_Totali'] if x['Volumi_Totali'] > 0 else 0, axis=1
        )
        df_subcat['Floor_Ponderato'] = df_subcat.apply(
            lambda x: x['Valore_Floor_Totale'] / x['Volumi_Totali'] if x['Volumi_Totali'] > 0 else 0, axis=1
        )
        df_subcat['Allarme_SubCat'] = df_subcat['Net_Net_Ponderato'] < df_subcat['Floor_Ponderato']
        
        # 3. Roll-up Totale Cliente
        tot_volumi = df['volumi'].sum()
        tot_fatturato = df['Fatturato_Simulato'].sum()
        tot_floor = df['Valore_Floor_Totale'].sum()
        
        totali = {
            'Volumi_Totali': tot_volumi,
            'Net_Net_Ponderato_Totale': tot_fatturato / tot_volumi if tot_volumi > 0 else 0,
            'Floor_Ponderato_Totale': tot_floor / tot_volumi if tot_volumi > 0 else 0
        }
        
        return df, df_subcat, totali

# ==========================================
# 3. INTERFACCIA UTENTE (UI) ISOLATA
# ==========================================
def render_simulazione_rinnovi():
    init_modulo_rinnovi()
    
    st.title("🔄 Simulazione Rinnovi Contrattuali (Analisi Sub-Categorie)")
    st.markdown("Questa sezione permette di simulare i rinnovi garantendo che le medie ponderate non nascondano perdite sui cluster premium.")
    
    conn = sqlite3.connect(DB_FILE)
    
    # Estrazione dati unendo l'anagrafica esistente con la nuova gerarchia isolata
    query = """
        SELECT a.ean, a.descrizione_commerciale, 
               COALESCE(e.macro_categoria, 'Non Assegnato') as macro_categoria,
               COALESCE(e.categoria, 'Non Assegnato') as categoria,
               COALESCE(e.sub_categoria, 'Non Assegnato') as sub_categoria,
               COALESCE(e.net_net_floor, 0.0) as net_net_floor
        FROM anagrafica_master a
        LEFT JOIN ext_gerarchia_prodotti e ON a.ean = e.ean
        WHERE e.sub_categoria IS NOT NULL
    """
    df_base = pd.read_sql_query(query, conn)
    conn.close()
    
    if df_base.empty:
        st.warning("Nessun prodotto mappato nella nuova gerarchia. Controllare il database.")
        return

    # Preparazione della griglia di input per l'utente
    df_base['volumi'] = 0
    df_base['listino'] = 0.0
    df_base['sconto_perc'] = 0.0
    
    st.subheader("1. Inserimento Dati di Simulazione")
    st.markdown("Inserisci i volumi previsti, il listino lordo e lo sconto totale stimato per le referenze in trattativa.")
    
    # Data Editor interattivo
    edited_df = st.data_editor(
        df_base[['sub_categoria', 'descrizione_commerciale', 'net_net_floor', 'volumi', 'listino', 'sconto_perc']],
        column_config={
            "sub_categoria": st.column_config.TextColumn("Sub-Categoria", disabled=True),
            "descrizione_commerciale": st.column_config.TextColumn("Prodotto", disabled=True),
            "net_net_floor": st.column_config.NumberColumn("Floor Minimo (€)", format="€ %.2f", disabled=True),
            "volumi": st.column_config.NumberColumn("Volumi Previsti (Pz)", min_value=0, step=100),
            "listino": st.column_config.NumberColumn("Listino Lordo (€)", format="€ %.2f", min_value=0.0, step=0.1),
            "sconto_perc": st.column_config.NumberColumn("Sconto Totale (%)", format="%.2f %%", min_value=0.0, max_value=100.0, step=0.5),
        },
        hide_index=True,
        use_container_width=True
    )
    
    if st.button("🚀 ESEGUI ROLL-UP E VERIFICA MARGINI", type="primary"):
        # Filtriamo solo le righe dove l'utente ha inserito dei volumi
        df_active = edited_df[edited_df['volumi'] > 0].copy()
        
        if df_active.empty:
            st.error("Inserisci dei volumi per almeno una referenza per avviare la simulazione.")
            return
            
        # Esecuzione Motore
        df_sku, df_subcat, totali = RenewalEngine.calculate_rollups(df_active)
        
        st.divider()
        st.subheader("2. Risultati Aggregati per Sub-Categoria")
        
        # Funzione di styling per evidenziare in rosso le perdite
        def highlight_alarms(row):
            if row['Allarme_SubCat']:
                return ['background-color: #FEF2F2; color: #991B1B; font-weight: bold'] * len(row)
            return [''] * len(row)

        # Formattazione tabella Sub-Categorie
        df_subcat_display = df_subcat[['sub_categoria', 'Volumi_Totali', 'Floor_Ponderato', 'Net_Net_Ponderato', 'Allarme_SubCat']]
        
        st.dataframe(
            df_subcat_display.style.apply(highlight_alarms, axis=1).format({
                'Floor_Ponderato': '€ {:.2f}',
                'Net_Net_Ponderato': '€ {:.2f}'
            }),
            column_config={
                "sub_categoria": "Cluster (Sub-Categoria)",
                "Volumi_Totali": "Volumi Totali",
                "Floor_Ponderato": "Floor Ponderato (€)",
                "Net_Net_Ponderato": "Net-Net Ponderato (€)",
                "Allarme_SubCat": "Allarme Margine"
            },
            hide_index=True,
            use_container_width=True
        )
        
        # Controllo Allarmi
        if df_subcat['Allarme_SubCat'].any():
            st.error("⚠️ ATTENZIONE: Una o più Sub-Categorie stanno distruggendo valore! Il Net-Net Ponderato è inferiore al Floor minimo consentito per quel cluster.")
        else:
            st.success("✅ OTTIMO: Tutte le Sub-Categorie rispettano i margini minimi aziendali.")
            
        st.divider()
        
        # KPI Totali Cliente
        st.subheader("3. KPI Totali Cliente (Media Ponderata Globale)")
        col1, col2, col3 = st.columns(3)
        col1.metric("Volumi Totali Rinnovo", f"{totali['Volumi_Totali']:,.0f} Pz")
        col2.metric("Floor Ponderato Globale", f"€ {totali['Floor_Ponderato_Totale']:.2f}")
        
        delta = totali['Net_Net_Ponderato_Totale'] - totali['Floor_Ponderato_Totale']
        col3.metric("Net-Net Ponderato Globale", f"€ {totali['Net_Net_Ponderato_Totale']:.2f}", f"{delta:.2f} € vs Floor")
