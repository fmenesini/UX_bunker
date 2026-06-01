import streamlit as st
import sqlite3
import pandas as pd
from decimal import Decimal
from datetime import date
from config import DB_FILE

# ==========================================
# 1. DATABASE ISOLATO
# ==========================================
def init_modulo_rinnovi():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS ext_gerarchia_prodotti (
        ean TEXT PRIMARY KEY,
        macro_categoria TEXT,
        categoria TEXT,
        sub_categoria TEXT,
        net_net_floor REAL DEFAULT 0.0
    )""")
    
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
# 2. MOTORE DI CALCOLO DIFFERENZIALE
# ==========================================
class RenewalEngine:
    @staticmethod
    def calc_net_net(listino, sconto_in_fattura, pfa):
        """Calcola il Net-Net simulato partendo da Listino, Sconti in fattura e PFA"""
        netto_fattura = listino * (1 - (sconto_in_fattura / 100))
        net_net = netto_fattura * (1 - (pfa / 100))
        return net_net

    @staticmethod
    def calculate_rollups(df_input):
        df = df_input.copy()
        
        # Calcolo Net-Net Anno N
        df['Net_Net_N'] = df.apply(lambda x: RenewalEngine.calc_net_net(x['[N] Listino €'], x['[N] Sc. Fattura %'], x['[N] PFA %']), axis=1)
        df['Fatturato_N'] = df['Net_Net_N'] * df['[N] Volumi']
        
        # Calcolo Net-Net Anno N+1
        df['Net_Net_N1'] = df.apply(lambda x: RenewalEngine.calc_net_net(x['[N+1] Listino €'], x['[N+1] Sc. Fattura %'], x['[N+1] PFA %']), axis=1)
        df['Fatturato_N1'] = df['Net_Net_N1'] * df['[N+1] Volumi']
        
        # Calcolo Delta e Spazio Promo
        df['Delta_Net_Net_€'] = df['Net_Net_N1'] - df['Net_Net_N']
        df['Spazio_Promo_€'] = df['Net_Net_N1'] - df['Floor Minimo €']
        df['Allarme_SKU'] = df['Net_Net_N1'] < df['Floor Minimo €']
        
        # Valore Floor Totale per N+1
        df['Valore_Floor_Totale_N1'] = df['Floor Minimo €'] * df['[N+1] Volumi']
        
        # Roll-up per Sub-Categoria (Media Ponderata)
        df_subcat = df.groupby('Sub-Categoria').agg(
            Volumi_N=('[N] Volumi', 'sum'),
            Fatturato_N=('Fatturato_N', 'sum'),
            Volumi_N1=('[N+1] Volumi', 'sum'),
            Fatturato_N1=('Fatturato_N1', 'sum'),
            Valore_Floor_Totale_N1=('Valore_Floor_Totale_N1', 'sum')
        ).reset_index()
        
        df_subcat['Net_Net_Pond_N'] = df_subcat.apply(lambda x: x['Fatturato_N'] / x['Volumi_N'] if x['Volumi_N'] > 0 else 0, axis=1)
        df_subcat['Net_Net_Pond_N1'] = df_subcat.apply(lambda x: x['Fatturato_N1'] / x['Volumi_N1'] if x['Volumi_N1'] > 0 else 0, axis=1)
        df_subcat['Floor_Pond_N1'] = df_subcat.apply(lambda x: x['Valore_Floor_Totale_N1'] / x['Volumi_N1'] if x['Volumi_N1'] > 0 else 0, axis=1)
        
        df_subcat['Delta_Pond_€'] = df_subcat['Net_Net_Pond_N1'] - df_subcat['Net_Net_Pond_N']
        df_subcat['Spazio_Promo_Pond_€'] = df_subcat['Net_Net_Pond_N1'] - df_subcat['Floor_Pond_N1']
        df_subcat['Allarme_SubCat'] = df_subcat['Net_Net_Pond_N1'] < df_subcat['Floor_Pond_N1']
        
        # Totali Globali
        tot_vol_n1 = df['[N+1] Volumi'].sum()
        totali = {
            'Volumi_N1': tot_vol_n1,
            'Net_Net_Pond_N': df['Fatturato_N'].sum() / df['[N] Volumi'].sum() if df['[N] Volumi'].sum() > 0 else 0,
            'Net_Net_Pond_N1': df['Fatturato_N1'].sum() / tot_vol_n1 if tot_vol_n1 > 0 else 0,
            'Floor_Pond_N1': df['Valore_Floor_Totale_N1'].sum() / tot_vol_n1 if tot_vol_n1 > 0 else 0
        }
        
        return df, df_subcat, totali

# ==========================================
# 3. INTERFACCIA UTENTE (UI)
# ==========================================
def render_simulazione_rinnovi():
    init_modulo_rinnovi()
    
    st.title("🔄 Simulazione Rinnovi Contrattuali (N vs N+1)")
    st.markdown("Analisi differenziale dei margini, calcolo dello Spazio Promo e Roll-up per Sub-Categorie.")
    
    anno_corrente = date.today().year
    
    # --- SETUP INIZIALE ---
    with st.container(border=True):
        col_s1, col_s2 = st.columns([1, 3])
        with col_s1:
            anni_storico = st.slider("Anni di Storico da visualizzare", min_value=0, max_value=5, value=1, help="0 = Mostra solo N e N+1")
        with col_s2:
            st.info(f"**Anno N (Attuale):** {anno_corrente} | **Anno N+1 (Rinnovo):** {anno_corrente + 1}")

    conn = sqlite3.connect(DB_FILE)
    query = """
        SELECT a.ean, a.descrizione_commerciale, 
               COALESCE(e.sub_categoria, 'Non Assegnato') as sub_categoria,
               COALESCE(e.net_net_floor, 0.0) as net_net_floor
        FROM anagrafica_master a
        LEFT JOIN ext_gerarchia_prodotti e ON a.ean = e.ean
        WHERE e.sub_categoria IS NOT NULL
    """
    df_base = pd.read_sql_query(query, conn)
    conn.close()
    
    if df_base.empty:
        st.warning("Nessun prodotto mappato nella gerarchia. Controllare il database.")
        return

    # Preparazione Colonne DataFrame
    df_base = df_base.rename(columns={
        'sub_categoria': 'Sub-Categoria',
        'descrizione_commerciale': 'Prodotto',
        'net_net_floor': 'Floor Minimo €'
    })
    
    # Colonne Storico
    storico_cols = []
    for i in range(anni_storico, 0, -1):
        col_name = f"[N-{i}] Net-Net €"
        df_base[col_name] = 0.0
        storico_cols.append(col_name)
        
    # Colonne Anno N e N+1
    operative_cols = [
        '[N] Volumi', '[N] Listino €', '[N] Sc. Fattura %', '[N] PFA %',
        '[N+1] Volumi', '[N+1] Listino €', '[N+1] Sc. Fattura %', '[N+1] PFA %'
    ]
    for col in operative_cols:
        df_base[col] = 0.0 if '€' in col or '%' in col else 0

    # --- TABS DI NAVIGAZIONE ---
    tab_storico, tab_simulazione, tab_risultati = st.tabs([
        "📅 1. Dati Storici", 
        "⚙️ 2. Simulazione N vs N+1", 
        "📊 3. Analisi & Spazio Promo"
    ])
    
    # TAB 1: STORICO
    with tab_storico:
        st.markdown("#### Inserimento Net-Net Storici (Opzionale)")
        st.markdown("Inserisci i valori net-net degli anni passati per avere un riferimento durante la trattativa.")
        if anni_storico > 0:
            cols_to_show = ['Sub-Categoria', 'Prodotto'] + storico_cols
            df_storico_edited = st.data_editor(
                df_base[cols_to_show],
                disabled=['Sub-Categoria', 'Prodotto'],
                hide_index=True,
                use_container_width=True,
                key="editor_storico"
            )
            # Aggiorniamo il df_base con i dati storici inseriti
            for col in storico_cols:
                df_base[col] = df_storico_edited[col]
        else:
            st.info("Hai scelto di non visualizzare lo storico. Vai alla scheda Simulazione.")

    # TAB 2: SIMULAZIONE
    with tab_simulazione:
        st.markdown("#### Griglia di Simulazione Contrattuale")
        st.markdown("Inserisci i dati per l'anno in corso [N] e per il rinnovo [N+1]. **Verranno analizzate solo le referenze con Volumi [N+1] > 0.**")
        
        # Configurazione colonne per l'editor
        col_config = {
            "Sub-Categoria": st.column_config.TextColumn(disabled=True),
            "Prodotto": st.column_config.TextColumn(disabled=True),
            "Floor Minimo €": st.column_config.NumberColumn(format="€ %.2f", disabled=True),
            "[N] Volumi": st.column_config.NumberColumn(step=100),
            "[N] Listino €": st.column_config.NumberColumn(format="€ %.2f", step=0.1),
            "[N] Sc. Fattura %": st.column_config.NumberColumn(format="%.2f %%", step=0.5),
            "[N] PFA %": st.column_config.NumberColumn(format="%.2f %%", step=0.5),
            "[N+1] Volumi": st.column_config.NumberColumn(step=100),
            "[N+1] Listino €": st.column_config.NumberColumn(format="€ %.2f", step=0.1),
            "[N+1] Sc. Fattura %": st.column_config.NumberColumn(format="%.2f %%", step=0.5),
            "[N+1] PFA %": st.column_config.NumberColumn(format="%.2f %%", step=0.5),
        }
        
        cols_to_edit = ['Sub-Categoria', 'Prodotto', 'Floor Minimo €'] + operative_cols
        
        df_sim_edited = st.data_editor(
            df_base[cols_to_edit],
            column_config=col_config,
            hide_index=True,
            use_container_width=True,
            height=500,
            key="editor_simulazione"
        )

    # TAB 3: RISULTATI E ROLL-UP
    with tab_risultati:
        # Filtriamo solo le referenze attive (Volumi N+1 > 0)
        df_active = df_sim_edited[df_sim_edited['[N+1] Volumi'] > 0].copy()
        
        if df_active.empty:
            st.warning("Nessuna referenza attiva. Inserisci dei volumi nella colonna '[N+1] Volumi' nella scheda precedente.")
        else:
            # Calcolo Motore
            df_res_sku, df_res_subcat, totali = RenewalEngine.calculate_rollups(df_active)
            
            # --- KPI GLOBALI ---
            st.markdown("#### KPI Totali Cliente (Media Ponderata)")
            col_k1, col_k2, col_k3, col_k4 = st.columns(4)
            col_k1.metric("Volumi Totali [N+1]", f"{totali['Volumi_N1']:,.0f} Pz")
            col_k2.metric("Net-Net Pond. [N]", f"€ {totali['Net_Net_Pond_N']:.3f}")
            
            delta_tot = totali['Net_Net_Pond_N1'] - totali['Net_Net_Pond_N']
            col_k3.metric("Net-Net Pond. [N+1]", f"€ {totali['Net_Net_Pond_N1']:.3f}", f"{delta_tot:+.3f} € vs [N]")
            
            spazio_tot = totali['Net_Net_Pond_N1'] - totali['Floor_Pond_N1']
            col_k4.metric("Spazio Promo Globale", f"€ {spazio_tot:+.3f}", "Margine residuo medio")
            
            st.divider()
            
            # --- ROLL-UP SUB-CATEGORIE ---
            st.markdown("#### 1. Analisi Aggregata per Sub-Categoria")
            st.markdown("Verifica se i cluster premium stanno perdendo margine rispetto all'anno precedente o rispetto al Floor.")
            
            def highlight_subcat(row):
                if row['Allarme_SubCat']: return ['background-color: #FEF2F2; color: #991B1B; font-weight: bold'] * len(row)
                if row['Delta_Pond_€'] < 0: return ['color: #D97706'] * len(row) # Arancione se perdo vs N ma sono sopra floor
                return [''] * len(row)

            df_subcat_disp = df_res_subcat[['Sub-Categoria', 'Volumi_N1', 'Net_Net_Pond_N', 'Net_Net_Pond_N1', 'Delta_Pond_€', 'Floor_Pond_N1', 'Spazio_Promo_Pond_€', 'Allarme_SubCat']]
            
            st.dataframe(
                df_subcat_disp.style.apply(highlight_subcat, axis=1).format({
                    'Net_Net_Pond_N': '€ {:.3f}', 'Net_Net_Pond_N1': '€ {:.3f}', 
                    'Delta_Pond_€': '€ {:+.3f}', 'Floor_Pond_N1': '€ {:.3f}', 'Spazio_Promo_Pond_€': '€ {:+.3f}'
                }),
                column_config={"Allarme_SubCat": "Sotto Floor!"},
                hide_index=True, use_container_width=True
            )
            
            st.divider()
            
            # --- DETTAGLIO REFERENZE (SKU) ---
            st.markdown("#### 2. Dettaglio Referenze (SKU) e Spazio Promo")
            st.markdown("Esplosione del Net-Net per singola referenza. La colonna **Spazio Promo (€)** indica quanto margine unitario hai a disposizione per finanziare volantini o tagli prezzo prima di andare in perdita.")
            
            def highlight_sku(row):
                if row['Allarme_SKU']: return ['background-color: #FEF2F2; color: #991B1B'] * len(row)
                return [''] * len(row)

            cols_sku_disp = ['Sub-Categoria', 'Prodotto', 'Net_Net_N', 'Net_Net_N1', 'Delta_Net_Net_€', 'Floor Minimo €', 'Spazio_Promo_€', 'Allarme_SKU']
            
            st.dataframe(
                df_res_sku[cols_sku_disp].style.apply(highlight_sku, axis=1).format({
                    'Net_Net_N': '€ {:.3f}', 'Net_Net_N1': '€ {:.3f}', 
                    'Delta_Net_Net_€': '€ {:+.3f}', 'Floor Minimo €': '€ {:.3f}', 'Spazio_Promo_€': '€ {:+.3f}'
                }),
                column_config={"Allarme_SKU": "Sotto Floor!"},
                hide_index=True, use_container_width=True
            )
