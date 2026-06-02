import streamlit as st
import sqlite3
import pandas as pd
from decimal import Decimal
from datetime import date
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode, GridUpdateMode, DataReturnMode
from config import DB_FILE

class RenewalEngine:
    @staticmethod
    def calc_net_net(listino, sconto_in_fattura, pfa):
        """Calcola il Net-Net simulato nel backend (Python) per il commit finale"""
        netto_fattura = listino * (1 - (sconto_in_fattura / 100))
        net_net = netto_fattura * (1 - (pfa / 100))
        return net_net

    @staticmethod
    def render_master_grid(df_active):
        """
        Motore V2: Rendering della Master Grid con calcoli front-end a latenza zero.
        Sostituisce il vecchio st.dataframe e st.data_editor per simulazioni real-time.
        """
        st.markdown("#### 🚀 Master Grid Interattiva (Motore AG Grid)")
        st.markdown("Modifica i Volumi e le leve [N+1]. La griglia calcolerà all'istante il Net-Net e lo Spazio Promo.")

        gb = GridOptionsBuilder.from_dataframe(df_active)

        # Raggruppamento gerarchico per la radiografia
        gb.configure_column("Categoria", rowGroup=True, hide=True)
        gb.configure_column("Sub-Categoria", rowGroup=True, hide=True)

        # Blocchiamo il nome prodotto a sinistra
        gb.configure_column("Prodotto", pinned='left', width=250)
        gb.configure_column("Floor Minimo €", type=["numericColumn"], valueFormatter="x.toFixed(3) + ' €'")

        # Colonne operative: permettiamo l'editing
        editable_cols = ['[N] Volumi', '[N] Listino €', '[N] Sc. Fattura %', '[N] PFA %',
                         '[N+1] Volumi', '[N+1] Listino €', '[N+1] Sc. Fattura %', '[N+1] PFA %']
        for col in editable_cols:
            gb.configure_column(
                col, 
                editable=True, 
                type=["numericColumn", "numberColumnFilter"], 
                # Colora leggermente le celle [N+1] per far capire dove agire
                cellStyle={'backgroundColor': '#F4F6F8', 'border': '1px solid #CBD5E1'} if '[N+1]' in col else {}
            )

        # INIEZIONE JS: Net-Net N+1
        net_net_jscode = JsCode('''
        function(params) {
            let data = params.data;
            if (!data) return 0;
            let listino = data['[N+1] Listino €'] || 0;
            let sc = data['[N+1] Sc. Fattura %'] || 0;
            let pfa = data['[N+1] PFA %'] || 0;
            let netto_fattura = listino * (1 - (sc / 100));
            return netto_fattura * (1 - (pfa / 100));
        }
        ''')

        # INIEZIONE JS: Spazio Promo
        spazio_promo_jscode = JsCode('''
        function(params) {
            let data = params.data;
            if (!data) return 0;
            let listino = data['[N+1] Listino €'] || 0;
            let sc = data['[N+1] Sc. Fattura %'] || 0;
            let pfa = data['[N+1] PFA %'] || 0;
            let netto_fattura = listino * (1 - (sc / 100));
            let net_net = netto_fattura * (1 - (pfa / 100));
            let floor = data['Floor Minimo €'] || 0;
            return net_net - floor;
        }
        ''')

        # INIEZIONE JS: Allarme Semaforico
        allarme_style_jscode = JsCode('''
        function(params) {
            if (params.value < 0) {
                return { 'backgroundColor': '#FEF2F2', 'color': '#991B1B', 'fontWeight': 'bold' };
            }
            return { 'backgroundColor': '#F0FDF4', 'color': '#166534' };
        }
        ''')

        # Applichiamo i JS
        gb.configure_column("Net_Net_N1", 
                            header_name="Net-Net [N+1] Sim.", 
                            valueGetter=net_net_jscode, 
                            type=["numericColumn"], 
                            valueFormatter="x.toFixed(3) + ' €'")
                            
        gb.configure_column("Spazio_Promo_€", 
                            header_name="Spazio Promo Residuo", 
                            valueGetter=spazio_promo_jscode, 
                            cellStyle=allarme_style_jscode, 
                            type=["numericColumn"], 
                            valueFormatter="x.toFixed(3) + ' €'")

        gb.configure_grid_options(
            enableRangeSelection=True,
            suppressAggFuncInHeader=True,
            groupDefaultExpanded=-1, # Espandi tutto di default
            animateRows=True,
            domLayout='normal'
        )

        gridOptions = gb.build()

        grid_response = AgGrid(
            df_active,
            gridOptions=gridOptions,
            update_mode=GridUpdateMode.MANUAL, # Aspetta il pulsante Conferma
            data_return_mode=DataReturnMode.AS_INPUT,
            allow_unsafe_jscode=True,
            theme='alpine',
            height=650,
            fit_columns_on_grid_load=True
        )

        return grid_response['data']
# modulo_rinnovi.py (PARTE 2: Setup Dati e Logica Originale)
def render_simulazione_rinnovi():
    st.title("🔄 Simulazione Rinnovi Contrattuali (N vs N+1)")
    st.markdown("Analisi differenziale dei margini, calcolo dello Spazio Promo e Roll-up per Sub-Categorie integrato con Master Grid.")
    
    anno_corrente = date.today().year
    
    with st.container(border=True):
        col_s1, col_s2 = st.columns([1, 3])
        with col_s1:
            anni_storico = st.slider("Anni di Storico da visualizzare", min_value=0, max_value=5, value=1, help="0 = Mostra solo N e N+1")
        with col_s2:
            st.info(f"**Anno N (Attuale):** {anno_corrente} | **Anno N+1 (Rinnovo):** {anno_corrente + 1}")

    # Estrazione DB come da tuo file originale
    conn = sqlite3.connect(DB_FILE)
    query = """
        SELECT a.ean, a.descrizione_commerciale, a.tipo_olio, COALESCE(g.min_net_net_g, 0.0) as min_net_net_g
        FROM anagrafica_master a
        LEFT JOIN guardrail_aziendali g ON a.ean = g.ean
    """
    df_base = pd.read_sql_query(query, conn)
    conn.close()
    
    # Classificazione Dinamica mantenuta intatta
    def get_subcat(row):
        desc = str(row['descrizione_commerciale']).upper()
        tipo = str(row['tipo_olio']).upper()
        if tipo == 'EXTRAVERGINE':
            if '100% ITA' in desc or '100%I' in desc or 'TOSC' in desc: return 'Extravergini Italiani'
            if 'BIO' in desc: return 'Extravergini Biologici'
            return 'Extravergini Comunitari'
        elif tipo == 'OLIVA': return 'Oli di Oliva Raffinati'
        elif tipo == 'SEMI':
            if 'ARACHIDE' in desc: return 'Semi di Arachide'
            if 'MAIS' in desc: return 'Semi di Mais'
            if 'GIRAS' in desc: return 'Semi di Girasole'
            if 'FRITT' in desc or 'FRIMX' in desc: return 'Oli per Frittura Specifici'
            if 'VINACC' in desc: return 'Semi di Vinacciolo'
            return 'Altri Oli di Semi'
        elif tipo == 'ACETO': return 'Aceto Balsamico'
        return 'Altro'
        
    df_base['Sub-Categoria'] = df_base.apply(get_subcat, axis=1)
    df_base['Categoria'] = df_base['tipo_olio'] # Aggiunto per il raggruppamento macro
    df_base = df_base.rename(columns={'descrizione_commerciale': 'Prodotto', 'min_net_net_g': 'Floor Minimo €'})
    
    storico_cols = []
    for i in range(anni_storico, 0, -1):
        col_name = f"[N-{i}] Net-Net €"
        df_base[col_name] = 0.0
        storico_cols.append(col_name)
        
    operative_cols = [
        '[N] Volumi', '[N] Listino €', '[N] Sc. Fattura %', '[N] PFA %',
        '[N+1] Volumi', '[N+1] Listino €', '[N+1] Sc. Fattura %', '[N+1] PFA %'
    ]
    for col in operative_cols:
        df_base[col] = 0.0 if '€' in col or '%' in col else 0
# modulo_rinnovi.py (PARTE 3: Tabs UI e Fusione AG Grid)
    # Riduciamo a 2 tab: Storico e la nuova Master Grid all-in-one
    tab_storico, tab_master_grid = st.tabs([
        "📅 1. Dati Storici", 
        "⚡ 2. Master Grid Simulazione (Real-Time)"
    ])
    
    with tab_storico:
        st.markdown("#### Inserimento Net-Net Storici (Opzionale)")
        st.markdown("Inserisci i valori net-net degli anni passati per avere un riferimento.")
        if anni_storico > 0:
            cols_to_show = ['Categoria', 'Sub-Categoria', 'Prodotto'] + storico_cols
            df_storico_edited = st.data_editor(
                df_base[cols_to_show],
                disabled=['Categoria', 'Sub-Categoria', 'Prodotto'],
                hide_index=True,
                use_container_width=True,
                key="editor_storico"
            )
            # Aggiorna il df_base con i dati storici
            for col in storico_cols:
                df_base[col] = df_storico_edited[col]
        else:
            st.info("Hai scelto di non visualizzare lo storico. Vai alla scheda Master Grid.")

    with tab_master_grid:
        st.markdown("#### Radiografia Cliente e Leve N+1")
        st.markdown("I calcoli avvengono nel tuo browser. Muovi le leve e controlla l'impatto sui margini. Quando hai chiuso la trattativa, premi Conferma.")
        
        cols_to_edit = ['Categoria', 'Sub-Categoria', 'Prodotto', 'Floor Minimo €'] + operative_cols
        df_for_grid = df_base[cols_to_edit]
        
        # Iniettiamo l'AG Grid al posto delle vecchie tabelle statiche
        df_sim_final = RenewalEngine.render_master_grid(df_for_grid)
        
        st.divider()
        
        # Pulsante di fuoco finale
        col_b1, col_b2 = st.columns([1, 4])
        with col_b1:
            if st.button("💾 Conferma e Salva Simulazione", type="primary", use_container_width=True):
                # Il DataFrame df_sim_final contiene ora le ultime modifiche fatte a video
                df_attivo = df_sim_final[df_sim_final['[N+1] Volumi'] > 0]
                
                if df_attivo.empty:
                    st.warning("Nessuna referenza con volumi inseriti.")
                else:
                    st.success("Simulazione validata! Pronta per il PricingEngine e il Database.")
                    # Qui attaccheremo il loop che chiama PricingEngine per l'inserimento
                    st.dataframe(df_attivo[['Prodotto', '[N+1] Volumi', '[N+1] Listino €', '[N+1] Sc. Fattura %', '[N+1] PFA %']])
