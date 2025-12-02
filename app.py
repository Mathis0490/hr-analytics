import streamlit as st
import pandas as pd
from datetime import datetime
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
import zipfile
import xlsxwriter

st.set_page_config(page_title="HR Analytics", page_icon="📊", layout="wide")

st.title("🚀 HR Analytics Tool")
st.markdown("**Excel hochladen → Diagramme bekommen!**")

# Tabs
tab1, tab2 = st.tabs(["📥 1. Vorlage holen", "📊 2. Analyse starten"])

with tab1:
    st.markdown("### Leere Excel-Vorlage herunterladen")
    st.markdown("Alles ist optional - füllen Sie nur aus was Sie haben!")
    
    # Vorlage erstellen
    spalten = ['Mitarbeiter_ID','Geburtsjahr','Eintrittsjahr','Geschlecht','Abteilung','Karrierelevel','Gehalt_Brutto_Jahr','Arbeitszeit']
    df_vorlage = pd.DataFrame(columns=spalten, index=range(100))
    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_vorlage.to_excel(writer, index=False, sheet_name='Daten')
    
    st.download_button(
        label="📥 Vorlage herunterladen",
        data=buffer.getvalue(),
        file_name="HR_Vorlage.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

with tab2:
    st.markdown("### Excel hochladen und analysieren")
    
    col1, col2 = st.columns([1, 1])
    
    with col1:
        uploaded_file = st.file_uploader("Excel hochladen", type=['xlsx', 'xls'])
        rentenalter = st.slider("Rentenalter", 60, 70, 67)
        
        analyse_button = st.button("🚀 ANALYSE STARTEN", type="primary")
    
    with col2:
        region = st.selectbox("Region für Benchmark", 
            ["Niedersachsen", "NRW", "Bayern", "Hessen", "Berlin", "Hamburg", "Deutschland"])
    
    if analyse_button and uploaded_file:
        with st.spinner("Analysiere..."):
            try:
                df = pd.read_excel(uploaded_file, engine='openpyxl')
                df = df.dropna(how='all')
                
                if len(df) == 0:
                    st.error("Excel ist leer!")
                else:
                    # Spalten finden
                    def find(names):
                        for c in df.columns:
                            for n in names:
                                if n in str(c).lower():
                                    return c
                        return None
                    
                    col_geb = find(['geburtsjahr','jahrgang'])
                    col_ein = find(['eintrittsjahr','eintritt'])
                    col_ges = find(['geschlecht'])
                    col_abt = find(['abteilung'])
                    col_lvl = find(['karrierelevel','level'])
                    
                    jahr = datetime.now().year
                    
                    if col_geb:
                        df['Alter'] = jahr - pd.to_numeric(df[col_geb], errors='coerce')
                    if col_ein:
                        df['DJ'] = jahr - pd.to_numeric(df[col_ein], errors='coerce')
                    
                    F = ['#3498db','#e74c3c','#2ecc71','#9b59b6','#f39c12','#1abc9c']
                    
                    st.success(f"✅ {len(df)} Mitarbeiter geladen!")
                    
                    # Diagramme in Spalten
                    charts_html = []
                    
                    # RENTE
                    if 'Alter' in df.columns:
                        d = df[df['Alter'].notna()].copy()
                        if len(d) > 0:
                            d['JbR'] = rentenalter - d['Alter']
                            d['Kat'] = pd.cut(d['JbR'], [-100,0,5,10,15,20,100], 
                                labels=['Rente','0-5J','5-10J','10-15J','15-20J','20+J'])
                            
                            st.markdown("### 🎯 Rentenanalyse")
                            c1, c2 = st.columns(2)
                            
                            with c1:
                                k = d['Kat'].value_counts()
                                fig = go.Figure(go.Pie(
                                    labels=[str(x) for x in k.index], 
                                    values=list(k.values),
                                    marker_colors=['#c0392b','#e74c3c','#f39c12','#f1c40f','#2ecc71','#27ae60']
                                ))
                                fig.update_layout(title="Jahre bis Rente", height=400)
                                st.plotly_chart(fig, use_container_width=True)
                                charts_html.append(('01_Rente_Kategorien.html', fig.to_html()))
                            
                            with c2:
                                fig = go.Figure(go.Histogram(x=list(d['Alter']), nbinsx=15, marker_color='#3498db'))
                                fig.update_layout(title="Altersverteilung", height=400)
                                st.plotly_chart(fig, use_container_width=True)
                                charts_html.append(('02_Altersverteilung.html', fig.to_html()))
                            
                            # Warnung
                            r5 = len(d[d['JbR'] <= 5])
                            if r5 > 0:
                                st.warning(f"⚠️ {r5} Mitarbeiter ({round(r5/len(d)*100,1)}%) gehen in 5 Jahren in Rente!")
                    
                    # TREUE
                    if 'DJ' in df.columns:
                        d = df[df['DJ'].notna()].copy()
                        if len(d) > 0:
                            st.markdown("### 🏆 Treue-Analyse")
                            c1, c2 = st.columns(2)
                            
                            with c1:
                                fig = go.Figure(go.Histogram(x=list(d['DJ']), nbinsx=15, marker_color='#9b59b6'))
                                fig.update_layout(title="Dienstjahre", height=400)
                                st.plotly_chart(fig, use_container_width=True)
                                charts_html.append(('03_Dienstjahre.html', fig.to_html()))
                            
                            with c2:
                                jubi = [int(len(d[(d['DJ']>=j-0.5)&(d['DJ']<j+0.5)])) for j in [5,10,15,20,25]]
                                fig = go.Figure(go.Bar(x=['5J','10J','15J','20J','25J'], y=jubi, marker_color=F[:5]))
                                fig.update_layout(title="Jubiläen", height=400)
                                st.plotly_chart(fig, use_container_width=True)
                                charts_html.append(('04_Jubilaeen.html', fig.to_html()))
                            
                            st.info(f"📊 Durchschnitt: {d['DJ'].mean():.1f} Jahre Betriebszugehörigkeit")
                    
                    # WISSEN
                    if 'Alter' in df.columns and 'DJ' in df.columns:
                        d = df[(df['Alter'].notna())&(df['DJ'].notna())].copy()
                        if len(d) > 0:
                            d['JbR'] = rentenalter - d['Alter']
                            d['R'] = d.apply(lambda r:'KRITISCH' if r['DJ']>=15 and r['JbR']<=5 else 
                                ('WARNUNG' if r['DJ']>=10 and r['JbR']<=10 else 'OK'), axis=1)
                            
                            st.markdown("### ⚠️ Wissensverlust-Prognose")
                            
                            fig = go.Figure()
                            cm = {'KRITISCH':'#e74c3c','WARNUNG':'#f39c12','OK':'#2ecc71'}
                            for r in ['OK','WARNUNG','KRITISCH']:
                                x = d[d['R']==r]
                                if len(x)>0:
                                    fig.add_trace(go.Scatter(
                                        x=list(x['JbR']), y=list(x['DJ']),
                                        mode='markers', name=r,
                                        marker=dict(color=cm[r], size=12)
                                    ))
                            fig.update_layout(
                                title="Erfahrung vs. Jahre bis Rente",
                                xaxis_title="Jahre bis Rente",
                                yaxis_title="Dienstjahre",
                                height=500
                            )
                            st.plotly_chart(fig, use_container_width=True)
                            charts_html.append(('05_Wissensverlust.html', fig.to_html()))
                            
                            krit = len(d[d['R']=='KRITISCH'])
                            warn = len(d[d['R']=='WARNUNG'])
                            if krit > 0:
                                st.error(f"🔴 {krit} KRITISCHE Mitarbeiter (viel Erfahrung + bald Rente)")
                            if warn > 0:
                                st.warning(f"🟡 {warn} Mitarbeiter mit Warnung")
                    
                    # GESCHLECHT & ABTEILUNG
                    st.markdown("### 📈 Weitere Analysen")
                    c1, c2 = st.columns(2)
                    
                    with c1:
                        if col_ges:
                            c = df[col_ges].value_counts()
                            fig = go.Figure(go.Pie(labels=[str(x) for x in c.index], values=list(c.values), marker_colors=F))
                            fig.update_layout(title="Geschlecht", height=400)
                            st.plotly_chart(fig, use_container_width=True)
                            charts_html.append(('06_Geschlecht.html', fig.to_html()))
                    
                    with c2:
                        if col_abt:
                            c = df[col_abt].value_counts()
                            fig = go.Figure(go.Bar(y=list(c.index), x=list(c.values), orientation='h', marker_color=F[:len(c)]))
                            fig.update_layout(title="Abteilungen", height=400)
                            st.plotly_chart(fig, use_container_width=True)
                            charts_html.append(('07_Abteilungen.html', fig.to_html()))
                    
                    # ZIP Download
                    if charts_html:
                        st.markdown("---")
                        st.markdown("### 📦 Alle Diagramme herunterladen")
                        
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                            for name, html in charts_html:
                                zf.writestr(name, html)
                        
                        st.download_button(
                            label="📥 Alle Diagramme als ZIP",
                            data=zip_buffer.getvalue(),
                            file_name="HR_Analyse_Ergebnisse.zip",
                            mime="application/zip"
                        )
                    
            except Exception as e:
                st.error(f"Fehler: {e}")
    
    elif analyse_button:
        st.warning("Bitte erst eine Excel-Datei hochladen!")

# Footer
st.markdown("---")
st.markdown("*HR Analytics Tool - Rentenanalyse • Jubiläums-Tracker • Wissensverlust-Prognose*")
