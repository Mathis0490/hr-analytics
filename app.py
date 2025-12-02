import streamlit as st
import pandas as pd
from datetime import datetime
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
import zipfile

st.set_page_config(page_title="HR Analytics", page_icon="📊", layout="wide")

st.title("🚀 HR Analytics Tool")
st.markdown("**Excel hochladen → Diagramme bekommen!**")

BENCHMARK = {
    "Niedersachsen": {"alter": 44.6, "frauen": 50.3, "teilzeit": 28.4, "gehalt": 3650},
    "NRW": {"alter": 44.2, "frauen": 50.8, "teilzeit": 27.8, "gehalt": 3850},
    "Bayern": {"alter": 43.8, "frauen": 50.1, "teilzeit": 26.5, "gehalt": 4200},
    "Baden-Wuerttemberg": {"alter": 43.5, "frauen": 49.8, "teilzeit": 27.2, "gehalt": 4350},
    "Hessen": {"alter": 43.9, "frauen": 50.5, "teilzeit": 26.8, "gehalt": 4150},
    "Berlin": {"alter": 42.8, "frauen": 51.2, "teilzeit": 29.5, "gehalt": 3750},
    "Hamburg": {"alter": 42.5, "frauen": 51.0, "teilzeit": 28.2, "gehalt": 4450},
    "Sachsen": {"alter": 46.2, "frauen": 50.6, "teilzeit": 25.8, "gehalt": 3150},
    "Deutschland": {"alter": 44.3, "frauen": 50.5, "teilzeit": 27.5, "gehalt": 3950}
}

F = ['#3498db','#e74c3c','#2ecc71','#9b59b6','#f39c12','#1abc9c','#e67e22','#34495e','#16a085','#c0392b','#8e44ad','#d35400']

tab1, tab2 = st.tabs(["📥 1. Vorlage holen", "📊 2. Analyse starten"])

with tab1:
    st.markdown("### Leere Excel-Vorlage herunterladen")
    st.markdown("**Alles ist optional** - füllen Sie nur aus was Sie haben!")
    
    spalten = ['Mitarbeiter_ID','Geburtsjahr','Eintrittsjahr','Geschlecht','Abteilung','Position','Karrierelevel','Gehalt_Brutto_Jahr','Arbeitszeit','Wochenstunden','Standort','Bildungsabschluss','Vertragsart']
    df_vorlage = pd.DataFrame(columns=spalten, index=range(200))
    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_vorlage.to_excel(writer, index=False, sheet_name='Daten')
    
    st.download_button(label="📥 Excel-Vorlage herunterladen", data=buffer.getvalue(), file_name="HR_Vorlage.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with tab2:
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        uploaded_file = st.file_uploader("Excel hochladen", type=['xlsx', 'xls'])
    with col2:
        rentenalter = st.slider("Rentenalter", 60, 70, 67)
    with col3:
        region = st.selectbox("Region", list(BENCHMARK.keys()))
    
    if st.button("🚀 ANALYSE STARTEN", type="primary", use_container_width=True) and uploaded_file:
        with st.spinner("Analysiere..."):
            try:
                df = pd.read_excel(uploaded_file, engine='openpyxl')
                df = df.dropna(how='all')
                
                if len(df) == 0:
                    st.error("Excel ist leer!")
                else:
                    def find(names):
                        for c in df.columns:
                            for n in names:
                                if n in str(c).lower().replace('_',''):
                                    return c
                        return None
                    
                    col_geb = find(['geburtsjahr','jahrgang'])
                    col_ein = find(['eintrittsjahr','eintritt'])
                    col_ges = find(['geschlecht'])
                    col_abt = find(['abteilung'])
                    col_lvl = find(['karrierelevel','level'])
                    col_geh = find(['gehalt','brutto'])
                    col_az = find(['arbeitszeit'])
                    col_ort = find(['standort'])
                    col_bil = find(['bildung','abschluss'])
                    col_ver = find(['vertragsart'])
                    
                    jahr = datetime.now().year
                    if col_geb: df['Alter'] = jahr - pd.to_numeric(df[col_geb], errors='coerce')
                    if col_ein: df['DJ'] = jahr - pd.to_numeric(df[col_ein], errors='coerce')
                    if col_geh: df['Gehalt'] = pd.to_numeric(df[col_geh].astype(str).str.replace('[^0-9.]','',regex=True), errors='coerce')
                    
                    charts_html = []
                    st.success(f"✅ **{len(df)} Mitarbeiter** geladen!")
                    
                    # RENTE
                    if 'Alter' in df.columns:
                        d = df[df['Alter'].notna()].copy()
                        if len(d) > 0:
                            d['JbR'] = rentenalter - d['Alter']
                            d['RJ'] = jahr + d['JbR']
                            d['Kat'] = pd.cut(d['JbR'], [-100,0,5,10,15,20,100], labels=['Rente','0-5J','5-10J','10-15J','15-20J','20+J'])
                            
                            st.markdown("---")
                            st.markdown("## 🎯 RENTENANALYSE")
                            r5 = len(d[d['JbR'] <= 5])
                            if r5 > 0: st.error(f"⚠️ **{r5} Mitarbeiter ({round(r5/len(d)*100,1)}%)** gehen in 5 Jahren in Rente!")
                            
                            c1, c2 = st.columns(2)
                            with c1:
                                k = d['Kat'].value_counts()
                                fig = go.Figure(go.Pie(labels=[str(x) for x in k.index], values=list(k.values), marker_colors=['#c0392b','#e74c3c','#f39c12','#f1c40f','#2ecc71','#27ae60'], textinfo='label+percent+value'))
                                fig.update_layout(title="Jahre bis Rente", height=450)
                                st.plotly_chart(fig, use_container_width=True)
                                charts_html.append(('01_Rente_Kategorien.html', fig.to_html()))
                            
                            with c2:
                                rj = d[(d['RJ']>=jahr)&(d['RJ']<=jahr+15)]['RJ'].value_counts().sort_index()
                                if len(rj) > 0:
                                    fig = go.Figure(go.Bar(x=[int(x) for x in rj.index], y=list(rj.values), marker_color=['#e74c3c' if j<=jahr+5 else '#f39c12' if j<=jahr+10 else '#2ecc71' for j in rj.index], text=list(rj.values), textposition='outside'))
                                    fig.update_layout(title="Renteneintritte pro Jahr", height=450)
                                    st.plotly_chart(fig, use_container_width=True)
                                    charts_html.append(('02_Rente_Timeline.html', fig.to_html()))
                            
                            c1, c2 = st.columns(2)
                            with c1:
                                fig = go.Figure(go.Histogram(x=list(d['Alter']), nbinsx=20, marker_color='#3498db'))
                                fig.update_layout(title="Altersverteilung", height=400)
                                st.plotly_chart(fig, use_container_width=True)
                                charts_html.append(('03_Alter.html', fig.to_html()))
                            
                            with c2:
                                kum = [int(len(d[d['JbR']<=j])) for j in range(1,11)]
                                fig = go.Figure(go.Bar(x=[f"{j}J" for j in range(1,11)], y=kum, marker_color=['#e74c3c']*3+['#f39c12']*3+['#2ecc71']*4, text=[f"{k} ({round(k/len(d)*100)}%)" for k in kum], textposition='outside'))
                                fig.update_layout(title="Kumulative Abgänge", height=400)
                                st.plotly_chart(fig, use_container_width=True)
                                charts_html.append(('04_Kumulativ.html', fig.to_html()))
                            
                            if col_abt:
                                c1, c2 = st.columns(2)
                                with c1:
                                    avg = d.groupby(col_abt)['Alter'].mean().sort_values()
                                    fig = go.Figure(go.Bar(y=list(avg.index), x=[round(float(x),1) for x in avg.values], orientation='h', marker_color=['#e74c3c' if a>=50 else '#f39c12' if a>=45 else '#2ecc71' for a in avg.values], text=[f"{x:.1f}" for x in avg.values], textposition='outside'))
                                    fig.update_layout(title="Ø Alter/Abteilung", height=450)
                                    st.plotly_chart(fig, use_container_width=True)
                                    charts_html.append(('05_Alter_Abt.html', fig.to_html()))
                                
                                with c2:
                                    ab5 = d[d['JbR']<=5].groupby(col_abt).size()
                                    ges = d.groupby(col_abt).size()
                                    pct = (ab5/ges*100).fillna(0).sort_values()
                                    fig = go.Figure(go.Bar(y=list(pct.index), x=[round(float(x),1) for x in pct.values], orientation='h', marker_color=['#e74c3c' if p>=30 else '#f39c12' if p>=15 else '#2ecc71' for p in pct.values], text=[f"{x:.0f}%" for x in pct.values], textposition='outside'))
                                    fig.update_layout(title="Abgang 5J (%)", height=450)
                                    st.plotly_chart(fig, use_container_width=True)
                                    charts_html.append(('06_Abgang_Abt.html', fig.to_html()))
                    
                    # TREUE
                    if 'DJ' in df.columns:
                        d = df[df['DJ'].notna()].copy()
                        if len(d) > 0:
                            st.markdown("---")
                            st.markdown("## 🏆 TREUE-ANALYSE")
                            st.info(f"📊 Ø Betriebszugehörigkeit: **{d['DJ'].mean():.1f} Jahre** | {len(d[d['DJ']>=20])} MA sind 20+ Jahre dabei")
                            
                            c1, c2 = st.columns(2)
                            with c1:
                                fig = go.Figure(go.Histogram(x=list(d['DJ']), nbinsx=20, marker_color='#9b59b6'))
                                fig.update_layout(title="Dienstjahre", height=400)
                                st.plotly_chart(fig, use_container_width=True)
                                charts_html.append(('07_DJ.html', fig.to_html()))
                            
                            with c2:
                                d['Gr'] = pd.cut(d['DJ'],[-1,2,5,10,15,20,100],labels=['0-2J','3-5J','6-10J','11-15J','16-20J','20+J'])
                                gr = d['Gr'].value_counts()
                                fig = go.Figure(go.Pie(labels=[str(x) for x in gr.index], values=list(gr.values), marker_colors=F, textinfo='label+percent+value'))
                                fig.update_layout(title="Gruppen", height=400)
                                st.plotly_chart(fig, use_container_width=True)
                                charts_html.append(('08_DJ_Gruppen.html', fig.to_html()))
                            
                            c1, c2 = st.columns(2)
                            with c1:
                                jubi = [int(len(d[(d['DJ']>=j-0.5)&(d['DJ']<j+0.5)])) for j in [5,10,15,20,25,30]]
                                fig = go.Figure(go.Bar(x=['5J','10J','15J','20J','25J','30J'], y=jubi, marker_color=F[:6], text=jubi, textposition='outside'))
                                fig.update_layout(title="Jubiläen", height=400)
                                st.plotly_chart(fig, use_container_width=True)
                                charts_html.append(('09_Jubi.html', fig.to_html()))
                            
                            with c2:
                                if col_abt:
                                    avg = d.groupby(col_abt)['DJ'].mean().sort_values()
                                    fig = go.Figure(go.Bar(y=list(avg.index), x=[round(float(x),1) for x in avg.values], orientation='h', marker_color=['#27ae60' if x>=10 else '#f39c12' if x>=5 else '#e74c3c' for x in avg.values], text=[f"{x:.1f}J" for x in avg.values], textposition='outside'))
                                    fig.update_layout(title="Ø DJ/Abteilung", height=400)
                                    st.plotly_chart(fig, use_container_width=True)
                                    charts_html.append(('10_DJ_Abt.html', fig.to_html()))
                    
                    # WISSEN
                    if 'Alter' in df.columns and 'DJ' in df.columns:
                        d = df[(df['Alter'].notna())&(df['DJ'].notna())].copy()
                        if len(d) > 0:
                            d['JbR'] = rentenalter - d['Alter']
                            d['RJ'] = jahr + d['JbR']
                            d['R'] = d.apply(lambda r:'KRITISCH' if r['DJ']>=15 and r['JbR']<=5 else ('WARNUNG' if r['DJ']>=10 and r['JbR']<=10 else 'OK'),axis=1)
                            
                            st.markdown("---")
                            st.markdown("## ⚠️ WISSENSVERLUST")
                            
                            krit = len(d[d['R']=='KRITISCH'])
                            warn = len(d[d['R']=='WARNUNG'])
                            verl5 = d[d['JbR']<=5]['DJ'].sum()
                            
                            m1, m2, m3 = st.columns(3)
                            m1.metric("🔴 KRITISCH", krit)
                            m2.metric("🟡 WARNUNG", warn)
                            m3.metric("📉 Verlust 5J", f"{verl5:.0f} Jahre")
                            
                            c1, c2 = st.columns(2)
                            with c1:
                                fig = go.Figure()
                                cm = {'KRITISCH':'#e74c3c','WARNUNG':'#f39c12','OK':'#2ecc71'}
                                for r in ['OK','WARNUNG','KRITISCH']:
                                    x = d[d['R']==r]
                                    if len(x)>0:
                                        fig.add_trace(go.Scatter(x=list(x['JbR']), y=list(x['DJ']), mode='markers', name=r, marker=dict(color=cm[r], size=12, opacity=0.7)))
                                fig.update_layout(title="Erfahrung vs Rente", xaxis_title="Jahre bis Rente", yaxis_title="Dienstjahre", height=500)
                                st.plotly_chart(fig, use_container_width=True)
                                charts_html.append(('11_Wissen_Scatter.html', fig.to_html()))
                            
                            with c2:
                                verlust = {j: d[d['RJ']==j]['DJ'].sum() for j in range(jahr, jahr+11)}
                                fig = go.Figure(go.Bar(x=list(verlust.keys()), y=list(verlust.values()), marker_color=['#e74c3c' if v>d['DJ'].sum()*0.1 else '#f39c12' for v in verlust.values()], text=[f"{v:.0f}J" for v in verlust.values()], textposition='outside'))
                                fig.update_layout(title="Verlust pro Jahr", height=500)
                                st.plotly_chart(fig, use_container_width=True)
                                charts_html.append(('12_Wissen_Jahr.html', fig.to_html()))
                            
                            if col_abt:
                                c1, c2 = st.columns(2)
                                with c1:
                                    rc = d['R'].value_counts()
                                    fig = go.Figure(go.Bar(x=list(rc.index), y=list(rc.values), marker_color=[cm.get(r,'#999') for r in rc.index], text=list(rc.values), textposition='outside'))
                                    fig.update_layout(title="Risiko", height=400)
                                    st.plotly_chart(fig, use_container_width=True)
                                    charts_html.append(('13_Risiko.html', fig.to_html()))
                                
                                with c2:
                                    va = d[d['JbR']<=5].groupby(col_abt)['DJ'].sum().sort_values()
                                    if len(va)>0:
                                        fig = go.Figure(go.Bar(y=list(va.index), x=list(va.values), orientation='h', marker_color='#e74c3c', text=[f"{v:.0f}J" for v in va.values], textposition='outside'))
                                        fig.update_layout(title="Verlust/Abteilung", height=400)
                                        st.plotly_chart(fig, use_container_width=True)
                                        charts_html.append(('14_Wissen_Abt.html', fig.to_html()))
                    
                    # WEITERE
                    st.markdown("---")
                    st.markdown("## 📈 WEITERE ANALYSEN")
                    
                    c1, c2 = st.columns(2)
                    with c1:
                        if col_ges:
                            c = df[col_ges].value_counts()
                            fig = go.Figure(go.Pie(labels=[str(x) for x in c.index], values=list(c.values), marker_colors=['#3498db','#e74c3c','#2ecc71'], textinfo='label+percent+value'))
                            fig.update_layout(title="Geschlecht", height=400)
                            st.plotly_chart(fig, use_container_width=True)
                            charts_html.append(('15_Geschlecht.html', fig.to_html()))
                    
                    with c2:
                        if col_abt:
                            c = df[col_abt].value_counts()
                            fig = go.Figure(go.Bar(y=list(c.index), x=list(c.values), orientation='h', marker_color=F[:len(c)], text=list(c.values), textposition='outside'))
                            fig.update_layout(title="Abteilungen", height=400)
                            st.plotly_chart(fig, use_container_width=True)
                            charts_html.append(('16_Abteilungen.html', fig.to_html()))
                    
                    c1, c2 = st.columns(2)
                    with c1:
                        if col_lvl:
                            c = df[col_lvl].value_counts()
                            fig = go.Figure(go.Bar(x=list(c.index), y=list(c.values), marker_color=F[:len(c)], text=list(c.values), textposition='outside'))
                            fig.update_layout(title="Karrierelevel", height=400)
                            st.plotly_chart(fig, use_container_width=True)
                            charts_html.append(('17_Level.html', fig.to_html()))
                    
                    with c2:
                        if col_az:
                            c = df[col_az].value_counts()
                            fig = go.Figure(go.Pie(labels=list(c.index), values=list(c.values), marker_colors=['#3498db','#f39c12'], textinfo='label+percent+value'))
                            fig.update_layout(title="Arbeitszeit", height=400)
                            st.plotly_chart(fig, use_container_width=True)
                            charts_html.append(('18_AZ.html', fig.to_html()))
                    
                    c1, c2 = st.columns(2)
                    with c1:
                        if 'Gehalt' in df.columns:
                            g = df['Gehalt'].dropna()
                            if len(g)>0:
                                fig = go.Figure(go.Histogram(x=list(g), nbinsx=20, marker_color='#2ecc71'))
                                fig.update_layout(title="Gehalt", height=400)
                                st.plotly_chart(fig, use_container_width=True)
                                charts_html.append(('19_Gehalt.html', fig.to_html()))
                    
                    with c2:
                        if 'Gehalt' in df.columns and col_abt:
                            ga = df.groupby(col_abt)['Gehalt'].mean().dropna().sort_values()
                            if len(ga)>0:
                                fig = go.Figure(go.Bar(y=list(ga.index), x=[round(float(x),0) for x in ga.values], orientation='h', marker_color='#3498db', text=[f"{x:,.0f}€" for x in ga.values], textposition='outside'))
                                fig.update_layout(title="Gehalt/Abteilung", height=400)
                                st.plotly_chart(fig, use_container_width=True)
                                charts_html.append(('20_Gehalt_Abt.html', fig.to_html()))
                    
                    c1, c2 = st.columns(2)
                    with c1:
                        if col_ort:
                            c = df[col_ort].value_counts()
                            fig = go.Figure(go.Pie(labels=list(c.index), values=list(c.values), marker_colors=F, textinfo='label+percent+value'))
                            fig.update_layout(title="Standorte", height=400)
                            st.plotly_chart(fig, use_container_width=True)
                            charts_html.append(('21_Standorte.html', fig.to_html()))
                    
                    with c2:
                        if col_bil:
                            c = df[col_bil].value_counts()
                            fig = go.Figure(go.Bar(x=list(c.index), y=list(c.values), marker_color=F[:len(c)], text=list(c.values), textposition='outside'))
                            fig.update_layout(title="Bildung", height=400)
                            st.plotly_chart(fig, use_container_width=True)
                            charts_html.append(('22_Bildung.html', fig.to_html()))
                    
                    # Sankey
                    if col_abt and col_lvl:
                        st.markdown("### 🌊 Karriere-Flow")
                        flow = df.groupby([col_abt, col_lvl]).size().reset_index(name='count')
                        flow = flow[flow['count'] > 0]
                        if len(flow) > 0:
                            abt_list = list(flow[col_abt].unique())
                            lvl_list = list(flow[col_lvl].unique())
                            nodes = abt_list + lvl_list
                            fig = go.Figure(go.Sankey(node=dict(pad=15, thickness=20, label=nodes, color=['#3498db']*len(abt_list)+['#e74c3c']*len(lvl_list)), link=dict(source=[nodes.index(a) for a in flow[col_abt]], target=[nodes.index(l) for l in flow[col_lvl]], value=list(flow['count']))))
                            fig.update_layout(title="Abteilung → Level", height=500)
                            st.plotly_chart(fig, use_container_width=True)
                            charts_html.append(('23_Sankey.html', fig.to_html()))
                    
                    # Benchmark
                    if region in BENCHMARK:
                        st.markdown("---")
                        st.markdown("## 📊 BENCHMARK")
                        b = BENCHMARK[region]
                        kat, u, bm = [], [], []
                        if 'Alter' in df.columns:
                            a = df['Alter'].mean()
                            if not pd.isna(a): kat.append('Alter'); u.append(round(float(a),1)); bm.append(b['alter'])
                        if col_ges:
                            g = df[col_ges].value_counts(normalize=True)*100
                            f = sum(float(v) for k,v in g.items() if str(k).lower() in ['w','weiblich'])
                            kat.append('Frauen%'); u.append(round(f,1)); bm.append(b['frauen'])
                        if col_az:
                            a = df[col_az].value_counts(normalize=True)*100
                            t = sum(float(v) for k,v in a.items() if 'teil' in str(k).lower())
                            kat.append('Teilzeit%'); u.append(round(t,1)); bm.append(b['teilzeit'])
                        if 'Gehalt' in df.columns:
                            a = df['Gehalt'].mean()
                            if not pd.isna(a): kat.append('Gehalt/M'); u.append(round(float(a)/12,0)); bm.append(b['gehalt'])
                        if kat:
                            fig = go.Figure([go.Bar(name='Sie', x=kat, y=u, marker_color='#3498db'), go.Bar(name=region, x=kat, y=bm, marker_color='#e74c3c')])
                            fig.update_layout(title=f"Sie vs {region}", barmode='group', height=450)
                            st.plotly_chart(fig, use_container_width=True)
                            charts_html.append(('24_Benchmark.html', fig.to_html()))
                    
                    # Download
                    if charts_html:
                        st.markdown("---")
                        st.success(f"✅ **{len(charts_html)} Diagramme** erstellt!")
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                            for name, html in charts_html:
                                zf.writestr(name, html)
                        st.download_button(label="📥 Alle Diagramme als ZIP", data=zip_buffer.getvalue(), file_name="HR_Ergebnisse.zip", mime="application/zip", use_container_width=True)
                    
            except Exception as e:
                st.error(f"Fehler: {e}")

st.markdown("---")
st.markdown("*HR Analytics v2.0 | 🎯 Rente • 🏆 Treue • ⚠️ Wissensverlust • 📊 Benchmark*")
