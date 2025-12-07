import streamlit as st
import pandas as pd
from datetime import datetime
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
import zipfile

# ============================================
# SEITEN-EINSTELLUNGEN
# ============================================
st.set_page_config(
    page_title="HR Analyse", 
    page_icon="📊", 
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ============================================
# GROSSER, LESBARER STIL
# ============================================
st.markdown("""
<style>
    /* Größere Schrift überall */
    html, body, [class*="css"] {
        font-size: 18px !important;
    }
    
    /* Große Überschriften */
    h1 {
        font-size: 42px !important;
        color: #2c3e50 !important;
    }
    h2 {
        font-size: 32px !important;
        color: #34495e !important;
        border-bottom: 3px solid #3498db !important;
        padding-bottom: 10px !important;
    }
    h3 {
        font-size: 26px !important;
    }
    
    /* Große Buttons */
    .stButton > button {
        font-size: 24px !important;
        padding: 20px 40px !important;
        border-radius: 15px !important;
        font-weight: bold !important;
    }
    
    /* Große Upload-Box */
    .stFileUploader {
        font-size: 20px !important;
    }
    
    /* Große Tabs */
    .stTabs [data-baseweb="tab"] {
        font-size: 22px !important;
        padding: 15px 30px !important;
    }
    
    /* Erfolgs/Warn-Boxen größer */
    .stAlert {
        font-size: 20px !important;
        padding: 20px !important;
    }
    
    /* Metriken größer */
    [data-testid="stMetricValue"] {
        font-size: 48px !important;
    }
    [data-testid="stMetricLabel"] {
        font-size: 20px !important;
    }
    
    /* Download-Button groß */
    .stDownloadButton > button {
        font-size: 22px !important;
        padding: 15px 30px !important;
        background-color: #27ae60 !important;
        color: white !important;
    }
    
    /* Slider größer */
    .stSlider {
        padding: 10px 0 !important;
    }
    
    /* Selectbox größer */
    .stSelectbox {
        font-size: 18px !important;
    }
    
    /* Info-Box Stil */
    .info-box {
        background-color: #e8f4f8;
        border-left: 6px solid #3498db;
        padding: 20px;
        margin: 20px 0;
        border-radius: 0 10px 10px 0;
        font-size: 18px;
    }
    
    /* Hilfe-Text */
    .help-text {
        background-color: #fef9e7;
        border: 2px solid #f39c12;
        padding: 20px;
        border-radius: 10px;
        margin: 15px 0;
        font-size: 18px;
    }
</style>
""", unsafe_allow_html=True)

# ============================================
# BENCHMARK DATEN
# ============================================
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

F = ['#3498db','#e74c3c','#2ecc71','#9b59b6','#f39c12','#1abc9c','#e67e22','#34495e']

# ============================================
# KOPFZEILE
# ============================================
st.title("📊 HR Analyse Tool")
st.markdown("### Ihre Mitarbeiterdaten einfach auswerten")

# ============================================
# HAUPT-TABS
# ============================================
tab1, tab2 = st.tabs(["📥 SCHRITT 1: Vorlage holen", "📊 SCHRITT 2: Analyse starten"])

# ============================================
# TAB 1: VORLAGE
# ============================================
with tab1:
    st.markdown("## 📥 Excel-Vorlage herunterladen")
    
    st.markdown("""
    <div class="help-text">
    <b>💡 Was ist das?</b><br><br>
    Diese Vorlage ist eine leere Excel-Tabelle mit den richtigen Spaltenüberschriften.<br>
    Sie können Ihre Mitarbeiterdaten dort einfügen und dann hier analysieren lassen.
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    st.markdown("### ✅ Diese Spalten sind in der Vorlage:")
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("""
        **Wichtigste Spalten:**
        - 📅 **Geburtsjahr** → für Rentenanalyse
        - 📅 **Eintrittsjahr** → für Jubiläen
        - 👤 **Geschlecht** (m/w/d)
        - 🏢 **Abteilung**
        """)
    with col2:
        st.markdown("""
        **Zusätzliche Spalten:**
        - 💼 Einstiegsposition
        - 💼 Aktuelle Position  
        - 📈 Karrierelevel
        - 💰 Gehalt
        - ⏰ Arbeitszeit
        """)
    
    st.markdown("---")
    
    st.markdown("""
    <div class="info-box">
    <b>ℹ️ Hinweis:</b> Sie müssen nicht alle Spalten ausfüllen!<br>
    Füllen Sie nur aus, was Sie haben. Das Tool zeigt dann passende Auswertungen.
    </div>
    """, unsafe_allow_html=True)
    
    # Vorlage erstellen
    spalten = ['Mitarbeiter_ID','Geburtsjahr','Eintrittsjahr','Geschlecht','Abteilung',
               'Einstiegsposition','Aktuelle_Position','Karrierelevel','Gehalt_Brutto_Jahr',
               'Arbeitszeit','Wochenstunden','Standort','Bildungsabschluss','Vertragsart']
    df_vorlage = pd.DataFrame(columns=spalten, index=range(500))
    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_vorlage.to_excel(writer, index=False, sheet_name='Mitarbeiter')
        ws = writer.sheets['Mitarbeiter']
        fmt = writer.book.add_format({'bold': True, 'bg_color': '#27AE60', 'font_color': 'white', 'font_size': 12})
        for i, col in enumerate(spalten):
            ws.write(0, i, col, fmt)
            ws.set_column(i, i, 20)
    
    st.markdown("### 👇 Klicken Sie hier zum Herunterladen:")
    
    st.download_button(
        label="📥  VORLAGE HERUNTERLADEN",
        data=buffer.getvalue(),
        file_name="HR_Vorlage.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

# ============================================
# TAB 2: ANALYSE
# ============================================
with tab2:
    st.markdown("## 📊 Ihre Daten analysieren")
    
    st.markdown("""
    <div class="help-text">
    <b>💡 So geht's:</b><br><br>
    1️⃣ Klicken Sie auf "Browse files" (oder "Dateien auswählen")<br>
    2️⃣ Wählen Sie Ihre ausgefüllte Excel-Datei aus<br>
    3️⃣ Stellen Sie das Rentenalter ein (normalerweise 67)<br>
    4️⃣ Klicken Sie auf den großen grünen Button "ANALYSE STARTEN"<br>
    5️⃣ Warten Sie einen Moment - die Diagramme erscheinen automatisch
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # DATEI UPLOAD
    st.markdown("### 1️⃣ Excel-Datei auswählen:")
    uploaded_file = st.file_uploader(
        "Klicken Sie hier oder ziehen Sie Ihre Datei hierher",
        type=['xlsx', 'xls'],
        help="Nur Excel-Dateien (.xlsx oder .xls)"
    )
    
    if uploaded_file:
        st.success(f"✅ Datei geladen: **{uploaded_file.name}**")
    
    st.markdown("---")
    
    # EINSTELLUNGEN
    st.markdown("### 2️⃣ Einstellungen:")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**Rentenalter:**")
        rentenalter = st.slider(
            "Ab welchem Alter gehen Ihre Mitarbeiter in Rente?",
            min_value=60,
            max_value=70,
            value=67,
            help="In Deutschland ist das Regelrentenalter 67 Jahre"
        )
        st.markdown(f"*Eingestellt: **{rentenalter} Jahre***")
    
    with col2:
        st.markdown("**Vergleichsregion:**")
        region = st.selectbox(
            "Mit welcher Region möchten Sie vergleichen?",
            list(BENCHMARK.keys()),
            help="Ihre Zahlen werden mit dem Durchschnitt dieser Region verglichen"
        )
    
    st.markdown("---")
    
    # START BUTTON
    st.markdown("### 3️⃣ Analyse starten:")
    
    analyse_button = st.button(
        "🚀  ANALYSE STARTEN",
        type="primary",
        use_container_width=True,
        disabled=uploaded_file is None
    )
    
    if uploaded_file is None:
        st.warning("⚠️ Bitte laden Sie zuerst eine Excel-Datei hoch (siehe Schritt 1)")
    
    # ============================================
    # ANALYSE DURCHFÜHREN
    # ============================================
    if analyse_button and uploaded_file:
        
        # Lade-Animation
        with st.spinner("🔄 Bitte warten... Ihre Daten werden analysiert..."):
            try:
                df = pd.read_excel(uploaded_file, engine='openpyxl')
                df = df.dropna(how='all')
                
                if len(df) == 0:
                    st.error("❌ Die Excel-Datei ist leer! Bitte füllen Sie zuerst Daten ein.")
                else:
                    # Spalten finden
                    def find(names):
                        for c in df.columns:
                            for n in names:
                                if n in str(c).lower().replace('_','').replace(' ',''):
                                    return c
                        return None
                    
                    col_geb = find(['geburtsjahr','jahrgang'])
                    col_ein = find(['eintrittsjahr','eintritt'])
                    col_ges = find(['geschlecht'])
                    col_abt = find(['abteilung'])
                    col_lvl = find(['karrierelevel','level'])
                    col_geh = find(['gehalt','brutto'])
                    col_az = find(['arbeitszeit'])
                    col_ein_pos = find(['einstiegsposition','einstieg'])
                    col_akt_pos = find(['aktuelleposition','aktuelle','position'])
                    col_ort = find(['standort'])
                    
                    jahr = datetime.now().year
                    
                    if col_geb: 
                        df['Alter'] = jahr - pd.to_numeric(df[col_geb], errors='coerce')
                    if col_ein: 
                        df['DJ'] = jahr - pd.to_numeric(df[col_ein], errors='coerce')
                    if col_geh: 
                        df['Gehalt'] = pd.to_numeric(df[col_geh].astype(str).str.replace('[^0-9.]','',regex=True), errors='coerce')
                    
                    charts_html = []
                    
                    # ============================================
                    # ERFOLGS-MELDUNG
                    # ============================================
                    st.markdown("---")
                    st.success(f"✅ **{len(df)} Mitarbeiter** wurden erfolgreich geladen!")
                    
                    # Übersicht
                    st.markdown("## 📋 Übersicht Ihrer Daten")
                    
                    m1, m2, m3, m4 = st.columns(4)
                    with m1:
                        st.metric("👥 Mitarbeiter", len(df))
                    with m2:
                        if 'Alter' in df.columns:
                            st.metric("🎂 Ø Alter", f"{df['Alter'].mean():.1f} Jahre")
                    with m3:
                        if 'DJ' in df.columns:
                            st.metric("🏆 Ø Betriebszugehörigkeit", f"{df['DJ'].mean():.1f} Jahre")
                    with m4:
                        if 'Gehalt' in df.columns:
                            st.metric("💰 Ø Gehalt", f"{df['Gehalt'].mean():,.0f} €")
                    
                    # ============================================
                    # DATENQUALITÄTS-CHECK
                    # ============================================
                    st.markdown("---")
                    st.markdown("## 🔍 Datenqualitäts-Check")
                    
                    st.markdown("""
                    <div class="help-text">
                    <b>💡 Was zeigt das?</b><br><br>
                    Hier sehen Sie auf einen Blick:<br>
                    • 🔴 <b>Fehlende Daten</b> = Leere Zellen in Ihrer Excel<br>
                    • 🟡 <b>Ausreißer</b> = Werte die ungewöhnlich hoch oder niedrig sind (könnten Tippfehler sein)
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # Alle Spalten für Qualitätscheck
                    qual_spalten = {
                        'Geburtsjahr': col_geb,
                        'Eintrittsjahr': col_ein,
                        'Geschlecht': col_ges,
                        'Abteilung': col_abt,
                        'Karrierelevel': col_lvl,
                        'Gehalt': col_geh,
                        'Arbeitszeit': col_az,
                        'Einstiegsposition': col_ein_pos,
                        'Aktuelle Position': col_akt_pos,
                        'Standort': col_ort
                    }
                    
                    # Fehlende Werte berechnen
                    fehlend_data = []
                    for name, col in qual_spalten.items():
                        if col and col in df.columns:
                            fehlend = df[col].isna().sum()
                            pct = round(fehlend / len(df) * 100, 1)
                            fehlend_data.append({
                                'Spalte': name,
                                'Fehlend': fehlend,
                                'Prozent': pct,
                                'Status': '🔴 Kritisch' if pct > 20 else ('🟡 Prüfen' if pct > 5 else '🟢 OK')
                            })
                    
                    # Ausreißer berechnen
                    ausreisser_data = []
                    ausreisser_details = []
                    
                    # Alter prüfen
                    if 'Alter' in df.columns:
                        alter = df['Alter'].dropna()
                        if len(alter) > 0:
                            zu_jung = df[df['Alter'] < 16]
                            zu_alt = df[df['Alter'] > 70]
                            ausreisser_data.append({
                                'Kategorie': 'Alter < 16 Jahre',
                                'Anzahl': len(zu_jung),
                                'Status': '🔴' if len(zu_jung) > 0 else '🟢'
                            })
                            ausreisser_data.append({
                                'Kategorie': 'Alter > 70 Jahre',
                                'Anzahl': len(zu_alt),
                                'Status': '🔴' if len(zu_alt) > 0 else '🟢'
                            })
                            if len(zu_jung) > 0:
                                for _, row in zu_jung.head(5).iterrows():
                                    ausreisser_details.append(f"• {row.get('Mitarbeiter_ID', '?')}: Alter {row['Alter']:.0f} (zu jung?)")
                            if len(zu_alt) > 0:
                                for _, row in zu_alt.head(5).iterrows():
                                    ausreisser_details.append(f"• {row.get('Mitarbeiter_ID', '?')}: Alter {row['Alter']:.0f} (zu alt?)")
                    
                    # Dienstjahre prüfen
                    if 'DJ' in df.columns:
                        dj = df['DJ'].dropna()
                        if len(dj) > 0:
                            negativ_dj = df[df['DJ'] < 0]
                            sehr_lang = df[df['DJ'] > 50]
                            ausreisser_data.append({
                                'Kategorie': 'Negative Dienstjahre',
                                'Anzahl': len(negativ_dj),
                                'Status': '🔴' if len(negativ_dj) > 0 else '🟢'
                            })
                            ausreisser_data.append({
                                'Kategorie': 'Dienstjahre > 50',
                                'Anzahl': len(sehr_lang),
                                'Status': '🟡' if len(sehr_lang) > 0 else '🟢'
                            })
                            if len(negativ_dj) > 0:
                                for _, row in negativ_dj.head(5).iterrows():
                                    ausreisser_details.append(f"• {row.get('Mitarbeiter_ID', '?')}: {row['DJ']:.0f} Dienstjahre (negativ!)")
                            if len(sehr_lang) > 0:
                                for _, row in sehr_lang.head(5).iterrows():
                                    ausreisser_details.append(f"• {row.get('Mitarbeiter_ID', '?')}: {row['DJ']:.0f} Dienstjahre (sehr lang)")
                    
                    # Gehalt prüfen
                    if 'Gehalt' in df.columns:
                        gehalt = df['Gehalt'].dropna()
                        if len(gehalt) > 0:
                            mean_g = gehalt.mean()
                            std_g = gehalt.std()
                            sehr_niedrig = df[(df['Gehalt'] < 15000) & (df['Gehalt'].notna())]
                            sehr_hoch = df[(df['Gehalt'] > 300000) & (df['Gehalt'].notna())]
                            # Statistische Ausreißer (mehr als 3 Standardabweichungen)
                            stat_ausreisser = df[(df['Gehalt'].notna()) & ((df['Gehalt'] < mean_g - 3*std_g) | (df['Gehalt'] > mean_g + 3*std_g))]
                            
                            ausreisser_data.append({
                                'Kategorie': 'Gehalt < 15.000€',
                                'Anzahl': len(sehr_niedrig),
                                'Status': '🟡' if len(sehr_niedrig) > 0 else '🟢'
                            })
                            ausreisser_data.append({
                                'Kategorie': 'Gehalt > 300.000€',
                                'Anzahl': len(sehr_hoch),
                                'Status': '🔴' if len(sehr_hoch) > 0 else '🟢'
                            })
                            ausreisser_data.append({
                                'Kategorie': 'Statistische Ausreißer (±3σ)',
                                'Anzahl': len(stat_ausreisser),
                                'Status': '🟡' if len(stat_ausreisser) > 0 else '🟢'
                            })
                            if len(sehr_niedrig) > 0:
                                for _, row in sehr_niedrig.head(3).iterrows():
                                    ausreisser_details.append(f"• {row.get('Mitarbeiter_ID', '?')}: {row['Gehalt']:,.0f}€ (sehr niedrig)")
                            if len(sehr_hoch) > 0:
                                for _, row in sehr_hoch.head(3).iterrows():
                                    ausreisser_details.append(f"• {row.get('Mitarbeiter_ID', '?')}: {row['Gehalt']:,.0f}€ (sehr hoch)")
                    
                    # Logik-Prüfung: Eintritt vor Geburt?
                    if col_geb and col_ein:
                        logik_fehler = df[df[col_ein] < df[col_geb]]
                        ausreisser_data.append({
                            'Kategorie': 'Eintritt vor Geburt (!)',
                            'Anzahl': len(logik_fehler),
                            'Status': '🔴' if len(logik_fehler) > 0 else '🟢'
                        })
                        if len(logik_fehler) > 0:
                            for _, row in logik_fehler.head(3).iterrows():
                                ausreisser_details.append(f"• {row.get('Mitarbeiter_ID', '?')}: Geb. {row[col_geb]}, Eintritt {row[col_ein]} (unmöglich!)")
                    
                    # Eintritt mit unter 14?
                    if 'Alter' in df.columns and 'DJ' in df.columns:
                        zu_frueh = df[(df['Alter'] - df['DJ']) < 14]
                        ausreisser_data.append({
                            'Kategorie': 'Eintritt unter 14 Jahren',
                            'Anzahl': len(zu_frueh),
                            'Status': '🔴' if len(zu_frueh) > 0 else '🟢'
                        })
                        if len(zu_frueh) > 0:
                            for _, row in zu_frueh.head(3).iterrows():
                                eintrittsalter = row['Alter'] - row['DJ']
                                ausreisser_details.append(f"• {row.get('Mitarbeiter_ID', '?')}: Eintritt mit {eintrittsalter:.0f} Jahren (zu jung)")
                    
                    # Visualisierung
                    c1, c2 = st.columns(2)
                    
                    with c1:
                        st.markdown("### 📊 Fehlende Daten pro Spalte")
                        if fehlend_data:
                            df_fehlend = pd.DataFrame(fehlend_data)
                            
                            # Farben basierend auf Prozent
                            farben = ['#e74c3c' if p > 20 else '#f39c12' if p > 5 else '#2ecc71' for p in df_fehlend['Prozent']]
                            
                            fig = go.Figure()
                            fig.add_trace(go.Bar(
                                y=df_fehlend['Spalte'],
                                x=df_fehlend['Prozent'],
                                orientation='h',
                                marker_color=farben,
                                text=[f"{p}% ({f})" for p, f in zip(df_fehlend['Prozent'], df_fehlend['Fehlend'])],
                                textposition='outside'
                            ))
                            fig.update_layout(
                                title="Anteil fehlender Werte (%)",
                                xaxis_title="Fehlend in %",
                                height=400,
                                font=dict(size=14),
                                xaxis=dict(range=[0, max(df_fehlend['Prozent'].max() * 1.3, 10)])
                            )
                            st.plotly_chart(fig, use_container_width=True)
                            charts_html.append(('00_Datenqualitaet_Fehlend.html', fig.to_html()))
                            
                            # Zusammenfassung
                            kritisch = len([x for x in fehlend_data if '🔴' in x['Status']])
                            warnung = len([x for x in fehlend_data if '🟡' in x['Status']])
                            if kritisch > 0:
                                st.error(f"🔴 {kritisch} Spalte(n) haben mehr als 20% fehlende Daten!")
                            elif warnung > 0:
                                st.warning(f"🟡 {warnung} Spalte(n) haben 5-20% fehlende Daten")
                            else:
                                st.success("🟢 Alle Spalten haben weniger als 5% fehlende Daten")
                        else:
                            st.info("Keine Spalten zum Prüfen gefunden")
                    
                    with c2:
                        st.markdown("### ⚠️ Ausreißer & Fehler")
                        if ausreisser_data:
                            df_aus = pd.DataFrame(ausreisser_data)
                            
                            # Nur Zeilen mit Problemen anzeigen
                            df_probleme = df_aus[df_aus['Anzahl'] > 0]
                            
                            if len(df_probleme) > 0:
                                farben = ['#e74c3c' if '🔴' in s else '#f39c12' for s in df_probleme['Status']]
                                
                                fig = go.Figure()
                                fig.add_trace(go.Bar(
                                    y=df_probleme['Kategorie'],
                                    x=df_probleme['Anzahl'],
                                    orientation='h',
                                    marker_color=farben,
                                    text=df_probleme['Anzahl'],
                                    textposition='outside'
                                ))
                                fig.update_layout(
                                    title="Gefundene Probleme",
                                    xaxis_title="Anzahl Datensätze",
                                    height=400,
                                    font=dict(size=14)
                                )
                                st.plotly_chart(fig, use_container_width=True)
                                charts_html.append(('00_Datenqualitaet_Ausreisser.html', fig.to_html()))
                            else:
                                st.success("🟢 Keine offensichtlichen Ausreißer gefunden!")
                                fig = go.Figure()
                                fig.add_annotation(
                                    text="✅ Alles OK!",
                                    xref="paper", yref="paper",
                                    x=0.5, y=0.5,
                                    showarrow=False,
                                    font=dict(size=40, color='#2ecc71')
                                )
                                fig.update_layout(height=400)
                                st.plotly_chart(fig, use_container_width=True)
                            
                            # Kritische Fehler zählen
                            krit_aus = len([x for x in ausreisser_data if '🔴' in x['Status'] and x['Anzahl'] > 0])
                            warn_aus = len([x for x in ausreisser_data if '🟡' in x['Status'] and x['Anzahl'] > 0])
                            
                            if krit_aus > 0:
                                st.error(f"🔴 {krit_aus} kritische(r) Fehler gefunden!")
                            elif warn_aus > 0:
                                st.warning(f"🟡 {warn_aus} mögliche(r) Ausreißer - bitte prüfen")
                            else:
                                st.success("🟢 Keine Ausreißer gefunden")
                    
                    # Details zu Ausreißern
                    if ausreisser_details:
                        with st.expander("📋 Details zu den gefundenen Problemen (klicken zum Öffnen)"):
                            st.markdown("**Betroffene Datensätze:**")
                            for detail in ausreisser_details[:15]:  # Max 15 anzeigen
                                st.markdown(detail)
                            if len(ausreisser_details) > 15:
                                st.markdown(f"*... und {len(ausreisser_details) - 15} weitere*")
                    
                    # Gesamtbewertung
                    st.markdown("### 📊 Gesamtbewertung Datenqualität")
                    
                    # Score berechnen
                    total_fehlend = sum([x['Fehlend'] for x in fehlend_data]) if fehlend_data else 0
                    max_fehlend = len(df) * len(fehlend_data) if fehlend_data else 1
                    fehlend_score = 100 - (total_fehlend / max_fehlend * 100) if max_fehlend > 0 else 100
                    
                    ausreisser_count = sum([x['Anzahl'] for x in ausreisser_data if '🔴' in x['Status']]) if ausreisser_data else 0
                    ausreisser_score = max(0, 100 - (ausreisser_count / len(df) * 500)) if len(df) > 0 else 100
                    
                    gesamt_score = (fehlend_score * 0.6 + ausreisser_score * 0.4)
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Vollständigkeit", f"{fehlend_score:.0f}%", 
                            help="Wie viele Felder sind ausgefüllt?")
                    with col2:
                        st.metric("Plausibilität", f"{ausreisser_score:.0f}%",
                            help="Wie viele Werte sind plausibel?")
                    with col3:
                        farbe = "🟢" if gesamt_score >= 80 else ("🟡" if gesamt_score >= 60 else "🔴")
                        st.metric(f"{farbe} Gesamtscore", f"{gesamt_score:.0f}%")
                    
                    if gesamt_score >= 80:
                        st.success("✅ **Gute Datenqualität!** Sie können die Analyse starten.")
                    elif gesamt_score >= 60:
                        st.warning("⚠️ **Mittlere Datenqualität.** Einige Analysen könnten ungenau sein.")
                    else:
                        st.error("❌ **Datenqualität verbesserungswürdig.** Bitte prüfen Sie die markierten Probleme.")
                    
                    # ============================================
                    # RENTENANALYSE
                    # ============================================
                    if 'Alter' in df.columns:
                        d = df[df['Alter'].notna()].copy()
                        if len(d) > 0:
                            d['JbR'] = rentenalter - d['Alter']
                            d['RJ'] = jahr + d['JbR']
                            d['Kat'] = pd.cut(d['JbR'], [-100,0,5,10,15,20,100], 
                                labels=['Bereits Rente','0-5 Jahre','5-10 Jahre','10-15 Jahre','15-20 Jahre','Mehr als 20 Jahre'])
                            
                            st.markdown("---")
                            st.markdown("## 🎯 Wann gehen Ihre Mitarbeiter in Rente?")
                            
                            r5 = len(d[d['JbR'] <= 5])
                            r10 = len(d[(d['JbR'] > 5) & (d['JbR'] <= 10)])
                            
                            if r5 > 0:
                                st.error(f"""
                                ⚠️ **ACHTUNG:** {r5} Mitarbeiter ({round(r5/len(d)*100,1)}%) 
                                erreichen in den nächsten **5 Jahren** das Rentenalter!
                                """)
                            
                            if r10 > 0:
                                st.warning(f"""
                                ℹ️ Weitere {r10} Mitarbeiter ({round(r10/len(d)*100,1)}%) 
                                gehen in **5-10 Jahren** in Rente.
                                """)
                            
                            c1, c2 = st.columns(2)
                            
                            with c1:
                                k = d['Kat'].value_counts()
                                fig = go.Figure(go.Pie(
                                    labels=[str(x) for x in k.index], 
                                    values=list(k.values),
                                    marker_colors=['#c0392b','#e74c3c','#f39c12','#f1c40f','#2ecc71','#27ae60'],
                                    textinfo='label+percent+value',
                                    textfont_size=16
                                ))
                                fig.update_layout(title="Wie lange noch bis zur Rente?", height=500, font=dict(size=16))
                                st.plotly_chart(fig, use_container_width=True)
                                charts_html.append(('01_Rente_Uebersicht.html', fig.to_html()))
                            
                            with c2:
                                rj = d[(d['RJ']>=jahr)&(d['RJ']<=jahr+15)]['RJ'].value_counts().sort_index()
                                if len(rj) > 0:
                                    fig = go.Figure(go.Bar(
                                        x=[int(x) for x in rj.index], 
                                        y=list(rj.values),
                                        marker_color=['#e74c3c' if j<=jahr+5 else '#f39c12' if j<=jahr+10 else '#2ecc71' for j in rj.index],
                                        text=list(rj.values),
                                        textposition='outside',
                                        textfont_size=14
                                    ))
                                    fig.update_layout(title="Renteneintritte pro Jahr", height=500, font=dict(size=14),
                                        xaxis_title="Jahr", yaxis_title="Anzahl Mitarbeiter")
                                    st.plotly_chart(fig, use_container_width=True)
                                    charts_html.append(('02_Rente_Pro_Jahr.html', fig.to_html()))
                            
                            c1, c2 = st.columns(2)
                            
                            with c1:
                                fig = go.Figure(go.Histogram(x=list(d['Alter']), nbinsx=20, marker_color='#3498db'))
                                fig.update_layout(title="Altersverteilung aller Mitarbeiter", height=450, font=dict(size=14),
                                    xaxis_title="Alter", yaxis_title="Anzahl")
                                st.plotly_chart(fig, use_container_width=True)
                                charts_html.append(('03_Altersverteilung.html', fig.to_html()))
                            
                            with c2:
                                kum = [int(len(d[d['JbR']<=j])) for j in range(1,11)]
                                fig = go.Figure(go.Bar(
                                    x=[f"In {j} Jahr(en)" for j in range(1,11)], 
                                    y=kum,
                                    marker_color=['#e74c3c']*3+['#f39c12']*3+['#2ecc71']*4,
                                    text=[f"{k} ({round(k/len(d)*100)}%)" for k in kum],
                                    textposition='outside'
                                ))
                                fig.update_layout(title="Wie viele gehen wann?", height=450, font=dict(size=14),
                                    yaxis_title="Anzahl Mitarbeiter (kumuliert)")
                                st.plotly_chart(fig, use_container_width=True)
                                charts_html.append(('04_Rente_Kumuliert.html', fig.to_html()))
                            
                            # Nach Abteilung
                            if col_abt:
                                st.markdown("### Nach Abteilung:")
                                c1, c2 = st.columns(2)
                                
                                with c1:
                                    avg = d.groupby(col_abt)['Alter'].mean().sort_values()
                                    fig = go.Figure(go.Bar(
                                        y=list(avg.index), 
                                        x=[round(float(x),1) for x in avg.values],
                                        orientation='h',
                                        marker_color=['#e74c3c' if a>=50 else '#f39c12' if a>=45 else '#2ecc71' for a in avg.values],
                                        text=[f"{x:.1f} Jahre" for x in avg.values],
                                        textposition='outside'
                                    ))
                                    fig.update_layout(title="Durchschnittsalter pro Abteilung", height=500, font=dict(size=14))
                                    st.plotly_chart(fig, use_container_width=True)
                                    charts_html.append(('05_Alter_Abteilung.html', fig.to_html()))
                                
                                with c2:
                                    ab5 = d[d['JbR']<=5].groupby(col_abt).size()
                                    ges = d.groupby(col_abt).size()
                                    pct = (ab5/ges*100).fillna(0).sort_values()
                                    fig = go.Figure(go.Bar(
                                        y=list(pct.index),
                                        x=[round(float(x),1) for x in pct.values],
                                        orientation='h',
                                        marker_color=['#e74c3c' if p>=30 else '#f39c12' if p>=15 else '#2ecc71' for p in pct.values],
                                        text=[f"{x:.0f}%" for x in pct.values],
                                        textposition='outside'
                                    ))
                                    fig.update_layout(title="Wer verliert in 5 Jahren wie viel?", height=500, font=dict(size=14))
                                    st.plotly_chart(fig, use_container_width=True)
                                    charts_html.append(('06_Abgang_Abteilung.html', fig.to_html()))
                    
                    # ============================================
                    # TREUE / JUBILÄEN
                    # ============================================
                    if 'DJ' in df.columns:
                        d = df[df['DJ'].notna()].copy()
                        if len(d) > 0:
                            st.markdown("---")
                            st.markdown("## 🏆 Wie lange sind Ihre Mitarbeiter dabei?")
                            
                            lang = len(d[d['DJ']>=20])
                            st.info(f"ℹ️ **{lang} Mitarbeiter** sind schon **20 Jahre oder länger** bei Ihnen!")
                            
                            c1, c2 = st.columns(2)
                            
                            with c1:
                                fig = go.Figure(go.Histogram(x=list(d['DJ']), nbinsx=20, marker_color='#9b59b6'))
                                fig.update_layout(title="Betriebszugehörigkeit (Jahre)", height=450, font=dict(size=14),
                                    xaxis_title="Jahre im Unternehmen", yaxis_title="Anzahl Mitarbeiter")
                                st.plotly_chart(fig, use_container_width=True)
                                charts_html.append(('07_Dienstjahre.html', fig.to_html()))
                            
                            with c2:
                                d['Gr'] = pd.cut(d['DJ'],[-1,2,5,10,15,20,100],
                                    labels=['Neu (0-2 J)','3-5 Jahre','6-10 Jahre','11-15 Jahre','16-20 Jahre','Über 20 Jahre'])
                                gr = d['Gr'].value_counts()
                                fig = go.Figure(go.Pie(
                                    labels=[str(x) for x in gr.index], 
                                    values=list(gr.values),
                                    marker_colors=F,
                                    textinfo='label+percent+value',
                                    textfont_size=14
                                ))
                                fig.update_layout(title="Gruppen nach Betriebszugehörigkeit", height=450, font=dict(size=14))
                                st.plotly_chart(fig, use_container_width=True)
                                charts_html.append(('08_Dienstjahre_Gruppen.html', fig.to_html()))
                            
                            # Jubiläen
                            st.markdown("### 🎉 Wer hat bald Jubiläum?")
                            c1, c2 = st.columns(2)
                            
                            with c1:
                                jubi_data = []
                                for j in [5,10,15,20,25,30]:
                                    count = int(len(d[(d['DJ']>=j-0.5)&(d['DJ']<j+0.5)]))
                                    jubi_data.append({'Jahre': f'{j} Jahre', 'Anzahl': count})
                                
                                fig = go.Figure(go.Bar(
                                    x=[f'{j} Jahre' for j in [5,10,15,20,25,30]], 
                                    y=[int(len(d[(d['DJ']>=j-0.5)&(d['DJ']<j+0.5)])) for j in [5,10,15,20,25,30]], 
                                    marker_color=F[:6],
                                    text=[int(len(d[(d['DJ']>=j-0.5)&(d['DJ']<j+0.5)])) for j in [5,10,15,20,25,30]],
                                    textposition='outside'
                                ))
                                fig.update_layout(title="Mitarbeiter mit rundem Jubiläum", height=400, font=dict(size=14))
                                st.plotly_chart(fig, use_container_width=True)
                                charts_html.append(('09_Jubilaeen.html', fig.to_html()))
                            
                            with c2:
                                if col_abt:
                                    avg = d.groupby(col_abt)['DJ'].mean().sort_values()
                                    fig = go.Figure(go.Bar(
                                        y=list(avg.index),
                                        x=[round(float(x),1) for x in avg.values],
                                        orientation='h',
                                        marker_color=['#27ae60' if x>=10 else '#f39c12' if x>=5 else '#e74c3c' for x in avg.values],
                                        text=[f"{x:.1f} J" for x in avg.values],
                                        textposition='outside'
                                    ))
                                    fig.update_layout(title="Durchschnitt pro Abteilung", height=400, font=dict(size=14))
                                    st.plotly_chart(fig, use_container_width=True)
                                    charts_html.append(('10_DJ_Abteilung.html', fig.to_html()))
                    
                    # ============================================
                    # WISSENSVERLUST
                    # ============================================
                    if 'Alter' in df.columns and 'DJ' in df.columns:
                        d = df[(df['Alter'].notna())&(df['DJ'].notna())].copy()
                        if len(d) > 0:
                            d['JbR'] = rentenalter - d['Alter']
                            d['RJ'] = jahr + d['JbR']
                            d['R'] = d.apply(lambda r:'KRITISCH' if r['DJ']>=15 and r['JbR']<=5 
                                else ('WARNUNG' if r['DJ']>=10 and r['JbR']<=10 else 'OK'),axis=1)
                            
                            st.markdown("---")
                            st.markdown("## ⚠️ Droht Ihnen Wissensverlust?")
                            
                            st.markdown("""
                            <div class="help-text">
                            <b>Was bedeutet das?</b><br><br>
                            🔴 <b>KRITISCH</b> = Mitarbeiter mit 15+ Jahren Erfahrung, die in 5 Jahren gehen<br>
                            🟡 <b>WARNUNG</b> = Mitarbeiter mit 10+ Jahren Erfahrung, die in 10 Jahren gehen<br>
                            🟢 <b>OK</b> = Noch genug Zeit für Wissenstransfer
                            </div>
                            """, unsafe_allow_html=True)
                            
                            krit = len(d[d['R']=='KRITISCH'])
                            warn = len(d[d['R']=='WARNUNG'])
                            verl5 = d[d['JbR']<=5]['DJ'].sum()
                            
                            m1, m2, m3 = st.columns(3)
                            with m1:
                                if krit > 0:
                                    st.error(f"🔴 **{krit}** KRITISCH")
                                else:
                                    st.success(f"🔴 **{krit}** KRITISCH")
                            with m2:
                                if warn > 0:
                                    st.warning(f"🟡 **{warn}** WARNUNG")
                                else:
                                    st.success(f"🟡 **{warn}** WARNUNG")
                            with m3:
                                st.metric("📉 Erfahrungsjahre die verloren gehen", f"{verl5:.0f} Jahre")
                            
                            c1, c2 = st.columns(2)
                            
                            with c1:
                                fig = go.Figure()
                                cm = {'KRITISCH':'#e74c3c','WARNUNG':'#f39c12','OK':'#2ecc71'}
                                for r in ['OK','WARNUNG','KRITISCH']:
                                    x = d[d['R']==r]
                                    if len(x)>0:
                                        fig.add_trace(go.Scatter(
                                            x=list(x['JbR']), y=list(x['DJ']),
                                            mode='markers', name=r,
                                            marker=dict(color=cm[r], size=14, opacity=0.7)
                                        ))
                                fig.update_layout(
                                    title="Jeder Punkt = 1 Mitarbeiter",
                                    xaxis_title="Jahre bis zur Rente →",
                                    yaxis_title="Jahre Erfahrung ↑",
                                    height=500, font=dict(size=14),
                                    legend=dict(font=dict(size=16))
                                )
                                st.plotly_chart(fig, use_container_width=True)
                                charts_html.append(('11_Wissensverlust.html', fig.to_html()))
                            
                            with c2:
                                verlust = {j: d[d['RJ']==j]['DJ'].sum() for j in range(jahr, jahr+11)}
                                fig = go.Figure(go.Bar(
                                    x=list(verlust.keys()),
                                    y=list(verlust.values()),
                                    marker_color=['#e74c3c' if v>d['DJ'].sum()*0.1 else '#f39c12' for v in verlust.values()],
                                    text=[f"{v:.0f}" for v in verlust.values()],
                                    textposition='outside'
                                ))
                                fig.update_layout(title="Wie viel Erfahrung geht wann verloren?", height=500, font=dict(size=14),
                                    xaxis_title="Jahr", yaxis_title="Verlorene Erfahrungsjahre")
                                st.plotly_chart(fig, use_container_width=True)
                                charts_html.append(('12_Verlust_Pro_Jahr.html', fig.to_html()))
                    
                    # ============================================
                    # KARRIEREENTWICKLUNG
                    # ============================================
                    if col_ein_pos and col_akt_pos:
                        st.markdown("---")
                        st.markdown("## 📈 Wie haben sich Ihre Mitarbeiter entwickelt?")
                        
                        # Beispiele zeigen
                        d = df[(df[col_ein_pos].notna()) & (df[col_akt_pos].notna())].copy()
                        if 'DJ' in d.columns:
                            beispiele = d[d['DJ'] >= 5].head(10)
                            if len(beispiele) > 0:
                                st.markdown("### Beispiele (mindestens 5 Jahre dabei):")
                                for _, row in beispiele.iterrows():
                                    dj = int(row['DJ']) if 'DJ' in row else '?'
                                    st.markdown(f"- **{row[col_ein_pos]}** → **{row[col_akt_pos]}** *(nach {dj} Jahren)*")
                        
                        # Sankey wenn Level vorhanden
                        if col_lvl and col_abt:
                            st.markdown("### Karriere-Flow: Von Abteilung zu Level")
                            flow = df.groupby([col_abt, col_lvl]).size().reset_index(name='count')
                            flow = flow[flow['count'] > 0]
                            if len(flow) > 0:
                                abt_list = list(flow[col_abt].unique())
                                lvl_list = list(flow[col_lvl].unique())
                                nodes = abt_list + lvl_list
                                fig = go.Figure(go.Sankey(
                                    node=dict(pad=15, thickness=20, label=nodes, 
                                        color=['#3498db']*len(abt_list)+['#e74c3c']*len(lvl_list)),
                                    link=dict(
                                        source=[nodes.index(a) for a in flow[col_abt]],
                                        target=[nodes.index(l) for l in flow[col_lvl]],
                                        value=list(flow['count'])
                                    )
                                ))
                                fig.update_layout(title="Wer arbeitet auf welchem Level?", height=600, font=dict(size=14))
                                st.plotly_chart(fig, use_container_width=True)
                                charts_html.append(('13_Karriere_Flow.html', fig.to_html()))
                    
                    # ============================================
                    # WEITERE ANALYSEN
                    # ============================================
                    st.markdown("---")
                    st.markdown("## 📊 Weitere Auswertungen")
                    
                    c1, c2 = st.columns(2)
                    
                    with c1:
                        if col_ges:
                            c = df[col_ges].value_counts()
                            labels = []
                            for x in c.index:
                                if str(x).lower() in ['m','männlich','male']:
                                    labels.append('Männlich')
                                elif str(x).lower() in ['w','weiblich','female']:
                                    labels.append('Weiblich')
                                else:
                                    labels.append(str(x))
                            fig = go.Figure(go.Pie(
                                labels=labels, 
                                values=list(c.values), 
                                marker_colors=['#3498db','#e74c3c','#2ecc71'],
                                textinfo='label+percent+value',
                                textfont_size=16
                            ))
                            fig.update_layout(title="Geschlechterverteilung", height=450, font=dict(size=16))
                            st.plotly_chart(fig, use_container_width=True)
                            charts_html.append(('14_Geschlecht.html', fig.to_html()))
                    
                    with c2:
                        if col_abt:
                            c = df[col_abt].value_counts()
                            fig = go.Figure(go.Bar(
                                y=list(c.index), 
                                x=list(c.values), 
                                orientation='h', 
                                marker_color=F[:len(c)],
                                text=list(c.values),
                                textposition='outside'
                            ))
                            fig.update_layout(title="Mitarbeiter pro Abteilung", height=450, font=dict(size=14))
                            st.plotly_chart(fig, use_container_width=True)
                            charts_html.append(('15_Abteilungen.html', fig.to_html()))
                    
                    c1, c2 = st.columns(2)
                    
                    with c1:
                        if col_lvl:
                            c = df[col_lvl].value_counts()
                            fig = go.Figure(go.Bar(
                                x=list(c.index), 
                                y=list(c.values), 
                                marker_color=F[:len(c)],
                                text=list(c.values),
                                textposition='outside'
                            ))
                            fig.update_layout(title="Karrierelevel", height=400, font=dict(size=14))
                            st.plotly_chart(fig, use_container_width=True)
                            charts_html.append(('16_Level.html', fig.to_html()))
                    
                    with c2:
                        if col_az:
                            c = df[col_az].value_counts()
                            fig = go.Figure(go.Pie(
                                labels=list(c.index), 
                                values=list(c.values), 
                                marker_colors=['#3498db','#f39c12'],
                                textinfo='label+percent+value',
                                textfont_size=16
                            ))
                            fig.update_layout(title="Vollzeit / Teilzeit", height=400, font=dict(size=16))
                            st.plotly_chart(fig, use_container_width=True)
                            charts_html.append(('17_Arbeitszeit.html', fig.to_html()))
                    
                    c1, c2 = st.columns(2)
                    
                    with c1:
                        if 'Gehalt' in df.columns:
                            g = df['Gehalt'].dropna()
                            if len(g)>0:
                                fig = go.Figure(go.Histogram(x=list(g), nbinsx=20, marker_color='#2ecc71'))
                                fig.update_layout(title="Gehaltsverteilung", height=400, font=dict(size=14),
                                    xaxis_title="Jahresgehalt in €", yaxis_title="Anzahl")
                                st.plotly_chart(fig, use_container_width=True)
                                charts_html.append(('18_Gehalt.html', fig.to_html()))
                    
                    with c2:
                        if col_ort:
                            c = df[col_ort].value_counts()
                            fig = go.Figure(go.Pie(
                                labels=list(c.index), 
                                values=list(c.values), 
                                marker_colors=F,
                                textinfo='label+percent+value',
                                textfont_size=14
                            ))
                            fig.update_layout(title="Standorte", height=400, font=dict(size=14))
                            st.plotly_chart(fig, use_container_width=True)
                            charts_html.append(('19_Standorte.html', fig.to_html()))
                    
                    # ============================================
                    # BENCHMARK
                    # ============================================
                    if region in BENCHMARK:
                        st.markdown("---")
                        st.markdown(f"## 📊 Vergleich: Sie vs. {region}")
                        
                        b = BENCHMARK[region]
                        kat, u, bm = [], [], []
                        
                        if 'Alter' in df.columns:
                            a = df['Alter'].mean()
                            if not pd.isna(a): 
                                kat.append('Durchschnittsalter')
                                u.append(round(float(a),1))
                                bm.append(b['alter'])
                        
                        if col_ges:
                            g = df[col_ges].value_counts(normalize=True)*100
                            f = sum(float(v) for k,v in g.items() if str(k).lower() in ['w','weiblich'])
                            kat.append('Frauenanteil %')
                            u.append(round(f,1))
                            bm.append(b['frauen'])
                        
                        if col_az:
                            a = df[col_az].value_counts(normalize=True)*100
                            t = sum(float(v) for k,v in a.items() if 'teil' in str(k).lower())
                            kat.append('Teilzeitquote %')
                            u.append(round(t,1))
                            bm.append(b['teilzeit'])
                        
                        if 'Gehalt' in df.columns:
                            a = df['Gehalt'].mean()
                            if not pd.isna(a): 
                                kat.append('Monatsgehalt €')
                                u.append(round(float(a)/12,0))
                                bm.append(b['gehalt'])
                        
                        if kat:
                            fig = go.Figure([
                                go.Bar(name='Ihr Unternehmen', x=kat, y=u, marker_color='#3498db', 
                                    text=u, textposition='outside'),
                                go.Bar(name=region, x=kat, y=bm, marker_color='#95a5a6',
                                    text=bm, textposition='outside')
                            ])
                            fig.update_layout(title=f"Ihre Zahlen im Vergleich zu {region}", 
                                barmode='group', height=500, font=dict(size=16),
                                legend=dict(font=dict(size=18)))
                            st.plotly_chart(fig, use_container_width=True)
                            charts_html.append(('20_Benchmark.html', fig.to_html()))
                    
                    # ============================================
                    # DOWNLOAD
                    # ============================================
                    if charts_html:
                        st.markdown("---")
                        st.markdown("## 📥 Alle Diagramme speichern")
                        
                        st.success(f"✅ **{len(charts_html)} Diagramme** wurden erstellt!")
                        
                        st.markdown("""
                        <div class="info-box">
                        <b>💡 Tipp:</b> Klicken Sie auf den Button unten, um alle Diagramme als ZIP-Datei zu speichern.<br>
                        Die Diagramme sind HTML-Dateien und können im Browser geöffnet werden.
                        </div>
                        """, unsafe_allow_html=True)
                        
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                            for name, html in charts_html:
                                zf.writestr(name, html)
                        
                        st.download_button(
                            label="📥  ALLE DIAGRAMME HERUNTERLADEN (ZIP)",
                            data=zip_buffer.getvalue(),
                            file_name="HR_Analyse_Ergebnisse.zip",
                            mime="application/zip",
                            use_container_width=True
                        )
                        
                        st.markdown("---")
                        st.balloons()
                        st.success("🎉 **Fertig!** Ihre Analyse ist abgeschlossen.")
                    
            except Exception as e:
                st.error(f"❌ Es ist ein Fehler aufgetreten: {e}")
                st.markdown("""
                <div class="help-text">
                <b>Was kann ich tun?</b><br><br>
                • Prüfen Sie, ob die Excel-Datei das richtige Format hat (.xlsx)<br>
                • Stellen Sie sicher, dass die Datei nicht geöffnet ist<br>
                • Versuchen Sie es mit der Vorlage aus Schritt 1
                </div>
                """, unsafe_allow_html=True)

# ============================================
# FOOTER
# ============================================
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #7f8c8d; font-size: 16px;">
<b>HR Analyse Tool</b><br>
🎯 Rentenanalyse • 🏆 Jubiläen • ⚠️ Wissensverlust • 📈 Karriereentwicklung • 📊 Benchmark
</div>
""", unsafe_allow_html=True)
