
import streamlit as st, os

st.set_page_config(page_title="HUB ANA - 2 Progetti", page_icon="🏠", layout="wide")

st.markdown("""
<style>
.stButton>button{border-radius:14px !important; font-weight:900 !important; height:150px !important; font-size:18px !important; border:3px solid #FFD700 !important; background:white !important;}
#MainMenu{visibility:hidden;} footer{visibility:hidden;}
.stApp{background:#e8f5e9 !important;}
</style>
""", unsafe_allow_html=True)

# SPLASH iniziale HUB
if "hub_splash_shown" not in st.session_state:
    st.session_state.hub_splash_shown = False

if not st.session_state.hub_splash_shown:
    st.markdown("""
    <div style="text-align:center;padding:40px;background:linear-gradient(135deg,#1A5D1A,#2e7d32);border-radius:16px;color:white;border:4px solid #FFD700;">
    <div style="font-size:100px;">🛡️</div>
    <h1>ANA Varese</h1>
    <h2>Protezione Civile - Caronno Pertusella</h2>
    <p>Gestionale ODV 2026 - 2 Progetti Separati</p>
    </div>
    """, unsafe_allow_html=True)
    st.write("")
    if st.button("🎬 ENTRA NEL HUB", type="primary", use_container_width=True):
        st.session_state.hub_splash_shown = True
        st.rerun()
    st.stop()

with st.sidebar:
    try:
        if os.path.exists("logo.png"):
            st.image("logo.png", width=110)
        else:
            st.markdown("<div style='width:110px;height:110px;background:#1A5D1A;border-radius:10px;display:flex;align-items:center;justify-content:center;color:white;font-weight:bold'>ANA</div>", unsafe_allow_html=True)
    except:
        st.write("ANA")
    st.markdown("<p style='font-weight:bold;color:#1A5D1A'>HUB - 2 Progetti</p>", unsafe_allow_html=True)
    if st.button("🎬 MOSTRA SPLASH", use_container_width=True):
        st.session_state.hub_splash_shown = False
        st.rerun()
    st.caption("Senza Geolocalizzazione - Con Splash")

st.markdown("""
<div style="text-align:center;padding:30px;background:linear-gradient(135deg,#1A5D1A,#2e7d32);border-radius:15px;color:white;">
<h1>🏠 HUB ANA - 2 Progetti Separati</h1>
<h3>Amministrativo (13 form con Splash) + Operativo (17 form con Splash)</h3>
<p>Senza Geolocalizzazione - File base App-Con-Spese_89.py - Dashboard fixata</p>
</div>
""", unsafe_allow_html=True)

st.divider()
c1,c2 = st.columns(2, gap="large")
with c1:
    st.markdown('<div style="text-align:center;background:#1565C0;color:white;padding:16px;border-radius:12px;font-weight:bold;font-size:18px">📋 AMMINISTRATIVO<br><small>13 form - Splash, Volontari, Ospiti, Documenti, Spese, Note Spese, Statistiche, Backup</small><br><small>SENZA Geoloc</small></div>', unsafe_allow_html=True)
    st.write("")
    if st.button("📋 APRI AMMINISTRATIVO\n13 form con Splash", key="open_admin", use_container_width=True):
        st.code("streamlit run App-ANA-AMMINISTRATIVO.py")
        st.success("File: App-ANA-AMMINISTRATIVO.py")

with c2:
    st.markdown('<div style="text-align:center;background:#1A5D1A;color:white;padding:16px;border-radius:12px;font-weight:bold;font-size:18px">🚒 OPERATIVO<br><small>17 form - Splash, Radio, Emergenze, Mezzi, Mappe, Turni</small><br><small>SENZA Geoloc</small></div>', unsafe_allow_html=True)
    st.write("")
    if st.button("🚒 APRI OPERATIVO\n17 form con Splash", key="open_oper", use_container_width=True):
        st.code("streamlit run App-ANA-OPERATIVO.py")
        st.success("File: App-ANA-OPERATIVO.py")

st.divider()
colA,colB = st.columns(2)
with colA:
    st.markdown("#### 📋 AMMINISTRATIVO (13) - CON SPLASH - SENZA GEOLOC")
    st.markdown("""
    - Splash (NUOVO)
    - Dashboard (tasti fixati)
    - Volontari (con foto)
    - Ospiti
    - Verbali
    - Archivio Documenti
    - Diplomi Attestati
    - Spese ODV per Evento
    - Note Spese
    - Report Filtro
    - Statistiche
    - Gestione Utenti
    - Backup
    """)
with colB:
    st.markdown("#### 🚒 OPERATIVO (17) - CON SPLASH - SENZA GEOLOC")
    st.markdown("""
    - Splash (NUOVO)
    - Dashboard (tasti fixati)
    - DB Radio
    - Consegna Radio
    - Alias Radio
    - Brogliaccio
    - Eventi
    - Emergenze
    - Tabella Emergenze
    - Check-in
    - Interventi Emergenza
    - Tabella Interventi Emergenza
    - Mezzi
    - Attrezzature
    - Mappe Postazioni
    - Libreria Icone
    - Turni
    - Chat
    """)

st.success("✅ Geolocalizzazione RIMOSSA + Form Splash AGGIUNTO")
