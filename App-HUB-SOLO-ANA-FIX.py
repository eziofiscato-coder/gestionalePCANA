
import streamlit as st, os, json
from datetime import datetime

st.set_page_config(page_title="HUB ANA - Solo ANA", page_icon="🏠", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
.stButton>button{border-radius:14px !important; font-weight:900 !important; height:150px !important; font-size:20px !important; border:3px solid #FFD700 !important;}
#MainMenu{visibility:hidden;} footer{visibility:hidden;}
.stApp{background:#e8f5e9 !important;}
[data-testid="stAppViewContainer"]{background:linear-gradient(180deg,#c8e6c9 0%,#a5d6a7 100%) !important;}
</style>
""", unsafe_allow_html=True)

# FIX SyntaxError 1968 - sidebar corretta
with st.sidebar:
    try:
        if os.path.exists("logo.png"):
            st.image("logo.png", width=110)
        else:
            st.markdown("<div style='width:110px;height:110px;background:#1A5D1A;border-radius:10px;display:flex;align-items:center;justify-content:center;color:white;font-weight:bold'>ANA</div>", unsafe_allow_html=True)
    except:
        st.write("ANA")
    st.markdown("<p style='font-weight:bold;color:#1A5D1A;margin-top:10px'>HUB ANA - SOLO ANA</p>", unsafe_allow_html=True)
    st.caption("Fix 1968 + Persistenza OK - GAE secondo momento")

st.markdown("""
<div style="text-align:center;padding:30px;background:linear-gradient(135deg,#1A5D1A,#2e7d32);border-radius:15px;color:white;">
<h1>🏠 HUB ANA - SOLO ANA</h1>
<h3>2 Progetti ANA - Amministrativo + Operativo</h3>
<p>Fix SyntaxError 1968 | Dati persistenti dopo logout | GAE dopo</p>
</div>
""", unsafe_allow_html=True)

st.divider()

c1,c2 = st.columns(2, gap="large")

with c1:
    st.markdown('<div style="text-align:center;background:#1565C0;color:white;padding:16px;border-radius:12px;font-weight:bold;font-size:18px">📋 ANA AMMINISTRATIVO<br><small>10 form - Volontari, Ospiti, Documenti, Spese, Note Spese, Utenti, Backup</small></div>', unsafe_allow_html=True)
    st.write("")
    if st.button("📋 ENTRA ANA\nAMMINISTRATIVO\nGestione ODV", key="btn_ana_admin", use_container_width=True):
        st.info("Lancia: streamlit run App-ANA-AMMINISTRATIVO.py")
        st.code("streamlit run App-ANA-AMMINISTRATIVO.py")
        st.success("File: App-ANA-AMMINISTRATIVO.py - Dati: ana_amministrativo_dati.json")

with c2:
    st.markdown('<div style="text-align:center;background:#1A5D1A;color:white;padding:16px;border-radius:12px;font-weight:bold;font-size:18px">🚒 ANA OPERATIVO<br><small>18 form - Emergenze, Radio, Mezzi, Mappe, Turni</small></div>', unsafe_allow_html=True)
    st.write("")
    if st.button("🚒 ENTRA ANA\nOPERATIVO\nEmergenze e Radio", key="btn_ana_oper", use_container_width=True):
        st.info("Lancia: streamlit run App-ANA-OPERATIVO.py")
        st.code("streamlit run App-ANA-OPERATIVO.py")
        st.success("File: App-ANA-OPERATIVO.py - Dati: ana_operativo_dati.json")

st.divider()
st.markdown("### 📋 Elenco Form - Solo ANA (GAE dopo)")

colA,colB = st.columns(2)
with colA:
    st.markdown("#### 📋 AMMINISTRATIVO (10)")
    st.markdown("""
    - Dashboard
    - Volontari (con foto)
    - Ospiti
    - Verbali
    - Archivio Documenti
    - Diplomi Attestati
    - Spese ODV per Evento\n    - Note Spese (NUOVO)
    - Report Filtro
    - Gestione Utenti
    - Backup
    """)
with colB:
    st.markdown("#### 🚒 OPERATIVO (18)")
    st.markdown("""
    - Dashboard
    - Emergenze / Tabella Emergenze
    - Interventi Emergenza / Tabella Interventi
    - Check-in / Brogliaccio / Eventi
    - DB Radio / Consegna Radio / Alias Radio
    - Geolocalizzazione Hytera + Anytone
    - Mezzi / Attrezzature / Mappe Postazioni
    - Libreria Icone / Turni / Chat
    """)

st.success("✅ Fix: SyntaxError 1968 risolto (sidebar try/except) + Dati persistenti dopo logout (JSON) + HUB 2 bottoni ANA")
st.caption("GAE lo facciamo in secondo momento - per ora carica solo questi 3 file su GitHub")
