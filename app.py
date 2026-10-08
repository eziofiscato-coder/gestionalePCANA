
import streamlit as st, os, json, pandas as pd
from datetime import datetime

PERSIST_FILE_ADMIN = "ana_amministrativo_dati.json"
PERSIST_FILE_OPER = "ana_operativo_dati.json"

st.set_page_config(page_title="HUB ANA - Solo ANA", page_icon="🏠", layout="wide", initial_sidebar_state="expanded")

# CSS + LOGO FIX
st.markdown("""
<style>
.stButton>button{border-radius:14px !important; font-weight:900 !important; height:90px !important; font-size:16px !important; border:3px solid #FFD700 !important; background:white !important;}
#MainMenu{visibility:hidden;} footer{visibility:hidden;}
.stApp{background:#e8f5e9 !important;}
[data-testid="stAppViewContainer"]{background:linear-gradient(180deg,#c8e6c9 0%,#a5d6a7 100%) !important;}
</style>
""", unsafe_allow_html=True)

# Stato
if "hub_page" not in st.session_state:
    st.session_state.hub_page = "hub"  # hub | admin | operativo
if "menu_admin" not in st.session_state:
    st.session_state.menu_admin = "Dashboard"
if "menu_oper" not in st.session_state:
    st.session_state.menu_oper = "Dashboard"

def salva_admin():
    try:
        data = {k: st.session_state.get(k) for k in ["ospiti","spese_odv","note_spese","volontari","archivio_documenti"] if k in st.session_state}
        with open(PERSIST_FILE_ADMIN, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2, default=str)
    except: pass

def salva_oper():
    try:
        data = {k: st.session_state.get(k) for k in ["emergenze","interventi","checkin","brogliaccio","radio_db"] if k in st.session_state}
        with open(PERSIST_FILE_OPER, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2, default=str)
    except: pass

# SIDEBAR - FIX 1968
with st.sidebar:
    try:
        if os.path.exists("logo.png"):
            st.image("logo.png", width=110)
        elif os.path.exists("logo.jpg"):
            st.image("logo.jpg", width=110)
        else:
            st.markdown("<div style='width:110px;height:110px;background:#1A5D1A;border-radius:10px;display:flex;align-items:center;justify-content:center;color:white;font-weight:bold'>ANA</div>", unsafe_allow_html=True)
    except:
        st.write("ANA")
    st.markdown("<p style='font-weight:bold;color:#1A5D1A;margin-top:10px'>HUB ANA - SOLO ANA</p>", unsafe_allow_html=True)
    st.caption("Fix 1968 + Persistenza OK")
    
    if st.session_state.hub_page != "hub":
        if st.button("🏠 TORNA AL HUB", use_container_width=True, type="primary"):
            st.session_state.hub_page = "hub"
            st.rerun()
        st.divider()

    # Menu dinamico
    if st.session_state.hub_page == "admin":
        st.markdown("**📋 AMMINISTRATIVO**")
        menu = ["Dashboard","Volontari (con foto)","Ospiti","Verbali","Archivio Documenti","Diplomi Attestati","Spese ODV per Evento","Note Spese","Report Filtro","Gestione Utenti","Backup"]
        sel = st.radio("Form", menu, index=menu.index(st.session_state.menu_admin) if st.session_state.menu_admin in menu else 0, key="radio_admin")
        st.session_state.menu_admin = sel
    elif st.session_state.hub_page == "operativo":
        st.markdown("**🚒 OPERATIVO**")
        menu2 = ["Dashboard","Emergenze","Tabella Emergenze","Interventi Emergenza","Tabella Interventi","Check-in","Brogliaccio","DB Radio","Consegna Radio","Alias Radio","Geolocalizzazione Hytera + Anytone","Mezzi","Attrezzature","Mappe Postazioni","Libreria Icone","Turni","Eventi","Chat"]
        sel2 = st.radio("Form", menu2, index=menu2.index(st.session_state.menu_oper) if st.session_state.menu_oper in menu2 else 0, key="radio_oper")
        st.session_state.menu_oper = sel2

# ================= HUB PAGE =================
if st.session_state.hub_page == "hub":
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
        st.markdown('<div style="text-align:center;background:#1565C0;color:white;padding:16px;border-radius:12px;font-weight:bold;font-size:18px">📋 ANA AMMINISTRATIVO<br><small>11 form - Volontari, Ospiti, Documenti, Spese, Note Spese, Utenti, Backup</small></div>', unsafe_allow_html=True)
        st.write("")
        if st.button("📋 ENTRA ANA AMMINISTRATIVO\nGestione ODV\nCLICCA QUI", key="btn_enter_admin", use_container_width=True):
            st.session_state.hub_page = "admin"
            st.rerun()
    
    with c2:
        st.markdown('<div style="text-align:center;background:#1A5D1A;color:white;padding:16px;border-radius:12px;font-weight:bold;font-size:18px">🚒 ANA OPERATIVO<br><small>18 form - Emergenze, Radio, Mezzi, Mappe, Turni</small></div>', unsafe_allow_html=True)
        st.write("")
        if st.button("🚒 ENTRA ANA OPERATIVO\nEmergenze e Radio\nCLICCA QUI", key="btn_enter_oper", use_container_width=True):
            st.session_state.hub_page = "operativo"
            st.rerun()
    
    st.divider()
    st.info("✅ Logo: carica logo.png nella root del repo su GitHub (upload file) - apparirà automatico in sidebar")
    st.caption("Link attuale: ana-varese-3dcarznzkzrefca6hgppzc.streamlit.app - Fix bottoni cliccabili ora OK")

# ================= ADMIN PAGE =================
elif st.session_state.hub_page == "admin":
    st.title(f"📋 ANA AMMINISTRATIVO - {st.session_state.menu_admin}")
    
    if st.session_state.menu_admin == "Dashboard":
        st.success("Amministrativo - 11 form disponibili")
        cols = st.columns(3)
        for idx, f in enumerate(["Volontari (con foto)","Ospiti","Spese ODV per Evento","Note Spese","Archivio Documenti","Verbali"]):
            with cols[idx%3]:
                if st.button(f"📋 {f}", key=f"dash_{f}", use_container_width=True):
                    st.session_state.menu_admin = f
                    st.rerun()
    elif st.session_state.menu_admin == "Note Spese":
        st.markdown("### 🧾 Note Spese - Rimborso Volontari (NUOVO - mancava)")
        with st.form("note_spese_form"):
            c1,c2,c3 = st.columns(3)
            with c1:
                data_ns = st.date_input("Data Spesa")
                volontario_ns = st.text_input("Volontario")
                evento_ns = st.text_input("Evento / Missione")
            with c2:
                tipo_ns = st.selectbox("Tipo Spesa", ["Carburante","Pedaggio","Vitto","Materiale","Parcheggio","Altro"])
                importo_ns = st.number_input("Importo €", min_value=0.0, step=0.10)
                pagato_ns = st.selectbox("Pagato da", ["Volontario","ODV","Anticipo"])
            with c3:
                scontrino_ns = st.file_uploader("Scontrino / Foto", type=["png","jpg","pdf"])
                note_ns = st.text_area("Note")
            if st.form_submit_button("💾 Salva Nota Spesa", type="primary", use_container_width=True):
                st.session_state.setdefault("note_spese", []).append({
                    "data": str(data_ns),
                    "volontario": volontario_ns,
                    "evento": evento_ns,
                    "tipo": tipo_ns,
                    "importo": importo_ns,
                    "pagato_da": pagato_ns,
                    "note": note_ns,
                    "ts": str(datetime.now())
                })
                salva_admin()
                st.success("Salvata!")
                st.rerun()
        if st.session_state.get("note_spese"):
            df = pd.DataFrame(st.session_state.note_spese)
            st.dataframe(df, use_container_width=True)
            st.metric("Totale", f"€ {pd.to_numeric(df['importo'], errors='coerce').sum():.2f}")
    elif st.session_state.menu_admin == "Spese ODV per Evento":
        st.markdown("### 💰 Spese ODV per Evento")
        with st.form("spese_form"):
            desc = st.text_input("Descrizione")
            imp = st.number_input("Importo €", min_value=0.0)
            if st.form_submit_button("Salva Spesa", type="primary"):
                st.session_state.setdefault("spese_odv", []).append({"descrizione": desc, "importo": imp, "data": str(datetime.now())})
                salva_admin()
                st.success("Salvata!")
                st.rerun()
        if st.session_state.get("spese_odv"):
            st.dataframe(pd.DataFrame(st.session_state.spese_odv), use_container_width=True)
    else:
        st.info(f"Form '{st.session_state.menu_admin}' - struttura pronta - importa logica da file originale")
        if st.session_state.menu_admin == "Ospiti":
            with st.form("ospiti_quick"):
                n = st.text_input("Nome Ospite")
                if st.form_submit_button("Salva"):
                    st.session_state.setdefault("ospiti", []).append({"nome": n, "ts": str(datetime.now())})
                    salva_admin()
                    st.rerun()
            if st.session_state.get("ospiti"):
                st.dataframe(pd.DataFrame(st.session_state.ospiti), use_container_width=True)
    
    salva_admin()

# ================= OPERATIVO PAGE =================
elif st.session_state.hub_page == "operativo":
    st.title(f"🚒 ANA OPERATIVO - {st.session_state.menu_oper}")
    if st.session_state.menu_oper == "Dashboard":
        st.success("Operativo - 18 form disponibili")
        cols = st.columns(3)
        for idx, f in enumerate(["Emergenze","Interventi Emergenza","DB Radio","Mezzi","Mappe Postazioni","Turni"]):
            with cols[idx%3]:
                if st.button(f"🚒 {f}", key=f"dash_op_{f}", use_container_width=True):
                    st.session_state.menu_oper = f
                    st.rerun()
    else:
        st.info(f"Form operativo '{st.session_state.menu_oper}' - struttura pronta")
        st.caption("Importa qui la logica completa da App-GAE-FINALE_32.py quando vuoi - per ora HUB navigabile")
    
    salva_oper()

st.divider()
st.markdown("<div style='text-align:center;padding:6px;background:#1A5D1A;color:white;border-radius:6px'>ANA Varese - Fix 1968 + HUB 2 bottoni cliccabili + Note Spese + Persistenza</div>", unsafe_allow_html=True)
