
import streamlit as st, os, json, pandas as pd
from datetime import datetime, date
import base64

st.set_page_config(page_title="ANA Varese - Dashboard Unica", page_icon="🛡️", layout="wide", initial_sidebar_state="expanded")

# Stato base
for k,v in {
    "page":"entra","logged":False,"menu":"Dashboard","hub_page":"hub",
    "volontari":[],"ospiti":[],"eventi":[],"emergenze":[],"checkin":[],"interventi":[],
    "mezzi":[],"attrezzature":[],"radio_db":[],"consegna_radio":[],"alias_radio":[],
    "brogliaccio":[],"turni":[],"chat":[],"spese_odv":[],"note_spese":[],"archivio_documenti":[],
    "verbali":[],"diplomi":[]
}.items():
    if k not in st.session_state:
        st.session_state[k]=v

MENU_ADMIN = ["Dashboard","Splash","Volontari (con foto)","Ospiti","Verbali","Archivio Documenti","Diplomi Attestati","Spese ODV per Evento","Note Spese","Report Filtro","Statistiche","Gestione Utenti","Backup"]
MENU_OPER = ["Dashboard","Splash","DB Radio","Consegna Radio","Alias Radio","Brogliaccio","Eventi","Emergenze","Tabella Emergenze","Check-in","Interventi Emergenza","Tabella Interventi Emergenza","Mezzi","Attrezzature","Mappe Postazioni","Libreria Icone","Turni","Chat"]

def hdr_form(tit):
    st.markdown(f"<div style='background:linear-gradient(135deg,#1A5D1A,#2e7d32);color:white;padding:10px;border-radius:8px;border:2px solid #FFD700;text-align:center'><b>{tit}</b></div>", unsafe_allow_html=True)
    st.write("")

def inject_popout_splash():
    try:
        st.markdown("""
        <div style="text-align:center;padding:30px;background:linear-gradient(135deg,#1A5D1A,#2e7d32);border-radius:16px;color:white;border:3px solid #FFD700;">
        <div style="font-size:80px;">🛡️</div><h1>ANA Varese</h1><h3>Protezione Civile - Caronno Pertusella</h3>
        </div>
        """, unsafe_allow_html=True)
    except:
        pass

# CSS
st.markdown("""
<style>
#MainMenu{visibility:hidden;} footer{visibility:hidden;}
.stApp{background:#e8f5e9 !important;}
.stButton>button{border-radius:8px;font-weight:600;}
</style>
""", unsafe_allow_html=True)

# PAGINA ENTRA - come primo file
if st.session_state.page == "entra":
    hdr_form("GESTIONALE DI PROTEZIONE CIVILE - Versione BETA")
    c1,c2,c3 = st.columns([1,2,1])
    with c2:
        try:
            if os.path.exists("copertina.png"):
                st.image("copertina.png", width=350)
            elif os.path.exists("logo.png"):
                st.image("logo.png", width=250)
            else:
                st.markdown('<div style="text-align:center;padding:40px;background:#c8e6c9;border-radius:16px;border:2px dashed #1A5D1A;"><div style="font-size:100px;">🛡️</div><p><b>Squadra Volontari di protezione civile</b></p></div>', unsafe_allow_html=True)
        except:
            pass
        st.markdown('<h2 style="text-align:center;">GESTIONALE DI PROTEZIONE CIVILE</h2><p style="text-align:center;">Versione BETA - Sviluppato per Ezio</p>', unsafe_allow_html=True)
        if st.button("ENTRA NEL GESTIONALE", type="primary", use_container_width=True):
            st.session_state.page = "login"
            st.rerun()
    st.divider()
    st.markdown('<div style="text-align:center;padding:6px;background:#1A5D1A;color:white;border-radius:6px">ANA Varese - Splash come primo file - Senza Geoloc</div>', unsafe_allow_html=True)
    st.stop()

# LOGIN SEMPLICE
if not st.session_state.logged:
    if st.session_state.page == "login":
        st.markdown("### Login")
        u = st.text_input("Username", value="admin")
        p = st.text_input("Password", type="password", value="ana2024")
        if st.button("Accedi", type="primary", use_container_width=True):
            if u=="admin" and p=="ana2024":
                st.session_state.logged=True
                st.session_state.page="dashboard"
                st.session_state.hub_page="hub"
                st.rerun()
            else:
                st.error("Credenziali errate")
        if st.button("Torna a Entra"):
            st.session_state.page="entra"
            st.rerun()
    st.stop()

# HUB - DIVISO ADMIN / OPERATIVO
if st.session_state.hub_page == "hub":
    inject_popout_splash()
    st.markdown("""
    <div style="text-align:center;padding:25px;background:linear-gradient(135deg,#1A5D1A,#2e7d32);border-radius:15px;color:white;border:3px solid #FFD700;">
    <h1>🏠 HUB ANA - Seleziona Progetto</h1>
    <p>Dashboard divisa - Solo bottoni amministrativi o operativi - Senza Geolocalizzazione</p>
    </div>
    """, unsafe_allow_html=True)
    st.divider()
    c1,c2 = st.columns(2, gap="large")
    with c1:
        st.markdown('<div style="text-align:center;background:#1565C0;color:white;padding:14px;border-radius:12px;font-weight:bold">📋 AMMINISTRATIVO<br><small>13 form</small></div>', unsafe_allow_html=True)
        st.write("")
        if st.button("📋 ENTRA AMMINISTRATIVO", key="hub_admin", use_container_width=True):
            st.session_state.hub_page="admin"
            st.session_state.menu="Dashboard"
            st.rerun()
    with c2:
        st.markdown('<div style="text-align:center;background:#1A5D1A;color:white;padding:14px;border-radius:12px;font-weight:bold">🚒 OPERATIVO<br><small>17 form</small></div>', unsafe_allow_html=True)
        st.write("")
        if st.button("🚒 ENTRA OPERATIVO", key="hub_oper", use_container_width=True):
            st.session_state.hub_page="operativo"
            st.session_state.menu="Dashboard"
            st.rerun()
    st.stop()

# SIDEBAR con filtro
with st.sidebar:
    try:
        if os.path.exists("logo.png"):
            st.image("logo.png", width=110)
        else:
            st.markdown('<div style="width:110px;height:110px;background:#1A5D1A;border-radius:10px;display:flex;align-items:center;justify-content:center;color:white;font-weight:bold">ANA</div>', unsafe_allow_html=True)
    except:
        st.write("ANA")
    st.markdown(f"<p style='font-weight:bold;color:#1A5D1A'>HUB: {st.session_state.hub_page.upper()}</p>", unsafe_allow_html=True)
    if st.button("🏠 TORNA AL HUB", use_container_width=True, type="primary"):
        st.session_state.hub_page="hub"
        st.session_state.menu="Dashboard"
        st.rerun()
    st.divider()
    if st.session_state.hub_page=="admin":
        menu_base = MENU_ADMIN
        st.markdown("**📋 AMMINISTRATIVO**")
    else:
        menu_base = MENU_OPER
        st.markdown("**🚒 OPERATIVO**")
    cur = st.radio("Seleziona form", menu_base, index=menu_base.index(st.session_state.menu) if st.session_state.menu in menu_base else 0)
    st.session_state.menu = cur
    st.divider()
    if st.button("Logout", use_container_width=True):
        st.session_state.logged=False
        st.session_state.page="entra"
        st.session_state.hub_page="hub"
        st.rerun()

# DASHBOARD con solo bottoni del progetto aperto
if cur == "Dashboard":
    hdr_form(f"DASHBOARD {st.session_state.hub_page.upper()} - Solo form del progetto")
    if st.session_state.hub_page=="admin":
        buttons = [("Splash","🎬 Splash"),("Volontari (con foto)","👤 Volontari"),("Ospiti","🧑‍🤝‍🧑 Ospiti"),("Verbali","📝 Verbali"),("Archivio Documenti","📁 Archivio"),("Diplomi Attestati","🏅 Diplomi"),("Spese ODV per Evento","💰 Spese ODV"),("Note Spese","🧾 Note Spese"),("Report Filtro","📊 Report"),("Statistiche","📈 Statistiche"),("Backup","💾 Backup")]
        st.success(f"📋 AMMINISTRATIVO - {len(buttons)} form - Solo amministrativi")
    else:
        buttons = [("Splash","🎬 Splash"),("Eventi","📅 Eventi"),("DB Radio","📻 DB Radio"),("Consegna Radio","🤝 Consegna"),("Alias Radio","🔖 Alias"),("Brogliaccio","📓 Brogliaccio"),("Emergenze","🚨 Emergenze"),("Tabella Emergenze","📋 Tab Emergenze"),("Check-in","✅ Check-in"),("Interventi Emergenza","🚒 Interventi"),("Tabella Interventi Emergenza","📋 Tab Interventi"),("Mezzi","🚐 Mezzi"),("Attrezzature","🧰 Attrezzature"),("Mappe Postazioni","🌍 Mappe"),("Libreria Icone","🎨 Icone"),("Turni","🕐 Turni"),("Chat","💬 Chat")]
        st.success(f"🚒 OPERATIVO - {len(buttons)} form - Solo operativi - Geoloc rimossa")
    cols = st.columns(3)
    for idx,(fkey,flabel) in enumerate(buttons):
        with cols[idx%3]:
            # Bottoni ATTIVI - cliccabili
            if st.button(flabel, key=f"dash_{st.session_state.hub_page}_{fkey}_{idx}_ATTIVO", use_container_width=True, type="primary" if st.session_state.hub_page=="operativo" else "secondary"):
                st.session_state.menu=fkey
                st.rerun()
    st.divider()
    st.caption(f"Dashboard {st.session_state.hub_page.upper()} - {len(buttons)} bottoni ATTIVI - A sinistra vedi solo elenco {st.session_state.hub_page}")

elif cur == "Splash":
    hdr_form("🎬 SPLASH - Come primo file")
    inject_popout_splash()
    try:
        if os.path.exists("copertina.png"):
            st.image("copertina.png", width=350)
    except:
        pass
    if st.button("🏠 Dashboard"):
        st.session_state.menu="Dashboard"
        st.rerun()

elif cur == "Eventi":
    hdr_form("📅 EVENTI - Ripristinato")
    with st.form("eventi_form"):
        c1,c2 = st.columns(2)
        with c1:
            nome = st.text_input("Nome Evento")
            data = st.date_input("Data Evento")
            luogo = st.text_input("Luogo")
        with c2:
            tipo = st.selectbox("Tipo", ["Esercitazione","Emergenza","Manifestazione","Formazione","Altro"])
            note = st.text_area("Note")
        if st.form_submit_button("💾 Salva Evento", type="primary", use_container_width=True):
            st.session_state.eventi.append({"nome":nome,"data":str(data),"luogo":luogo,"tipo":tipo,"note":note,"ts":str(datetime.now())})
            st.success("Salvato!")
            st.rerun()
    if st.session_state.eventi:
        st.dataframe(pd.DataFrame(st.session_state.eventi), use_container_width=True)

elif cur == "Note Spese":
    hdr_form("🧾 NOTE SPESE")
    with st.form("note_spese"):
        c1,c2,c3 = st.columns(3)
        with c1:
            d = st.date_input("Data")
            vol = st.text_input("Volontario")
            ev = st.text_input("Evento")
        with c2:
            tipo = st.selectbox("Tipo", ["Carburante","Pedaggio","Vitto","Materiale","Parcheggio","Altro"])
            imp = st.number_input("Importo €", min_value=0.0, step=0.1)
        with c3:
            pag = st.selectbox("Pagato da", ["Volontario","ODV","Anticipo"])
            note = st.text_area("Note")
        if st.form_submit_button("💾 Salva", type="primary", use_container_width=True):
            st.session_state.note_spese.append({"data":str(d),"volontario":vol,"evento":ev,"tipo":tipo,"importo":imp,"pagato_da":pag,"note":note})
            st.success("Salvata!")
            st.rerun()
    if st.session_state.note_spese:
        df = pd.DataFrame(st.session_state.note_spese)
        st.dataframe(df, use_container_width=True)
        st.metric("Totale", f"€ {df['importo'].sum():.2f}")

elif cur == "Statistiche":
    hdr_form("📊 STATISTICHE")
    spese = st.session_state.get("spese_odv",[])
    note = st.session_state.get("note_spese",[])
    tot_s = sum([float(x.get("importo",0)) for x in spese]) if spese else 0
    tot_n = sum([float(x.get("importo",0)) for x in note]) if note else 0
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Spese ODV", f"€ {tot_s:.2f}")
    c2.metric("Note Spese", f"€ {tot_n:.2f}")
    c3.metric("Totale", f"€ {tot_s+tot_n:.2f}")
    c4.metric("Eventi", len(st.session_state.eventi))
    if spese or note:
        st.bar_chart(pd.DataFrame([{"tipo":"Spese ODV","importo":tot_s},{"tipo":"Note Spese","importo":tot_n}]).set_index("tipo"))

elif cur == "Spese ODV per Evento":
    hdr_form("💰 SPESE ODV")
    with st.form("spese"):
        desc = st.text_input("Descrizione")
        imp = st.number_input("Importo €", min_value=0.0)
        if st.form_submit_button("Salva", type="primary"):
            st.session_state.spese_odv.append({"descrizione":desc,"importo":imp,"data":str(datetime.now())})
            st.rerun()
    if st.session_state.spese_odv:
        st.dataframe(pd.DataFrame(st.session_state.spese_odv), use_container_width=True)

else:
    hdr_form(f"{cur} - Form base")
    st.info(f"Form '{cur}' pronto - struttura base come primo file - Geolocalizzazione rimossa")
    st.caption("Per i form complessi (Volontari con foto, DB Radio, Brogliaccio, Mezzi ecc.) usa i file completi App-ANA-AMMINISTRATIVO.py e App-ANA-OPERATIVO.py che hai già")

st.divider()
st.markdown('<div style="text-align:center;padding:6px;background:#1A5D1A;color:white;border-radius:6px">ANA Varese - UNICO compatto - Splash + Entra + Dashboard divisa + Eventi + Senza Geoloc - Fix line 7513</div>', unsafe_allow_html=True)
