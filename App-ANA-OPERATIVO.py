
import streamlit as st
import pandas as pd
import os, json, io, base64, pathlib
from datetime import datetime, date, time
from io import BytesIO

PERSIST_FILE = "ana_operativo_dati.json"
DATA_KEYS = ["ospiti","spese_odv","volontari","radio_db","consegna_radio","alias_radio","brogliaccio","eventi","emergenze","checkin","interventi","tabella_interventi","mezzi","attrezzature","mappa_avanzata_markers","icone","turni","chat","posizioni_pd785","posizioni_anytone","archivio_documenti","verbali","diplomi","geoloc"]

def salva():
    try:
        d = {k: st.session_state.get(k) for k in DATA_KEYS if k in st.session_state}
        with open(PERSIST_FILE, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2, default=str)
    except Exception as e:
        print(e)

def carica():
    if os.path.exists(PERSIST_FILE):
        try:
            with open(PERSIST_FILE, "r", encoding="utf-8") as ff:
                dd = json.load(ff)
            for k,v in dd.items():
                if k not in st.session_state or (isinstance(st.session_state[k], list) and len(st.session_state[k])==0):
                    st.session_state[k] = v
        except: pass

carica()

if "page" not in st.session_state: st.session_state.page = "entra"
if "logged" not in st.session_state: st.session_state.logged = False
if "menu" not in st.session_state: st.session_state.menu = "Dashboard"
if "hub_modulo" not in st.session_state: st.session_state.hub_modulo = "operativo"

st.set_page_config(page_title="ANA Operativo - 🚒", page_icon="🚒", layout="wide")

# CSS
st.markdown("""
<style>
#MainMenu{visibility:hidden;} footer{visibility:hidden;}
.stApp{background:#e8f5e9 !important;}
[data-testid="stAppViewContainer"]{background:linear-gradient(180deg,#c8e6c9 0%,#a5d6a7 100%) !important;}
</style>
""", unsafe_allow_html=True)

MENU_OPERATIVO = [
    "Dashboard",
            "Emergenze",
            "Tabella Emergenze",
            "Interventi Emergenza",
            "Tabella Interventi Emergenza",
            "Check-in",
            "Brogliaccio",
            "DB Radio",
            "Consegna Radio",
            "Alias Radio",
            "Geolocalizzazione Hytera + Anytone",
            "Posizioni PD785",
            "Posizioni Anytone",
            "Mezzi",
            "Attrezzature",
            "Mappe Postazioni",
            "Libreria Icone",
            "Turni",
            "Eventi",
            "Chat"
]

# ---- ENTRA + LOGIN ----
if st.session_state.page == "entra":
    st.markdown(f"<div style='text-align:center;padding:30px;background:#1A5D1A;border-radius:15px;color:white;'><h1>🚒 ANA OPERATIVO</h1><p>Progetto Operativo - Dati persistenti dopo logout</p></div>", unsafe_allow_html=True)
    c1,c2,c3 = st.columns([1,2,1])
    with c2:
        with st.form("login"):
            u = st.text_input("Utente", value="admin")
            p = st.text_input("Password", type="password", value="ana2024")
            if st.form_submit_button("Entra in Operativo", type="primary", use_container_width=True):
                st.session_state.logged = True
                st.session_state.page = "dashboard"
                st.rerun()

if not st.session_state.logged and st.session_state.page != "entra":
    st.session_state.page = "entra"
    st.rerun()

# ---- SIDEBAR FIX SYNTAX 1968 ----
if st.session_state.page != "entra":
    with st.sidebar:
        try:
            if os.path.exists("logo.png"):
                st.image("logo.png", width=110)
            else:
                st.markdown(f"<div style='width:110px;height:110px;background:#1A5D1A;border-radius:10px;display:flex;align-items:center;justify-content:center;color:white;font-weight:bold'>OPE</div>", unsafe_allow_html=True)
        except:
            st.write("ANA")
        # FIX riga 1968: markdown fuori try/except
        st.markdown(f"<p style='font-weight:bold;color:#1A5D1A;margin-top:10px'>MENU OPERATIVO</p>", unsafe_allow_html=True)
        st.caption(f"Persist: {PERSIST_FILE}")
        sel = st.radio("Form", MENU_OPERATIVO, index=0, key="menu_sel_Operativo")
        st.session_state.menu = sel
        st.divider()
        if st.button("🏠 HUB Principale", use_container_width=True):
            st.session_state.page = "entra"
            st.session_state.logged = False
            salva()
            st.rerun()
        if st.button("🚪 Logout (salva dati)", type="primary", use_container_width=True):
            salva()
            st.session_state.page = "entra"
            st.session_state.logged = False
            st.rerun()

    # ---- DASHBOARD ----
    st.title(f"🚒 Operativo - {st.session_state.menu}")
    if st.session_state.menu == "Dashboard":
        st.success(f"Progetto Operativo - Form disponibili: {len(MENU_OPERATIVO)}")
        cols = st.columns(3)
        for idx, f in enumerate(MENU_OPERATIVO):
            with cols[idx % 3]:
                if st.button(f"📋 {f}", key=f"btn_{f}_{project_name}", use_container_width=True):
                    st.session_state.menu = f
                    st.rerun()
        st.divider()
        st.markdown("### Dati persistenti")
        st.json({"ospiti": len(st.session_state.get("ospiti",[])), "spese": len(st.session_state.get("spese_odv",[])), "volontari": len(st.session_state.get("volontari",[]))})
    else:
        st.info(f"Form '{st.session_state.menu}' - struttura presa da App-GAE-FINALE_32.py - integra qui il codice specifico del form")
        st.warning("Questo è lo scheletro - importa la funzione render del form originale da App-GAE-FINALE_32.py")
        # placeholder per test persistenza
        if st.session_state.menu == "Spese ODV per Evento":
            with st.form("spesa_quick"):
                d = st.text_input("Descrizione")
                im = st.number_input("Importo")
                if st.form_submit_button("Salva Spesa"):
                    st.session_state.setdefault("spese_odv", []).append({"descrizione": d, "importo": im, "data": str(datetime.now())})
                    salva()
                    st.success("Salvato!")
                    st.rerun()
            if st.session_state.get("spese_odv"):
                st.dataframe(pd.DataFrame(st.session_state.spese_odv), use_container_width=True)
        elif st.session_state.menu in ["Ospiti", "Volontari (con foto)"]:
            with st.form("osp_quick"):
                n = st.text_input("Nome")
                if st.form_submit_button("Salva"):
                    key = "ospiti" if "Ospiti" in st.session_state.menu else "volontari"
                    st.session_state.setdefault(key, []).append({"nome": n, "ts": str(datetime.now())})
                    salva()
                    st.success("Salvato!")
                    st.rerun()
            st.dataframe(pd.DataFrame(st.session_state.get("ospiti" if "Ospiti" in st.session_state.menu else "volontari", [])), use_container_width=True)

    try:
        salva()
    except:
        pass
    st.divider()
    st.markdown(f"<div style='text-align:center;padding:6px;background:#1A5D1A;color:white;border-radius:6px'>ANA Operativo - Fix 1968 + Persistenza + HUB</div>", unsafe_allow_html=True)
