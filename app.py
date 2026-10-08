import streamlit as st
import pandas as pd
import os
import io
import json
import base64
from io import BytesIO
from datetime import datetime, date, time
import requests
import sys

# --- Dipendenze opzionali - gestite senza pip a runtime ---
try:
    from reportlab.lib.pagesizes import landscape, A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    REPORTLAB_OK = True
except ImportError:
    REPORTLAB_OK = False

try:
    import openpyxl
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False

try:
    import xlsxwriter
    XLSXWRITER_OK = True
except ImportError:
    XLSXWRITER_OK = False

try:
    import xlrd
    XLRD_OK = True
except ImportError:
    XLRD_OK = False

if "ospiti" not in st.session_state:
    st.session_state.ospiti = []
if "osp_edit_index" not in st.session_state:
    st.session_state.osp_edit_index = None
if "spese_odv" not in st.session_state:
    st.session_state.spese_odv = []
if "spese_edit_index" not in st.session_state:
    st.session_state.spese_edit_index = None

if "page" not in st.session_state:
    st.session_state.page = "entra"
if "logged" not in st.session_state:
    st.session_state.logged = False

if "menu" not in st.session_state:
    st.session_state.menu = "Dashboard"
if "hub_page" not in st.session_state:
    st.session_state.hub_page = "hub"
if "menu_admin" not in st.session_state:
    st.session_state.menu_admin = "Dashboard"
if "menu_oper" not in st.session_state:
    st.session_state.menu_oper = "Dashboard"

MENU_ADMINISTRATIVO = [
    "Dashboard","Splash","Volontari (con foto)","Ospiti","Verbali","Archivio Documenti","Diplomi Attestati",
    "Spese ODV per Evento","Note Spese","Report Filtro","Statistiche","Gestione Utenti","Backup"
]
MENU_OPERATIVO = [
    "Dashboard","Splash","DB Radio","Consegna Radio","Alias Radio","Brogliaccio","Eventi","Emergenze",
    "Tabella Emergenze","Check-in","Interventi Emergenza","Tabella Interventi Emergenza",
    "Mezzi","Attrezzature","Mappe Postazioni","Libreria Icone","Turni","Chat"
]


st.set_page_config(
    page_title="ANA Varese - Caronno Pertusella - Dashboard",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CSS globale VERDE ANA + font nero Times bold + FIX - versione definitiva ieri ---
def inject_global_css():
    st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header[data-testid="stHeader"] {background: transparent;}
    .stApp {padding-top: 12px; background: #e8f5e9 !important;}
    [data-testid="stAppViewContainer"] {background: linear-gradient(180deg, #c8e6c9 0%, #a5d6a7 100%) !important;}
    .stButton>button {border-radius: 8px; font-weight: 600;}
    [data-testid="stForm"], .stForm {
        background: #1A5D1A !important;
        border: 3px solid #FFD700 !important;
        border-radius: 12px !important;
        padding: 16px !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.3) !important;
    }
    /* FONT TIMES ROMAN GRASSETTO DEFINITIVO IERI - NERO */
    /* FONT NERO GRASSETTO TIMES ROMAN PER TUTTE LE CELLE FORM - RICHIESTA EZIO DEFINITIVA */
    .stTextInput > div > div > input,
    .stTextArea > div > div > textarea,
    .stSelectbox > div > div > div,
    .stNumberInput > div > div > input,
    .stDateInput > div > div > input,
    .stTimeInput > div > div > input,
    input[type="text"], input[type="number"], textarea, select {
        color: black !important;
        font-family: 'Times New Roman', Times, serif !important;
        font-weight: bold !important;
        font-size: 15px !important;
        background: white !important;
        border: 1.5px solid #1A5D1A !important;
        border-radius: 6px !important;
    }
    .stTextInput label, .stTextArea label, .stSelectbox label, .stNumberInput label, .stDateInput label, .stTimeInput label, .stRadio label, label {
        color: black !important;
        font-family: 'Times New Roman', Times, serif !important;
        font-weight: bold !important;
        font-size: 14px !important;
    }
    div[data-testid="column"] {
        background: rgba(46,125,50,0.85) !important;
        border: 1px solid #FFD700 !important;
        border-radius: 8px !important;
        padding: 8px !important;
    }
    [data-testid="stDataFrame"] div {
        font-family: 'Times New Roman', serif !important;
        font-weight: bold !important;
        color: black !important;
    }
    /* --- COLORI LINGUETTE VOLONTARI + OSPITI - DEFINITIVO - BOTTONI + stTabs --- */
    div[data-testid="stTabs"] button[role="tab"]:nth-child(1),
    div[data-baseweb="tab-list"] button:nth-child(1),
    button[data-baseweb="tab"]:nth-of-type(1) {background:#1A5D1A !important; color:white !important; border-radius:8px 8px 0 0 !important;}
    div[data-testid="stTabs"] button[role="tab"]:nth-child(2),
    div[data-baseweb="tab-list"] button:nth-child(2),
    button[data-baseweb="tab"]:nth-of-type(2) {background:#1976d2 !important; color:white !important; border-radius:8px 8px 0 0 !important;}
    div[data-testid="stTabs"] button[role="tab"]:nth-child(3),
    div[data-baseweb="tab-list"] button:nth-child(3),
    button[data-baseweb="tab"]:nth-of-type(3) {background:#d32f2f !important; color:white !important; border-radius:8px 8px 0 0 !important;}
    div[data-testid="stTabs"] button[role="tab"]:nth-child(4),
    div[data-baseweb="tab-list"] button:nth-child(4),
    button[data-baseweb="tab"]:nth-of-type(4) {background:#ff9800 !important; color:black !important; border-radius:8px 8px 0 0 !important;}
    div[data-testid="stTabs"] button[role="tab"]:nth-child(5),
    div[data-baseweb="tab-list"] button:nth-child(5),
    button[data-baseweb="tab"]:nth-of-type(5) {background:#6a1b9a !important; color:white !important; border-radius:8px 8px 0 0 !important;}
    div[data-testid="stTabs"] button[role="tab"]:nth-child(6),
    div[data-baseweb="tab-list"] button:nth-child(6),
    button[data-baseweb="tab"]:nth-of-type(6) {background:#00796b !important; color:white !important; border-radius:8px 8px 0 0 !important;}
    div[data-testid="stTabs"] button[role="tab"] p,
    button[data-baseweb="tab"] div {color:inherit !important; font-weight:bold !important;}
    button[data-baseweb="tab"][aria-selected="true"],
    div[data-testid="stTabs"] button[aria-selected="true"] {border:3px solid #FFD700 !important; font-weight:bold !important; box-shadow:0 -2px 8px rgba(0,0,0,0.4) !important;}
    /* BOTTONI COLORATI DEFINITIVO - Volontari + Ospiti - FIX BIANCO */
    div[data-testid="stHorizontalBlock"] div[data-testid="column"]:nth-child(1) div[data-testid="stButton"] button,
    div[data-testid="stHorizontalBlock"] div[data-testid="column"]:nth-child(1) button,
    .linguette-box div[data-testid="column"]:nth-child(1) button {background:#1A5D1A !important; background-color:#1A5D1A !important; color:white !important; border:2px solid #FFD700 !important; border-radius:8px 8px 0 0 !important; font-weight:bold !important;}
    div[data-testid="stHorizontalBlock"] div[data-testid="column"]:nth-child(2) div[data-testid="stButton"] button,
    div[data-testid="stHorizontalBlock"] div[data-testid="column"]:nth-child(2) button,
    .linguette-box div[data-testid="column"]:nth-child(2) button {background:#1976d2 !important; background-color:#1976d2 !important; color:white !important; border:2px solid #FFD700 !important; border-radius:8px 8px 0 0 !important; font-weight:bold !important;}
    div[data-testid="stHorizontalBlock"] div[data-testid="column"]:nth-child(3) div[data-testid="stButton"] button,
    div[data-testid="stHorizontalBlock"] div[data-testid="column"]:nth-child(3) button,
    .linguette-box div[data-testid="column"]:nth-child(3) button {background:#d32f2f !important; background-color:#d32f2f !important; color:white !important; border:2px solid #FFD700 !important; border-radius:8px 8px 0 0 !important; font-weight:bold !important;}
    div[data-testid="stHorizontalBlock"] div[data-testid="column"]:nth-child(4) div[data-testid="stButton"] button,
    div[data-testid="stHorizontalBlock"] div[data-testid="column"]:nth-child(4) button,
    .linguette-box div[data-testid="column"]:nth-child(4) button {background:#ff9800 !important; background-color:#ff9800 !important; color:black !important; border:2px solid #FFD700 !important; border-radius:8px 8px 0 0 !important; font-weight:bold !important;}
    div[data-testid="stHorizontalBlock"] div[data-testid="column"]:nth-child(5) div[data-testid="stButton"] button,
    div[data-testid="stHorizontalBlock"] div[data-testid="column"]:nth-child(5) button,
    .linguette-box div[data-testid="column"]:nth-child(5) button {background:#6a1b9a !important; background-color:#6a1b9a !important; color:white !important; border:2px solid #FFD700 !important; border-radius:8px 8px 0 0 !important; font-weight:bold !important;}
    div[data-testid="stHorizontalBlock"] div[data-testid="column"]:nth-child(6) div[data-testid="stButton"] button,
    div[data-testid="stHorizontalBlock"] div[data-testid="column"]:nth-child(6) button,
    .linguette-box div[data-testid="column"]:nth-child(6) button {background:#00796b !important; background-color:#00796b !important; color:white !important; border:2px solid #FFD700 !important; border-radius:8px 8px 0 0 !important; font-weight:bold !important;}
    .linguette-box button {min-height:40px !important;}
    </style>
    """, unsafe_allow_html=True)



def inject_linguette_js_fix():
    st.components.v1.html("""
    <script>
    function colorLinguette() {
        try {
            const rootDoc = window.parent ? window.parent.document : document;
            const colors = {
                "Anagrafica": "#1A5D1A",
                "Contatti": "#1976d2",
                "Ruolo": "#d32f2f",
                "Dotazione": "#ff9800",
                "Documenti": "#6a1b9a",
                "Foto": "#00796b"
            };
            rootDoc.querySelectorAll('button').forEach(btn => {
                const txt = (btn.innerText || btn.textContent || "").trim();
                if (txt.length > 30) return;
                for (const [key, col] of Object.entries(colors)) {
                    if (txt.includes(key)) {
                        btn.style.setProperty('background', col, 'important');
                        btn.style.setProperty('background-color', col, 'important');
                        btn.style.setProperty('color', (key==="Dotazione" ? "black" : "white"), 'important');
                        btn.style.setProperty('border', '2px solid #FFD700', 'important');
                        btn.style.setProperty('border-radius', '8px 8px 0 0', 'important');
                        btn.style.setProperty('font-weight', 'bold', 'important');
                    }
                }
            });
        } catch(e) {}
    }
    setTimeout(colorLinguette, 200);
    setInterval(colorLinguette, 600);
    try { window.parent.addEventListener('load', colorLinguette); } catch(e) {}
    </script>
    """, height=0)


def inject_first_page_green():
    pass

def inject_fullscreen_kiosk():
    # Versione SICURA senza parent.document - funziona su Streamlit Cloud
    st.markdown("""
    <style>
    #fs-kiosk-btn {
        position: fixed; top: 10px; right: 10px; z-index: 999999;
        background: #d32f2f; color: white; border: none;
        padding: 10px 18px; font-weight: bold; border-radius: 8px;
        cursor: pointer; box-shadow: 0 2px 8px rgba(0,0,0,0.3);
    }
    </style>
    """, unsafe_allow_html=True)
    st.components.v1.html("""
    <button id="fs-kiosk-btn" onclick="
        if(!document.fullscreenElement){
            document.documentElement.requestFullscreen().catch(()=>{});
            this.innerText='✕ ESCI';
            this.style.background='#1A5D1A';
        } else {
            document.exitFullscreen();
            this.innerText='⛶ FULLSCREEN';
            this.style.background='#d32f2f';
        }
    ">⛶ FULLSCREEN</button>
    <script>
    document.addEventListener('keydown', (e)=>{
        if(e.key.toLowerCase()==='f' && !['INPUT','TEXTAREA','SELECT'].includes(document.activeElement.tagName)){
            if(!document.fullscreenElement) document.documentElement.requestFullscreen();
        }
        if(e.key==='Escape' && document.fullscreenElement) document.exitFullscreen();
    });
    </script>
    """, height=0)


inject_global_css()
inject_linguette_js_fix()
inject_fullscreen_kiosk()

def to_excel_bytes(dfs_dict: dict, engine="openpyxl") -> bytes:
    """Esporta dizionario di DataFrame in un unico Excel multi-foglio - engine unico"""
    output = BytesIO()
    # Scegli engine disponibile
    if OPENPYXL_OK:
        eng = "openpyxl"
    elif XLSXWRITER_OK:
        eng = "xlsxwriter"
    else:
        eng = None
    with pd.ExcelWriter(output, engine=eng) as writer:
        for sheet_name, df in dfs_dict.items():
            # tronca nome foglio a 31 char (limite Excel)
            safe_name = sheet_name[:31]
            if isinstance(df, list):
                df = pd.DataFrame(df)
            df.to_excel(writer, sheet_name=safe_name, index=False)
    return output.getvalue()

def safe_read_excel(file):
    """Legge Excel provando tutti gli engine disponibili"""
    last_err = ""
    for eng in [None, "openpyxl", "xlrd"]:
        try:
            file.seek(0)
            if eng is None:
                return pd.read_excel(file), ""
            else:
                return pd.read_excel(file, engine=eng), ""
        except Exception as e:
            last_err = str(e)
            continue
    return None, last_err


def inject_fullscreen_all_maps():

    st.components.v1.html('''
    <script>
    function addFullscreenToAllLeafletMaps() {
        document.querySelectorAll('.leaflet-container').forEach(function(container) {
            if (container.querySelector('.fullscreen-btn-all')) return;
            var zoomCtrl = container.querySelector('.leaflet-control-zoom');
            if (!zoomCtrl) return;
            var fsDiv = document.createElement('div');
            fsDiv.className = 'leaflet-bar leaflet-control';
            fsDiv.style.marginTop = '5px';
            var btn = document.createElement('a');
            btn.className = 'fullscreen-btn-all';
            btn.innerHTML = '⛶';
            btn.href = '#';
            btn.title = 'Schermo intero 100%';
            btn.onclick = function(e) {
                e.preventDefault();
                e.stopPropagation();
                var mapEl = this.closest('.leaflet-container');
                if (!document.fullscreenElement) {
                    if (mapEl.requestFullscreen) mapEl.requestFullscreen();
                    else if (mapEl.webkitRequestFullscreen) mapEl.webkitRequestFullscreen();
                    else if (mapEl.msRequestFullscreen) mapEl.msRequestFullscreen();
                    setTimeout(function(){ 
                        try { mapEl._leaflet_map && mapEl._leaflet_map.invalidateSize(); } catch(e){}
                    }, 600);
                } else {
                    if (document.exitFullscreen) document.exitFullscreen();
                }
            };
            fsDiv.appendChild(btn);
            zoomCtrl.parentNode.insertBefore(fsDiv, zoomCtrl.nextSibling);
        });
    }
    setTimeout(addFullscreenToAllLeafletMaps, 800);
    setInterval(addFullscreenToAllLeafletMaps, 1500);
    </script>
    ''', height=0)


COMUNI_ITALIA = [
    "Varese", "Busto Arsizio", "Gallarate", "Saronno", "Cassano Magnago",
    "Tradate", "Malnate", "Somma Lombardo", "Gavirate", "Laveno-Mombello",
    "Luino", "Sesto Calende", "Samarate", "Lonate Pozzolo", "Fagnano Olona",
    "Castellanza", "Caronno Pertusella", "Gerenzano", "Origgio", "Uboldo",
    "Cislago", "Gorla Minore", "Gorla Maggiore", "Marnate", "Olgiate Olona",
    "Solbiate Olona", "Solbiate Arno", "Albizzate", "Cairate", "Carnago",
    "Caravate", "Besozzo", "Besnate", "Brebbia", "Bregano", "Brenta",
    "Bardello", "Biandronno", "Bodio Lomnago", "Buguggiate", "Casale Litta",
    "Casciago", "Castelseprio", "Castiglione Olona", "Cavaria con Premezzo",
    "Cazzago Brabbia", "Cislago", "Cittiglio", "Comabbio", "Comerio",
    "Cremenaga", "Cuasso al Monte", "Cugliate-Fabiasco", "Cunardo", "Curiglia",
    "Daverio", "Dumenza", "Duno", "Ferrera di Varese", "Gazzada Schianno",
    "Gemonio", "Gornate Olona", "Inarzo", "Induno Olona", "Ispra",
    "Jerago con Orago", "Lavena Ponte Tresa", "Lozza", "Maccagno",
    "Malgesso", "Marchirolo", "Marzio", "Masciago Primo", "Mercallo",
    "Montegrino Valtravaglia", "Morazzone", "Mornago", "Oggiona con Santo Stefano",
    "Porto Ceresio", "Porto Valtravaglia", "Rancio Valcuvia", "Saltrio",
    "Sangiano", "Travedona Monate", "Vedano Olona", "Venegono Inferiore",
    "Venegono Superiore", "Vergiate", "Viggiu"
]

VIE_STANDARD = [
    "Via Roma", "Via Garibaldi", "Via Matteotti", "Via Verdi",
    "Via Manzoni", "Via Milano", "Via Varese", "Via Dante",
    "Via Mazzini", "Via Cavour", "Corso Italia", "Piazza Libertà",
    "Via San Martino", "Via XXV Aprile", "Via IV Novembre",
    "Via Risorgimento", "Via Volta", "Via Marconi", "Via De Gasperi",
    "Viale Europa"
]


# ===== BASE 2032 + PATCH CSV COMUNE VIA - INIZIO =====
BASE_2032_PATHS = [
    "/mnt/data/base_2032.csv",
    "/mnt/data/base2032.csv",
    "base_2032.csv",
    "base2032.csv"
]
PATCH_PATHS = [
    "/mnt/data/patch_comune_via.csv",
    "/mnt/data/patch_comuni_vie.csv",
    "patch_comune_via.csv"
]

def load_base_2032_df():
    for p in BASE_2032_PATHS:
        if os.path.exists(p):
            try:
                df = pd.read_csv(p, dtype=str, keep_default_na=False)
                df.columns = [c.strip().lower() for c in df.columns]
                if 'comune' in df.columns:
                    return df
            except Exception as e:
                print(f"Errore lettura base {p}: {e}")
    return None

def load_patch_df():
    if "patch_df" in st.session_state and st.session_state.patch_df is not None:
        return st.session_state.patch_df
    for p in PATCH_PATHS:
        if os.path.exists(p):
            try:
                df = pd.read_csv(p, dtype=str, keep_default_na=False)
                df.columns = [c.strip().lower() for c in df.columns]
                if 'comune' in df.columns:
                    return df
            except:
                pass
    return None

def ui_patch_loader_sidebar():
    with st.sidebar.expander("🛠️ BASE 2032 + PATCH CSV", expanded=False):
        st.caption("CSV con colonne `comune,via` - sovrascrive base")
        up_base = st.file_uploader("BASE 2032 (opzionale)", type=["csv"], key="up_base_2032")
        if up_base:
            try:
                df = pd.read_csv(up_base, dtype=str, keep_default_na=False)
                df.to_csv("/mnt/data/base_2032.csv", index=False, encoding='utf-8-sig')
                st.success(f"Base 2032 caricata: {len(df)} righe")
                st.cache_data.clear()
            except Exception as e:
                st.error(f"Errore base: {e}")
        up_patch = st.file_uploader("PATCH comune via", type=["csv"], key="up_patch_comune_via")
        if up_patch:
            try:
                df = pd.read_csv(up_patch, dtype=str, keep_default_na=False)
                df.to_csv("/mnt/data/patch_comune_via.csv", index=False, encoding='utf-8-sig')
                df.columns = [c.strip().lower() for c in df.columns]
                st.session_state.patch_df = df
                st.success(f"Patch: {len(df)} righe - {df['comune'].nunique() if 'comune' in df.columns else '?'} comuni")
                st.cache_data.clear()
            except Exception as e:
                st.error(f"Errore patch: {e}")
        patch_df = load_patch_df()
        base_df = load_base_2032_df()
        c1,c2 = st.columns(2)
        with c1:
            st.metric("Base", f"{len(base_df) if base_df is not None else 0} righe")
        with c2:
            st.metric("Patch", f"{len(patch_df) if patch_df is not None else 0} righe")
        if patch_df is not None:
            st.dataframe(patch_df.head(20), use_container_width=True)
            if st.button("❌ Rimuovi patch", key="btn_remove_patch"):
                if os.path.exists("/mnt/data/patch_comune_via.csv"):
                    os.remove("/mnt/data/patch_comune_via.csv")
                st.session_state.patch_df = None
                st.cache_data.clear()
                st.rerun()
# ===== BASE 2032 + PATCH - FINE =====




def get_stato_color(stato):
    """
    Modifica 4: Colora fondo campo stato intervento emergenze
    Return bg, txt, label
    """
    bg_color = "#ffffff"
    txt_color = "#000000"
    label = stato

    if stato == "Operativo":
        bg_color = "#ff0000"
        txt_color = "white"
        label = "Operativo"
    elif stato == "In Corso":
        bg_color = "#ffff00"
        txt_color = "black"
        label = "In Corso"
    elif stato == "Completato":
        bg_color = "#00ff00"
        txt_color = "black"
        label = "Completato"
    elif stato == "Chiuso":
        bg_color = "#808080"
        txt_color = "white"
        label = "Chiuso"
    elif stato == "In Stand By":
        bg_color = "#ff8c00"
        txt_color = "white"
        label = "In Stand By"
    elif stato == "Sospeso":
        bg_color = "#87ceeb"
        txt_color = "black"
        label = "Sospeso"
    elif stato == "Annullato":
        bg_color = "#000000"
        txt_color = "white"
        label = "Annullato"
    elif stato == "In Attesa":
        bg_color = "#ffd700"
        txt_color = "black"
        label = "In Attesa"
    else:
        bg_color = "#ffffff"
        txt_color = "#000000"
        label = stato

    return bg_color, txt_color, label


def get_comuni():
    """
    BASE 2032 + PATCH CSV comune via
    Priorità: PATCH > BASE_2032.csv > GitHub > COMUNI_ITALIA
    """
    comuni_set = set()
    # 1. BASE 2032
    base_df = load_base_2032_df()
    if base_df is not None and 'comune' in base_df.columns:
        for c in base_df['comune'].dropna().unique():
            c = str(c).strip()
            if c:
                comuni_set.add(c)
    else:
        # GitHub come base 2032 online
        try:
            url = "https://raw.githubusercontent.com/matteocontrini/comuni-json/master/comuni.json"
            resp = requests.get(url, timeout=5)
            if resp.status_code == 200:
                data = resp.json()
                for cc in data:
                    nome = cc.get("nome", "")
                    if nome:
                        comuni_set.add(nome)
        except:
            pass
        if not comuni_set:
            comuni_set.update(COMUNI_ITALIA)

    # 2. PATCH aggiunge comuni
    patch_df = load_patch_df()
    if patch_df is not None and 'comune' in patch_df.columns:
        for c in patch_df['comune'].dropna().unique():
            c = str(c).strip()
            if c:
                # normalizza titolo
                comuni_set.add(c)

    return sorted(list(comuni_set))


def get_vie(comune):
    """
    BASE 2032 + PATCH CSV comune via
    Priorità: PATCH per comune > BASE_2032 > Overpass > VIE_STANDARD
    """
    vie = []
    if not comune:
        return VIE_STANDARD
    comune_norm = comune.strip().lower()

    # 1. PATCH
    patch_df = load_patch_df()
    if patch_df is not None and 'comune' in patch_df.columns and 'via' in patch_df.columns:
        mask = patch_df['comune'].astype(str).str.strip().str.lower() == comune_norm
        vie_patch = patch_df[mask]['via'].dropna().astype(str).str.strip().unique().tolist()
        vie_patch = [v for v in vie_patch if v]
        if vie_patch:
            vie.extend(vie_patch)

    # 2. BASE 2032
    base_df = load_base_2032_df()
    if base_df is not None and 'comune' in base_df.columns and 'via' in base_df.columns:
        mask = base_df['comune'].astype(str).str.strip().str.lower() == comune_norm
        vie_base = base_df[mask]['via'].dropna().astype(str).str.strip().unique().tolist()
        for v in vie_base:
            if v not in vie:
                vie.append(v)
        if vie:
            return sorted(list(set(vie)))[:150]

    if vie:
        return sorted(list(set(vie)))[:150]

    # 3. Overpass API
    try:
        query = """
        [out:json][timeout:10];
        area["name"="%s"]->.a;
        way(area.a)["highway"]["name"];
        out tags;
        """ % comune
        overpass_url = "https://overpass-api.de/api/interpreter"
        resp = requests.get(
            overpass_url,
            params={"data": query},
            timeout=8
        )
        if resp.status_code == 200:
            data = resp.json()
            for el in data.get("elements", []):
                tags = el.get("tags", {})
                nome_via = tags.get("name", "")
                if nome_via and nome_via not in vie:
                    vie.append(nome_via)
            if len(vie) > 3:
                return sorted(vie[:80])
    except:
        pass

    pref = "Via " + comune
    custom = [pref + " Centro", pref + " Nord", pref + " Sud"]
    return VIE_STANDARD + custom



def emoji_sicura(emoji_char, fallback_text="📍"):
    """Ritorna emoji se valida, altrimenti fallback testo - evita quadratini"""
    if not emoji_char or emoji_char in ["□", "☐", "�", ""]:
        return fallback_text
    # Lista emoji sicure 100% compatibili su Streamlit Cloud
    sicure = ["🚨", "📅", "🚐", "👤", "🏥", "🔥", "💧", "📻", "📍", "✅", "❌", "⚠️", "💾", "📋", "🗑️", "🔄", "🗺️"]
    if emoji_char in sicure or len(emoji_char) <= 2:
        return emoji_char
    return fallback_text


def combo_comune(label, key, default=""):
    """
    Selectbox COMUNI ITALIA con default
    """
    comuni_list = get_comuni()
    comuni_list = sorted(list(set(comuni_list)))

    if default and default not in comuni_list:
        comuni_list = [default] + comuni_list

    idx_default = 0
    if default:
        try:
            idx_default = comuni_list.index(default)
        except:
            idx_default = 0

    selected = st.selectbox(
        label,
        comuni_list,
        index=idx_default,
        key=key
    )
    return selected


def combo_vie(label, comune, key, default=""):
    """
    Se comune, get_vie e selectbox VIE DI COMUNE + checkbox via manuale
    """
    vie_list = []
    if comune:
        vie_list = get_vie(comune)
    else:
        vie_list = VIE_STANDARD

    vie_list = sorted(list(set(vie_list)))

    if default and default not in vie_list:
        vie_list = [default] + vie_list

    idx_default = 0
    if default:
        try:
            idx_default = vie_list.index(default)
        except:
            idx_default = 0

    col1, col2 = st.columns([3, 1])
    with col1:
        selected_via = st.selectbox(
            label,
            vie_list,
            index=idx_default,
            key=key
        )
    with col2:
        manuale = st.checkbox(
            "Via manuale",
            key=key + "_manuale_chk"
        )

    if manuale:
        via_manuale = st.text_input(
            "Inserisci via manuale",
            value=default if default else "",
            key=key + "_manuale_txt"
        )
        if via_manuale:
            return via_manuale
        else:
            return selected_via
    else:
        return selected_via


def to_excel(df):
    """
    Esporta DataFrame in Excel - FIX Office 2016 100% compatibile - Ezio
    Usa openpyxl, compatibile Office 2016/2019/365 - MAI CSV travestito
    """
    buf = io.BytesIO()
    df_copy = df.copy()

    cols_to_exclude = [
        "FotoBytes",
        "FileBytes",
        "FotoConsegnaBytes",
        "Foto",
        "FotoBytesObj"
    ]

    for col in cols_to_exclude:
        if col in df_copy.columns:
            df_copy = df_copy.drop(columns=[col])

    # Assicura che df non sia None
    if df_copy is None or not isinstance(df_copy, pd.DataFrame):
        df_copy = pd.DataFrame()

    # Prova openpyxl prima (100% Office 2016 compatibile) - FIX CLOUD: prova sempre, ignora flag
    last_error = ""
    for engine_try in ["openpyxl", "xlsxwriter"]:
        try:
            buf = io.BytesIO()
            # FIX: NON saltare per OPENPYXL_OK - prova sempre!
            with pd.ExcelWriter(buf, engine=engine_try) as writer:
                df_copy.to_excel(writer, index=False, sheet_name="Dati")
            buf.seek(0)
            data = buf.getvalue()
            # Verifica che sia un vero xlsx (PK zip header)
            if data[:2] == b'PK' and len(data) > 100:
                return data
            else:
                last_error = f"Engine {engine_try} non ha prodotto xlsx valido - len {len(data) if data else 0}"
        except Exception as e:
            last_error = str(e)
            continue

    # Ultimo tentativo: forza openpyxl diretto - FIX CLOUD
    try:
        import openpyxl
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df_copy.to_excel(writer, index=False, sheet_name="Dati")
        buf.seek(0)
        data = buf.getvalue()
        if data[:2] == b'PK' and len(data) > 100:
            return data
        last_error = f"openpyxl diretto len {len(data)} non PK"
    except Exception as e:
        last_error = str(e)

    # Se proprio fallisce, crea file Excel minimo con openpyxl diretto - MAI CSV!
    try:
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Dati"
        for c_idx, col_name in enumerate(df_copy.columns, 1):
            ws.cell(row=1, column=c_idx, value=str(col_name))
        for r_idx, row in enumerate(df_copy.itertuples(index=False), 2):
            for c_idx, val in enumerate(row, 1):
                try:
                    if val is None:
                        ws.cell(row=r_idx, column=c_idx, value="")
                    else:
                        ws.cell(row=r_idx, column=c_idx, value=val)
                except:
                    try:
                        ws.cell(row=r_idx, column=c_idx, value=str(val)[:30000])
                    except:
                        ws.cell(row=r_idx, column=c_idx, value="")
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        data = buf.getvalue()
        if data[:2] == b'PK':  # Verifica sia xlsx valido
            return data
        else:
            raise Exception("Fallback openpyxl non ha prodotto PK")
    except Exception as e:
        # ULTIMA SPIAGGIA: Crea Excel vuoto ma VALIDO xlsx, MAI CSV!
        try:
            import openpyxl
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Dati"
            ws.cell(row=1, column=1, value="Errore Export")
            ws.cell(row=2, column=1, value=str(e)[:100])
            ws.cell(row=3, column=1, value="Verifica requirements.txt contenga openpyxl==3.1.5")
            if not df_copy.empty:
                for c_idx, col_name in enumerate(df_copy.columns, 2):
                    ws.cell(row=1, column=c_idx, value=str(col_name))
            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)
            return buf.getvalue()
        except Exception as e2:
            # FIX: DEVE ESSERE SOLO EXCEL XLSX - MAI CSV - Richiesta Ezio
            # Riprova installazione runtime openpyxl
            try:
                _try_install_excel_deps()
            except:
                pass
            # Riprova openpyxl dopo installazione
            try:
                import openpyxl
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Dati"
                for c_idx, col_name in enumerate(df_copy.columns, 1):
                    ws.cell(row=1, column=c_idx, value=str(col_name))
                for r_idx, row in enumerate(df_copy.itertuples(index=False), 2):
                    for c_idx, val in enumerate(row, 1):
                        try:
                            ws.cell(row=r_idx, column=c_idx, value=val if val is not None else "")
                        except:
                            try:
                                ws.cell(row=r_idx, column=c_idx, value=str(val)[:30000])
                            except:
                                ws.cell(row=r_idx, column=c_idx, value="")
                buf = io.BytesIO()
                wb.save(buf)
                buf.seek(0)
                data = buf.getvalue()
                if data[:2] == b'PK' and len(data) > 100:
                    return data
            except Exception as e3:
                last_error = str(e3)
            
            # Ultima spiaggia - crea XLSX valido vuoto ma MAI CSV
            try:
                import openpyxl
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Dati"
                ws.cell(row=1, column=1, value="Export Volontari - Excel")
                ws.cell(row=2, column=1, value=f"Errore originale: {str(e2)[:100]}")
                ws.cell(row=3, column=1, value="Verifica requirements.txt: openpyxl==3.1.5 su GitHub + Reboot Cloud")
                if not df_copy.empty:
                    for c_idx, col_name in enumerate(df_copy.columns, 1):
                        ws.cell(row=1, column=c_idx+1, value=str(col_name))
                    for r_idx, row in enumerate(df_copy.itertuples(index=False), 2):
                        for c_idx, val in enumerate(row, 1):
                            try:
                                ws.cell(row=r_idx, column=c_idx+1, value=str(val)[:32000])
                            except:
                                pass
                buf = io.BytesIO()
                wb.save(buf)
                buf.seek(0)
                return buf.getvalue()
            except:
                # Se proprio tutto fallisce, ritorna XLSX vuoto valido - MAI CSV
                try:
                    import openpyxl
                    wb = openpyxl.Workbook()
                    buf = io.BytesIO()
                    wb.save(buf)
                    buf.seek(0)
                    return buf.getvalue()
                except:
                    # Fallback finale: crea file XLSX minimo con zip - non CSV!
                    import openpyxl
                    wb = openpyxl.Workbook()
                    buf = io.BytesIO()
                    wb.save(buf)
                    buf.seek(0)
                    return buf.getvalue()


def to_excel_multi(datasets):
    """
    datasets = dict nome_sheet -> df - FIX Win7 - NON CRASHA CLOUD
    """
    buf = io.BytesIO()
    try:
        engine = "openpyxl" if OPENPYXL_OK else ("xlsxwriter" if XLSXWRITER_OK else "openpyxl")
        with pd.ExcelWriter(buf, engine=engine) as writer:
            for sheet_name, df in datasets.items():
                df_copy = df.copy()
                for col in ["FotoBytes", "FileBytes", "FotoConsegnaBytes", "Foto"]:
                    if col in df_copy.columns:
                        df_copy = df_copy.drop(columns=[col])
                safe_name = sheet_name[:30]
                df_copy.to_excel(writer, index=False, sheet_name=safe_name)
        buf.seek(0)
        data = buf.getvalue()
        if data and len(data) > 100:
            return data
        else:
            return to_excel(list(datasets.values())[0] if datasets else pd.DataFrame())
    except Exception as e:
        try:
            data = to_excel(list(datasets.values())[0] if datasets else pd.DataFrame())
            if data and data[:2] == b'PK':
                return data
            else:
                # Crea xlsx valido vuoto
                import openpyxl
                wb = openpyxl.Workbook()
                buf = io.BytesIO()
                wb.save(buf)
                buf.seek(0)
                return buf.getvalue()
        except:
            try:
                import openpyxl
                wb = openpyxl.Workbook()
                buf = io.BytesIO()
                wb.save(buf)
                buf.seek(0)
                return buf.getvalue()
            except:
                return b''








def show_leaflet_map(map_data, height=450, zoom=12, key_suffix=""):
    """Mappa Leaflet OSM - sempre visibile - fix Ezio - no Mapbox - con Fullscreen 100% come Mappe Postazioni"""
    try:
        if not map_data:
            avg_lat, avg_lon = 45.8167, 8.90417
            markers_js = """
            L.marker([45.75217, 8.90417]).addTo(map).bindPopup('Demo Varese IR2DO');
            L.marker([45.77, 9.02]).addTo(map).bindPopup('Demo Caronno');
            """
        else:
            try:
                avg_lat = sum([p['lat'] for p in map_data])/len(map_data)
                avg_lon = sum([p['lon'] for p in map_data])/len(map_data)
            except:
                avg_lat, avg_lon = 45.8167, 8.90417
            markers_js = ""
            for p in map_data:
                try:
                    lat = p.get('lat', 0)
                    lon = p.get('lon', 0)
                    idr = str(p.get('ID','')).replace("'", "").replace('"','')
                    col = p.get('color','#FF0000')
                    markers_js += f"L.circleMarker([{lat}, {lon}], {{color: '{col}', radius: 12, fillColor: '{col}', fillOpacity: 0.9, weight: 2}}).addTo(map).bindPopup('{idr} - {lat},{lon}');\n"
                except:
                    continue
        
        html = f"""
        <div id='map_{key_suffix}' style='width:100%;height:{height}px;border:3px solid #1A5D1A;border-radius:10px;'></div>
        <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
        <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
        <script>
        var map = L.map('map_{key_suffix}').setView([{avg_lat}, {avg_lon}], {zoom});
        L.tileLayer('https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{maxZoom: 19, attribution: 'OSM ANA Varese'}}).addTo(map);
        {markers_js}
        var fsControl = L.Control.extend({{
            onAdd: function(map) {{
                var cEl=L.DomUtil.create('div','leaflet-bar leaflet-control');
                cEl.innerHTML='FULL';
                cEl.title='Fullscreen 100%';
                cEl.style.background='white';cEl.style.width='50px';cEl.style.height='34px';
                cEl.style.lineHeight='34px';cEl.style.textAlign='center';cEl.style.cursor='pointer';
                cEl.style.fontSize='14px';cEl.style.fontWeight='bold';
                cEl.onclick=function(){{
                    var el = document.getElementById('map_{key_suffix}');
                    if(el.requestFullscreen) el.requestFullscreen();
                    else if(document.documentElement.requestFullscreen) document.documentElement.requestFullscreen();
                    el.style.height='100vh';
                    setTimeout(function(){{ map.invalidateSize(); }}, 500);
                }};
                return cEl;
            }}
        }});
        map.addControl(new fsControl({{position:'topright'}}));
        </script>
        """
        st.components.v1.html(html, height=height+20)
    except Exception as e:
        st.error(f"Errore mappa Leaflet: {e}")
        try:
            df_fallback = pd.DataFrame([{"lat": p['lat'], "lon": p['lon']} for p in map_data])
            st.map(df_fallback)
        except:
            st.map(pd.DataFrame([{"lat": 45.75217, "lon": 8.90417}]))


def excel_import_inline(form_key, form_label):
    """
    Import/Export inline per ogni form - Ezio richiesta - ORA CON PDF
    Ogni form ha Excel + PDF + Template ODV + Import - Office 2016 compatibile
    Lascia tutti campi aggiunti OK
    """
    st.divider()
    st.markdown(f"Import/Export")

    c1, c2, c3, c4 = st.columns(4)

    # Export Excel corrente - SOLO XLSX - MAI CSV - Richiesta Ezio - Fix per tutti i form
    with c1:
        data = st.session_state.get(form_key, [])
        if data:
            clean = [{kk: vv for kk, vv in r.items() if "Bytes" not in kk and "Foto" not in kk and "File" not in kk} for r in data if isinstance(r, dict)]
            if clean:
                df_exp = pd.DataFrame(clean)
                try:
                    excel_data = to_excel(df_exp)
                    if excel_data and len(excel_data) > 100 and excel_data[:2] == b'PK':
                        st.download_button(
                            f"⬇️ Excel {form_label}",
                            data=excel_data,
                            file_name=f"{form_key}_export_{datetime.now().strftime('%Y%m%d')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                            key=f"exp_inline_{form_key}"
                        )
                        st.caption(f"Excel XLSX: {len(data)} record - {len(excel_data)} bytes")
                    else:
                        st.error(f"Excel non valido - len {len(excel_data) if excel_data else 0} - Verifica openpyxl su Cloud")
                        st.info("Su Streamlit Cloud: Manage app -> Reboot - Verifica requirements.txt contenga openpyxl==3.1.5")
                        # Tentativo emergenza XLSX
                        try:
                            import openpyxl
                            wb = openpyxl.Workbook()
                            ws = wb.active
                            ws.title = "Dati"
                            for c_idx, col in enumerate(df_exp.columns, 1):
                                ws.cell(row=1, column=c_idx, value=str(col))
                            for r_idx, row in enumerate(df_exp.itertuples(index=False), 2):
                                for c_idx, val in enumerate(row, 1):
                                    ws.cell(row=r_idx, column=c_idx, value=str(val)[:32000])
                            from io import BytesIO
                            buf = io.BytesIO()
                            wb.save(buf)
                            buf.seek(0)
                            st.download_button(
                                f"⬇️ Excel Emergenza {form_label}",
                                data=buf.getvalue(),
                                file_name=f"{form_key}_EMERGENZA_{datetime.now().strftime('%Y%m%d')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True,
                                key=f"exp_emerg_{form_key}_xlsx"
                            )
                        except Exception as e_em:
                            st.error(f"Emergenza XLSX fallita: {e_em}")
                except Exception as e:
                    st.error(f"Excel errore: {str(e)[:200]}")
                    st.info("Su Streamlit Cloud: verifica requirements.txt contenga openpyxl==3.1.5 poi Reboot - Vedi log: Manage app -> Logs")
            else:
                st.info("Nessun dato")
        else:
            st.info(f"{form_label} vuoto")

    # Export PDF corrente - NUOVO RICHIESTA EZIO
    with c2:
        data_pdf = st.session_state.get(form_key, [])
        if data_pdf:
            clean_pdf = [{kk: vv for kk, vv in r.items() if "Bytes" not in kk and "Foto" not in kk and "File" not in kk} for r in data_pdf if isinstance(r, dict)]
            if clean_pdf:
                df_pdf = pd.DataFrame(clean_pdf)
                if REPORTLAB_OK:
                    try:
                        pdf_data = to_pdf(df_pdf, form_label.upper())
                        st.download_button(
                            f"📄 PDF {form_label}",
                            data=pdf_data,
                            file_name=f"{form_key}_{datetime.now().strftime('%Y%m%d')}.pdf",
                            mime="application/pdf",
                            use_container_width=True,
                            key=f"pdf_inline_{form_key}",
                            type="primary"
                        )
                        st.caption(f"PDF: {len(data_pdf)} record")
                    except Exception as e:
                        st.error(f"PDF errore: {e}")
                else:
                    st.warning("Reportlab non installato - aggiungi in requirements.txt: reportlab")
            else:
                st.info("Nessun dato PDF")
        else:
            st.info("Nessun dato per PDF")

    # Template vuoto per ODV
    with c3:
        data_existing = st.session_state.get(form_key, [])
        if data_existing and len(data_existing) > 0:
            first = data_existing[0]
            cols = [k for k in first.keys() if "Bytes" not in k and "Foto" not in k and "File" not in k]
            if not cols:
                cols = ["Nome","Cognome","Note"]
        else:
            if form_key == "volontari":
                cols = ["Nome","Cognome","Comune","Via","CapoODV","ODVAppartenenza","DataNascita","CodFisc","Cellulare","Email","TelEmergenza","Ruolo","Squadra","RadioID","Documento","ScadDoc","Note"]
            elif form_key == "radio_db":
                cols = ["ID","Modello","Frequenza","Canale","Note"]
            elif form_key == "consegna_radio":
                cols = ["Data","Volontario","RadioID","Note"]
            elif form_key == "mezzi":
                cols = ["Targa","Modello","Tipo","ODV","Stato","Note"]
            elif form_key == "attrezzature":
                cols = ["Nome","Tipo","Quantita","ODV","Stato","Note"]
            elif form_key == "brogliaccio":
                cols = ["Data","Evento","Descrizione","Operatore","Note"]
            elif form_key == "eventi":
                cols = ["Data","Titolo","Luogo","Descrizione","Note"]
            elif form_key == "emergenze":
                cols = ["Data","Tipo","Luogo","Descrizione","Note"]
            else:
                cols = ["Campo1","Campo2","Campo3","Note"]

        df_template = pd.DataFrame(columns=cols)
        try:
            tpl_data = to_excel(df_template)
            if tpl_data and len(tpl_data) > 100:
                st.download_button(
                    f"📋 Template {form_label} ODV",
                    data=tpl_data,
                    file_name=f"TEMPLATE_{form_key}_ODV_Office2016.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key=f"tpl_inline_{form_key}",
                    help="File .xlsx puro compatibile Office 2016 - solo intestazioni"
                )
            else:
                st.warning("Template Excel non disponibile - openpyxl mancante su Cloud")
                st.info("Su GitHub verifica requirements.txt contenga openpyxl, poi Manage app -> Reboot")
        except Exception as e:
            st.error(f"Template errore: {str(e)[:100]}")
            st.info("Fix: requirements.txt deve contenere openpyxl")
        st.caption("Template vuoto per ODV")

    # Import
    with c4:
        up_mode = st.radio("Modalità import", ["Aggiungi","Sostituisci"], key=f"mode_inline_{form_key}", horizontal=True)
        up_file = st.file_uploader(f"Carica Excel per {form_label}", type=["xlsx","xls"], key=f"up_inline_{form_key}")

        if up_file:
            try:
                df_imp = None
                last_err = ""
                # Prova tutti gli engine per compatibilità Office 2016
                for eng in [None, "openpyxl", "xlrd"]:
                    try:
                        up_file.seek(0)
                        if eng is None:
                            df_imp = pd.read_excel(up_file)
                        else:
                            df_imp = pd.read_excel(up_file, engine=eng)
                        if df_imp is not None and len(df_imp.columns) > 0:
                            break
                    except Exception as e:
                        last_err = str(e)
                        continue

                # FIX: accetta anche file con solo header + dati, e mostra anche se vuoto
                if df_imp is not None:
                    # Pulisci
                    try:
                        df_imp = df_imp.dropna(how='all')
                        # Pulisci colonne Unnamed
                        if not df_imp.columns.empty:
                            df_imp = df_imp.loc[:, ~df_imp.columns.astype(str).str.contains('^Unnamed', na=False)]
                    except:
                        pass

                    if df_imp.empty:
                        # File ha solo intestazioni o vuoto - mostra colonne
                        if len(df_imp.columns) > 0:
                            st.warning(f"File letto: {len(df_imp.columns)} colonne trovate ma 0 righe dati")
                            st.write(f"Colonne: {list(df_imp.columns)}")
                            st.info("💡 Aggiungi righe dati sotto intestazione in Excel e ricarica")
                            st.dataframe(pd.DataFrame(columns=df_imp.columns).head(), use_container_width=True)
                        else:
                            st.warning("Excel vuoto - solo intestazioni? Aggiungi righe e ricarica")
                    else:
                        st.success(f"✅ {len(df_imp)} righe lette da Excel - {len(df_imp.columns)} colonne")
                        st.write(f"Colonne: {list(df_imp.columns)}")
                        # Richiesta Ezio: non mostrare tabella sotto, solo con tasto
                        if st.button(f"📋 Mostra anteprima {len(df_imp)} righe", key=f"btn_preview_{form_key}_{len(df_imp)}"):
                            st.session_state[f"show_preview_{form_key}"] = not st.session_state.get(f"show_preview_{form_key}", False)
                        if st.session_state.get(f"show_preview_{form_key}", False):
                            st.dataframe(df_imp.head(20), use_container_width=True)

                        if st.button(f"✅ Importa {len(df_imp)} righe in {form_label}", type="primary", use_container_width=True, key=f"btn_imp_inline_{form_key}"):
                            imported = df_imp.to_dict(orient="records")
                            # Pulisci NaN
                            cleaned = []
                            for r in imported:
                                nr = {}
                                for k,v in r.items():
                                    if pd.isna(v):
                                        continue
                                    if isinstance(v, (pd.Timestamp, datetime, date)):
                                        nr[k] = v.strftime("%d/%m/%Y")
                                    else:
                                        nr[k] = str(v).strip() if isinstance(v, str) else v
                                # Salta righe vuote
                                if any(nr.values()):
                                    cleaned.append(nr)

                            if up_mode.startswith("Sostituisci"):
                                st.session_state[form_key] = cleaned
                            else:
                                st.session_state[form_key] = st.session_state.get(form_key, []) + cleaned

                            st.success(f"Importati {len(cleaned)} in {form_label}!")
                            st.balloons()
                            st.rerun()
                elif df_imp is not None:
                    st.warning("Excel vuoto - solo intestazioni? Aggiungi righe e ricarica")
                else:
                    st.error(f"Errore: {last_err}")
                    if "openpyxl" in last_err.lower():
                        st.error("Office 2016 FIX: Salva file come .xlsx (non .xls) in Office 2016 -> File -> Salva con nome -> Cartella di lavoro Excel (*.xlsx)")
            except Exception as e:
                st.error(f"Errore import: {e}")



def to_pdf(df, tit):
    """
    PDF con logo a SX del titolo + totale solo per Spese ODV
    """
    global REPORTLAB_OK
    import io
    try:
        buf = io.BytesIO()
        if not REPORTLAB_OK:
            raise ImportError("reportlab non disponibile")

        from reportlab.lib.pagesizes import landscape, A4
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.lib import colors
        from reportlab.lib.units import cm

        doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=1*cm, rightMargin=1*cm, topMargin=1.5*cm, bottomMargin=1*cm)
        styles = getSampleStyleSheet()
        story = []
        # LOGO A SX DEL TITOLO
        try:
            if os.path.exists("logo.png"):
                logo_img = Image("logo.png", width=70, height=70)
            else:
                logo_img = Paragraph("", styles["Normal"])
        except:
            logo_img = Paragraph("", styles["Normal"])
        
        title_para = Paragraph(f"<b>{tit} - ANA Varese Protezione Civile</b><br/><font size=9>Generato il {datetime.now().strftime('%d/%m/%Y %H:%M')}</font>", styles["Title"])
        try:
            header_table = Table([[logo_img, title_para]], colWidths=[3*cm, landscape(A4)[0] - 5*cm])
            header_table.setStyle(TableStyle([
                ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                ("ALIGN", (0,0), (0,0), "LEFT"),
                ("ALIGN", (1,0), (1,0), "LEFT"),
                ("LEFTPADDING", (0,0), (-1,-1), 0),
                ("RIGHTPADDING", (0,0), (-1,-1), 0),
            ]))
            story.append(header_table)
        except:
            story.append(title_para)
        story.append(Spacer(1, 16))

        if df is not None and not df.empty:
            df_clean = df.copy()
            for col in df_clean.columns:
                if df_clean[col].dtype == object:
                    df_clean[col] = df_clean[col].apply(lambda x: str(x)[:80].replace("<","").replace(">","") if not isinstance(x, (bytes, bytearray)) else "")
            cols = list(df_clean.columns)[:12]
            if len(cols) == 0:
                cols = ["Dato"]
            header = cols
            rows = []
            for _, r in df_clean.iterrows():
                row_vals = []
                for c in cols:
                    try:
                        val = r.get(c, "")
                        row_vals.append(str(val)[:80].replace("<","").replace(">",""))
                    except:
                        row_vals.append("")
                rows.append(row_vals)
            data = [header] + rows
            available_width = landscape(A4)[0] - 2 * cm
            col_width = available_width / len(cols) if cols else available_width
            col_widths = [col_width] * len(cols)
            t = Table(data, colWidths=col_widths, repeatRows=1)
            t.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1A5D1A")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 7),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#e8f5e9")]),
            ]))
            story.append(t)
        else:
            story.append(Paragraph("Nessun dato disponibile", styles["Normal"]))

        doc.build(story)
        buf.seek(0)
        return buf.getvalue()
    except Exception as e:
        try:
            buf2 = io.BytesIO()
            from reportlab.lib.pagesizes import landscape, A4
            from reportlab.platypus import SimpleDocTemplate, Paragraph
            from reportlab.lib.styles import getSampleStyleSheet
            doc2 = SimpleDocTemplate(buf2, pagesize=landscape(A4))
            styles2 = getSampleStyleSheet()
            story2 = [Paragraph(f"{tit} - Errore PDF: {str(e)[:200]}", styles2["Title"])]
            doc2.build(story2)
            buf2.seek(0)
            return buf2.getvalue()
        except:
            # Fallback PDF minimo valido
            return b"%PDF-1.4 fallback PDF"


def hdr():

    """
    Header con 2 loghi affiancati + intestazione centrata - FIX etichetta verde scuro centrata nel verde chiaro
    """
    # Colonne con allineamento verticale centrato - FIX etichetta verde scuro centrata
    try:
        c1, c2, c3 = st.columns([0.8, 0.8, 4.4], vertical_alignment="center")
    except:
        c1, c2, c3 = st.columns([0.8, 0.8, 4.4])
    # CSS specifico solo per header - non globale - FIX etichetta verde scuro centrata
    st.markdown("""
    <style>
    /* Solo header - centra verticalmente */
    div[data-testid="column"]:has(> div > div > img) {
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
    }
    </style>
    """, unsafe_allow_html=True)
    with c1:
        try:
            if os.path.exists("logo.png"):
                st.image("logo.png", width=110)
            else:
                st.markdown(
                    """
                    <div style="width:80px;height:80px;background:#1A5D1A;
                    border-radius:10px;display:flex;align-items:center;
                    justify-content:center;color:white;font-weight:bold;
                    font-size:24px;text-align:center;line-height:80px;">
                    ANA
                    </div>
                    """,
                    unsafe_allow_html=True
                )
        except:
            st.markdown("**ANA**")
    
    with c2:
        try:
            # Secondo logo - Gruppo Caronno Pertusella - allegato Ezio - a fianco pcana
            logo2_path = None
            for p in ["gruppo_CPB.jpeg", "logo2.png", "gruppo_caronno.png", "logo_gruppo.png", "Gruppo_Caronno.png", "/mnt/data/gruppo_CPB.jpeg", "/mnt/data/logo2.png", "/mnt/data/gruppo_caronno.png"]:
                if os.path.exists(p):
                    logo2_path = p
                    break
            if logo2_path:
                st.image(logo2_path, width=110)
            else:
                st.markdown(
                    """
                    <div style="width:80px;height:80px;background:#0D47A1;
                    border-radius:10px;display:flex;align-items:center;
                    justify-content:center;color:white;font-weight:bold;
                    font-size:20px;text-align:center;line-height:80px;">
                    GRUPPO
                    </div>
                    """,
                    unsafe_allow_html=True
                )
        except:
            st.markdown("**GRUPPO**")

    with c3:
        st.markdown(
            """
            <div style="background:linear-gradient(135deg,#1A5D1A 0%,#2e7d32 100%);
            padding:12px 16px;border-radius:10px;color:white;
            text-align:center;border:3px solid #1A5D1A;
            display:flex;flex-direction:column;justify-content:center;align-items:center;
            min-height:115px;box-sizing:border-box;">
                                <p style="margin:0;font-family:Times New Roman;
                font-weight:bold;font-size:22px;color:white;line-height:1.3;letter-spacing:0.5px;text-align:center;">
                Squadra Volontari di protezione civile - Gruppo Alpini di Caronno Pertusella Bariola
                </p>
                <p style="margin:6px 0 0 0;font-family:Times New Roman;
                font-weight:bold;font-size:20px;color:white;line-height:1.25;letter-spacing:0.5px;text-align:center;">
                NUCLEO VOLONTARI DI P.C. A.N.A. - SEZIONE DI VARESE
                </p>
                <p style="margin:6px 0 0 0;font-family:Times New Roman;
                font-weight:bold;font-size:18px;color:white;line-height:1.2;letter-spacing:0.8px;text-align:center;">
                ASSOCIAZIONE NAZIONALE ALPINI
                </p>
            </div>
            """,
            unsafe_allow_html=True
        )


def hdr_form(t):
    st.markdown(
        f"""
        <h2 style="font-family:'Times New Roman',Times,serif;font-weight:bold;color:white;background: linear-gradient(135deg,#1A5D1A 0%,#2e7d32 100%);border: 3px solid #FFD700;padding:12px 16px;margin-top:0px;margin-bottom:14px;font-size:22px;border-radius:10px;">
        {t}
        </h2>
        """,
        unsafe_allow_html=True
    )

def get_form_key(base, form_ver_key):
    """Restituisce chiave versionata per pulisci maschera che funziona - Ezio"""
    ver = st.session_state.get(form_ver_key, 0)
    return f"{base}_v{ver}"

def pulisci_maschera_form(form_ver_key, keys_prefixes):
    """Pulisci maschera con versionamento - funziona davvero - Ezio"""
    # Incrementa versione
    st.session_state[form_ver_key] = st.session_state.get(form_ver_key, 0) + 1
    # Cancella chiavi con prefissi
    for k in list(st.session_state.keys()):
        for pref in keys_prefixes:
            if k.startswith(pref):
                try:
                    del st.session_state[k]
                except:
                    pass
    try:
        # query_params disattivato per stabilità
        pass
    except:
        pass
    st.rerun()


    st.markdown("""
    <style>
    .block-container { padding-top: 1rem !important; }
    [data-testid="stVerticalBlock"] { gap: 0.8rem !important; }
    </style>
    """, unsafe_allow_html=True)




# POPOUT INIZIALE SOLO ICONA MANIFESTO - Richiesta Ezio - TEMPO CONFIGURABILE
SPLASH_SECONDS = 5  # Fix Ezio - 5 secondi splash manifesto

def inject_popout_splash():
    try:
        import base64, os
        b64 = None
        for p in ["/mnt/data/manifesto_GAE.jpg", "manifesto_protezione_civile_GAE.jpg", "/mnt/data/copertina_GAE_splash.jpg", "manifesto_GAE.jpg", "manifesto_protezione_civile_GAE.jpg", "/mnt/data/manifesto_protezione_civile_ANA.jpg", "manifesto_protezione_civile_ANA.jpg", "/mnt/data/copertina.png", "copertina.png", "logo_GAE.png", "logo.png"]:
            if os.path.exists(p):
                with open(p, "rb") as fh:
                    b64 = base64.b64encode(fh.read()).decode()
                break
        if not b64:
            return
        tempo = globals().get("SPLASH_SECONDS", 3)
        html = """
        <div id="ph"></div>
        <script>
        (function(){
            var parentDoc = window.parent.document;
            if(parentDoc.getElementById('popout-splash-ezio')) return;
            var overlay = parentDoc.createElement('div');
            overlay.id = 'popout-splash-ezio';
            overlay.style.cssText = 'position:fixed;top:0;left:0;width:100vw;height:100vh;background:rgba(0,0,0,0.92);z-index:99999999;display:flex;flex-direction:column;align-items:center;justify-content:center;cursor:pointer;';
            overlay.innerHTML = '<div style="display:flex;flex-direction:column;align-items:center;justify-content:center;"><img src="__IMG_SRC__" style="max-width:90vw;max-height:85vh;width:auto;height:auto;object-fit:contain;border-radius:16px;box-shadow:0 15px 50px rgba(0,0,0,0.8);border:4px solid #FFD700;"><div style="margin-top:15px;background:rgba(0,0,0,0.6);padding:8px 18px;border-radius:20px;display:flex;align-items:center;gap:10px;border:1px solid #FFD700;"><span style="color:white;font-family:Times New Roman, serif;font-size:13px;">Chiusura tra <span id="countdown-ezio" style="font-weight:bold;font-size:16px;color:#FFD700;">__SECS__</span>s</span><div style="width:80px;height:4px;background:rgba(255,255,255,0.3);border-radius:2px;overflow:hidden;"><div id="progress-ezio" style="background:#FFD700;height:100%;width:100%;transition:width 1s linear;"></div></div></div></div>';
            parentDoc.body.appendChild(overlay);
            var seconds = __SECS__;
            var totalSecs = __SECS__;
            var countdownEl = parentDoc.getElementById('countdown-ezio');
            var progressEl = parentDoc.getElementById('progress-ezio');
            function chiudi(){ overlay.style.opacity='0'; overlay.style.transition='opacity 0.5s'; setTimeout(function(){ if(overlay.parentNode) overlay.parentNode.removeChild(overlay); }, 500); }
            var interval = setInterval(function(){
                seconds--;
                if(countdownEl) countdownEl.textContent = seconds;
                if(progressEl) progressEl.style.width = (seconds*(100/totalSecs)) + '%';
                if(seconds <= 0){ clearInterval(interval); chiudi(); }
            }, 1000);
            overlay.addEventListener('click', function(){ clearInterval(interval); chiudi(); });
            parentDoc.addEventListener('keydown', function escHandler(e){ if(e.key==='Escape'){ clearInterval(interval); chiudi(); parentDoc.removeEventListener('keydown', escHandler); } });
        })();
        </script>
        """.replace("__IMG_SRC__", "data:image/jpeg;base64," + b64).replace("__SECS__", str(tempo))
        st.components.v1.html(html, height=0)
    except:
        pass


def init_session():


    defaults = {
        "page": "entra",
        "logged": False,
        "menu": "Dashboard",
        "volontari": [],
        "radio_db": [],
        "consegna_radio": [],
        "eventi": [],
        "emergenze": [],
        "checkin": [],
        "icone": [],
        "postazioni": [],
        "temp_markers": [],
        "brogliaccio": [],
        "mezzi": [],
        "attrezzature": [],
        "map_fullscreen": False,
        "vol_form_data": {},
        "alias_radio": [],
        "brog_evento_blindato": None,
        "brog_emergenza_blindata": None,
        "brog_blindato": False,
        "check_evento_blindato": None,
        "check_emergenza_blindata": None,
        "check_blindato": False,
        "interventi": [],
        "interventi_emergenza_blindata": None,
        "interventi_blindato": False,
        "chat": [],
        "tabella_interventi": [],
        "json_visualizzato": None,
        "posizioni_pd785": [],
        "posizioni_anytone": [],
        "vol_edit_index": None,
        "int_edit_index": None,
        "mappe": []
    }

    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

# --- GESTIONE UTENTI + PRESENZA CHAT - Ezio richiesta - FIX accesso negato ---
import json
import hashlib

UTENTI_FILE = "utenti.json"
PRESENZA_FILE = "presenza.json"

def hash_pwd(pwd):
    return hashlib.sha256(pwd.encode()).hexdigest()

def load_utenti():
    default_utenti = [
        {"username": "admin", "password": hash_pwd("ana2024"), "nome": "Amministratore ANA", "ruolo": "amministratore", "attivo": True, "permessi": [], "creato_da": "sistema", "data_creazione": "01/01/2026"},
        {"username": "operatore1", "password": hash_pwd("operatore1"), "nome": "Operatore 1", "ruolo": "operatore", "attivo": True, "permessi": [], "creato_da": "admin", "data_creazione": "01/01/2026"},
        {"username": "lettore1", "password": hash_pwd("lettore1"), "nome": "Lettore 1", "ruolo": "lettore", "attivo": True, "permessi": [], "creato_da": "admin", "data_creazione": "01/01/2026"},
    ]
    try:
        if os.path.exists(UTENTI_FILE):
            with open(UTENTI_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if data:
                    return data
    except:
        pass
    try:
        with open(UTENTI_FILE, 'w', encoding='utf-8') as f:
            json.dump(default_utenti, f, indent=2, ensure_ascii=False)
    except:
        pass
    return default_utenti

def save_utenti(utenti_list):
    try:
        with open(UTENTI_FILE, 'w', encoding='utf-8') as f:
            json.dump(utenti_list, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"Errore save utenti: {e}")
        return False

def load_presenza():
    try:
        if os.path.exists(PRESENZA_FILE):
            with open(PRESENZA_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
    except:
        pass
    return []

def save_presenza(presenza_list):
    try:
        now = datetime.now()
        fresh = []
        for p in presenza_list:
            try:
                t = datetime.fromisoformat(p.get("ultimo_accesso", ""))
                if (now - t).total_seconds() < 1800:
                    fresh.append(p)
            except:
                fresh.append(p)
        with open(PRESENZA_FILE, 'w', encoding='utf-8') as f:
            json.dump(fresh, f, indent=2, ensure_ascii=False)
    except:
        pass

def aggiorna_presenza(username, nome, ruolo):
    try:
        presenza = load_presenza()
        presenza = [p for p in presenza if p.get("username") != username]
        presenza.append({
            "username": username,
            "nome": nome,
            "ruolo": ruolo,
            "ultimo_accesso": datetime.now().isoformat(),
            "ora": datetime.now().strftime("%H:%M:%S")
        })
        save_presenza(presenza)
    except:
        pass

def rimuovi_presenza(username):
    try:
        presenza = load_presenza()
        presenza = [p for p in presenza if p.get("username") != username]
        save_presenza(presenza)
    except:
        pass


init_session()

# === HUB UNICO - 2 PROGETTI + SPLASH COME PRIMO FILE ===
if st.session_state.get("page") != "entra" and st.session_state.get("logged", False):
    if st.session_state.hub_page == "hub":
        try:
            inject_popout_splash()
        except:
            pass
        st.markdown("""
        <div style="text-align:center;padding:30px;background:linear-gradient(135deg,#1A5D1A,#2e7d32);border-radius:15px;color:white;border:3px solid #FFD700;">
        <h1>🏠 HUB ANA - UNICO - 2 Progetti</h1>
        <h3>Amministrativo 13 form + Operativo 17 form - Senza Geolocalizzazione</h3>
        <p>File base App-Con-Spese_89.py - Con Splash, Eventi ripristinato, Statistiche, Note Spese</p>
        </div>
        """, unsafe_allow_html=True)
        st.divider()
        c1,c2 = st.columns(2, gap="large")
        with c1:
            st.markdown('<div style="text-align:center;background:#1565C0;color:white;padding:16px;border-radius:12px;font-weight:bold;font-size:18px">📋 AMMINISTRATIVO<br><small>13 form - Splash, Volontari, Ospiti, Eventi, Spese, Note Spese, Statistiche, Backup</small></div>', unsafe_allow_html=True)
            st.write("")
            if st.button("📋 ENTRA AMMINISTRATIVO\n13 form\nGestionale ODV", key="hub_admin_unico", use_container_width=True):
                st.session_state.hub_page = "admin"
                st.session_state.menu = "Dashboard"
                st.session_state.menu_admin = "Dashboard"
                st.rerun()
        with c2:
            st.markdown('<div style="text-align:center;background:#1A5D1A;color:white;padding:16px;border-radius:12px;font-weight:bold;font-size:18px">🚒 OPERATIVO<br><small>17 form - Splash, Eventi, Radio, Emergenze, Mezzi, Mappe, Turni</small></div>', unsafe_allow_html=True)
            st.write("")
            if st.button("🚒 ENTRA OPERATIVO\n17 form\nEmergenze e Radio", key="hub_oper_unico", use_container_width=True):
                st.session_state.hub_page = "operativo"
                st.session_state.menu = "Dashboard"
                st.session_state.menu_oper = "Dashboard"
                st.rerun()
        st.divider()
        st.info("Form Eventi ripristinato - Geolocalizzazione rimossa - Splash come primo file mantenuto")
        st.stop()


# FIX GLOBALE DISABILITATO - ora icona è vero tasto bottone, non link
# Gestione query param disabilitata - icona ora è vero bottone st.button, non link <a>

# PAGINA ENTRA - con footer fisso in basso Developed by Ezio F. 2026 Vers 1.0 - Logo cartoon
if st.session_state.page == "entra":
    hdr()
    try:
        inject_popout_splash()
    except:
        pass
    st.write("")
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        try:
            if os.path.exists("copertina.png"):
                st.image("copertina.png", width=350)
            else:
                st.markdown(
                    """
                    <div style="text-align:center;padding:40px;
                    background:linear-gradient(135deg,#e8f5e9,#c8e6c9);
                    border-radius:16px;border:2px dashed #1A5D1A;">
                    <div style="font-size:100px;">🛡️</div>
                    <p style="font-weight:bold;">Squadra Volontari di protezione civile - Gruppo Alpini di Caronno Pertusella Bariola</p>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
        except:
            st.markdown("### ANA Varese")

        st.markdown(
            """
            <h2 style="text-align:center;font-family:Times New Roman; font-weight:bold; color:black; margin-top:20px; white-space:nowrap; font-size:26px;">
            GESTIONALE DI PROTEZIONE CIVILE
            </h2>
            <p style="text-align:center;">
            Versione BETA
            </p>
            """,
            unsafe_allow_html=True
        )

        if st.button(
            "ENTRA NEL GESTIONALE",
            type="primary",
            use_container_width=True
        ):
            st.session_state.page = "login"
            st.rerun()
    
    # FOOTER FISSO IN BASSO - fuori da c2, centrato in tutta pagina - Fix posizione Ezio
    # CSS fixed bottom
    try:
        import base64
        logo_path = None
        if os.path.exists("logo_dev_ezio.png"):
            logo_path = "logo_dev_ezio.png"
        elif os.path.exists("/mnt/data/logo_dev_ezio.png"):
            logo_path = "/mnt/data/logo_dev_ezio.png"
        
        if logo_path:
            with open(logo_path, "rb") as f:
                b64_logo = base64.b64encode(f.read()).decode()
            # Cerca copertina.png per icona piccola a sx - richiesta Ezio
            copertina_b64 = None
            cop_path = None
            for p in ["copertina.png", "/mnt/data/copertina.png"]:
                if os.path.exists(p):
                    cop_path = p
                    break
            if cop_path:
                with open(cop_path, "rb") as cf:
                    copertina_b64 = base64.b64encode(cf.read()).decode()
            
            # Footer fisso con copertina.png piccola a sx + scritta - richiesta Ezio
            if copertina_b64:
                # Usa copertina.png piccola 20px a sx
                st.markdown(f"""
                <style>
                .footer-developed {{
                    position: fixed;
                    left: 0;
                    bottom: 0;
                    width: 100%;
                    background: linear-gradient(135deg,#1A5D1A 0%,#2e7d32 100%);
                    color: white;
                    text-align: center;
                    padding: 6px 0px;
                    z-index: 9999;
                    border-top: 3px solid #FFD700;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    gap: 8px;
                    font-family: 'Times New Roman', serif;
                }}
                .footer-developed .icon-copertina {{
                    width: 20px;
                    height: 20px;
                    object-fit: contain;
                    background: white;
                    border-radius: 3px;
                    padding: 1px;
                }}
                </style>
                <div class="footer-developed">
                    <img class="icon-copertina" src="data:image/png;base64,{copertina_b64}" alt="copertina">
                    <span style="font-weight:bold;font-size:13px;">Developed by Ezio F. 2026 Vers. 1.0 - ANA Varese Protezione Civile</span>
                </div>
                <div style="height:50px;"></div>
                """, unsafe_allow_html=True)
            else:
                # Fallback con logo_dev_ezio piccolo
                st.markdown(f"""
                <style>
                .footer-developed {{
                    position: fixed;
                    left: 0;
                    bottom: 0;
                    width: 100%;
                    background: linear-gradient(135deg,#1A5D1A 0%,#2e7d32 100%);
                    color: white;
                    text-align: center;
                    padding: 6px 0px;
                    z-index: 9999;
                    border-top: 3px solid #FFD700;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    gap: 8px;
                    font-family: 'Times New Roman', serif;
                }}
                .footer-developed img {{
                    width: 20px;
                    height: 20px;
                    border-radius: 3px;
                    object-fit: cover;
                    background: white;
                }}
                </style>
                <div class="footer-developed">
                    <img src="data:image/png;base64,{b64_logo}" alt="Logo Ezio">
                    <span style="font-weight:bold;font-size:13px;">Developed by Ezio F. 2026 Vers. 1.0 - ANA Varese Protezione Civile</span>
                </div>
                <div style="height:50px;"></div>
                """, unsafe_allow_html=True)
        else:
            # Fallback con copertina.png piccola a sx - richiesta Ezio
            cop_b64 = None
            for pp in ["copertina.png", "/mnt/data/copertina.png"]:
                if os.path.exists(pp):
                    with open(pp, "rb") as ff:
                        cop_b64 = base64.b64encode(ff.read()).decode()
                    break
            if cop_b64:
                st.markdown(f"""
                <style>
                .footer-developed {{
                    position: fixed;
                    left: 0;
                    bottom: 0;
                    width: 100%;
                    background: #1A5D1A;
                    color: white;
                    text-align: center;
                    padding: 6px;
                    z-index: 9999;
                    border-top: 3px solid #FFD700;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    gap: 8px;
                    font-family: 'Times New Roman', serif;
                }}
                .footer-developed .icon-copertina {{
                    width: 20px;
                    height: 20px;
                    object-fit: contain;
                    background: white;
                    border-radius: 3px;
                    padding: 1px;
                }}
                </style>
                <div class="footer-developed">
                    <img class="icon-copertina" src="data:image/png;base64,{cop_b64}" alt="copertina">
                    <span style="font-weight:bold;font-size:13px;">Developed by Ezio F. 2026 Vers. 1.0 - ANA Varese Protezione Civile</span>
                </div>
                <div style="height:50px;"></div>
                """, unsafe_allow_html=True)
            else:
                st.markdown("""
                <style>
                .footer-developed {{
                    position: fixed;
                    left: 0;
                    bottom: 0;
                    width: 100%;
                    background: #1A5D1A;
                    color: white;
                    text-align: center;
                    padding: 6px;
                    z-index: 9999;
                    border-top: 3px solid #FFD700;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    gap: 8px;
                    font-family: 'Times New Roman', serif;
                }}
                .footer-developed .icon-copertina {{
                    width: 20px;
                    height: 20px;
                    object-fit: contain;
                    background: white;
                    border-radius: 3px;
                }}
                </style>
                <div class="footer-developed">
                    <span style="font-weight:bold;font-size:13px;">Developed by Ezio F. 2026 Vers. 1.0 - ANA Varese Protezione Civile</span>
                </div>
                <div style="height:50px;"></div>
                """, unsafe_allow_html=True)
    except Exception as e:
        st.markdown(f"""
        <div style="position:fixed;bottom:0;left:0;width:100%;background:#1A5D1A;color:white;text-align:center;padding:8px;z-index:9999;border-top:2px solid #FFD700;">
        Developed by Ezio F. 2026 Vers. 1.0 - ANA Varese
        </div>
        <div style="height:50px;"></div>
        """, unsafe_allow_html=True)

    st.stop()

# PAGINA LOGIN - CON INTESTAZIONE - richiesta Ezio: metti intestazione prima pagina anche in login
if st.session_state.page == "login":
    hdr()  # Intestazione con 2 loghi
    # CSS login labels bianche - richiesta Ezio
    st.markdown("""
    <style>
    /* Login labels bianche */
    div[data-testid="stForm"] label, 
    div[data-testid="stForm"] .stTextInput label,
    div[data-testid="stForm"] p {
        color: white !important;
        font-weight: bold !important;
    }
    /* Forza bianco per Utente e Password nella pagina login */
    #login_form label, 
    [data-testid="stTextInput"] label {
        color: white !important;
    }
    /* Solo nella pagina login - form verde scuro con labels bianche */
    </style>
    """, unsafe_allow_html=True)
    # Override specifico per login
    st.markdown("""
    <style>
    /* Login form - labels bianche */
    section[data-testid="stSidebar"] ~ div label {
    }
    /* Login page specific */
    div:has(> div > #login_utente_form) label,
    div:has(> div > #login_pwd_form) label {
        color: white !important;
    }
    </style>
    """, unsafe_allow_html=True)
    # Etichetta LOGIN - Gestione Utenti Multi Livello rimossa - richiesta Ezio
    st.write("")
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        # Solo login pulito senza etichetta multiuso
        utenti_list = load_utenti()
        # st.info rimosso - non mostrare utenti configurati demo
        
        with st.form("login_form"):
            st.markdown('<p style="color:white !important;font-family:Times New Roman;font-weight:bold;margin-bottom:2px;">Utente</p>', unsafe_allow_html=True)
            utente = st.text_input("Utente", label_visibility="collapsed", key="login_utente_form")
            st.markdown('<p style="color:white !important;font-family:Times New Roman;font-weight:bold;margin-bottom:2px;margin-top:8px;">Password</p>', unsafe_allow_html=True)
            pwd = st.text_input("Password", label_visibility="collapsed", type="password", key="login_pwd_form")
            submitted = st.form_submit_button("Accedi", type="primary", use_container_width=True)
            
            if submitted:
                found = False
                for u in utenti_list:
                    if u.get("username") == utente.strip().lower() and u.get("password") == hash_pwd(pwd) and u.get("attivo", True):
                        st.session_state.logged = True
                        st.session_state.page = "dashboard"
                        st.session_state.menu = "Dashboard"
                        st.session_state.username = u.get("username")
                        st.session_state.nome_utente = u.get("nome")
                        st.session_state.ruolo_utente = u.get("ruolo")
                        # Aggiorna presenza per chat
                        try:
                            aggiorna_presenza(u.get("username"), u.get("nome"), u.get("ruolo"))
                        except:
                            pass
                        st.success(f"Accesso effettuato - Benvenuto {u.get('nome')} - Ruolo: {u.get('ruolo')}")
                        st.balloons()
                        st.rerun()
                        found = True
                        break
                if not found:
                    st.error("Credenziali errate o utente disattivato - Verifica username/password e che utente sia attivo")

        if st.button("Torna a Entra", use_container_width=True):
            st.session_state.page = "entra"
            st.rerun()

    st.stop()

# CONTROLLO LOGIN
if not st.session_state.logged:
    st.session_state.page = "login"
    st.rerun()

# SIDEBAR - MENU
with st.sidebar:
    try:
        if os.path.exists("logo.png"):
            st.image("logo.png", width=120)
        else:
            st.markdown(
                """
                <div style="width:120px;height:120px;background:#1A5D1A;
                border-radius:12px;display:flex;align-items:center;
                justify-content:center;color:white;font-weight:bold;font-size:18px;">
                ANA
                </div>
                """,
                unsafe_allow_html=True
            )
    except:
        st.write("ANA")

    st.markdown(
        """
        <p style="font-weight:bold;font-family:Times New Roman;
        margin-top:12px;color:#1A5D1A;">
        MENU 950+ MODIFICHE
        </p>
        """,
        unsafe_allow_html=True
    )

    # Menu base - Gestione Utenti solo per amministratore
    ruolo_corrente = st.session_state.get("ruolo_utente", "amministratore")
    if ruolo_corrente == "amministratore":
        menu_base = [
            "Dashboard",
            "Volontari (con foto)",
            "Ospiti",
            "DB Radio",
            "Consegna Radio",
            "Alias Radio",
            "Brogliaccio",
            "Eventi",
            "Emergenze",
            "Tabella Emergenze",
            "Check-in",
            "Interventi Emergenza",
            "Tabella Interventi Emergenza",
            "Mezzi",
            "Attrezzature",
            "Mappe Postazioni",
            "Libreria Icone",
            "Turni",
            "Chat",
            "Verbali",
            "Archivio Documenti",
            "Diplomi Attestati",
            "Spese ODV per Evento",
            "Report Filtro",
            "Geolocalizzazione Hytera + Anytone",
            "Gestione Utenti",
            "Backup"
        ]
    else:
        menu_base = [
            "Dashboard",
            "Volontari (con foto)",
            "Ospiti",
            "DB Radio",
            "Consegna Radio",
            "Alias Radio",
            "Brogliaccio",
            "Eventi",
            "Emergenze",
            "Tabella Emergenze",
            "Check-in",
            "Interventi Emergenza",
            "Tabella Interventi Emergenza",
            "Mezzi",
            "Attrezzature",
            "Mappe Postazioni",
            "Libreria Icone",
            "Turni",
            "Chat",
            "Verbali",
            "Archivio Documenti",
            "Diplomi Attestati",
            "Spese ODV per Evento",
            "Report Filtro",
            "Geolocalizzazione Hytera + Anytone",
            "Backup"
        ]

    # FIX: Se force_menu presente (click icona tabella emergenze), forza apertura Interventi Emergenza per modifiche
    if "force_menu" in st.session_state:
        try:
            forced = st.session_state["force_menu"]
            if forced in menu_base:
                st.session_state.menu = forced
            del st.session_state["force_menu"]
        except:
            pass

    # FIX: Assicura logged e page dashboard quando si clicca icona - evita pagina iniziale
    if st.session_state.get("int_edit_index") is not None and st.session_state.menu == "Interventi Emergenza":
        st.session_state.logged = True
        st.session_state.page = "dashboard"

    cur = st.radio(
        "Seleziona form",
        menu_base,
        index=menu_base.index(st.session_state.menu) if st.session_state.menu in menu_base else 0,
    )
    st.session_state.menu = cur

    st.divider()
    
    # Utente loggato + ruolo
    username = st.session_state.get("username", "admin")
    nome_utente = st.session_state.get("nome_utente", "Amministratore")
    ruolo_utente = st.session_state.get("ruolo_utente", "amministratore")
    
    # Colore ruolo
    colore_ruolo = {"amministratore": "#d32f2f", "operatore": "#1A5D1A", "lettore": "#1976d2"}.get(ruolo_utente, "#1A5D1A")
    
    st.markdown(f"""
    <div style="background:{colore_ruolo};color:white;padding:8px;border-radius:8px;text-align:center;">
    <b>{nome_utente}</b><br>
    <small>{username} - {ruolo_utente.upper()}</small>
    </div>
    """, unsafe_allow_html=True)
    
    # Aggiorna presenza ogni volta che naviga
    try:
        aggiorna_presenza(username, nome_utente, ruolo_utente)
    except:
        pass

    # MODIFICA 6 - LOGOUT RIPRISTINATO con rimozione presenza
    if st.button("🏠 TORNA AL HUB - 2 Progetti", use_container_width=True, key="back_hub_unico"):
        st.session_state.hub_page = "hub"
        st.session_state.menu = "Dashboard"
        st.rerun()
    st.divider()
    if st.button("Logout", type="primary", use_container_width=True, key="logout_btn"):
        try:
            rimuovi_presenza(st.session_state.get("username", ""))
        except:
            pass
        st.session_state.page = "entra"
        st.session_state.logged = False
        st.session_state.menu = "Dashboard"
        st.session_state.username = ""
        st.session_state.nome_utente = ""
        st.session_state.ruolo_utente = ""
        st.rerun()

    st.divider()
    st.markdown(
        """
        <div style="background:#e8f5e9;padding:8px;border-radius:8px;">
        <p style="font-size:12px;font-weight:bold;margin:0;">INFO RADIO HYTERA</p>
        <p style="font-size:11px;margin:4px 0 0 0;">
        PD785 - Anytone 878<br>
        Varese 950+ attivo<br>
        6 Modifiche implementate<br>
        Fusione Mappe: SI ottima idea
        </p>
        </div>
        """,
        unsafe_allow_html=True
    )

# DASHBOARD - SENZA INTESTAZIONE - richiesta Ezio: togli intestazione dalla dashboard, metti in login
if cur == "Dashboard":
    # hdr() rimosso da dashboard - messo in login - richiesta Ezio
    hdr_form("MENU'")

    # Etichetta "Seleziona un form" rimossa dalla dashboard - richiesta Ezio
    # st.markdown Seleziona un form rimosso

    form_buttons = [
        ("Splash", "🎬 Splash"),
        ("Volontari (con foto)", "👤 Volontari"),
        ("Ospiti", "🧑‍🤝‍🧑 Ospiti"),
        ("DB Radio", "📻 DB Radio"),
        ("Consegna Radio", "🤝 Consegna Radio"),
        ("Alias Radio", "🔖 Alias Radio"),
        ("Brogliaccio", "📓 Brogliaccio"),
        ("Eventi", "📅 Eventi"),
        ("Emergenze", "🚨 Emergenze"),
        ("Check-in", "✅ Check-in"),
        ("Interventi Emergenza", "🚒 Interventi Emergenza"),
        ("Tabella Interventi Emergenza", "📋 Tabella Interventi"),
        ("Mezzi", "🚐 Mezzi"),
        ("Attrezzature", "🧰 Attrezzature"),
        ("Mappe Postazioni", "🌍 Mappe Postazioni"),
        ("Libreria Icone", "🎨 Libreria Icone"),
        ("Turni", "🕐 Turni"),
        ("Chat", "💬 Chat"),
        ("Verbali", "📝 Verbali"),
        ("Archivio Documenti", "📁 Archivio Documenti"),
        ("Diplomi Attestati", "🏅 Diplomi"),
        ("Spese ODV per Evento", "💰 Spese ODV"),
        ("Note Spese", "🧾 Note Spese"),
        ("Report Filtro", "📊 Report Filtro"),
        ("Statistiche", "📈 Statistiche"),
        ("Backup", "💾 Backup")
    ]
    st.markdown("### Dashboard - Clicca per aprire (fix + Eventi ripristinato)")
    _cols = st.columns(3)
    for _idx, (_fkey, _flabel) in enumerate(form_buttons):
        if st.session_state.hub_page == "admin" and _fkey not in MENU_ADMINISTRATIVO:
            continue
        if st.session_state.hub_page == "operativo" and _fkey not in MENU_OPERATIVO:
            continue
        with _cols[_idx % 3]:
            if st.button(_flabel, key=f"dash_unico_{_fkey}_{_idx}", use_container_width=True):
                st.session_state.menu = _fkey
                if st.session_state.hub_page == "admin":
                    st.session_state.menu_admin = _fkey
                else:
                    st.session_state.menu_oper = _fkey
                st.rerun()
    st.divider()
    if True:
        st.markdown("Seleziona un form dalla sidebar o dalla dashboard sopra - Eventi presente - Geoloc rimossa")
        st.stop()
    # dashboard vecchia disabilitata
    if False:
        pass
    # FINE
    # TASTO ROSSO FULLSCREEN
    c_fs1, c_fs2 = st.columns([1,3])
    with c_fs1:
        if st.button("⛶ SCHERMO INTERO", key="btn_fullscreen_dash", use_container_width=True, type="primary"):
            st.session_state["fs_active"] = True
    with c_fs2:
        st.markdown('<span style="background:red;color:white;padding:6px 12px;border-radius:6px;font-weight:bold;">🔴 FULLSCREEN - ESC per uscire</span>', unsafe_allow_html=True)

    if st.session_state.get("fs_active"):
        st.components.v1.html(
            """
            <script>
            (function(){
                try {
                    const docEl = window.parent.document.documentElement;
                    if (docEl.requestFullscreen) docEl.requestFullscreen();
                } catch(e){}
                try {
                    if (document.documentElement.requestFullscreen) document.documentElement.requestFullscreen();
                } catch(e){}
            })();
            </script>
            <div style="background:#ff0000;color:white;padding:8px;border-radius:6px;text-align:center;font-weight:bold;">
            🔴 FULLSCREEN ATTIVO - premi ESC per uscire
            </div>
            """,
            height=70
        )
        if st.button("❌ Esci Fullscreen", key="btn_exit_fs", use_container_width=True):
            st.components.v1.html("<script>try{document.exitFullscreen(); parent.document.exitFullscreen();}catch(e){}</script>", height=0)
            st.session_state["fs_active"] = False
            st.rerun()

    st.write("")

    def vai_a_form_callback(form_name):
        st.session_state.menu = form_name
        st.session_state["cur"] = form_name
        st.session_state["cur"] = form_name

    # CSS bottoni verde ANA - SFONDO PIENO VERDE - FIX DEFINITIVO
    st.markdown(
        """
        <style>
        /* Reset e forza verde ANA su TUTTI i bottoni dashboard */
        [data-testid="column"] .stButton > button,
        [data-testid="stColumn"] .stButton > button,
        div[data-testid="stVerticalBlock"] .stButton > button {
            background-color: #1A5D1A !important;
            background-image: none !important;
            background: #1A5D1A !important;
            color: white !important;
            border: 2px solid #1A5D1A !important;
            font-weight: bold !important;
            font-family: 'Times New Roman', serif !important;
            border-radius: 8px !important;
        }
        [data-testid="column"] .stButton > button:hover {
            background-color: #2e7d32 !important;
            background: #2e7d32 !important;
            border-color: #2e7d32 !important;
            color: white !important;
        }
        [data-testid="column"] .stButton > button:active,
        [data-testid="column"] .stButton > button:focus,
        [data-testid="column"] .stButton > button:focus-visible {
            background-color: #1A5D1A !important;
            background: #1A5D1A !important;
            color: white !important;
            box-shadow: 0 0 0 2px rgba(26,93,26,0.3) !important;
            outline: none !important;
        }
        /* Forza anche su kind secondary */
        button[kind="secondary"] {
            background-color: #1A5D1A !important;
            background: #1A5D1A !important;
            color: white !important;
            border-color: #1A5D1A !important;
        }
        /* Solo i bottoni primary generici lasciali verdi, tranne fullscreen */
        button[kind="primary"]:not([data-testid*="fullscreen"]) {
            background-color: #1A5D1A !important;
            background: #1A5D1A !important;
            border-color: #1A5D1A !important;
        }
        /* Fullscreen dashboard rosso - specifico */
        button[key="btn_fullscreen_dash"], button[key="btn_exit_fs"] {
            background-color: #ff0000 !important;
            background: #ff0000 !important;
            border-color: #ff0000 !important;
        }
        /* Testo dentro bottone bianco */
        .stButton > button div, .stButton > button p, .stButton > button span {
            color: white !important;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    cols = st.columns(3)
    for i, (menu_name, btn_label) in enumerate(form_buttons):
        col = cols[i % 3]
        with col:
            if st.button(
                btn_label, 
                key=f"dash_btn_{i}_{menu_name}_FINAL2", 
                use_container_width=True, 
                help=f"Vai a {menu_name}"
            ):
                st.session_state.menu = menu_name
                st.rerun()

    st.divider()

    # Blocco informativo rimosso - pulizia etichette - richiesta Ezio

    # Statistiche semplici
    st.write("")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("Volontari", len(st.session_state.volontari))
    with c2:
        st.metric("Radio DB", len(st.session_state.radio_db))
    with c3:
        st.metric("Emergenze", len(st.session_state.emergenze))
    with c4:
        st.metric("Mappe Postazioni", len(st.session_state.mappe))

# VOLONTARI FORM CON SOTTOMASCHERE A LINGUETTE + CAMPO ODV - NUOVA VERSIONE FINALE
elif cur == "Volontari (con foto)":
    hdr_form("DB VOLONTARI")

    edit_mode = False
    edit_data = {}
    if st.session_state.vol_edit_index is not None:
        try:
            edit_data = st.session_state.volontari[st.session_state.vol_edit_index]
            edit_mode = True
        except:
            edit_data = {}
            edit_mode = False

    if edit_mode:
        st.warning(f"✏️ Modifica: {edit_data.get('Nome','')} {edit_data.get('Cognome','')} - Capo ODV: {edit_data.get('CapoODV','')} - I dati sono caricati nelle maschere sotto")
        # Forza caricamento dati nelle chiavi se non già presenti (per modifica da selectbox)
        try:
            if "vol_nome_tab" not in st.session_state:
                st.session_state["vol_nome_tab"] = edit_data.get("Nome","")
            if "vol_cognome_tab" not in st.session_state:
                st.session_state["vol_cognome_tab"] = edit_data.get("Cognome","")
            if "vol_capo_odv" not in st.session_state:
                st.session_state["vol_capo_odv"] = edit_data.get("CapoODV","")
            if "vol_cell_tab" not in st.session_state:
                st.session_state["vol_cell_tab"] = edit_data.get("Cellulare","")
        except:
            pass

    # SOTTOMASCHERE A LINGUETTE - 6 TAB - BOTTONI COLORATI DEFINITIVO - Chrome OK - Ezio
    if "vol_tab_active" not in st.session_state:
        st.session_state.vol_tab_active = 0
    st.markdown('<div class="linguette-box">', unsafe_allow_html=True)
    st.markdown("""
    <style>
    div[data-testid="stHorizontalBlock"] div[data-testid="column"]:nth-child(1) div[data-testid="stButton"] button {background:#1A5D1A !important; color:white !important; border:2px solid #FFD700 !important; border-radius:8px 8px 0 0 !important; font-weight:bold !important;}
    div[data-testid="stHorizontalBlock"] div[data-testid="column"]:nth-child(2) div[data-testid="stButton"] button {background:#1976d2 !important; color:white !important; border:2px solid #FFD700 !important; border-radius:8px 8px 0 0 !important; font-weight:bold !important;}
    div[data-testid="stHorizontalBlock"] div[data-testid="column"]:nth-child(3) div[data-testid="stButton"] button {background:#d32f2f !important; color:white !important; border:2px solid #FFD700 !important; border-radius:8px 8px 0 0 !important; font-weight:bold !important;}
    div[data-testid="stHorizontalBlock"] div[data-testid="column"]:nth-child(4) div[data-testid="stButton"] button {background:#ff9800 !important; color:black !important; border:2px solid #FFD700 !important; border-radius:8px 8px 0 0 !important; font-weight:bold !important;}
    div[data-testid="stHorizontalBlock"] div[data-testid="column"]:nth-child(5) div[data-testid="stButton"] button {background:#6a1b9a !important; color:white !important; border:2px solid #FFD700 !important; border-radius:8px 8px 0 0 !important; font-weight:bold !important;}
    div[data-testid="stHorizontalBlock"] div[data-testid="column"]:nth-child(6) div[data-testid="stButton"] button {background:#00796b !important; color:white !important; border:2px solid #FFD700 !important; border-radius:8px 8px 0 0 !important; font-weight:bold !important;}
    </style>
    """, unsafe_allow_html=True)
    c_tab1, c_tab2, c_tab3, c_tab4, c_tab5, c_tab6 = st.columns(6)
    with c_tab1:
        if st.button("📋 Anagrafica", key="tab_btn_0", use_container_width=True):
            st.session_state.vol_tab_active = 0
            st.rerun()
    with c_tab2:
        if st.button("📞 Contatti", key="tab_btn_1", use_container_width=True):
            st.session_state.vol_tab_active = 1
            st.rerun()
    with c_tab3:
        if st.button("🛡️ Ruolo", key="tab_btn_2", use_container_width=True):
            st.session_state.vol_tab_active = 2
            st.rerun()
    with c_tab4:
        if st.button("📻 Dotazione", key="tab_btn_3", use_container_width=True):
            st.session_state.vol_tab_active = 3
            st.rerun()
    with c_tab5:
        if st.button("📄 Documenti", key="tab_btn_4", use_container_width=True):
            st.session_state.vol_tab_active = 4
            st.rerun()
    with c_tab6:
        if st.button("📸 Foto", key="tab_btn_5", use_container_width=True):
            st.session_state.vol_tab_active = 5
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('<div style="background:white; border:3px solid #1A5D1A; border-radius:0 12px 12px 12px; padding:16px; margin-top:-8px;">', unsafe_allow_html=True)
    vol_active = st.session_state.vol_tab_active
    show_tab1 = (vol_active==0)
    show_tab2 = (vol_active==1)
    show_tab3 = (vol_active==2)
    show_tab4 = (vol_active==3)
    show_tab5 = (vol_active==4)

    # CSS colori per ogni bottone per Chrome
    st.markdown(f"""
    <style>
    button[kind="secondary"][data-testid="stBaseButton-secondary"]:has(div:contains("Anagrafica")), 
    div [data-testid="stButton"] button[key="tab_btn_0"] {{background:#1A5D1A !important; color:white !important;}}
    /* Iniezione diretta colori tramite key */
    </style>
    """, unsafe_allow_html=True)
    # Colori fissi con CSS mirato su key
    st.markdown("""
    <style>
    [data-testid="stColumn"]:nth-of-type(1) button {background:#1A5D1A !important; color:white !important;}
    [data-testid="stColumn"]:nth-of-type(2) button {background:#1976d2 !important; color:white !important;}
    [data-testid="stColumn"]:nth-of-type(3) button {background:#d32f2f !important; color:white !important;}
    [data-testid="stColumn"]:nth-of-type(4) button {background:#ff9800 !important; color:black !important;}
    [data-testid="stColumn"]:nth-of-type(5) button {background:#6a1b9a !important; color:white !important;}
    [data-testid="stColumn"]:nth-of-type(6) button {background:#00796b !important; color:white !important;}
    [data-testid="stColumn"]:nth-of-type(1) button[kind="primary"],
    [data-testid="stColumn"]:nth-of-type(2) button[kind="primary"],
    [data-testid="stColumn"]:nth-of-type(3) button[kind="primary"],
    [data-testid="stColumn"]:nth-of-type(4) button[kind="primary"],
    [data-testid="stColumn"]:nth-of-type(5) button[kind="primary"],
    [data-testid="stColumn"]:nth-of-type(6) button[kind="primary"] {border:4px solid #FFD700 !important; transform:scale(1.05);}
    </style>
    """, unsafe_allow_html=True)

    # contenitore per contenuto tab
    st.markdown('<div style="background:white; border:3px solid #1A5D1A; border-radius:0 12px 12px 12px; padding:16px; margin-top:-8px;">', unsafe_allow_html=True)
    active = st.session_state.vol_tab_active
    show_tab1 = (active==0)
    show_tab2 = (active==1)
    show_tab3 = (active==2)
    show_tab4 = (active==3)
    show_tab5 = (active==4)
    show_tab6 = (active==5)
    

    # Valori default da edit
    nome_def = edit_data.get("Nome", "")
    cognome_def = edit_data.get("Cognome", "")
    comune_def = edit_data.get("Comune", "Varese")
    via_def = edit_data.get("Via", "")
    capo_odv_def = edit_data.get("CapoODV", "")
    odv_app_def = edit_data.get("ODVAppartenenza", "ANA Varese")
    cell_def = edit_data.get("Cellulare", "")
    email_def = edit_data.get("Email", "")
    tel_em_def = edit_data.get("TelEmergenza", "")
    ruolo_def = edit_data.get("Ruolo", "Volontario")
    squadra_def = edit_data.get("Squadra", "Squadra A")
    radio_id_def = edit_data.get("RadioID", "")
    doc_def = edit_data.get("Documento", "")
    scad_def = edit_data.get("ScadDoc", "")

    if show_tab1:
        st.markdown("#### 📋 Anagrafica + Capo ODV")
        c1, c2 = st.columns(2)
        with c1:
            nome = st.text_input("Nome *", value=nome_def, key="vol_nome_tab")
            cognome = st.text_input("Cognome *", value=cognome_def, key="vol_cognome_tab")
            comune_res = combo_comune("Comune Residenza", "vol_comune_tab", comune_def)
            via = combo_vie("Via", comune_res, "vol_via_tab", via_def)
        with c2:
            capo_odv = st.text_input("Capo ODV *", value=capo_odv_def, key="vol_capo_odv", help="Nome del Capo ODV di riferimento")
            # CASELLA COMBO PER SCEGLIERE LE ODV - lista completa
            odv_lista = [
                "ANA Varese", "ANA Milano", "ANA Como", "ANA Bergamo", "ANA Brescia", "ANA Torino", "ANA Sezione Varese",
                "Protezione Civile Varese", "Protezione Civile Lombardia", "Protezione Civile Nazionale",
                "Croce Rossa Italiana - Varese", "Croce Rossa Italiana - Milano", "Misericordia", "ANPAS",
                "Associazione Nazionale Alpini", "Gruppo Comunale Volontari", "AIB - Antincendio Boschivo",
                "Altro"
            ]
            # Se valore esistente non in lista, aggiungilo
            if odv_app_def and odv_app_def not in odv_lista:
                odv_lista = [odv_app_def] + odv_lista
            
            odv_app = st.selectbox("ODV Associazione di Appartenenza *", odv_lista, index=odv_lista.index(odv_app_def) if odv_app_def in odv_lista else 0, key="vol_odv_app")
            if odv_app == "Altro":
                odv_app_custom = st.text_input("Specifica ODV - Inserisci nome", value="" if odv_app_def in odv_lista else odv_app_def, key="vol_odv_custom", placeholder="Es: Protezione Civile Busto Arsizio")
                if odv_app_custom:
                    odv_app = odv_app_custom
            data_nascita = st.date_input("Data Nascita", value=date(1970,1,1), min_value=date(1950,1,1), max_value=date.today(), format="DD/MM/YYYY", key="vol_data_nasc")
            codice_fisc = st.text_input("Codice Fiscale", value=edit_data.get("CodFisc",""), key="vol_cf")
            
            # FOTO NELLA PRIMA MASCHERA + DOWNLOAD
            st.markdown("**📸 Foto Volontario - Prima Maschera**")
            foto_file_prima = st.file_uploader("Carica foto (prima maschera)", type=["jpg", "jpeg", "png"], key="vol_foto_prima")
            if foto_file_prima:
                foto_bytes_prima = foto_file_prima.getvalue()
                st.image(foto_bytes_prima, width=120, caption="Preview prima maschera")
                # Salva in session per uso globale
                st.session_state["foto_temp_prima"] = foto_bytes_prima
                st.download_button("⬇️ Download Foto", data=foto_bytes_prima, file_name=f"foto_{nome}_{cognome}.jpg", mime="image/jpeg", use_container_width=True, key="download_foto_prima")
            elif edit_mode and edit_data.get("FotoBytes"):
                try:
                    st.image(edit_data.get("FotoBytes"), width=120, caption="Foto esistente")
                    st.download_button("⬇️ Download Foto Esistente", data=edit_data.get("FotoBytes"), file_name=f"foto_{edit_data.get('Cognome','')}_{edit_data.get('Nome','')}.jpg", mime="image/jpeg", use_container_width=True, key="download_foto_esistente_prima")
                except:
                    pass

    if show_tab2:
        st.markdown("#### 📞 Contatti")
        c1, c2 = st.columns(2)
        with c1:
            cellulare = st.text_input("Cellulare *", value=cell_def, key="vol_cell_tab")
            email = st.text_input("Email", value=email_def, key="vol_email_tab")
        with c2:
            tel_emerg = st.text_input("Telefono Emergenza", value=tel_em_def, key="vol_tel_em")
            note_cont = st.text_area("Note Contatti", value=edit_data.get("NoteContatti",""), key="vol_note_cont")

    if show_tab3:
        st.markdown("#### 🛡️ Ruolo e Squadra")
        c1, c2 = st.columns(2)
        with c1:
            ruolo = st.selectbox("Ruolo *", ["Volontario", "Capo Squadra", "Coordinatore", "Autista", "Radio Operatore", "Capo ODV", "Vice Capo ODV"], index=["Volontario", "Capo Squadra", "Coordinatore", "Autista", "Radio Operatore", "Capo ODV", "Vice Capo ODV"].index(ruolo_def) if ruolo_def in ["Volontario", "Capo Squadra", "Coordinatore", "Autista", "Radio Operatore", "Capo ODV", "Vice Capo ODV"] else 0, key="vol_ruolo_tab")
            squadra = st.selectbox("Squadra *", ["Squadra A", "Squadra B", "Squadra C", "Logistica", "Segreteria", "ODV Centrale"], index=["Squadra A", "Squadra B", "Squadra C", "Logistica", "Segreteria", "ODV Centrale"].index(squadra_def) if squadra_def in ["Squadra A", "Squadra B", "Squadra C", "Logistica", "Segreteria", "ODV Centrale"] else 0, key="vol_squadra_tab")
        with c2:
            data_iscriz = st.date_input("Data Iscrizione ODV", value=date.today(), format="DD/MM/YYYY", key="vol_data_iscr")
            stato_vol = st.selectbox("Stato", ["Attivo", "Inattivo", "In Formazione", "Sospeso"], key="vol_stato")

    if show_tab4:
        st.markdown("#### 📻 Dotazione Radio")
        radio_id = st.text_input("ID Radio / Matricola", value=radio_id_def, key="vol_radio_id")
        modello_radio = st.selectbox("Modello Radio", ["Hytera PD785", "Anytone 878", "Motorola", "Altro"], key="vol_radio_mod")
        note_dot = st.text_area("Note Dotazione", value=edit_data.get("NoteDotazione",""), key="vol_note_dot")

    if show_tab5:
        st.markdown("#### 📄 Documenti")
        doc_tipo = st.text_input("Tipo Documento", value=doc_def, key="vol_doc_tipo")
        doc_num = st.text_input("Numero Documento", value=edit_data.get("DocNum",""), key="vol_doc_num")
        doc_scad = st.date_input("Scadenza Documento", value=date.today(), format="DD/MM/YYYY", key="vol_doc_scad")
        st.caption("Formato data gg/mm/aaaa - es: 23/09/2026")

    if show_tab6:
        st.markdown("#### 📸 Foto Volontario")
        foto_file = st.file_uploader("Carica foto", type=["jpg", "jpeg", "png"], key="vol_foto_tab")
        foto_preview = st.session_state.get("foto_temp_prima", None)
        if foto_file:
            foto_bytes = foto_file.getvalue()
            st.image(foto_bytes, width=150, caption="Preview")
            foto_preview = foto_bytes
            st.download_button("⬇️ Download Foto Volontario", data=foto_bytes, file_name=f"foto_{nome}_{cognome}.jpg", mime="image/jpeg", use_container_width=True, key="download_foto_tab")
        elif edit_mode and edit_data.get("FotoBytes"):
            try:
                st.image(edit_data.get("FotoBytes"), width=150, caption="Foto esistente")
                foto_preview = edit_data.get("FotoBytes")
                st.download_button("⬇️ Download Foto", data=edit_data.get("FotoBytes"), file_name=f"foto_{edit_data.get('Cognome','')}.jpg", mime="image/jpeg", use_container_width=True, key="download_foto_tab_edit")
            except:
                pass
        elif foto_preview:
            st.image(foto_preview, width=150, caption="Foto da prima maschera")
            st.download_button("⬇️ Download Foto da Prima Maschera", data=foto_preview, file_name=f"foto_{nome}_{cognome}.jpg", mime="image/jpeg", use_container_width=True, key="download_foto_da_prima")

    st.markdown('</div>', unsafe_allow_html=True)
    st.divider()

    col_btn1, col_btn2, col_btn3 = st.columns([1,1,2])
    with col_btn3:
        # Pulisci maschera Volontari - FIX che funziona - Ezio
        if st.button("🧹 Pulisci maschera", use_container_width=True, key="btn_pulisci_volontari"):
            for k in [k for k in list(st.session_state.keys()) if k.startswith("vol_")]:
                try:
                    del st.session_state[k]
                except:
                    pass
            for k in ["foto_temp_prima", "foto_temp", "vol_edit_index"]:
                if k in st.session_state:
                    try:
                        del st.session_state[k]
                    except:
                        pass
            try:
                st.query_params.clear()
            except:
                pass
            st.rerun()
    if edit_mode:
        with col_btn1:
            if st.button("🔄 AGGIORNA VOLONTARIO", type="primary", use_container_width=True):
                if nome and cognome and cellulare and capo_odv:
                    updated = {
                        "Nome": nome,
                        "Cognome": cognome,
                        "Comune": comune_res,
                        "Via": via,
                        "CapoODV": capo_odv,
                        "ODVAppartenenza": odv_app,
                        "DataNascita": str(data_nascita),
                        "CodFisc": codice_fisc,
                        "Cellulare": cellulare,
                        "Email": email,
                        "TelEmergenza": tel_emerg,
                        "Ruolo": ruolo,
                        "Squadra": squadra,
                        "RadioID": radio_id,
                        "ModelloRadio": modello_radio,
                        "Documento": doc_tipo,
                        "DocNum": doc_num,
                        "ScadDoc": str(doc_scad),
                        "FotoBytes": foto_preview,
                        "DataAgg": datetime.now().strftime("%d/%m/%Y %H:%M")
                    }
                    st.session_state.volontari[st.session_state.vol_edit_index] = updated
                    st.session_state.vol_edit_index = None
                    st.success("Volontario aggiornato con ODV OK")
                    st.rerun()
                else:
                    st.error("Compila Nome, Cognome, Cellulare e Capo ODV *")
        with col_btn2:
            if st.button("❌ ANNULLA", use_container_width=True):
                st.session_state.vol_edit_index = None
                st.rerun()
    else:
        with col_btn1:
            if st.button("💾 SALVA VOLONTARIO", type="primary", use_container_width=True):
                if nome and cognome and cellulare and capo_odv:
                    nuovo = {
                        "Nome": nome,
                        "Cognome": cognome,
                        "Comune": comune_res,
                        "Via": via,
                        "CapoODV": capo_odv,
                        "ODVAppartenenza": odv_app,
                        "DataNascita": str(data_nascita),
                        "CodFisc": codice_fisc,
                        "Cellulare": cellulare,
                        "Email": email,
                        "TelEmergenza": tel_emerg,
                        "Ruolo": ruolo,
                        "Squadra": squadra,
                        "RadioID": radio_id,
                        "ModelloRadio": modello_radio,
                        "Documento": doc_tipo,
                        "DocNum": doc_num,
                        "ScadDoc": str(doc_scad),
                        "FotoBytes": foto_preview,
                        "DataIns": datetime.now().strftime("%d/%m/%Y %H:%M")
                    }
                    st.session_state.volontari.append(nuovo)
                    # PULISCI CAMPI PER NUOVO INSERIMENTO - Richiesta Ezio - TUTTI I CAMPI
                    chiavi_da_pulire = [k for k in list(st.session_state.keys()) if k.startswith("vol_")]
                    for k in chiavi_da_pulire:
                        try:
                            del st.session_state[k]
                        except:
                            pass
                    # Pulisci anche foto temp
                    for k in ["foto_temp_prima", "foto_temp"]:
                        if k in st.session_state:
                            try:
                                del st.session_state[k]
                            except:
                                pass
                    st.success(f"✅ Volontario {cognome} {nome} - Capo ODV {capo_odv} salvato! Maschera pulita per nuovo inserimento.")
                    st.balloons()
                    st.rerun()
                else:
                    st.error("Compila campi obbligatori * (Nome, Cognome, Cellulare, Capo ODV)")

    # Elenco volontari - COMBO invece di lista lunga 200 - richiesta Ezio + tasto mostra tabella
    st.divider()
    if st.session_state.volontari:
        st.markdown(f"#### Elenco Volontari ({len(st.session_state.volontari)}) - Usa combo per modifica")
        
        # COMBO per selezionare volontario - non mostra tutti i 200 sotto
        vol_options = []
        for i, v in enumerate(st.session_state.volontari):
            label = f"{i+1:03d}: {v.get('Cognome','')} {v.get('Nome','')} - {v.get('ODVAppartenenza','')} - Capo {v.get('CapoODV','')} - {v.get('Comune','')}"
            vol_options.append((i, label))
        
        # Selectbox combo
        combo_labels = ["-- Seleziona volontario per modifica --"] + [lbl for _, lbl in vol_options]
        sel_combo = st.selectbox("🔍 Cerca e seleziona volontario (combo) - per 200 volontari", combo_labels, key="vol_combo_modifica")
        
        if sel_combo != "-- Seleziona volontario per modifica --":
            try:
                # Trova indice
                idx_sel = None
                for idx, lbl in vol_options:
                    if lbl == sel_combo:
                        idx_sel = idx
                        break
                if idx_sel is not None:
                    c_mod1, c_mod2 = st.columns([1,1])
                    with c_mod1:
                        if st.button("✏️ Carica in maschera per modifica", type="primary", use_container_width=True, key=f"btn_load_vol_{idx_sel}"):
                            try:
                                vd = st.session_state.volontari[idx_sel]
                                st.session_state.vol_edit_index = idx_sel
                                for k in [k for k in list(st.session_state.keys()) if k.startswith("vol_")]:
                                    try:
                                        del st.session_state[k]
                                    except:
                                        pass
                                st.session_state["vol_nome_tab"] = vd.get("Nome","")
                                st.session_state["vol_cognome_tab"] = vd.get("Cognome","")
                                st.session_state["vol_comune_tab"] = vd.get("Comune","Varese")
                                st.session_state["vol_via_tab"] = vd.get("Via","")
                                st.session_state["vol_capo_odv"] = vd.get("CapoODV","")
                                st.session_state["vol_odv_app"] = vd.get("ODVAppartenenza","ANA Varese")
                                st.session_state["vol_cell_tab"] = vd.get("Cellulare","")
                                st.session_state["vol_email_tab"] = vd.get("Email","")
                                st.session_state["vol_tel_em"] = vd.get("TelEmergenza","")
                                st.session_state["vol_ruolo_tab"] = vd.get("Ruolo","Volontario")
                                st.session_state["vol_squadra_tab"] = vd.get("Squadra","Squadra A")
                                st.session_state["vol_radio_id"] = vd.get("RadioID","")
                                st.session_state["vol_cf"] = vd.get("CodFisc","")
                                st.session_state["vol_doc_tipo"] = vd.get("Documento","")
                                st.session_state["vol_doc_num"] = vd.get("DocNum","")
                                if vd.get("FotoBytes"):
                                    st.session_state["foto_temp_prima"] = vd.get("FotoBytes")
                                st.success(f"Caricato {vd.get('Cognome','')} {vd.get('Nome','')} in maschera")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Errore caricamento: {e}")
                    with c_mod2:
                        if st.button("🗑️ Elimina volontario", use_container_width=True, key=f"btn_del_vol_{idx_sel}"):
                            try:
                                st.session_state.volontari.pop(idx_sel)
                                st.session_state.vol_edit_index = None
                                st.success("Volontario eliminato")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Errore eliminazione: {e}")
            except Exception as e:
                st.error(f"Errore combo: {e}")
        
        st.divider()
        # Tasto per vedere tabella volontari - richiesta Ezio: non mostrare automaticamente sotto
        if st.button(f"📋 Mostra/Nascondi Tabella Volontari ({len(st.session_state.volontari)} record)", key="btn_toggle_tabella_vol"):
            st.session_state["show_vol_table"] = not st.session_state.get("show_vol_table", False)
        
        if st.session_state.get("show_vol_table", False):
            df_vol = pd.DataFrame([{k:v for k,v in vol.items() if "Bytes" not in k} for vol in st.session_state.volontari])
            cols_show = ["Cognome", "Nome", "CapoODV", "ODVAppartenenza", "Comune", "Cellulare", "Ruolo", "Squadra"]
            cols_show = [c for c in cols_show if c in df_vol.columns]
            st.dataframe(df_vol[cols_show] if cols_show else df_vol, use_container_width=True)
            st.caption("Tabella nascosta dietro bottone per non occupare spazio con 200 volontari")

        # Export Excel volontari - FIX sempre visibile - DEBUG OPENPYXL
        try:
            df_export = pd.DataFrame([{k:v for k,v in vol.items() if "Bytes" not in k} for vol in st.session_state.volontari])
            if not df_export.empty:
                # Debug
                st.caption(f"Debug: OPENPYXL_OK={OPENPYXL_OK} XLSXWRITER_OK={XLSXWRITER_OK} - {len(df_export)} record - {len(df_export.columns)} colonne")
                try:
                    import openpyxl
                    st.caption(f"openpyxl import OK v{openpyxl.__version__}")
                except Exception as e:
                    st.error(f"openpyxl import FAIL: {e} - Verifica requirements.txt su GitHub!")
                
                excel_bytes = to_excel(df_export)
                if excel_bytes and len(excel_bytes) > 100 and excel_bytes[:2] == b'PK':
                    st.download_button(
                        "⬇️ EXPORT EXCEL VOLONTARI - FIX VALIDO",
                        data=excel_bytes,
                        file_name=f"volontari_export_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="export_volontari_fix",
                        type="primary"
                    )
                    st.success(f"Excel pronto: {len(excel_bytes)} bytes - PK valido")
                else:
                    st.error(f"Excel non valido: len={len(excel_bytes) if excel_bytes else 0} - header={excel_bytes[:10] if excel_bytes else b''}")
                    st.warning("Su Streamlit Cloud: Vai su Manage app -> Clear cache -> Reboot - Verifica requirements.txt contenga openpyxl==3.1.5")
                    # Prova export diretto emergenza
                    try:
                        import openpyxl
                        wb = openpyxl.Workbook()
                        ws = wb.active
                        ws.title="Volontari"
                        for c_idx, col in enumerate(df_export.columns, 1):
                            ws.cell(row=1, column=c_idx, value=str(col))
                        for r_idx, row in enumerate(df_export.itertuples(index=False), 2):
                            for c_idx, val in enumerate(row, 1):
                                try:
                                    ws.cell(row=r_idx, column=c_idx, value=str(val)[:32000])
                                except:
                                    ws.cell(row=r_idx, column=c_idx, value="")
                        from io import BytesIO
                        buf = io.BytesIO()
                        wb.save(buf)
                        buf.seek(0)
                        st.download_button(
                            "⬇️ EXPORT EMERGENZA - Volontari",
                            data=buf.getvalue(),
                            file_name=f"volontari_EMERGENZA_{datetime.now().strftime('%Y%m%d')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                            key="export_vol_emerg"
                        )
                    except Exception as e2:
                        st.error(f"Export emergenza fallito: {e2}")
        except Exception as e:
            st.error(f"Errore export volontari: {e}")
            import traceback
            st.code(traceback.format_exc())

        # Selectbox fallback per modifica
        cognomi = [f"{i}: {v.get('Cognome','')} {v.get('Nome','')} - ODV {v.get('ODVAppartenenza','')} - Capo {v.get('CapoODV','')}" for i, v in enumerate(st.session_state.volontari)]
        sel = st.selectbox("Oppure seleziona volontario per modifica (metodo vecchio)", ["--"] + cognomi, key="sel_vol_mod")
        if sel != "--":
            try:
                idx = int(sel.split(":")[0])
                st.session_state.vol_edit_index = idx
                # Stesso fix caricamento
                try:
                    vd = st.session_state.volontari[idx]
                    for k in list(st.session_state.keys()):
                        if k.startswith("vol_"):
                            try:
                                del st.session_state[k]
                            except:
                                pass
                    st.session_state["vol_nome_tab"] = vd.get("Nome","")
                    st.session_state["vol_cognome_tab"] = vd.get("Cognome","")
                    st.session_state["vol_capo_odv"] = vd.get("CapoODV","")
                    st.session_state["vol_cell_tab"] = vd.get("Cellulare","")
                except:
                    pass
                st.rerun()
            except:
                pass
    else:
        pass  # istruzione rimossa
        pass
        # Istruzione rimossa - form pulito

    st.divider()

# DB RADIO
    # IMPORT/EXPORT
    excel_import_inline("volontari", "Volontari (con foto)")



elif cur == "Ospiti":
    hdr_form("DB OSPITI - Visitatori e Ospiti Esterni - Font Nero Times Bold - Verde ANA")
    edit_mode_osp = False
    edit_data_osp = {}
    if st.session_state.get("osp_edit_index") is not None:
        try:
            edit_data_osp = st.session_state.ospiti[st.session_state.osp_edit_index]
            edit_mode_osp = True
        except:
            edit_data_osp = {}
            edit_mode_osp = False
    if edit_mode_osp:
        st.warning(f"✏️ Modifica Ospite: {edit_data_osp.get('Nome','')} {edit_data_osp.get('Cognome','')}")
    # OSPITI - BOTTONI COLORATI IDENTICI A VOLONTARI - DEFINITIVO
    if "osp_tab_active" not in st.session_state:
        st.session_state.osp_tab_active = 0
    st.markdown('<div class="linguette-box">', unsafe_allow_html=True)
    co1, co2, co3, co4, co5, co6 = st.columns(6)
    with co1:
        if st.button("📋 Anagrafica", key="osp_tab_btn_0", use_container_width=True):
            st.session_state.osp_tab_active = 0
            st.rerun()
    with co2:
        if st.button("📞 Contatti", key="osp_tab_btn_1", use_container_width=True):
            st.session_state.osp_tab_active = 1
            st.rerun()
    with co3:
        if st.button("🛡️ Ruolo", key="osp_tab_btn_2", use_container_width=True):
            st.session_state.osp_tab_active = 2
            st.rerun()
    with co4:
        if st.button("📻 Dotazione", key="osp_tab_btn_3", use_container_width=True):
            st.session_state.osp_tab_active = 3
            st.rerun()
    with co5:
        if st.button("📄 Documenti", key="osp_tab_btn_4", use_container_width=True):
            st.session_state.osp_tab_active = 4
            st.rerun()
    with co6:
        if st.button("📸 Foto", key="osp_tab_btn_5", use_container_width=True):
            st.session_state.osp_tab_active = 5
            st.rerun()
    st.markdown('<div style="background:white; border:3px solid #1A5D1A; border-radius:0 12px 12px 12px; padding:16px; margin-top:-8px;">', unsafe_allow_html=True)
    osp_active = st.session_state.osp_tab_active
    show_osp_tab1 = (osp_active==0)
    show_osp_tab2 = (osp_active==1)
    show_osp_tab3 = (osp_active==2)
    show_osp_tab4 = (osp_active==3)
    show_osp_tab5 = (osp_active==4)
    show_osp_tab6 = (osp_active==5)
    nome_def = edit_data_osp.get("Nome", "")
    cognome_def = edit_data_osp.get("Cognome", "")
    comune_def = edit_data_osp.get("Comune", "Varese")
    via_def = edit_data_osp.get("Via", "")
    capo_odv_def = edit_data_osp.get("CapoODV", "")
    odv_app_def = edit_data_osp.get("ODVAppartenenza", "Ospite - Esterno")
    cell_def = edit_data_osp.get("Cellulare", "")
    email_def = edit_data_osp.get("Email", "")
    tel_em_def = edit_data_osp.get("TelEmergenza", "")
    if show_osp_tab1:
        c1, c2 = st.columns(2)
        with c1:
            nome = st.text_input("Nome *", value=nome_def, key="osp_nome_tab")
            cognome = st.text_input("Cognome *", value=cognome_def, key="osp_cognome_tab")
            comune_res = combo_comune("Comune Residenza", "osp_comune_tab", comune_def)
            via = combo_vie("Via", comune_res, "osp_via_tab", via_def)
        with c2:
            capo_odv = st.text_input("Referente / Capo ODV *", value=capo_odv_def, key="osp_capo_odv")
            odv_lista = ["Ospite - Esterno", "Visitatore", "Stampa", "Istituzioni", "ANA Varese", "Croce Rossa", "Altro"]
            if odv_app_def and odv_app_def not in odv_lista:
                odv_lista = [odv_app_def] + odv_lista
            odv_app = st.selectbox("Ente / Associazione *", odv_lista, index=odv_lista.index(odv_app_def) if odv_app_def in odv_lista else 0, key="osp_odv_app")
            data_nascita = st.date_input("Data Nascita", value=date(1970,1,1), min_value=date(1950,1,1), max_value=date.today(), format="DD/MM/YYYY", key="osp_data_nasc")
            codice_fisc = st.text_input("Codice Fiscale", value=edit_data_osp.get("CodFisc",""), key="osp_cf")
            foto_file_prima = st.file_uploader("Carica foto ospite", type=["jpg", "jpeg", "png"], key="osp_foto_prima")
            if foto_file_prima:
                st.session_state["osp_foto_temp_prima"] = foto_file_prima.getvalue()
                st.image(st.session_state["osp_foto_temp_prima"], width=120)
    if show_osp_tab2:
        c1, c2 = st.columns(2)
        with c1:
            cellulare = st.text_input("Cellulare *", value=cell_def, key="osp_cell_tab")
            email = st.text_input("Email", value=email_def, key="osp_email_tab")
        with c2:
            tel_emerg = st.text_input("Telefono Emergenza", value=tel_em_def, key="osp_tel_em")
            note_cont = st.text_area("Note Contatti", value=edit_data_osp.get("NoteContatti",""), key="osp_note_cont")
    if show_osp_tab3:
        c1, c2 = st.columns(2)
        with c1:
            ruolo = st.selectbox("Ruolo *", ["Ospite", "Visitatore", "Stampa", "Istituzioni", "Osservatore"], key="osp_ruolo_tab")
            squadra = st.selectbox("Gruppo *", ["Ospiti", "Visitatori", "Stampa"], key="osp_squadra_tab")
        with c2:
            data_iscriz = st.date_input("Data Visita", value=date.today(), format="DD/MM/YYYY", key="osp_data_iscr")
            stato_vol = st.selectbox("Stato", ["Attivo", "In Visita", "Uscito"], key="osp_stato")
    if show_osp_tab4:
        radio_id = st.text_input("ID Badge / Radio Ospite", value=edit_data_osp.get("RadioID",""), key="osp_radio_id")
        modello_radio = st.selectbox("Tipo Badge", ["Badge Ospite", "Badge Visitatore", "Nessuno"], key="osp_radio_mod")
        note_dot = st.text_area("Note Dotazione", value=edit_data_osp.get("NoteDotazione",""), key="osp_note_dot")
    if show_osp_tab5:
        doc_tipo = st.text_input("Tipo Documento", value=edit_data_osp.get("Documento",""), key="osp_doc_tipo")
        doc_num = st.text_input("Numero Documento", value=edit_data_osp.get("DocNum",""), key="osp_doc_num")
        doc_scad = st.date_input("Scadenza Documento", value=date.today(), format="DD/MM/YYYY", key="osp_doc_scad")
    if show_osp_tab6:
        foto_file = st.file_uploader("Carica foto", type=["jpg", "jpeg", "png"], key="osp_foto_tab")
        foto_preview = st.session_state.get("osp_foto_temp_prima", None)
        if foto_file:
            foto_preview = foto_file.getvalue()
            st.image(foto_preview, width=150)
        elif edit_mode_osp and edit_data_osp.get("FotoBytes"):
            foto_preview = edit_data_osp.get("FotoBytes")
            st.image(foto_preview, width=150)
    st.markdown('</div>', unsafe_allow_html=True)
    st.divider()
    if st.button("💾 Salva Ospite", type="primary", use_container_width=True, key="btn_save_osp"):
        if not nome or not cognome:
            st.error("Nome e Cognome obbligatori")
        else:
            nuovo_osp = {"Nome": nome, "Cognome": cognome, "Comune": comune_res, "Via": via, "CapoODV": capo_odv, "ODVAppartenenza": odv_app, "DataNascita": str(data_nascita), "CodFisc": codice_fisc, "Cellulare": cellulare, "Email": email, "TelEmergenza": tel_emerg, "NoteContatti": note_cont, "Ruolo": ruolo, "Squadra": squadra, "DataIscrizione": str(data_iscriz), "Stato": stato_vol, "RadioID": radio_id, "ModelloRadio": modello_radio, "NoteDotazione": note_dot, "Documento": doc_tipo, "DocNum": doc_num, "ScadDoc": str(doc_scad), "FotoBytes": foto_preview, "Timestamp": datetime.now().strftime("%d/%m/%Y %H:%M")}
            if edit_mode_osp:
                st.session_state.ospiti[st.session_state.osp_edit_index] = nuovo_osp
                st.session_state.osp_edit_index = None
                st.success("✅ Ospite modificato!")
            else:
                st.session_state.ospiti.append(nuovo_osp)
                st.success("✅ Ospite aggiunto!")
            st.rerun()
    st.divider()
    st.markdown(f"#### 📋 Elenco Ospiti ({len(st.session_state.get('ospiti', []))})")
    if st.session_state.get("ospiti"):
        df_osp = pd.DataFrame([{k:v for k,v in o.items() if "Bytes" not in k} for o in st.session_state.ospiti])
        st.dataframe(df_osp, use_container_width=True)
    excel_import_inline("ospiti", "Ospiti")



elif cur == "DB Radio":
    hdr_form("DB RADIO - Gestione Apparati")

    c1, c2, c3 = st.columns(3)
    with c1:
        modello = st.selectbox("Modello Radio", ["Hytera PD785", "Anytone 878", "Motorola", "Altro"], key="radio_modello_db_2026")
        matricola = st.text_input("Matricola / ID", key="radio_mat_db_2026")
        # CAMPO BANDA - Richiesta Ezio - VHF,UHF,VHF/UHF,HF
        banda = st.selectbox("Banda *", ["VHF", "UHF", "VHF/UHF", "HF"], key="radio_banda_db_2026")
        # CAMPO TIPO - Ripristinato - Richiesta Ezio - DMR,PMR446 ecc
        tipo_radio = st.selectbox("Tipo *", ["DMR", "TETRA", "PMR446", "NAUTICHE", "VARIE"], key="radio_tipo_db_2026")

    with c2:
        alias_r = st.text_input("Alias Radio", key="radio_alias_db_2026")
        stato_r = st.selectbox("Stato Radio", ["Operativa", "In Manutenzione", "Fuori Servizio", "Assegnata"], key="radio_stato_db_2026")
        note_r = st.text_area("Note", key="radio_note_db_2026")

    with c3:
        st.write("Foto Radio")
        foto_r = st.file_uploader("Foto", type=["jpg", "png"], key="radio_foto_db_2026")
        if foto_r:
            st.image(foto_r.getvalue(), width=100)

    if st.button("Salva Radio in DB", type="primary", use_container_width=True, key="btn_salva_radio_db_2026"):
        if matricola:
            st.session_state.radio_db.append({
                "Modello": modello,
                "Matricola": matricola,
                "Banda": banda,
                "Tipo": tipo_radio,
                "Alias": alias_r,
                "Stato": stato_r,
                "Note": note_r,
                "Data": datetime.now().strftime("%d/%m/%Y")
            })
            st.success("Radio salvata")
            st.rerun()

    if st.session_state.radio_db:
        df_r = pd.DataFrame(st.session_state.radio_db)
        st.dataframe(df_r, use_container_width=True)
        st.download_button("Excel Radio", to_excel(df_r), "radio_db.xlsx", use_container_width=True, key="dl_excel_radio_db_2026")
        if REPORTLAB_OK:
            st.download_button("PDF Logo Estesa", to_pdf(df_r, "DB RADIO"), "radio_db.pdf", use_container_width=True, key="dl_pdf_radio_db_2026")

    # IMPORT/EXPORT INLINE - Ezio - TUTTI I FORM - Excel + PDF + Template ODV
    excel_import_inline("radio_db", "DB Radio")

# CONSEGNA RADIO


elif cur == "Consegna Radio":
    # hdr() rimosso - richiesta Ezio: su form non mettere in alto intestazione con i loghi
    hdr_form("CONSEGNA RADIO - Consegna e Riconsegna con note problemi + Scanner CF")

    # CAMPI BARCODE - CF + MATRICOLA RADIO - Richiesta Ezio - Matricola associata a DB Radio
    # Etichetta DOPPIO SCANNER rimossa - pulizia Ezio
    
    c_scan1, c_scan2 = st.columns(2)
    with c_scan1:
        cf_radio = st.text_input("🔍 CF Volontario - SCANNER TESSERA SANITARIA", key="cons_cf", placeholder="Spara su tessera sanitaria...", help="Pistola barcode - legge CF e trova volontario")
    with c_scan2:
        barcode_radio = st.text_input("📻 Matricola Radio - SCANNER BARCODE RADIO (DB Radio)", key="cons_barcode_radio", placeholder="Spara su barcode matricola radio...", help="Associato a DB Radio - legge matricola e trova radio in DB")
    
    vol_radio_trovato = None
    radio_trovata_barcode = None
    
    if cf_radio:
        cf_clean_radio = cf_radio.strip().upper()
        for v in st.session_state.volontari:
            cf_v = (v.get('CodiceFiscale','') or v.get('CF','') or '').strip().upper()
            if cf_v and cf_v == cf_clean_radio:
                vol_radio_trovato = f"{v.get('Cognome','').strip()} {v.get('Nome','').strip()}"
                st.success(f"✅ Volontario trovato da CF per radio: {vol_radio_trovato}")
                break
        if not vol_radio_trovato:
            if len(cf_clean_radio) >= 16:
                for v in st.session_state.volontari:
                    cf_v = (v.get('CodiceFiscale','') or v.get('CF','') or '').strip().upper()
                    if cf_v and cf_v in cf_clean_radio:
                        vol_radio_trovato = f"{v.get('Cognome','').strip()} {v.get('Nome','').strip()}"
                        st.success(f"✅ Volontario trovato (estratto CF): {vol_radio_trovato}")
                        break
            if not vol_radio_trovato:
                st.warning(f"⚠️ CF {cf_clean_radio} non trovato - inserisci manuale")
    
    if barcode_radio:
        mat_clean = barcode_radio.strip().upper()
        for r in st.session_state.radio_db:
            mat_r = (r.get('Matricola','') or r.get('Seriale','') or '').strip().upper()
            if mat_r and (mat_r == mat_clean or mat_clean in mat_r or mat_r in mat_clean):
                radio_trovata_barcode = f"{r.get('Matricola','')} - {r.get('Modello','')} - {r.get('Alias','')}"
                st.success(f"✅ Radio trovata da matricola barcode: {radio_trovata_barcode} (DB Radio)")
                break
        if not radio_trovata_barcode:
            # Prova match parziale
            for r in st.session_state.radio_db:
                mat_r = (r.get('Matricola','') or '').strip().upper()
                if mat_r and mat_clean and (mat_clean.startswith(mat_r) or mat_r.startswith(mat_clean)):
                    radio_trovata_barcode = f"{r.get('Matricola','')} - {r.get('Modello','')} - {r.get('Alias','')}"
                    st.success(f"✅ Radio trovata (parziale): {radio_trovata_barcode}")
                    break
        if not radio_trovata_barcode:
            st.warning(f"⚠️ Matricola {mat_clean} non trovata in DB Radio - inserisci manuale")

    # Maschera verde come altri form
    st.markdown('<div style="background:#C8E6C9;padding:15px;border-radius:10px;border:2px solid #1A5D1A;margin-bottom:15px;">', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### 📻 Consegna")
        vol_list = [f"{v.get('Cognome','').strip()} {v.get('Nome','').strip()}" for v in st.session_state.volontari if v.get('Cognome') or v.get('Nome')]
        vol_list = sorted(list(set([x for x in vol_list if x.strip()])))
        if vol_list:
            if vol_radio_trovato and vol_radio_trovato in vol_list:
                idx_vol_radio = vol_list.index(vol_radio_trovato)
                sel_vol = st.selectbox("Volontario *", vol_list, index=idx_vol_radio, key="cons_vol")
            else:
                sel_vol = st.selectbox("Volontario *", vol_list, key="cons_vol")
        else:
            sel_vol = st.text_input("Volontario * (manuale)", value=vol_radio_trovato if vol_radio_trovato else "", key="cons_vol_man", placeholder="Cognome Nome")

        radio_list = [f"{r.get('Matricola','')} - {r.get('Modello','')} - {r.get('Alias','')}" for r in st.session_state.radio_db]
        if radio_list:
            if radio_trovata_barcode and radio_trovata_barcode in radio_list:
                idx_radio = radio_list.index(radio_trovata_barcode)
                sel_radio = st.selectbox("Radio * (da DB Radio - associata a barcode matricola)", radio_list, index=idx_radio, key="cons_radio")
            else:
                sel_radio = st.selectbox("Radio * (da DB Radio)", radio_list, key="cons_radio")
        else:
            sel_radio = st.text_input("Radio * manuale", value=radio_trovata_barcode if radio_trovata_barcode else "", key="cons_radio_man", placeholder="Matricola - Modello")

        data_cons = st.date_input("Data Consegna *", value=date.today(), format="DD/MM/YYYY", key="cons_data")
        ora_cons = st.time_input("Ora Consegna", value=datetime.now().time(), key="cons_ora")
        motivo = st.text_input("Motivo / Evento", key="cons_motivo", placeholder="Esercitazione, Emergenza...")

    with c2:
        st.markdown("#### 🔄 Riconsegna")
        data_ricons = st.date_input("Data Riconsegna", value=date.today(), format="DD/MM/YYYY", key="ricons_data")
        ora_ricons = st.time_input("Ora Riconsegna", value=datetime.now().time(), key="ricons_ora")
        stato_ricons = st.selectbox("Stato alla Riconsegna", ["Da riconsegnare", "Riconsegnata - OK", "Riconsegnata - Guasta", "Riconsegnata - Batteria scarica", "Riconsegnata - Antenna rotta", "Riconsegnata - Problemi audio", "Persa", "In Manutenzione"], key="ricons_stato")
        # Campo note problemi radio - richiesta Ezio
        note_problemi = st.text_area("Note - Problemi Radio *", key="cons_note_problemi", placeholder="Descrivi se radio ha avuto problemi: es. batteria scarica dopo 2h, audio gracchiante, tasto PTT bloccato, antenna piegata, display spento...", height=120)

    st.markdown("</div>", unsafe_allow_html=True)

    if st.button("Registra Consegna + Riconsegna", type="primary", use_container_width=True, key="btn_cons_reg"):
        if sel_vol and sel_radio:
            cf_to_save_radio = cf_radio.strip().upper() if cf_radio else ""
            barcode_to_save = barcode_radio.strip().upper() if barcode_radio else ""
            if vol_radio_trovato:
                sel_vol_final = vol_radio_trovato
            else:
                sel_vol_final = sel_vol
            if radio_trovata_barcode:
                sel_radio_final = radio_trovata_barcode
            else:
                sel_radio_final = sel_radio
            st.session_state.consegna_radio.append({
                "Volontario": sel_vol_final,
                "CodiceFiscale": cf_to_save_radio,
                "Radio": sel_radio_final,
                "MatricolaBarcode": barcode_to_save,
                "Matricola": barcode_to_save or sel_radio_final.split(" - ")[0],
                "DataConsegna": str(data_cons),
                "OraConsegna": str(ora_cons),
                "Motivo": motivo,
                "DataRiconsegna": str(data_ricons),
                "OraRiconsegna": str(ora_ricons),
                "StatoRiconsegna": stato_ricons,
                "NoteProblemi": note_problemi,
                "Stato": stato_ricons,
                "Data": str(data_cons),
                "Ora": str(ora_cons)
            })
            st.success(f"Consegna registrata: {sel_vol_final} - CF: {cf_to_save_radio} - Radio: {sel_radio_final} - Matricola: {barcode_to_save} - {stato_ricons}")
            if "cons_cf" in st.session_state:
                del st.session_state["cons_cf"]
            if "cons_barcode_radio" in st.session_state:
                del st.session_state["cons_barcode_radio"]
            st.rerun()
        else:
            st.error("Seleziona Volontario e Radio *")

    if st.session_state.consegna_radio:
        st.divider()
        st.markdown(f"#### 📋 Elenco Consegne ({len(st.session_state.consegna_radio)})")
        df_cr = pd.DataFrame(st.session_state.consegna_radio)
        # Mostra colonne importanti
        cols_show = [c for c in ["Volontario","Radio","DataConsegna","OraConsegna","DataRiconsegna","StatoRiconsegna","NoteProblemi","Motivo"] if c in df_cr.columns]
        if cols_show:
            st.dataframe(df_cr[cols_show], use_container_width=True)
        else:
            st.dataframe(df_cr, use_container_width=True)
        
        c_exp1, c_exp2 = st.columns(2)
        with c_exp1:
            st.download_button("Excel Consegne", to_excel(df_cr), "consegne.xlsx", use_container_width=True, key="dl_excel_cons")
        with c_exp2:
            if REPORTLAB_OK:
                st.download_button("PDF Consegne", to_pdf(df_cr, "CONSEGNA RADIO - RICONSEGNA + NOTE PROBLEMI"), "consegne.pdf", use_container_width=True, key="dl_pdf_cons")

# ALIAS RADIO


    # IMPORT/EXPORT INLINE - Ezio - TUTTI I FORM - Excel + PDF + Template ODV
    excel_import_inline("consegna_radio", "Consegna Radio")


elif cur == "Alias Radio":
    hdr_form("ALIAS RADIO - Alias + Volontario Cognome Nome + Canale Radio")

    # Maschera verde come altri form - spostata in alto - RIGA 2867
    st.markdown('<div style="background:#C8E6C9;padding:12px;border-radius:10px;border:2px solid #1A5D1A;margin-bottom:12px;margin-top:0px;">', unsafe_allow_html=True)

    # Canali radio di lavoro - lista combo
    canali_radio_list = ["Canale 1 - Emergenza", "Canale 2 - Coordinamento", "Canale 3 - Logistica", "Canale 4 - Squadra A", "Canale 5 - Squadra B", "Canale 6 - Squadra C", "Canale 7 - Protezione Civile", "Canale 8 - Volontari", "Canale 9 - Mezzi", "Canale 10 - Base", "Canale 11 - Ponte Radio", "Canale 12 - Riserva"]

    c1, c2 = st.columns(2)
    with c1:
        alias_n = st.text_input("Alias *", key="alias_n", placeholder="Es: Centrale, Squadra A...")
        id_r = st.text_input("ID Radio *", key="alias_id", placeholder="Es: 101, 202...")
        # RIGA 2874 - Volontario agganciato a form Volontari campo Cognome e Nome in combo - richiesta Ezio
        vol_list_alias = [f"{v.get('Cognome','').strip()} {v.get('Nome','').strip()}" for v in st.session_state.volontari if v.get('Cognome') or v.get('Nome')]
        vol_list_alias = sorted(list(set([x for x in vol_list_alias if x.strip()])))
        if vol_list_alias:
            sel_vol_alias = st.selectbox("Volontario * (Cognome Nome da Volontari)", vol_list_alias, key="alias_vol", help="Lista da form Volontari - campo Cognome e Nome in combo")
        else:
            sel_vol_alias = st.text_input("Volontario * (manuale - aggiungi volontari in Volontari)", key="alias_vol_man", placeholder="Cognome Nome")

    with c2:
        gruppo = st.selectbox("Gruppo", ["Squadra A", "Squadra B", "Squadra C", "Coordinamento", "Logistica"], key="alias_gruppo")
        # RIGA 2883 - Canale Radio di lavoro combo - richiesta Ezio
        canale_radio = st.selectbox("Canale Radio Lavoro * (combo)", canali_radio_list, key="alias_canale", help="Assegna canale radio di lavoro - combo canali")
        desc = st.text_input("Descrizione", key="alias_desc", placeholder="Descrizione alias")
        note_alias = st.text_area("Note", key="alias_note", placeholder="Note alias radio...", height=80)

    st.markdown("</div>", unsafe_allow_html=True)

    def clear_alias():
        for k in ["alias_n", "alias_id", "alias_vol", "alias_vol_man", "alias_gruppo", "alias_canale", "alias_desc", "alias_note"]:
            if k in st.session_state:
                del st.session_state[k]

    c_save_alias1, c_save_alias2 = st.columns([3,1])
    with c_save_alias1:
        save_alias_btn = st.button("💾 Salva Alias", type="primary", use_container_width=True, key="btn_alias_save")
    with c_save_alias2:
        st.button("🔄 Pulisci maschera", use_container_width=True, key="btn_pulisci_alias", on_click=clear_alias)

    if save_alias_btn:
        if alias_n and id_r and sel_vol_alias:
            st.session_state.alias_radio.append({
                "Alias": alias_n,
                "ID Radio": id_r,
                "Volontario": sel_vol_alias,
                "Gruppo": gruppo,
                "Canale Radio": canale_radio,
                "Descrizione": desc,
                "Note": note_alias,
                "Data": datetime.now().strftime("%d/%m/%Y")
            })
            st.success(f"Alias salvato: {alias_n} - {sel_vol_alias} - {canale_radio}")
            st.rerun()
        else:
            st.error("Compila Alias *, ID Radio * e Volontario *")

    if st.session_state.alias_radio:
        st.divider()
        st.markdown(f"#### 📋 Elenco Alias Radio ({len(st.session_state.alias_radio)})")
        df_al = pd.DataFrame(st.session_state.alias_radio)
        cols_alias = [c for c in ["Alias","ID Radio","Volontario","Gruppo","Descrizione","Note"] if c in df_al.columns]
        if cols_alias:
            st.dataframe(df_al[cols_alias], use_container_width=True)
        else:
            st.dataframe(df_al, use_container_width=True)
        c_a1, c_a2 = st.columns(2)
        with c_a1:
            st.download_button("Excel Alias", to_excel(df_al), "alias.xlsx", use_container_width=True, key="dl_alias_excel")
        with c_a2:
            if REPORTLAB_OK:
                st.download_button("PDF Alias", to_pdf(df_al, "ALIAS RADIO - CON VOLONTARIO"), "alias.pdf", use_container_width=True, key="dl_alias_pdf")

# BROGLIACCIO


    # IMPORT/EXPORT INLINE - Ezio - TUTTI I FORM - Excel + PDF + Template ODV
    excel_import_inline("alias_radio", "Alias Radio")


elif cur == "Brogliaccio":
    hdr_form("BROGLIACCIO RADIO")

    # Etichetta Brogliaccio rimossa

    # Prepara liste combo da Alias Radio e Volontari
    # Alias Radio -> lista Alias
    alias_list = []
    try:
        alias_data = st.session_state.get("alias_radio", [])
        for a in alias_data:
            al = a.get("Alias", "").strip() if isinstance(a, dict) else ""
            if al and al not in alias_list:
                alias_list.append(al)
    except:
        pass
    if not alias_list:
        alias_list = ["Centrale Operativa", "Squadra A", "Squadra B", "Squadra C", "Coordinamento"]

    # Volontari -> lista Nome Cognome
    volontari_list = []
    try:
        vol_data = st.session_state.get("volontari", [])
        for v in vol_data:
            if isinstance(v, dict):
                nome = v.get("Nome", "").strip()
                cognome = v.get("Cognome", "").strip()
                full = f"{nome} {cognome}".strip()
                if full and full not in volontari_list:
                    volontari_list.append(full)
    except:
        pass
    if not volontari_list:
        volontari_list = ["Ezio Fiscato", "Operatore 1", "Operatore 2"]

    c1, c2 = st.columns(2)
    with c1:
        data_b = st.date_input("Data", value=date.today(), format="DD/MM/YYYY", key="brog_data")
        ora_b = st.time_input("Ora", value=datetime.now().time(), key="brog_ora")
        # OPERATORE COMBO DA VOLONTARI NOME COGNOME - Richiesta Ezio
        operatore = st.selectbox("Operatore (da Volontari Nome Cognome)", volontari_list, key="brog_op_combo", help="Lista agganciata a form Volontari - Nome Cognome")
        # Permetti anche inserimento manuale se non in lista
        operatore_custom = st.text_input("Oppure inserisci Operatore manuale", key="brog_op_custom", placeholder="Se non in lista volontari")
        if operatore_custom.strip():
            operatore = operatore_custom.strip()

        # CHIAMATE COMBO DA ALIAS RADIO - Richiesta Ezio
        chiamate = st.selectbox("Chiamate (da Alias Radio - Alias)", alias_list, key="brog_chiamate", help="Lista agganciata a form Alias Radio - campo Alias")
        chiamate_custom = st.text_input("Oppure Chiamate manuale", key="brog_chiamate_custom", placeholder="Alias non in lista")
        if chiamate_custom.strip():
            chiamate = chiamate_custom.strip()

    with c2:
        evento_b = st.text_input("Evento Riferimento", key="brog_evento")
        emerg_b = st.text_input("Emergenza Riferimento", key="brog_emerg")
        # RICEVENTE COMBO DA ALIAS RADIO - Richiesta Ezio
        ricevente = st.selectbox("Ricevente (da Alias Radio - Alias)", alias_list, key="brog_ricevente", help="Lista agganciata a form Alias Radio - campo Alias")
        ricevente_custom = st.text_input("Oppure Ricevente manuale", key="brog_ricevente_custom", placeholder="Alias non in lista")
        if ricevente_custom.strip():
            ricevente = ricevente_custom.strip()
        blindato = st.checkbox("Blinda Evento/Emergenza", key="brog_blind")

    testo_b = st.text_area("Testo Brogliaccio *", height=150, key="brog_testo")

    if st.button("Salva Brogliaccio", type="primary", use_container_width=True):
        if testo_b:
            st.session_state.brogliaccio.append({
                "Data": str(data_b),
                "Ora": str(ora_b),
                "Operatore": operatore,
                "Chiamate": chiamate,
                "Ricevente": ricevente,
                "Evento": evento_b,
                "Emergenza": emerg_b,
                "Testo": testo_b,
                "Blindato": blindato
            })
            st.success(f"Brogliaccio salvato - Op: {operatore} - Chiamate: {chiamate} -> Ricevente: {ricevente}")
            st.rerun()
        else:
            st.error("Compila Testo Brogliaccio *")

    if st.session_state.brogliaccio:
        df_br = pd.DataFrame(st.session_state.brogliaccio)
        st.markdown(f"#### Elenco Brogliaccio ({len(st.session_state.brogliaccio)})")
        st.dataframe(df_br, use_container_width=True)
        # Tabella con colonne importanti
        cols_show = ["Data","Ora","Operatore","Chiamate","Ricevente","Testo","Evento"]
        cols_show = [c for c in cols_show if c in df_br.columns]
        if cols_show:
            st.dataframe(df_br[cols_show], use_container_width=True)

# EVENTI


    # IMPORT/EXPORT INLINE - Ezio - TUTTI I FORM - Excel + PDF + Template ODV
    excel_import_inline("brogliaccio", "Brogliaccio")


elif cur == "Eventi":
    hdr_form("EVENTI - Gestione Eventi Programmati")

    c1, c2, c3 = st.columns(3)
    with c1:
        nome_ev = st.text_input("Nome Evento *", key="ev_nome")
        tipo_ev = st.selectbox("Tipo Evento", ["Esercitazione", "Manifestazione", "Formazione", "Riunione", "Altro"], key="ev_tipo")
        data_ev = st.date_input("Data Evento", value=date.today(), format="DD/MM/YYYY", key="ev_data")

    with c2:
        comune_ev = combo_comune("Comune Evento", "ev_comune", "Varese")
        via_ev = combo_vie("Via Evento", comune_ev, "ev_via", "")
        ora_ev = st.time_input("Ora Inizio", value=time(9, 0), key="ev_ora")

    with c3:
        resp_ev = st.text_input("Responsabile", key="ev_resp")
        stato_ev = st.selectbox("Stato", ["Programmato", "In Corso", "Completato", "Annullato"], key="ev_stato")
        note_ev = st.text_area("Note Evento", key="ev_note")

    col_save_ev1, col_save_ev2 = st.columns([3,1])
    with col_save_ev1:
        save_ev = st.button("💾 Salva Evento", type="primary", use_container_width=True, key="btn_salva_ev")
    with col_save_ev2:
        if st.button("🔄 Pulisci campi", use_container_width=True, key="btn_pulisci_ev", help="Pulisce campi maschera"):
            for k in ["ev_nome", "ev_tipo", "ev_data", "ev_comune", "ev_via", "ev_ora", "ev_resp", "ev_stato", "ev_note"]:
                if k in st.session_state:
                    try:
                        del st.session_state[k]
                    except:
                        pass
            st.rerun()
    if save_ev:
        if nome_ev:
            st.session_state.eventi.append({
                "Nome": nome_ev,
                "Tipo": tipo_ev,
                "Data": str(data_ev),
                "Comune": comune_ev,
                "Via": via_ev,
                "Ora": str(ora_ev),
                "Responsabile": resp_ev,
                "Stato": stato_ev,
                "Note": note_ev
            })
            st.success("Evento salvato")
            st.rerun()

    if st.session_state.eventi:
        df_ev = pd.DataFrame(st.session_state.eventi)
        st.dataframe(df_ev, use_container_width=True)
        st.download_button("Excel Eventi", to_excel(df_ev), "eventi.xlsx", use_container_width=True)
        if REPORTLAB_OK:
            st.download_button("PDF Logo Estesa", to_pdf(df_ev, "EVENTI"), "eventi.pdf", use_container_width=True)

# EMERGENZE


    # IMPORT/EXPORT INLINE - Ezio - TUTTI I FORM - Excel + PDF + Template ODV
    excel_import_inline("eventi", "Eventi")


elif cur == "Emergenze":
    hdr_form("EMERGENZE - Gestione Emergenze Attive")

    c1, c2, c3 = st.columns(3)
    with c1:
        nome_em = st.text_input("Nome Emergenza *", key="em_nome")
        tipo_em = st.selectbox("Tipo Emergenza", ["Alluvione", "Incendio", "Frana", "Neve", "Ricerca Persona", "Altro"], key="em_tipo")
        data_em = st.date_input("Data Emergenza", value=date.today(), format="DD/MM/YYYY", key="em_data")

    with c2:
        comune_em = combo_comune("Comune Emergenza", "em_comune", "Varese")
        via_em = combo_vie("Via Emergenza", comune_em, "em_via", "")
        prior_em = st.selectbox("Priorità", ["Bassa", "Media", "Alta", "Critica"], key="em_prior")

    with c3:
        stato_em = st.selectbox("Stato", ["Operativo", "In Corso", "Completato", "Chiuso"], key="em_stato")
        bg_c, txt_c, lab_c = get_stato_color(stato_em)
    note_em = st.text_area("Descrizione Emergenza", key="em_note")

    col_em_save1, col_em_save2 = st.columns([3,1])
    with col_em_save1:
        save_em = st.button("💾 Salva Emergenza", type="primary", use_container_width=True, key="btn_salva_em")
    with col_em_save2:
        if st.button("🔄 Pulisci campi", use_container_width=True, key="btn_pulisci_em", help="Pulisce campi maschera"):
            for k in ["em_nome", "em_tipo", "em_data", "em_comune", "em_via", "em_prior", "em_stato", "em_coord", "em_note"]:
                if k in st.session_state:
                    try:
                        del st.session_state[k]
                    except:
                        pass
            st.rerun()
    if save_em:
        if nome_em:
            try:
                coord_val = coord_em
            except NameError:
                coord_val = ""
            st.session_state.emergenze.append({
                "Nome": nome_em,
                "Tipo": tipo_em,
                "Data": str(data_em),
                "Comune": comune_em,
                "Via": via_em,
                "Priorita": prior_em,
                "Stato": stato_em,
                "StatoColoreBg": bg_c,
                "StatoColoreTxt": txt_c,
                "Coordinate": coord_val,
                "Note": note_em
            })
            st.success("Emergenza salvata")
            st.rerun()

    if st.session_state.emergenze:
        df_em = pd.DataFrame(st.session_state.emergenze)
        st.dataframe(df_em, use_container_width=True)
        st.download_button("Excel Emergenze", to_excel(df_em), "emergenze.xlsx", use_container_width=True)
        if REPORTLAB_OK:
            st.download_button("PDF Logo Estesa", to_pdf(df_em, "EMERGENZE"), "emergenze.pdf", use_container_width=True)

# MAPPE (Emergenze+Eventi) FUSIONE - SI OTTIMA IDEA


    # IMPORT/EXPORT INLINE - Ezio - TUTTI I FORM - Excel + PDF + Template ODV
    excel_import_inline("emergenze", "Emergenze")


# TABELLA EMERGENZE - FORM TABELLA - Richiesta Ezio - Formato tabella per vedere emergenze - FIX non vedo niente
elif cur == "Tabella Emergenze":
    hdr_form("TABELLA EMERGENZE - GRIGLIA EXCEL")

    st.markdown("""
    <div style="background:#e3f2fd;padding:10px;border-radius:8px;border-left:4px solid #1976d2;margin-bottom:12px;">
    <b>📋 Tabella Emergenze - Griglia Excel con celle evidenziate</b><br>
    Icona è un vero tasto - Clicca immagine per aprire Interventi Emergenza e aggiornare - Celle come Excel
    </div>
    """, unsafe_allow_html=True)

    st.info(f"Debug: {len(st.session_state.interventi)} interventi di emergenza - Icona cliccabile come tasto")

    if not st.session_state.interventi:
        st.warning("⚠️ Nessun intervento - Vai in Interventi Emergenza e crea con icona")
        st.dataframe(pd.DataFrame(columns=["Icona","Tipo","Squadra","Data","Comune","Via","Stato","Descrizione"]).head(), use_container_width=True)
    else:
        df_tab_em = pd.DataFrame(st.session_state.interventi)
        
        # Filtri
        squadre_list_em = sorted(list(set([str(x) for x in df_tab_em["Squadra"].tolist() if x]))) if "Squadra" in df_tab_em.columns else []
        comuni_list_em = sorted(list(set([str(x) for x in df_tab_em["Comune"].tolist() if x]))) if "Comune" in df_tab_em.columns else []
        stati_list_em = sorted(list(set([str(x) for x in df_tab_em["Stato"].tolist() if x]))) if "Stato" in df_tab_em.columns else []
        tipi_list_em = sorted(list(set([str(x) for x in df_tab_em["Tipo"].tolist() if x]))) if "Tipo" in df_tab_em.columns else []

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            filtro_sq_em = st.selectbox("Filtra Squadra", ["Tutte"] + squadre_list_em, key="tab_em_sq_excel")
        with c2:
            filtro_com_em = st.selectbox("Filtra Comune", ["Tutti"] + comuni_list_em, key="tab_em_com_excel")
        with c3:
            filtro_stato_em = st.selectbox("Filtra Stato", ["Tutti"] + stati_list_em, key="tab_em_stato_excel")
        with c4:
            filtro_tipo_em = st.selectbox("Filtra Tipo", ["Tutti"] + tipi_list_em, key="tab_em_tipo_excel")

        df_filtrato_em = df_tab_em.copy()
        if filtro_sq_em != "Tutte" and "Squadra" in df_filtrato_em.columns:
            df_filtrato_em = df_filtrato_em[df_filtrato_em["Squadra"] == filtro_sq_em]
        if filtro_com_em != "Tutti" and "Comune" in df_filtrato_em.columns:
            df_filtrato_em = df_filtrato_em[df_filtrato_em["Comune"] == filtro_com_em]
        if filtro_stato_em != "Tutti" and "Stato" in df_filtrato_em.columns:
            df_filtrato_em = df_filtrato_em[df_filtrato_em["Stato"] == filtro_stato_em]
        if filtro_tipo_em != "Tutti" and "Tipo" in df_filtrato_em.columns:
            df_filtrato_em = df_filtrato_em[df_filtrato_em["Tipo"] == filtro_tipo_em]

        st.write(f"**Risultati: {len(df_filtrato_em)} su {len(df_tab_em)} - Celle come Excel**")

        # STILE EXCEL per celle
        st.markdown("""
        <style>
        .excel-header {
            background:#1A5D1A !important;
            color:white !important;
            font-weight:bold;
            border:1px solid #000 !important;
            padding:8px !important;
            text-align:center;
            font-size:12px;
        }
        .excel-cell {
            border:1px solid #999 !important;
            padding:6px 8px !important;
            background:white !important;
            font-size:11px;
            min-height:40px;
            display:flex;
            align-items:center;
        }
        .excel-cell-icon {
            border:2px solid #1A5D1A !important;
            background:#e8f5e9 !important;
            cursor:pointer;
        }
        </style>
        """, unsafe_allow_html=True)

        # HEADER Excel-like
        h_cols = st.columns([0.9, 1, 1, 1.1, 1.2, 1.2, 1, 2.2])
        headers = ["ICONA (Tasto)", "Tipo", "Squadra", "Data/Ora", "Comune", "Via", "Stato", "Descrizione"]
        for col, header in zip(h_cols, headers):
            with col:
                st.markdown(f'<div class="excel-header">{header}</div>', unsafe_allow_html=True)
        
        # RIGHE con celle Excel - FIX icona centrata, click apre Interventi Emergenza, no pennetta
        for idx_orig, row in df_filtrato_em.iterrows():
            r_cols = st.columns([0.9, 1, 1, 1.1, 1.2, 1.2, 1, 2.2])
            emoji = row.get("IconaEmoji", "📍")
            nome_ico = row.get("IconaNome", "")
            has_file = row.get("HasFile", False)
            file_bytes = row.get("FileBytes", None)
            tipo_val = row.get("Tipo","")
            squadra_val = row.get("Squadra","")
            data_val = row.get("Data","")[:10] if row.get("Data") else ""
            ora_val = row.get("Ora","")[:5] if row.get("Ora") else ""
            comune_val = row.get("Comune","")
            via_val = row.get("Via","")[:25] if row.get("Via") else ""
            stato_val = row.get("Stato","")
            desc_val = str(row.get("Descrizione",""))[:90] if row.get("Descrizione") else ""

            # COLONNA ICONA - FIX: icona centrata, tasto vero senza pennetta, apre Interventi Emergenza per bonifica
            with r_cols[0]:
                # Contenitore cella Excel con bordo
                with st.container(border=True):
                    if has_file and file_bytes:
                        try:
                            st.image(file_bytes, width=50)
                        except:
                            st.markdown(f"<div style='font-size:28px;text-align:center;'>{emoji}</div>", unsafe_allow_html=True)
                    else:
                        # Emoji grande centrata
                        st.markdown(f"<div style='font-size:32px;text-align:center;background:white;border-radius:8px;padding:4px;'>{emoji}</div>", unsafe_allow_html=True)
                    # Tasto sotto icona - FIX: apre Interventi Emergenza per modifiche, non pagina iniziale
                    btn_label = f"{nome_ico[:12]}" if nome_ico else f"{tipo_val[:10]}"
                    # Chiave stabile senza spazi - usa solo idx_orig
                    btn_key = f"open_int_em_{int(idx_orig)}"
                    if st.button(btn_label, key=btn_key, help=f"Apri Interventi Emergenza per modificare - {tipo_val} - {comune_val}", use_container_width=True, type="primary"):
                        # FIX: Assicura di rimanere loggato e in dashboard, non pagina iniziale
                        st.session_state.logged = True
                        st.session_state.page = "dashboard"
                        # Pulisci campi int_ vecchi tranne edit_index
                        for k in [k for k in list(st.session_state.keys()) if k.startswith("int_") and k != "int_edit_index"]:
                            try:
                                del st.session_state[k]
                            except:
                                pass
                        # Imposta indice per aprire proprio quell'emergenza vista in tabella per modifiche
                        st.session_state.int_edit_index = int(idx_orig)
                        # Pre-popola
                        try:
                            st.session_state["int_tipo"] = tipo_val
                            st.session_state["int_squadra"] = squadra_val
                            st.session_state["int_comune"] = comune_val
                            st.session_state["int_via"] = via_val
                            st.session_state["int_desc"] = row.get("Descrizione","")
                            st.session_state["stato_int"] = stato_val
                        except:
                            pass
                        # Vai a Interventi Emergenza
                        st.session_state.menu = "Interventi Emergenza"
                        # Salva anche in query per sicurezza
                        try:
                            st.session_state["force_menu"] = "Interventi Emergenza"
                        except:
                            pass
                        st.rerun()

            # COLONNE DATI con celle Excel evidenziate - contenitore border per effetto Excel
            with r_cols[1]:
                with st.container(border=True):
                    st.write(tipo_val)
            with r_cols[2]:
                with st.container(border=True):
                    st.write(squadra_val)
            with r_cols[3]:
                with st.container(border=True):
                    st.write(f"{data_val} {ora_val}")
            with r_cols[4]:
                with st.container(border=True):
                    st.markdown(f"**{comune_val}**")
            with r_cols[5]:
                with st.container(border=True):
                    st.write(via_val)
            with r_cols[6]:
                with st.container(border=True):
                    bg = row.get("StatoColoreBg", "#e8f5e9")
                    txt_c = row.get("StatoColoreTxt", "black")
                    st.markdown(f"<span style='background:{bg};color:{txt_c};padding:2px 8px;border-radius:12px;font-size:11px;font-weight:bold;border:1px solid black;'>{stato_val}</span>", unsafe_allow_html=True)
            with r_cols[7]:
                with st.container(border=True):
                    st.write(desc_val)

        # Export
        st.divider()
        cols_to_remove_em = ["StatoColoreBg", "StatoColoreTxt", "IconaLabel", "IconaNome", "IconaColore", "IconaTipo", "IconaEmoji", "FileBytes", "FileName", "HasFile"]
        df_export_em = df_filtrato_em.copy()
        for col in cols_to_remove_em:
            if col in df_export_em.columns:
                df_export_em = df_export_em.drop(columns=[col])
        c_exp1, c_exp2 = st.columns(2)
        with c_exp1:
            st.download_button("Excel Griglia Excel", data=to_excel(df_export_em), file_name="tabella_emergenze_excel.xlsx", use_container_width=True, key="exp_tab_em_excel")
        with c_exp2:
            if REPORTLAB_OK:
                st.download_button("PDF Griglia Excel", data=to_pdf(df_export_em, "TABELLA EMERGENZE - EXCEL"), file_name="tabella_emergenze_excel.pdf", use_container_width=True, key="pdf_tab_em_excel")

    excel_import_inline("interventi", "Tabella Emergenze - Excel")





elif cur == "# RIMOSSO":
    hdr_form("MAPPE - Fusione Emergenze + Eventi - Proposta Ezio - SI OTTIMA IDEA")

    st.markdown(
        """
        <div style="background:linear-gradient(135deg,#e3f2fd,#bbdefb);
        padding:16px;border-radius:12px;border:2px solid #1976d2;
        margin-bottom:16px;">
        <h4 style="margin:0;color:#0d47a1;">Fusione Emergenze+Eventi in Mappe: SI ottima idea</h4>
        <p style="margin:8px 0 0 0;color:black;">
        Unico form georeferenziato con Tipo Emergenza/Evento, filtri, priorità e stato colorato.
        Soluzione ottimale approvata - gestione unificata su mappa unica.
        </p>
        </div>
        """,
        unsafe_allow_html=True
    )

    c1, c2, c3 = st.columns(3)
    with c1:
        tipo_mappa = st.selectbox("Tipo Mappa *", ["Emergenza", "Evento"], key="mappa_tipo")
        nome_mappa = st.text_input("Nome / Titolo *", key="mappa_nome")
        data_mappa = st.date_input("Data", value=date.today(), format="DD/MM/YYYY", key="mappa_data")

    with c2:
        comune_mappa = combo_comune("Comune", "mappa_comune", "Varese")
        via_mappa = combo_vie("Via", comune_mappa, "mappa_via", "")
        prior_mappa = st.selectbox("Priorità", ["Bassa", "Media", "Alta", "Critica"], key="mappa_prior")

    with c3:
        stato_mappa = st.selectbox(
            "STATO *",
            ["Operativo", "In Corso", "Completato", "Chiuso", "In Stand By", "Sospeso", "Annullato", "In Attesa"],
            key="stato_mappa_fusione"
        )
        bg_m, txt_m, lab_m = get_stato_color(stato_mappa)
        st.markdown(
            f"""
            <div style="background:{bg_m};color:{txt_m};padding:8px;
            border-radius:8px;text-align:center;font-weight:bold;
            border:3px solid black;margin-top:8px;">
            STATO SELEZIONATO: {lab_m} - Fondo campo colorato come richiesto Modifica 4
            </div>
            """,
            unsafe_allow_html=True
        )
        st.markdown(
            f"""
            <style>
            div[data-testid='stSelectbox']:has(#stato_mappa_fusione) div[data-baseweb='select'] {{
                background-color:{bg_m} !important;
            }}
            </style>
            """,
            unsafe_allow_html=True
        )

    c4, c5 = st.columns(2)
    with c4:
        lat_mappa = st.text_input("Latitudine", value="45.8167", key="mappa_lat")
        lon_mappa = st.text_input("Longitudine", value="8.8333", key="mappa_lon")
        icona_mappa = st.selectbox("Icona", ["🚨", "📅", "🚒", "👷", "📍", "⚠️"], key="mappa_icona")

    with c5:
        desc_mappa = st.text_area("Descrizione", key="mappa_desc")
        note_mappa = st.text_area("Note Coordinate", key="mappa_note")

    if st.button("Salva in Mappe Postazioni", type="primary", use_container_width=True):
        if nome_mappa:
            st.session_state.mappe.append({
                "Tipo": tipo_mappa,
                "Nome": nome_mappa,
                "Data": str(data_mappa),
                "Comune": comune_mappa,
                "Via": via_mappa,
                "Priorita": prior_mappa,
                "Stato": stato_mappa,
                "StatoBg": bg_m,
                "StatoTxt": txt_m,
                "Lat": lat_mappa,
                "Lon": lon_mappa,
                "Icona": icona_mappa,
                "Descrizione": desc_mappa,
                "Note": note_mappa
            })
            st.success("Mappa salvata - Fusione OK")
            st.rerun()

    if st.session_state.mappe:
        st.divider()
        st.markdown("**Riepilogo Mappe Postazioni con Filtri Tipo**")
        filtro_tipo = st.selectbox("Filtra per Tipo", ["Tutti", "Emergenza", "Evento"], key="filtro_mappa_tipo")
        df_map = pd.DataFrame(st.session_state.mappe)
        if filtro_tipo != "Tutti":
            df_map = df_map[df_map["Tipo"] == filtro_tipo]

        st.dataframe(df_map, use_container_width=True)

        # Mappa semplice con st.map se coordinate valide
        try:
            map_df = pd.DataFrame([
                {"lat": float(m.get("Lat", 0)), "lon": float(m.get("Lon", 0))}
                for m in st.session_state.mappe
                if m.get("Lat") and m.get("Lon")
            ])
            if not map_df.empty:
                st.map(map_df)
        except:
            st.info("Mappa coordinate non disponibili per visualizzazione")

        st.download_button("Excel Mappe Postazioni", to_excel(df_map), "mappe_fusione.xlsx", use_container_width=True)
        if REPORTLAB_OK:
            st.download_button("PDF Logo Estesa Tutto Foglio", to_pdf(df_map, "MAPPE FUSIONE EMERGENZE+EVENTI"), "mappe_fusione.pdf", use_container_width=True)

# CHECK-IN
elif cur == "Mappe":
    hdr_form("MAPPE PER POSTAZIONI/CANTIERI/EMERGENZE IN CORSO")
    c1, c2 = st.columns(2)
    with c1:
        tipo_mappa = st.selectbox("Tipo Mappa *", ["Emergenza", "Evento"], key="mappa_tipo_old2")
        nome_mappa = st.text_input("Nome / Titolo *", key="mappa_nome_old2")
    with c2:
        comune_mappa = combo_comune("Comune", "mappa_comune_old2", "Varese")
        via_mappa = combo_vie("Via", comune_mappa, "mappa_via_old2", "")
        lat_mappa = st.text_input("Latitudine", value="45.8167", key="mappa_lat_old2")
        lon_mappa = st.text_input("Longitudine", value="8.8333", key="mappa_lon_old2")
    if st.button("Salva in Mappe", type="primary", use_container_width=True):
        if nome_mappa:
            st.session_state.mappe.append({"Tipo": tipo_mappa, "Nome": nome_mappa, "Comune": comune_mappa, "Via": via_mappa, "Lat": lat_mappa, "Lon": lon_mappa})
            st.success("Salvata")
            st.rerun()
    if st.session_state.mappe:
        st.dataframe(pd.DataFrame(st.session_state.mappe), use_container_width=True)

elif cur == "Check-in":
    hdr_form("CHECK-IN - Presenze Operative - Scanner Tessera Sanitaria")

    # CAMPO CODICE FISCALE CON PISTOLA BARCODE - Richiesta Ezio
    st.markdown("""
    <div style="background:#FFD700;padding:10px;border-radius:8px;border:2px solid #1A5D1A;margin-bottom:10px;">
    <b>📷 SCANNER TESSERA SANITARIA - Pistola Barcode CF</b><br>
    Posiziona cursore nel campo sotto e spara con pistola - Legge CF da barcode tessera sanitaria
    </div>
    """, unsafe_allow_html=True)
    
    cf_check = st.text_input("🔍 Codice Fiscale - SCANNER PISTOLA (tessera sanitaria) *", key="check_cf", placeholder="Scansiona barcode tessera sanitaria con pistola...", help="Spara con pistola barcode sulla tessera sanitaria - Legge CF e cerca volontario")
    
    # Auto-lookup volontario da CF
    vol_check_trovato = None
    if cf_check:
        cf_clean = cf_check.strip().upper()
        # Cerca CF in volontari
        for v in st.session_state.volontari:
            cf_v = (v.get('CodiceFiscale','') or v.get('CF','') or '').strip().upper()
            if cf_v and cf_v == cf_clean:
                vol_check_trovato = f"{v.get('Cognome','')} {v.get('Nome','')}"
                st.success(f"✅ Volontario trovato da CF: {vol_check_trovato}")
                break
        if not vol_check_trovato:
            # Prova anche se CF è lungo (barcode potrebbe avere altri dati)
            if len(cf_clean) >= 16:
                possible_cf = cf_clean[-16:] if len(cf_clean) > 16 else cf_clean[:16]
                for v in st.session_state.volontari:
                    cf_v = (v.get('CodiceFiscale','') or v.get('CF','') or '').strip().upper()
                    if cf_v and possible_cf in cf_clean:
                        vol_check_trovato = f"{v.get('Cognome','')} {v.get('Nome','')}"
                        st.success(f"✅ Volontario trovato da CF (estratto): {vol_check_trovato}")
                        break
            if not vol_check_trovato:
                st.warning(f"⚠️ CF {cf_clean} non trovato in anagrafica volontari - Inserisci manuale sotto")

    c1, c2 = st.columns(2)
    with c1:
        vol_check_list = [f"{v.get('Cognome','')} {v.get('Nome','')}" for v in st.session_state.volontari]
        if vol_check_list:
            # Se trovato da CF, preseleziona
            if vol_check_trovato and vol_check_trovato in vol_check_list:
                idx_vol = vol_check_list.index(vol_check_trovato)
                sel_check_vol = st.selectbox("Volontario", vol_check_list, index=idx_vol, key="check_vol")
            else:
                sel_check_vol = st.selectbox("Volontario", vol_check_list, key="check_vol")
        else:
            sel_check_vol = st.text_input("Volontario", value=vol_check_trovato if vol_check_trovato else "", key="check_vol_man")

        ev_check_list = [e.get("Nome", "") for e in st.session_state.eventi]
        em_check_list = [em.get("Nome", "") for em in st.session_state.emergenze]
        all_ref = ev_check_list + em_check_list
        if all_ref:
            sel_check_ref = st.selectbox("Evento/Emergenza", all_ref, key="check_ref")
        else:
            sel_check_ref = st.text_input("Evento/Emergenza", key="check_ref_man")

    with c2:
        data_check = st.date_input("Data Check-in", value=date.today(), format="DD/MM/YYYY", key="check_data")
        ora_check = st.time_input("Ora Check-in", value=datetime.now().time(), key="check_ora")
        stato_check = st.selectbox("Stato", ["Presente", "Assente", "Ritardo"], key="check_stato")

    if st.button("Registra Check-in", type="primary", use_container_width=True):
        # Usa CF scanner se presente
        cf_to_save = cf_check.strip().upper() if cf_check else ""
        if vol_check_trovato:
            sel_vol_final = vol_check_trovato
        else:
            sel_vol_final = sel_check_vol
        st.session_state.checkin.append({
            "Volontario": sel_vol_final,
            "CodiceFiscale": cf_to_save,
            "Riferimento": sel_check_ref,
            "Data": str(data_check),
            "Ora": str(ora_check),
            "Stato": stato_check
        })
        st.success(f"Check-in registrato: {sel_vol_final} - CF: {cf_to_save}")
        # Pulisci CF per prossimo scan
        if "check_cf" in st.session_state:
            del st.session_state["check_cf"]
        st.rerun()

    if st.session_state.checkin:
        df_ch = pd.DataFrame(st.session_state.checkin)
        st.dataframe(df_ch, use_container_width=True)
        st.download_button("Excel Check-in", to_excel(df_ch), "checkin.xlsx", use_container_width=True)

# INTERVENTI EMERGENZA - MODIFICA 4 STATO COLORE FONDO CAMPO
    # IMPORT/EXPORT INLINE - Ezio - TUTTI I FORM - Excel + PDF + Template ODV
    excel_import_inline("checkin", "Check-in")


elif cur == "Interventi Emergenza":
    hdr_form("INTERVENTI EMERGENZA")

    st.markdown(
        """
        <div style="background:#e8f5e9;padding:10px;border-radius:8px;border-left:4px solid #1A5D1A;margin-bottom:10px;">
        Interventi - Puoi caricare un'icona dalla Libreria Icone - Clicca icona in Tabella Interventi per modificare qui
        </div>
        """,
        unsafe_allow_html=True
    )
    
    # Edit mode - come volontari - richiesta Ezio: clicca icona in tabella per vedere scheda in maschera
    edit_mode_int = False
    edit_data_int = {}
    if st.session_state.get("int_edit_index") is not None:
        try:
            edit_data_int = st.session_state.interventi[st.session_state.int_edit_index]
            edit_mode_int = True
        except:
            edit_data_int = {}
            edit_mode_int = False
    
    if edit_mode_int:
        st.warning(f"✏️ Modifica Intervento: {edit_data_int.get('Tipo','')} - {edit_data_int.get('Comune','')} - {edit_data_int.get('Data','')} - Icona {edit_data_int.get('IconaEmoji','')} {edit_data_int.get('IconaNome','')}")
        c_w1, c_w2 = st.columns(2)
        with c_w1:
            if st.button("❌ Annulla Modifica - Torna a nuovo intervento", key="annulla_edit_int"):
                st.session_state.int_edit_index = None
                for k in list(st.session_state.keys()):
                    if k.startswith("int_"):
                        try:
                            del st.session_state[k]
                        except:
                            pass
                st.rerun()
        with c_w2:
            if st.button("🗑️ Elimina questo intervento", key="elimina_edit_int"):
                try:
                    st.session_state.interventi.pop(st.session_state.int_edit_index)
                    st.session_state.int_edit_index = None
                    st.success("Intervento eliminato")
                    st.rerun()
                except Exception as e:
                    st.error(f"Errore eliminazione: {e}")

    # Prepara lista icone da Libreria Icone - Richiesta Ezio
    icone_lib = st.session_state.get("icone", [])
    if not icone_lib:
        icone_lib = [
            {"Nome": "Emergenza", "Emoji": "🚨", "Tipo": "Emergenza", "Colore": "red"},
            {"Nome": "Soccorso", "Emoji": "👷", "Tipo": "Emergenza", "Colore": "red"},
            {"Nome": "Mezzo", "Emoji": "🚐", "Tipo": "Mezzo", "Colore": "green"},
            {"Nome": "Logistica", "Emoji": "📦", "Tipo": "Logistica", "Colore": "blue"},
        ]
    
    # Crea opzioni per selectbox: Emoji + Nome
    icone_options = ["-- Nessuna Icona --"]
    icone_map = {"-- Nessuna Icona --": None}
    for ico in icone_lib:
        label = f"{ico.get('Emoji','📍')} {ico.get('Nome','')} - {ico.get('Tipo','')} ({ico.get('Colore','')})"
        icone_options.append(label)
        icone_map[label] = ico

    # Valori default da edit se in modifica
    tipo_def = edit_data_int.get("Tipo", "Soccorso") if edit_mode_int else "Soccorso"
    squadra_def = edit_data_int.get("Squadra", "Squadra A") if edit_mode_int else "Squadra A"
    comune_def_int = edit_data_int.get("Comune", "Varese") if edit_mode_int else "Varese"
    via_def_int = edit_data_int.get("Via", "") if edit_mode_int else ""
    stato_def_int = edit_data_int.get("Stato", "Operativo") if edit_mode_int else "Operativo"
    desc_def_int = edit_data_int.get("Descrizione", "") if edit_mode_int else ""
    mezzi_def_int = edit_data_int.get("Mezzi", "") if edit_mode_int else ""
    vol_def_int = edit_data_int.get("Volontari", "") if edit_mode_int else ""
    
    c1, c2, c3 = st.columns(3)
    with c1:
        tipo_int = st.selectbox("Tipo Intervento", ["Soccorso", "Logistica", "Monitoraggio", "Bonifica", "Altro"], index=["Soccorso", "Logistica", "Monitoraggio", "Bonifica", "Altro"].index(tipo_def) if tipo_def in ["Soccorso", "Logistica", "Monitoraggio", "Bonifica", "Altro"] else 0, key="int_tipo")
        squadra_int = st.selectbox("Squadra", ["Squadra A", "Squadra B", "Squadra C", "Logistica"], index=["Squadra A", "Squadra B", "Squadra C", "Logistica"].index(squadra_def) if squadra_def in ["Squadra A", "Squadra B", "Squadra C", "Logistica"] else 0, key="int_squadra")
        try:
            data_val = date.fromisoformat(edit_data_int.get("Data","")) if edit_mode_int and edit_data_int.get("Data") else date.today()
        except:
            data_val = date.today()
        data_int = st.date_input("Data Intervento", value=data_val, format="DD/MM/YYYY", key="int_data")
        # ICONA DA LIBRERIA - Richiesta Ezio - FIX mantiene icona salvata in edit
        # Trova index di default da edit_data_int
        default_icon_idx = 0
        if edit_mode_int and edit_data_int.get("IconaLabel"):
            try:
                if edit_data_int.get("IconaLabel") in icone_options:
                    default_icon_idx = icone_options.index(edit_data_int.get("IconaLabel"))
                elif edit_data_int.get("IconaNome"):
                    # Cerca per nome
                    for i, opt in enumerate(icone_options):
                        if edit_data_int.get("IconaNome") in opt:
                            default_icon_idx = i
                            break
            except:
                default_icon_idx = 0
        icona_sel_label = st.selectbox("Icona da Libreria Icone", icone_options, index=default_icon_idx, key="int_icona", help="Scegli icona creata in Libreria Icone")
        sel_ico_obj = icone_map.get(icona_sel_label)

    with c2:
        # FIX: usa comune_def_int e via_def_int da edit per mantenere valore salvato - altrimenti aggiorna con Varese
        comune_int = combo_comune("Comune Intervento", "int_comune", comune_def_int)
        via_int = combo_vie("Via Intervento", comune_int, "int_via", via_def_int)
        ora_int = st.time_input("Ora Intervento", value=datetime.now().time(), key="int_ora")
        # Anteprima icona scelta - richiesta Ezio: vedere in anteprima icona scelte - FIX più grande e visibile
        st.markdown("**🔍 Anteprima Icona Scelta**")
        if sel_ico_obj:
            # Se ha file immagine, mostra file, altrimenti emoji
            if sel_ico_obj.get("HasFile") and sel_ico_obj.get("FileBytes"):
                try:
                    st.image(sel_ico_obj.get("FileBytes"), caption=f"Icona file: {sel_ico_obj.get('FileName','')} - {sel_ico_obj.get('Nome','')}", width=150)
                except:
                    st.markdown(f"""
                    <div style="background:white;padding:12px;border-radius:10px;border:3px solid #1A5D1A;text-align:center;margin-top:8px;box-shadow:0 2px 8px rgba(0,0,0,0.2);">
                    <div style="font-size:48px;">{sel_ico_obj.get('Emoji','📍')}</div>
                    <b style="font-size:16px;">{sel_ico_obj.get('Nome','')}</b><br>
                    <small style="color:#1A5D1A;">{sel_ico_obj.get('Tipo','')} - {sel_ico_obj.get('Colore','')}</small><br>
                    <small>{sel_ico_obj.get('Descrizione','')}</small>
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.markdown(f"""
                <div style="background:white;padding:12px;border-radius:10px;border:3px solid #1A5D1A;text-align:center;margin-top:8px;box-shadow:0 2px 8px rgba(0,0,0,0.2);">
                <div style="font-size:48px;">{sel_ico_obj.get('Emoji','📍')}</div>
                <b style="font-size:16px;">{sel_ico_obj.get('Nome','')}</b><br>
                <small style="color:#1A5D1A;font-weight:bold;">{sel_ico_obj.get('Tipo','')} - {sel_ico_obj.get('Colore','')}</small><br>
                <small>{sel_ico_obj.get('Descrizione','')}</small>
                </div>
                """, unsafe_allow_html=True)
            st.success(f"✅ Icona selezionata: {sel_ico_obj.get('Emoji','')} {sel_ico_obj.get('Nome','')}")
        else:
            # Etichetta nessuna icona rimossa - Ezio
            pass
            st.markdown("""
            <div style="background:#fff3e0;padding:10px;border-radius:8px;border:2px dashed #ff9800;text-align:center;">
            <div style="font-size:32px;">❓</div>
            <small>Seleziona icona sopra per vedere anteprima</small>
            </div>
            """, unsafe_allow_html=True)

    with c3:
        # MODIFICA 4 - STATO CON COLORE FONDO CAMPO
        stato_int = st.selectbox(
            "STATO *",
            ["Operativo", "In Corso", "Completato", "Chiuso", "In Stand By", "Sospeso", "Annullato", "In Attesa"],
            key="stato_int"
        )
        bg_color, txt_color, label = get_stato_color(stato_int)

        st.markdown(
            f"""
            <div style="background-color:{bg_color};color:{txt_color};
            padding:15px;border:3px solid black;border-radius:8px;
            text-align:center;font-weight:bold;font-size:16px;
            margin-top:10px;box-shadow:0 2px 8px rgba(0,0,0,0.3);">
            STATO SELEZIONATO: {label}<br>
            Fondo campo colorato - Modifica 4
            </div>
            """,
            unsafe_allow_html=True
        )

    # Campi descrizione con default edit - come volontari
    desc_int = st.text_area("Descrizione Intervento *", value=desc_def_int, key="int_desc")
    mezzi_int = st.text_input("Mezzi Utilizzati", value=mezzi_def_int, key="int_mezzi")
    volontari_int = st.text_input("Volontari Coinvolti", value=vol_def_int, key="int_vol")

    if st.button("💾 Aggiorna Intervento" if edit_mode_int else "💾 Salva Intervento Emergenza con Icona", type="primary", use_container_width=True):
        if desc_int:
            new_intervento = {
                "Tipo": tipo_int,
                "Squadra": squadra_int,
                "Data": str(data_int),
                "Comune": comune_int,
                "Via": via_int,
                "Ora": str(ora_int),
                "Stato": stato_int,
                "StatoColoreBg": bg_color,
                "StatoColoreTxt": txt_color,
                "Descrizione": desc_int,
                "Mezzi": mezzi_int,
                "Volontari": volontari_int,
                "IconaLabel": icona_sel_label,
                "IconaNome": sel_ico_obj.get("Nome","") if sel_ico_obj else edit_data_int.get("IconaNome",""),
                "IconaEmoji": sel_ico_obj.get("Emoji","") if sel_ico_obj else edit_data_int.get("IconaEmoji",""),
                "IconaColore": sel_ico_obj.get("Colore","") if sel_ico_obj else edit_data_int.get("IconaColore",""),
                "IconaTipo": sel_ico_obj.get("Tipo","") if sel_ico_obj else edit_data_int.get("IconaTipo",""),
                "HasFile": sel_ico_obj.get("HasFile", False) if sel_ico_obj and sel_ico_obj.get("HasFile") else edit_data_int.get("HasFile", False),
                "FileBytes": sel_ico_obj.get("FileBytes", None) if sel_ico_obj and sel_ico_obj.get("FileBytes") else edit_data_int.get("FileBytes", None),
                "FileName": sel_ico_obj.get("FileName", "") if sel_ico_obj and sel_ico_obj.get("FileName") else edit_data_int.get("FileName", "")
            }
            if edit_mode_int:
                st.session_state.interventi[st.session_state.int_edit_index] = new_intervento
                st.session_state.int_edit_index = None
                for k in list(st.session_state.keys()):
                    if k.startswith("int_") and k != "int_edit_index":
                        try:
                            del st.session_state[k]
                        except:
                            pass
                icona_msg_edit = f" {new_intervento.get('IconaEmoji','')} {new_intervento.get('IconaNome','')}"
                if new_intervento.get("HasFile"):
                    icona_msg_edit += f" + file {new_intervento.get('FileName','')}"
                st.success(f"✅ Intervento aggiornato - {icona_msg_edit}")
                st.info(f"📋 Dettagli aggiornati: {new_intervento.get('Tipo','')} - {new_intervento.get('Squadra','')} - {new_intervento.get('Comune','')} - {new_intervento.get('Stato','')} - {new_intervento.get('Descrizione','')[:60]}")
                st.balloons()
                # Mostra tabella aggiornata prima di tornare - richiesta Ezio: vedere aggiornamento
                st.markdown("#### 📊 Tabella aggiornata - Torno in Tabella Emergenze tra 2 secondi...")
                df_updated = pd.DataFrame(st.session_state.interventi)
                # Rimuovi colonne interne per preview
                cols_remove_preview = ["StatoColoreBg", "StatoColoreTxt", "IconaLabel", "IconaNome", "IconaColore", "IconaTipo", "IconaEmoji", "FileBytes", "FileName", "HasFile"]
                df_preview = df_updated.copy()
                for c in cols_remove_preview:
                    if c in df_preview.columns:
                        df_preview = df_preview.drop(columns=[c])
                st.dataframe(df_preview.tail(5), use_container_width=True)
                # FIX richiesta Ezio: dopo aggiornamento salva e torna su tabella emergenze - con delay per vedere aggiornamento
                import time
                time.sleep(2)
                st.session_state.menu = "Tabella Emergenze"
                st.rerun()
            else:
                st.session_state.interventi.append(new_intervento)
                icona_msg = f" con icona {sel_ico_obj.get('Emoji','')} {sel_ico_obj.get('Nome','')}" if sel_ico_obj else ""
                if new_intervento.get("HasFile"):
                    icona_msg += f" + file {new_intervento.get('FileName','')}"
                st.success(f"Intervento salvato con stato {label} colorato {bg_color}{icona_msg}")
                st.balloons()
                st.rerun()
        else:
            st.error("Compila Descrizione Intervento *")

    if st.session_state.interventi:
        st.divider()
        st.markdown("**Interventi Salvati - con Icona all'inizio**")

        # Tabella interventi con icona all'inizio - senza colonne interne - richiesta Ezio
        df_int_full = pd.DataFrame(st.session_state.interventi)
        # Rimuovi colonne interne per visualizzazione
        cols_remove_int = ["StatoColoreBg", "StatoColoreTxt", "IconaLabel", "IconaNome", "IconaColore", "IconaTipo", "IconaEmoji", "FileBytes", "FileName", "HasFile"]
        df_int_display = df_int_full.copy()
        for c in cols_remove_int:
            if c in df_int_display.columns:
                df_int_display = df_int_display.drop(columns=[c])
        
        # Mostra con icona all'inizio
        st.markdown("**Interventi con Icona**")
        h1, h2, h3, h4, h5 = st.columns([0.8, 1, 1, 1.5, 2.5])
        with h1: st.markdown("**ICONA**")
        with h2: st.markdown("**Tipo**")
        with h3: st.markdown("**Squadra**")
        with h4: st.markdown("**Comune**")
        with h5: st.markdown("**Descrizione**")
        for _, r in df_int_full.iterrows():
            c1, c2, c3, c4, c5 = st.columns([0.8, 1, 1, 1.5, 2.5])
            with c1:
                emo = r.get("IconaEmoji","📍")
                nome = r.get("IconaNome","")
                fb = r.get("FileBytes")
                hf = r.get("HasFile")
                # Mostra immagine vera dell'icona assegnata - fix Ezio: vedeva scritta icona
                if hf and fb:
                    try:
                        st.image(fb, width=50, caption=nome[:10] if nome else "")
                    except Exception as e:
                        st.markdown(f"<div style='font-size:32px;text-align:center;background:#e8f5e9;border:2px solid #1A5D1A;border-radius:8px;padding:4px;'>{emo}</div>", unsafe_allow_html=True)
                else:
                    # Emoji grande visibile - non scritta icona
                    st.markdown(f"<div style='font-size:32px;text-align:center;background:white;border:2px solid #1A5D1A;border-radius:8px;padding:6px;'>{emo}</div>", unsafe_allow_html=True)
                    if nome:
                        st.caption(nome[:12])
            with c2: st.write(r.get("Tipo",""))
            with c3: st.write(r.get("Squadra",""))
            with c4: st.write(r.get("Comune",""))
            with c5: st.write(str(r.get("Descrizione",""))[:80])
        
        st.dataframe(df_int_display, use_container_width=True)
        st.download_button("Excel Interventi (pulito)", to_excel(df_int_display), "interventi.xlsx", use_container_width=True)
        if REPORTLAB_OK:
            st.download_button("PDF Logo Estesa", to_pdf(df_int_display, "INTERVENTI EMERGENZA"), "interventi.pdf", use_container_width=True)

# TABELLA INTERVENTI EMERGENZA
    # IMPORT/EXPORT INLINE - Ezio - TUTTI I FORM - Excel + PDF + Template ODV
    excel_import_inline("interventi", "Interventi Emergenza")


elif cur == "Tabella Interventi Emergenza":
    # Gestione click immagine come tasto via ?edit_idx2=
    try:
        qp2 = st.query_params
        if "edit_idx2" in qp2:
            idx_q2 = int(qp2.get("edit_idx2"))
            for k in [k for k in list(st.session_state.keys()) if k.startswith("int_") and k != "int_edit_index"]:
                try:
                    del st.session_state[k]
                except:
                    pass
            st.session_state.int_edit_index = idx_q2
            st.session_state.menu = "Interventi Emergenza"
            st.session_state["scroll_top"] = True
            del st.query_params["edit_idx2"]
            st.rerun()
    except:
        pass
    hdr_form("TABELLA INTERVENTI EMERGENZA - ICONA COME TASTO - GRIGLIA")

    st.markdown("""
    <div style="background:#e3f2fd;padding:10px;border-radius:8px;border-left:4px solid #1976d2;margin-bottom:12px;">
    <b>📋 Tabella Interventi Emergenza - Dati dal form Interventi Emergenza</b><br>
    Questa tabella legge direttamente da <b>Interventi Emergenza</b> (st.session_state.interventi) - Clicca icona per aprire scheda in maschera
    </div>
    """, unsafe_allow_html=True)

    # Debug - quanti interventi da Interventi Emergenza
    st.info(f"Debug: {len(st.session_state.interventi)} interventi da form Interventi Emergenza in memoria - Se 0, vai in Interventi Emergenza e crea intervento con icona")

    if not st.session_state.interventi:
        st.warning("⚠️ Nessun intervento in memoria - Vai in Interventi Emergenza, compila con icona e salva - Apparirà qui")
        st.markdown("""
        <div style="background:#e8f5e9;padding:12px;border-radius:8px;text-align:center;">
        <b>Come creare intervento:</b><br>
        1. Vai in <b>Interventi Emergenza</b><br>
        2. Scegli Tipo, Squadra, Data, Comune, Icona da Libreria<br>
        3. Compila Descrizione * e salva<br>
        4. Torna qui - tabella si crea da quel form
        </div>
        """, unsafe_allow_html=True)
        # Tabella vuota esempio
        st.dataframe(pd.DataFrame(columns=["Icona","Tipo","Squadra","Data","Comune","Stato","Descrizione"]).head(), use_container_width=True)
    else:
        # Tabella creata dal form Interventi Emergenza - OK richiesta Ezio
        st.success(f"✅ Tabella creata dal form Interventi Emergenza - {len(st.session_state.interventi)} interventi")
        df_tab = pd.DataFrame(st.session_state.interventi)

        # Filtri con variabili intermedie corrette parentesi chiuse
        squadre_list = sorted(list(set([str(x) for x in df_tab["Squadra"].tolist() if x])))
        comuni_list = sorted(list(set([str(x) for x in df_tab["Comune"].tolist() if x])))
        stati_list = sorted(list(set([str(x) for x in df_tab["Stato"].tolist() if x])))
        tipi_list = sorted(list(set([str(x) for x in df_tab["Tipo"].tolist() if x])))

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            filtro_squadra = st.selectbox(
                "Filtra Squadra",
                ["Tutte"] + squadre_list,
                key="tab_f_sq"
            )
        with c2:
            filtro_comune = st.selectbox(
                "Filtra Comune",
                ["Tutti"] + comuni_list,
                key="tab_f_com"
            )
        with c3:
            filtro_stato = st.selectbox(
                "Filtra Stato",
                ["Tutti"] + stati_list,
                key="tab_f_stato"
            )
        with c4:
            filtro_tipo = st.selectbox(
                "Filtra Tipo",
                ["Tutti"] + tipi_list,
                key="tab_f_tipo"
            )

        df_filtrato = df_tab.copy()

        if filtro_squadra != "Tutte":
            df_filtrato = df_filtrato[df_filtrato["Squadra"] == filtro_squadra]

        if filtro_comune != "Tutti":
            df_filtrato = df_filtrato[df_filtrato["Comune"] == filtro_comune]

        if filtro_stato != "Tutti":
            df_filtrato = df_filtrato[df_filtrato["Stato"] == filtro_stato]

        if filtro_tipo != "Tutti":
            df_filtrato = df_filtrato[df_filtrato["Tipo"] == filtro_tipo]

        st.write(f"Risultati filtrati: {len(df_filtrato)} su {len(df_tab)}")
        
        # RICHIESTA EZIO: togliere colonne interne e aggiungere colonna icona all'inizio con immagine
        # Colonne da togliere: stato colore, stato colore txt, icona nome, icona label, Icona nome, icona colore, icona tipo
        cols_to_remove = ["StatoColoreBg", "StatoColoreTxt", "IconaLabel", "IconaNome", "IconaColore", "IconaTipo", "IconaEmoji", "FileBytes", "FileName", "HasFile", "StatoColore", "Icona nome", "Icona label", "Icona colore", "Icona tipo"]
        
        st.markdown("#### 📋 Tabella Interventi con Icona assegnata all'inizio")
        
        # Header tabella
        h1, h2, h3, h4, h5, h6, h7 = st.columns([0.9, 1, 1, 1.2, 1.5, 1, 2])
        with h1: st.markdown("**ICONA**")
        with h2: st.markdown("**Tipo**")
        with h3: st.markdown("**Squadra**")
        with h4: st.markdown("**Data/Ora**")
        with h5: st.markdown("**Comune/Via**")
        with h6: st.markdown("**Stato**")
        with h7: st.markdown("**Descrizione**")
        st.divider()
        
        # Righe tabella con icona CLICCABILE - richiesta Ezio: cliccare su icona (non penna) per aprire form Interventi Emergenza
        for idx_f, (idx_orig, row) in enumerate(df_filtrato.iterrows()):
            c1, c2, c3, c4, c5, c6, c7 = st.columns([0.9, 1, 1, 1.2, 1.5, 1, 2])
            with c1:
                emoji = row.get("IconaEmoji", "📍")
                nome_ico = row.get("IconaNome", "")
                has_file = row.get("HasFile", False)
                file_bytes = row.get("FileBytes", None)
                # ICONA COME TASTO - Richiesta Ezio: icona diventi un tasto, non tasto sotto icona
                if has_file and file_bytes:
                    try:
                        b64 = base64.b64encode(file_bytes).decode()
                        html_btn = f"""
                        <a href="?edit_idx2={int(idx_orig)}" target="_self" style="text-decoration:none;">
                            <img src="data:image/png;base64,{b64}" width="60" height="60" 
                            style="border:3px solid #1A5D1A;border-radius:12px;cursor:pointer;box-shadow:0 2px 6px rgba(0,0,0,0.2);object-fit:contain;background:white;padding:3px;display:block;margin:auto;"
                            title="Clicca immagine {nome_ico} per aprire Interventi Emergenza">
                        </a>
                        <div style="text-align:center;font-size:10px;color:green;font-weight:bold;">Clicca immagine</div>
                        """
                        st.markdown(html_btn, unsafe_allow_html=True)
                    except:
                        # Fallback bottone emoji come tasto
                        if st.button(f"{emoji}", key=f"edit_int_icon_{idx_orig}", help=f"Clicca icona {nome_ico} per aprire", use_container_width=True):
                            idx_to_edit = int(idx_orig)
                            for k in [k for k in list(st.session_state.keys()) if k.startswith("int_") and k != "int_edit_index"]:
                                try:
                                    del st.session_state[k]
                                except:
                                    pass
                            st.session_state.int_edit_index = idx_to_edit
                            st.session_state.menu = "Interventi Emergenza"
                            st.session_state["scroll_top"] = True
                            st.rerun()
                else:
                    safe_emoji = emoji if emoji and emoji.strip() else "📍"
                    # Emoji stessa come tasto grande
                    if st.button(f"{safe_emoji}", key=f"edit_int_icon_{idx_orig}", help=f"Clicca icona {nome_ico} per aprire Interventi Emergenza", use_container_width=True):
                        idx_to_edit = int(idx_orig)
                        for k in [k for k in list(st.session_state.keys()) if k.startswith("int_") and k != "int_edit_index"]:
                            try:
                                del st.session_state[k]
                            except:
                                pass
                        st.session_state.int_edit_index = idx_to_edit
                        st.session_state.menu = "Interventi Emergenza"
                        st.session_state["scroll_top"] = True
                        st.rerun()
                if nome_ico:
                    st.caption(nome_ico[:12])
            with c2:
                st.write(row.get("Tipo",""))
            with c3:
                st.write(row.get("Squadra",""))
            with c4:
                st.write(f"{row.get('Data','')} {row.get('Ora','')}")
            with c5:
                st.write(f"{row.get('Comune','')} {row.get('Via','')}")
            with c6:
                bg = row.get("StatoColoreBg", "#e8f5e9")
                txt_c = row.get("StatoColoreTxt", "black")
                st.markdown(f"<span style='background:{bg};color:{txt_c};padding:2px 8px;border-radius:10px;font-weight:bold;border:1px solid black;font-size:11px;'>{row.get('Stato','')}</span>", unsafe_allow_html=True)
            with c7:
                st.write(str(row.get("Descrizione",""))[:100])
                # Bottone alternativo
                if st.button(f"📋 Vedi scheda", key=f"edit_int_text_{idx_orig}", use_container_width=True):
                    idx_to_edit = int(idx_orig)
                    for k in [k for k in list(st.session_state.keys()) if k.startswith("int_") and k != "int_edit_index"]:
                        try:
                            del st.session_state[k]
                        except:
                            pass
                    st.session_state.int_edit_index = idx_to_edit
                    st.session_state.menu = "Interventi Emergenza"
                    st.session_state["scroll_top"] = True
                    st.rerun()
            st.divider()
        
        # Prepara df per export senza colonne interne
        df_export = df_filtrato.copy()
        for col in cols_to_remove:
            if col in df_export.columns:
                df_export = df_export.drop(columns=[col])
        
        st.download_button(
            "Excel Filtrato (senza colonne interne)",
            to_excel(df_export),
            "tabella_interventi_filtrata.xlsx",
            use_container_width=True
        )
        if REPORTLAB_OK:
            st.download_button(
                "PDF Tabella",
                to_pdf(df_export, "TABELLA INTERVENTI EMERGENZA - CON ICONA"),
                "tabella_interventi.pdf",
                use_container_width=True
            )

# MEZZI
    # IMPORT/EXPORT INLINE - Ezio - TUTTI I FORM - Excel + PDF + Template ODV
    excel_import_inline("tabella_interventi", "Tabella Interventi Emergenza")


elif cur == "Mezzi":
    hdr_form("MEZZI - Parco Automezzi")

    c1, c2, c3 = st.columns(3)
    with c1:
        targa = st.text_input("Targa", key="mez_targa")
        modello_m = st.text_input("Modello Mezzo", key="mez_modello")
        tipo_m = st.selectbox("Tipo", ["Fuoristrada", "Furgone", "Autocarro", "Auto", "Moto"], key="mez_tipo")

    with c2:
        stato_m = st.selectbox("Stato Mezzo", ["Operativo", "In Manutenzione", "Fuori Servizio"], key="mez_stato")
        km = st.text_input("Km", key="mez_km")
        scadenza = st.date_input("Scadenza Revisione", value=date.today(), format="DD/MM/YYYY", key="mez_scad")

    with c3:
        note_mez = st.text_area("Note Mezzo", key="mez_note")
        foto_mez = st.file_uploader("Foto Mezzo", type=["jpg", "png"], key="mez_foto")

    def clear_mezzi():
        for k in ["mez_targa", "mez_modello", "mez_tipo", "mez_stato", "mez_km", "mez_scad", "mez_note", "mez_foto"]:
            if k in st.session_state:
                del st.session_state[k]
    c_mez_save1, c_mez_save2 = st.columns([3,1])
    with c_mez_save1:
        save_mez = st.button("💾 Salva Mezzo", type="primary", use_container_width=True, key="btn_salva_mezzi")
    with c_mez_save2:
        st.button("🔄 Pulisci maschera", use_container_width=True, key="btn_pulisci_mezzi", on_click=clear_mezzi)
    if save_mez:
        if targa:
            st.session_state.mezzi.append({
                "Targa": targa,
                "Modello": modello_m,
                "Tipo": tipo_m,
                "Stato": stato_m,
                "Km": km,
                "Scadenza": str(scadenza),
                "Note": note_mez
            })
            st.success("Mezzo salvato")
            st.rerun()

    if st.session_state.mezzi:
        df_mez = pd.DataFrame(st.session_state.mezzi)
        st.dataframe(df_mez, use_container_width=True)
        st.download_button("Excel Mezzi", to_excel(df_mez), "mezzi.xlsx", use_container_width=True)
        if REPORTLAB_OK:
            st.download_button("PDF Logo Estesa", to_pdf(df_mez, "MEZZI"), "mezzi.pdf", use_container_width=True)

# ATTREZZATURE


    # IMPORT/EXPORT INLINE - Ezio - TUTTI I FORM - Excel + PDF + Template ODV
    excel_import_inline("mezzi", "Mezzi")


elif cur == "Attrezzature":
    hdr_form("ATTREZZATURE - Magazzino")

    c1, c2 = st.columns(2)
    with c1:
        nome_att = st.text_input("Nome Attrezzatura", key="att_nome")
        cat_att = st.selectbox("Categoria", ["DPI", "Utensili", "Elettrico", "Idraulico", "Altro"], key="att_cat")
        qta_att = st.number_input("Quantità", min_value=1, value=1, key="att_qta")

    with c2:
        stato_att = st.selectbox("Stato", ["Disponibile", "In Uso", "Guasto", "Esaurito"], key="att_stato")
        ubic_att = st.text_input("Ubicazione Magazzino", key="att_ubic")
        note_att = st.text_area("Note", key="att_note")

    def clear_att():
        for k in ["att_nome", "att_cat", "att_qta", "att_stato", "att_ubic", "att_note"]:
            if k in st.session_state:
                del st.session_state[k]

    c_att1, c_att2 = st.columns([3,1])
    with c_att1:
        save_att = st.button("💾 Salva Attrezzatura", type="primary", use_container_width=True, key="btn_salva_att")
    with c_att2:
        st.button("🔄 Pulisci maschera", use_container_width=True, key="btn_pulisci_att", on_click=clear_att)
    if save_att:
        if nome_att:
            st.session_state.attrezzature.append({
                "Nome": nome_att,
                "Categoria": cat_att,
                "Quantita": qta_att,
                "Stato": stato_att,
                "Ubicazione": ubic_att,
                "Note": note_att
            })
            st.success("Attrezzatura salvata")
            st.rerun()

    if st.session_state.attrezzature:
        df_att = pd.DataFrame(st.session_state.attrezzature)
        st.dataframe(df_att, use_container_width=True)
        st.download_button("Excel Attrezzature", to_excel(df_att), "attrezzature.xlsx", use_container_width=True)
        if REPORTLAB_OK:
            st.download_button("PDF Logo Estesa", to_pdf(df_att, "ATTREZZATURE"), "attrezzature.pdf", use_container_width=True)

# MAPPE POSTAZIONI - STABILE - MARKER RIMANGONO - TABELLA SOTTO - ANTEPRIMA SOTTO TABELLA - NOME EMERGENZA/EVENTO COMBO


    # IMPORT/EXPORT INLINE - Ezio - TUTTI I FORM - Excel + PDF + Template ODV
    excel_import_inline("attrezzature", "Attrezzature")


elif cur == "Mappe Postazioni":
    hdr_form("MAPPE POSTAZIONI")

    if "mappa_avanzata_markers" not in st.session_state:
        st.session_state.mappa_avanzata_markers = []
    if "last_clicked_lat" not in st.session_state:
        st.session_state.last_clicked_lat = ""
    if "last_clicked_lon" not in st.session_state:
        st.session_state.last_clicked_lon = ""
    if "map_focus" not in st.session_state:
        st.session_state.map_focus = None
    if "selected_icon_label" not in st.session_state:
        st.session_state.selected_icon_label = ""

    emergenze_list = st.session_state.get("emergenze", [])
    eventi_list = st.session_state.get("eventi", [])
    nomi_emergenze = ["-- Nessuna --"] + [f"{e.get('Nome','')} - {e.get('Data','')}" for e in emergenze_list[-20:]] if emergenze_list else ["-- Nessuna --"]
    nomi_eventi = ["-- Nessuno --"] + [f"{ev.get('Nome','')} - {ev.get('Data','')}" for ev in eventi_list[-20:]] if eventi_list else ["-- Nessuno --"]

    c1, c2, c3, c4 = st.columns([1,1,1,2])
    with c1:
        if st.button("⛶ Fullscreen", key="btn_fs_mappa", use_container_width=True, type="primary"):
            st.session_state["fs_mappa_active"] = True
    with c2:
        if st.button("🧹 Pulisci TUTTI", key="btn_clear_all_markers", use_container_width=True):
            st.session_state.mappa_avanzata_markers = []
            st.session_state.map_focus = None
            st.success("Tutti i marker rimossi")
            st.rerun()
    with c3:
        if st.button("🎯 Mostra tutte", key="btn_fit_all", use_container_width=True):
            st.session_state.map_focus = None
            st.rerun()
    with c4:
        st.markdown(f'<span style="background:#1A5D1A;color:white;padding:6px 12px;border-radius:6px;font-weight:bold;">🗺️ {len(st.session_state.mappa_avanzata_markers)} postazioni</span>', unsafe_allow_html=True)

    icone_disponibili = st.session_state.get("icone", [])
    # Icone di default tolte - le carichi tu in Libreria Icone - richiesta Ezio - etichette rimosse
    if not icone_disponibili:
        # Etichetta Libreria Icone vuota rimossa - richiesta Ezio - non mostrare
        icone_disponibili = []
    else:
        st.caption(f"📚 Libreria: {len(icone_disponibili)} icone")
    icone_options = []
    icone_map = {}
    # Se libreria vuota, aggiungi placeholder per evitare errori ma non mostra icone di default
    if not icone_disponibili:
        icone_disponibili_placeholder = []
    for ico in icone_disponibili:
        has_file = ico.get("HasFile", False)
        file_info = " [FILE]" if has_file else ""
        label = f"{ico.get('Emoji','📍')} {ico.get('Nome','')} - {ico.get('Tipo','')} ({ico.get('Colore','')}){file_info}"
        icone_options.append(label)
        icone_map[label] = ico
    if not st.session_state.selected_icon_label and icone_options:
        st.session_state.selected_icon_label = icone_options[0]

    # Griglia Scegli icona dalla Libreria tolta definitivamente - richiesta Ezio
    # Solo combo Icona Libreria in maschera, caricata dal form Libreria Icone
    try:
        default_idx = icone_options.index(st.session_state.selected_icon_label) if st.session_state.selected_icon_label in icone_options else 0
    except:
        default_idx = 0
    # Se libreria vuota, default_idx 0 ma icone_options vuoto - gestito sotto

    st.markdown("#### 📍 Maschera Postazione")
    # Alias da form Alias Radio - per assegnarlo alla postazione - richiesta Ezio - RIGA 4207
    alias_form_list = [a.get('Alias','').strip() for a in st.session_state.get('alias_radio', []) if a.get('Alias','').strip()]
    alias_form_list = sorted(list(set(alias_form_list)))
    alias_postazioni = [m.get('Alias','') for m in st.session_state.get('mappa_avanzata_markers', []) if m.get('Alias')]
    alias_esistenti = sorted(list(set(alias_form_list + alias_postazioni)))
    if alias_form_list:
        st.info(f"📻 Alias da form Alias Radio ({len(alias_form_list)}): {', '.join(alias_form_list[:20])} - Assegnalo alla postazione")
    elif alias_esistenti:
        st.caption(f"📻 Alias già usati in postazioni: {', '.join(alias_esistenti[:15])}")

    # FIX DEFINITIVO Pulisci maschera + no default prima posizione marker - versione con versionamento chiavi
    # Inizializza versione maschera se non esiste
    if "map_form_version" not in st.session_state:
        st.session_state["map_form_version"] = 0
    
    # Se flag pulisci, incrementa versione e cancella tutto
    if st.session_state.get("do_clear_maschera"):
        st.session_state["map_form_version"] += 1
        # Cancella tutte le chiavi last_clicked e query params
        for k in list(st.session_state.keys()):
            if k.startswith("adv_marker_") or k.startswith("last_clicked_") or k.startswith("adv_alias_") or k.startswith("nome_emergenza") or k.startswith("nome_evento"):
                try:
                    del st.session_state[k]
                except:
                    pass
        try:
            st.query_params.clear()
        except:
            pass
        st.session_state["do_clear_maschera"] = False
        st.rerun()
    
    # Versione corrente per chiavi - così pulisci cambia tutte le chiavi e maschera si svuota davvero
    ver = st.session_state.get("map_form_version", 0)
    def k_map(base):
        return f"{base}_v{ver}"
    
    # Pre-fill lat/lon/comune/via da mappa - SOLO se appena cliccato mappa, NON di default prima posizione
    # Se non hai appena cliccato, non riempire con vecchia posizione
    # Solo se last_clicked esiste e campo vuoto
    if not st.session_state.get("do_clear_maschera"):
        # Non riempire di default la prima posizione marker - solo se last_clicked è recente e campo vuoto
        if st.session_state.get("last_clicked_lat"):
            # Solo se adv_marker_lat con versione corrente non esiste
            key_lat = k_map("adv_marker_lat")
            if not st.session_state.get(key_lat):
                st.session_state[key_lat] = str(st.session_state["last_clicked_lat"])
        if st.session_state.get("last_clicked_lon"):
            key_lon = k_map("adv_marker_lon")
            if not st.session_state.get(key_lon):
                st.session_state[key_lon] = str(st.session_state["last_clicked_lon"])
        if st.session_state.get("last_clicked_comune"):
            key_com = k_map("adv_marker_comune")
            if not st.session_state.get(key_com):
                st.session_state[key_com] = str(st.session_state["last_clicked_comune"])
        if st.session_state.get("last_clicked_via"):
            key_via = k_map("adv_marker_via")
            if not st.session_state.get(key_via):
                st.session_state[key_via] = str(st.session_state["last_clicked_via"])

    c1, c2, c3 = st.columns(3)
    with c1:
        marker_nome = st.text_input("Nome Postazione *", key=k_map("adv_marker_nome"), placeholder="Es: Postazione 1")
        # Alias combo da form Alias Radio
        alias_options = ["-- Nessun Alias --"] + alias_form_list
        if alias_postazioni and not alias_form_list:
            alias_options = ["-- Nessun Alias --"] + alias_esistenti
        alias_options = alias_options + ["-- Nuovo Alias --"]
        alias_options_unique = []
        for o in alias_options:
            if o not in alias_options_unique:
                alias_options_unique.append(o)
        alias_options = alias_options_unique
        sel_alias_combo = st.selectbox("Alias (combo) da form Alias Radio", alias_options, key=k_map("adv_alias_combo"), help="Seleziona alias da form Alias Radio")
        if sel_alias_combo == "-- Nuovo Alias --":
            marker_alias = st.text_input("Nuovo Alias", key=k_map("adv_marker_alias_new"), placeholder="Es: Alfa 1, Base, 01")
        elif sel_alias_combo == "-- Nessun Alias --":
            marker_alias = ""
        else:
            marker_alias = sel_alias_combo
        # Lat/Lon - senza etichette via - richiesta Ezio - togli etichette via su form map
        marker_lat = st.text_input("Latitudine *", key=k_map("adv_marker_lat"), placeholder="Clicca mappa - es: 45.8167")
        marker_lon = st.text_input("Longitudine *", key=k_map("adv_marker_lon"), placeholder="Clicca mappa - es: 8.8333")
    with c2:
        # Comune e Via - senza etichette extra via su form map - richiesta Ezio
        marker_comune = st.text_input("Comune *", key=k_map("adv_marker_comune"), placeholder="Es: Varese")
        marker_via = st.text_input("Via *", key=k_map("adv_marker_via"), placeholder="Via + civico - Es: Via Rossi 10")
        nome_emergenza = st.selectbox("Nome Emergenza (combo)", nomi_emergenze, index=0, key=k_map("nome_emergenza_combo"))
        nome_evento = st.selectbox("Nome Evento (combo)", nomi_eventi, index=0, key=k_map("nome_evento_combo"))
    with c3:
        # Carica icone dal form Libreria Icone
        if icone_options:
            marker_icona_label = st.selectbox("Icona Libreria (caricate da te)", icone_options, index=default_idx, key=k_map("adv_marker_icona_select"))
            st.session_state.selected_icon_label = marker_icona_label
            selected_ico_obj = icone_map.get(marker_icona_label, {"Emoji":"👷","Nome":"Postazione","Colore":"green","Tipo":"Postazione"})
        else:
            # Etichetta rimossa - richiesta Ezio - togli queste etichette anche negli altri form
            marker_icona_label = ""
            selected_ico_obj = {"Emoji":"👷","Nome":"Postazione","Colore":"green","Tipo":"Postazione","HasFile":False}
            st.session_state.selected_icon_label = ""
        has_file_sel = selected_ico_obj.get("HasFile", False)
        file_bytes_sel = selected_ico_obj.get("FileBytes")
        if has_file_sel and file_bytes_sel:
            try:
                st.image(file_bytes_sel, width=90, caption=f"{selected_ico_obj.get('Nome','')} - FILE")
            except:
                st.markdown(f"<div style='font-size:32px;text-align:center;background:#e8f5e9;padding:10px;border-radius:10px;border:2px solid #1A5D1A;'>{selected_ico_obj.get('Emoji','👷')}<br><small>{selected_ico_obj.get('Nome','')}</small></div>", unsafe_allow_html=True)
        else:
            if icone_options:
                st.markdown(f"<div style='font-size:32px;text-align:center;background:#e8f5e9;padding:10px;border-radius:10px;border:2px solid #1A5D1A;'>{selected_ico_obj.get('Emoji','👷')}<br><small>{selected_ico_obj.get('Nome','Postazione')}</small></div>", unsafe_allow_html=True)
            else:
                # Etichetta rimossa - richiesta Ezio
                pass
        marker_tipo = st.selectbox("Tipo", ["Postazione", "Emergenza", "Evento", "Mezzo", "Volontario"], key=k_map("adv_marker_tipo"))
        marker_desc = st.text_input("Descrizione", key=k_map("adv_marker_desc"))

    col_save1, col_save2 = st.columns([3,1])
    with col_save1:
        save_clicked = st.button("💾 SALVA POSTAZIONE", type="primary", use_container_width=True, key=f"btn_salva_postazione_v{ver}")
    with col_save2:
        # Pulisci maschera - FIX definitivo con versionamento chiavi - pulisce davvero
        if st.button("🧹 Pulisci maschera", type="primary", use_container_width=True, key=f"btn_pulisci_maschera_v{ver}", help="Pulisce maschera postazione - cancella tutti i campi - FIX definitivo"):
            st.session_state["do_clear_maschera"] = True
            try:
                st.query_params.clear()
            except:
                pass
            st.rerun()


    try:
        qp_lat = st.query_params.get("lat", "")
        qp_lon = st.query_params.get("lon", "")
        qp_comune = st.query_params.get("comune", "")
        qp_via = st.query_params.get("via", "")
        if qp_lat and qp_lon:
            st.session_state.last_clicked_lat = str(qp_lat)
            st.session_state.last_clicked_lon = str(qp_lon)
            # Se arrivano anche comune/via da JS, usali
            if qp_comune:
                st.session_state.last_clicked_comune = str(qp_comune)
            if qp_via:
                st.session_state.last_clicked_via = str(qp_via)
            # Se non arrivano, fai reverse geocode qui per comune/via
            if not qp_comune or not qp_via:
                try:
                    import requests as req_geo
                    r = req_geo.get(f"https://nominatim.openstreetmap.org/reverse?format=json&lat={qp_lat}&lon={qp_lon}", headers={"User-Agent": "ANA-Varese"}, timeout=3)
                    if r.status_code == 200:
                        d = r.json()
                        addr = d.get("address", {})
                        com = addr.get("city") or addr.get("town") or addr.get("village") or addr.get("municipality") or ""
                        via = addr.get("road") or ""
                        if com and not st.session_state.get("last_clicked_comune"):
                            st.session_state.last_clicked_comune = com
                        if via and not st.session_state.get("last_clicked_via"):
                            st.session_state.last_clicked_via = via
                except:
                    pass
    except:
        pass

    if save_clicked:
        eff_lat = marker_lat or st.session_state.get("last_clicked_lat") or ""
        eff_lon = marker_lon or st.session_state.get("last_clicked_lon") or ""
        if not marker_nome:
            st.error("❌ Inserisci Nome Postazione")
        elif not eff_lat or not eff_lon:
            st.error("❌ Manca Latitudine o Longitudine - Clicca mappa")
        else:
            try:
                lat_f = float(str(eff_lat).replace(",", "."))
                lon_f = float(str(eff_lon).replace(",", "."))
                sel_obj = icone_map.get(st.session_state.selected_icon_label, selected_ico_obj)
                nuovo = {
                    "Nome": marker_nome,
                    "Alias": marker_alias,
                    "Lat": lat_f,
                    "Lon": lon_f,
                    "Comune": marker_comune,
                    "Via": marker_via,
                    "NomeEmergenza": nome_emergenza,
                    "NomeEvento": nome_evento,
                    "Emoji": sel_obj.get("Emoji","👷"),
                    "Colore": sel_obj.get("Colore","green"),
                    "IconaNome": sel_obj.get("Nome","Postazione"),
                    "Icona": st.session_state.selected_icon_label,
                    "Tipo": marker_tipo,
                    "Descrizione": marker_desc,
                    "DataIns": datetime.now().strftime("%d/%m/%Y %H:%M"),
                    "HasFile": sel_obj.get("HasFile", False),
                    "FileBytes": sel_obj.get("FileBytes"),
                    "FileName": sel_obj.get("FileName",""),
                    "IconaTipo": sel_obj.get("Tipo","Postazione")
                }
                st.session_state.mappa_avanzata_markers.append(nuovo)
                st.session_state.map_focus = nuovo
                # Pulisci last_clicked dopo salvataggio per non tenere default prima posizione marker
                for k in ["last_clicked_lat", "last_clicked_lon", "last_clicked_comune", "last_clicked_via"]:
                    if k in st.session_state:
                        try:
                            del st.session_state[k]
                        except:
                            pass
                st.success(f"✅ SALVATA {marker_nome} - Totale {len(st.session_state.mappa_avanzata_markers)}")
                try:
                    st.query_params.clear()
                except:
                    pass
                # Incrementa versione per pulire maschera dopo salvataggio
                st.session_state["map_form_version"] = st.session_state.get("map_form_version", 0) + 1
                st.rerun()
            except Exception as e:
                st.error(f"❌ Errore: {e}")

    st.divider()
    st.markdown("#### 🗺️ Anteprima - Tutti i marker visibili - +/- sotto")
    all_markers = st.session_state.get("mappa_avanzata_markers", [])
    focus_marker = st.session_state.get("map_focus")
    import json as json_lib
    def clean_marker_for_json(m):
        b64 = ""
        if m.get("HasFile") and m.get("FileBytes"):
            try:
                import base64 as b64lib
                fb = m.get("FileBytes")
                if isinstance(fb, bytes):
                    b64 = b64lib.b64encode(fb).decode()
            except:
                b64 = ""
        return {"lat": m.get("Lat"), "lon": m.get("Lon"), "nome": m.get("Nome",""), "alias": m.get("Alias",""), "emoji": m.get("Emoji","👷"), "colore": m.get("Colore","green"), "hasFile": m.get("HasFile", False), "fileB64": b64, "comune": m.get("Comune",""), "via": m.get("Via",""), "emergenza": m.get("NomeEmergenza",""), "evento": m.get("NomeEvento","")}
    markers_for_js = json_lib.dumps([clean_marker_for_json(m) for m in all_markers])
    if focus_marker:
        focus_clean = clean_marker_for_json(focus_marker)
        focus_clean["Lat"] = focus_marker.get("Lat")
        focus_clean["Lon"] = focus_marker.get("Lon")
        focus_for_js = json_lib.dumps(focus_clean)
    else:
        focus_for_js = "null"

    # ANTEPRIMA - marker fissati qui e visionabili in mappa grande - lat/lon automatico in maschera senza bottone
    preview_html = """
    <div style="border:3px solid #1A5D1A;border-radius:8px;overflow:hidden;">
    <div style="background:#1A5D1A;color:white;padding:6px;text-align:center;">ANTEPRIMA - Clicca qui per mettere marker - Si fissano e si vedono in mappa grande - """ + str(len(all_markers)) + """ marker</div>
    <div id="preview_map_top" style="height:400px;width:100%;"></div>
    </div>
    <div id="preview_coords" style="background:#fffde7;padding:6px;border-radius:6px;margin-top:6px;font-weight:bold;border-left:4px solid #FFD700;">📍 Clicca anteprima per aggiungere marker - Lat/Lon vanno in maschera automatico</div>
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    <script>
    var markersPreview = MARKERS_PREVIEW_PLACEHOLDER;
    var focusPreview = FOCUS_PREVIEW_PLACEHOLDER;
    var selEmojiPrev = SELECTED_EMOJI_PREVIEW;
    var selColorPrev = SELECTED_COLOR_PREVIEW;
    var selIconHasFilePrev = SELECTED_HASFILE_PREVIEW;
    var selIconFileB64Prev = SELECTED_FILEB64_PREVIEW;
    var pMap = L.map('preview_map_top', {zoomControl: false}).setView([45.8167, 8.8333], 12);
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png').addTo(pMap);
    L.control.zoom({position: 'bottomleft'}).addTo(pMap);
    var FsTop = L.Control.extend({onAdd: function(m){
        var cEl=L.DomUtil.create('div','leaflet-bar leaflet-control');
        cEl.style.background='white';cEl.style.width='34px';cEl.style.height='34px';cEl.style.lineHeight='34px';cEl.style.textAlign='center';cEl.style.cursor='pointer';cEl.style.fontSize='22px';cEl.style.fontWeight='bold';cEl.style.border='2px solid rgba(0,0,0,0.2)';cEl.style.borderRadius='4px';
        cEl.innerHTML='⛶';cEl.title='Fullscreen 100% tutto schermo - come prima';
        cEl.onclick=function(){
            var mapCont = document.getElementById('preview_map_top');
            var parentCont = document.getElementById('preview_map_top').parentElement;
            try{
                if(!document.fullscreenElement){
                    if(parentCont.requestFullscreen) parentCont.requestFullscreen();
                    else if(document.documentElement.requestFullscreen) document.documentElement.requestFullscreen();
                    try{var pd=window.parent.document; if(pd.documentElement.requestFullscreen) pd.documentElement.requestFullscreen();}catch(e){}
                    setTimeout(function(){
                        parentCont.style.width='100vw';parentCont.style.height='100vh';parentCont.style.position='fixed';parentCont.style.top='0';parentCont.style.left='0';parentCont.style.zIndex='9999';parentCont.style.background='white';
                        mapCont.style.height='100vh';mapCont.style.width='100vw';
                        m.invalidateSize();
                    },100);
                } else {
                    if(document.exitFullscreen) document.exitFullscreen();
                    try{var pd=window.parent.document; if(pd.exitFullscreen) pd.exitFullscreen();}catch(e){}
                    setTimeout(function(){
                        parentCont.style.width='100%';parentCont.style.height='';parentCont.style.position='';parentCont.style.top='';parentCont.style.left='';parentCont.style.zIndex='';parentCont.style.background='';
                        mapCont.style.height='400px';mapCont.style.width='100%';
                        m.invalidateSize();
                    },100);
                }
            }catch(err){}
            setTimeout(function(){m.invalidateSize();},600);
        };return cEl;
    }}); new FsTop({position: 'bottomleft'}).addTo(pMap);
    document.addEventListener('fullscreenchange', function(){
        setTimeout(function(){ pMap.invalidateSize(); }, 600);
    });
    function getColorCodePrev(c){var m={'red':'#d32f2f','blue':'#1976d2','green':'#388e3c','orange':'#f57c00','purple':'#7b1fa2'};return m[c]||'#388e3c';}
    var allPrev=[];
    markersPreview.forEach(function(md){
        // FIX: Vedi tutti i marker nelle mappe - richiesta Ezio - non solo con file
        var mk;
        if(md.hasFile && md.fileB64){
            var ic = L.icon({iconUrl:'data:image/png;base64,'+md.fileB64,iconSize:[36,36],iconAnchor:[18,18]});
            mk=L.marker([md.lat,md.lon],{icon:ic}).addTo(pMap).bindPopup("<b>📍 "+md.nome+"</b><br>📻 Alias Radio: <b>"+(md.alias||"--")+"</b> - Chiamata<br>🏙️ Comune: "+md.comune+"<br>📍 Via: "+md.via+"<br>Lat: "+md.lat+" Lon: "+md.lon+"<br><small>Emergenza: "+(md.emergenza||"--")+" | Evento: "+(md.evento||"--")+"</small>");
        } else {
            // Fallback emoji/colore se non ha file - mostra comunque marker
            var colorMap = {'red':'#d32f2f','blue':'#1976d2','green':'#388e3c','orange':'#f57c00','purple':'#7b1fa2'};
            var col = colorMap[md.colore] || '#388e3c';
            var divIcon = L.divIcon({html:"<div style='background:white;border:2px solid "+col+";width:32px;height:32px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:18px;box-shadow:0 2px 4px rgba(0,0,0,0.3);'>"+(md.emoji||"📍")+"</div>",iconSize:[32,32],iconAnchor:[16,16]});
            mk=L.marker([md.lat,md.lon],{icon:divIcon}).addTo(pMap).bindPopup("<b>📍 "+md.nome+"</b><br>📻 Alias Radio: <b>"+(md.alias||"--")+"</b> - Chiamata<br>🏙️ Comune: "+md.comune+"<br>📍 Via: "+md.via+"<br>Lat: "+md.lat+" Lon: "+md.lon+"<br><small>Emergenza: "+(md.emergenza||"--")+" | Evento: "+(md.evento||"--")+"</small>");
        }
        allPrev.push(mk);
    });
    // Se c'è focus da vedi su mappa, ingrandisce su mappa grande (non qui) - solo centra
    if(focusPreview && focusPreview.Lat){
        pMap.setView([focusPreview.Lat, focusPreview.Lon], 17);
        if(focusPreview.hasFile && focusPreview.fileB64){
            var focusIcon = L.icon({iconUrl: 'data:image/png;base64,'+focusPreview.fileB64, iconSize: [44,44], iconAnchor: [22,22]});
            L.marker([focusPreview.Lat, focusPreview.Lon], {icon: focusIcon}).addTo(pMap).bindPopup("<b>📍 "+focusPreview.nome+"</b>").openPopup();
        }
        document.getElementById('preview_coords').innerHTML = "📍 Focus: "+focusPreview.nome+" - "+focusPreview.comune+" "+focusPreview.via;
    } else if(allPrev.length>0){
        var g=L.featureGroup(allPrev);pMap.fitBounds(g.getBounds().pad(0.4));
    }
    // Click su anteprima per mettere marker fissi - icona scelta non cerchio giallo - richiesta Ezio
    pMap.on('click', function(e){
        var lat = e.latlng.lat.toFixed(6);
        var lon = e.latlng.lng.toFixed(6);
        var tmpIcon;
        if(selIconHasFilePrev && selIconFileB64Prev){
            tmpIcon = L.icon({iconUrl: 'data:image/png;base64,'+selIconFileB64Prev, iconSize: [38,38], iconAnchor: [19,19]});
        } else {
            tmpIcon = L.divIcon({html:"<div style='background:white;border:2px solid "+getColorCodePrev(selColorPrev)+";width:32px;height:32px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:18px;box-shadow:0 2px 4px rgba(0,0,0,0.3);'>"+selEmojiPrev+"</div>",iconSize:[32,32],iconAnchor:[16,16]});
        }
        L.marker([lat, lon], {icon: tmpIcon}).addTo(pMap).bindPopup("Nuova "+selEmojiPrev+"<br>"+lat+","+lon).openPopup();
        document.getElementById('preview_coords').innerHTML = "📍 Nuova da anteprima: "+selEmojiPrev+" Lat: "+lat+" Lon: "+lon+" - Inserita in maschera automatico";
        // FIX: Quando metti marker deve compilarmi maschera con lat long comune e via - richiesta Ezio
        var currentComune = "";
        var currentVia = "";
        try{
            fetch('https://nominatim.openstreetmap.org/reverse?format=json&lat='+lat+'&lon='+lon)
                .then(r=>r.json()).then(d=>{
                    var com = d.address.city||d.address.town||d.address.village||"";
                    var via = d.address.road||"";
                    currentComune = com;
                    currentVia = via;
                    document.getElementById('preview_coords').innerHTML = "📍 Nuova: Lat "+lat+" Lon "+lon+"<br>Comune: "+com+" Via: "+via+" - Compilo maschera automatico";
                    try{
                        var url = new URL(window.parent.location.href);
                        url.searchParams.set('lat', lat);
                        url.searchParams.set('lon', lon);
                        url.searchParams.set('comune', com);
                        url.searchParams.set('via', via);
                        window.parent.history.replaceState(null, '', url.toString());
                    }catch(e){}
                });
        }catch(e){
            try{
                var url = new URL(window.parent.location.href);
                url.searchParams.set('lat', lat);
                url.searchParams.set('lon', lon);
                window.parent.history.replaceState(null, '', url.toString());
            }catch(err){}
        }
        // Fallback immediato lat/lon
        try{
            var url = new URL(window.parent.location.href);
            url.searchParams.set('lat', lat);
            url.searchParams.set('lon', lon);
            window.parent.history.replaceState(null, '', url.toString());
        }catch(err){}
        // Marker spostabile
        var draggableMarker = L.marker([lat, lon], {icon: tmpIcon, draggable: true}).addTo(pMap);
        draggableMarker.on('dragend', function(ev){
            var newLat = ev.target.getLatLng().lat.toFixed(6);
            var newLon = ev.target.getLatLng().lng.toFixed(6);
            document.getElementById('preview_coords').innerHTML = "📍 Marker spostato - Lat: "+newLat+" Lon: "+newLon;
            try{
                var url2 = new URL(window.parent.location.href);
                url2.searchParams.set('lat', newLat);
                url2.searchParams.set('lon', newLon);
                url2.searchParams.set('comune', currentComune);
                url2.searchParams.set('via', currentVia);
                window.parent.history.replaceState(null, '', url2.toString());
            }catch(e){}
        });
    });
    </script>
    """
    import json as json_lib_prev
    import base64 as b64lib_prev
    sel_e_prev = selected_ico_obj.get('Emoji','👷')
    sel_c_prev = selected_ico_obj.get('Colore','green')
    sel_hasfile_prev = selected_ico_obj.get("HasFile", False)
    sel_fileb64_prev = ""
    if sel_hasfile_prev and selected_ico_obj.get("FileBytes"):
        try:
            fb = selected_ico_obj.get("FileBytes")
            if isinstance(fb, bytes):
                sel_fileb64_prev = b64lib_prev.b64encode(fb).decode()
        except:
            sel_fileb64_prev = ""
    preview_html = preview_html.replace("MARKERS_PREVIEW_PLACEHOLDER", markers_for_js)
    preview_html = preview_html.replace("FOCUS_PREVIEW_PLACEHOLDER", focus_for_js)
    preview_html = preview_html.replace("SELECTED_EMOJI_PREVIEW", json_lib_prev.dumps(sel_e_prev))
    preview_html = preview_html.replace("SELECTED_COLOR_PREVIEW", json_lib_prev.dumps(sel_c_prev))
    preview_html = preview_html.replace("SELECTED_HASFILE_PREVIEW", json_lib_prev.dumps(sel_hasfile_prev))
    preview_html = preview_html.replace("SELECTED_FILEB64_PREVIEW", json_lib_prev.dumps(sel_fileb64_prev))
    st.components.v1.html(preview_html, height=450)

    st.divider()
    st.markdown("#### 🌍 Mappa Grande - Tutti i marker rimangono")

    html_code = """
    <div id="map-container" style="position:relative; background:white; border-radius:12px;">
        <div id="map" style="height:650px; width:100%; border-radius:12px; border:3px solid #1A5D1A;"></div>
    </div>
    <div id="coords" style="background:#fffde7;padding:8px;border-radius:6px;margin-top:8px;font-weight:bold;border-left:4px solid #FFD700;">📍 Clicca per aggiungere marker</div>
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    <script>
    var markersData = MARKERS_JSON_PLACEHOLDER;
    var focusMarker = FOCUS_JSON_PLACEHOLDER;
    var selectedIconEmoji = SELECTED_EMOJI_PLACEHOLDER;
    var selectedIconColor = SELECTED_COLOR_PLACEHOLDER;
    var selectedIconHasFile = SELECTED_HASFILE_PLACEHOLDER;
    var selectedIconFileB64 = SELECTED_FILEB64_PLACEHOLDER;
    var map = L.map('map', {zoomControl: false}).setView([45.8167, 8.8333], 13);
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png').addTo(map);
    L.control.zoom({position: 'bottomleft'}).addTo(map);
    var FullscreenControl = L.Control.extend({
        onAdd: function(map) {
            var container = L.DomUtil.create('div', 'leaflet-bar leaflet-control');
            container.style.backgroundColor = 'white';
            container.style.width = '34px';
            container.style.height = '34px';
            container.style.lineHeight = '34px';
            container.style.textAlign = 'center';
            container.style.cursor = 'pointer';
            container.style.fontSize = '22px';
            container.style.fontWeight = 'bold';
            container.style.border = '2px solid rgba(0,0,0,0.2)';
            container.style.borderRadius = '4px';
            container.innerHTML = '⛶';
            container.title = 'Fullscreen 100% tutto schermo - come prima faceva';
            container.onclick = function(){
                var mapContainer = document.getElementById('map');
                var parentContainer = document.getElementById('map-container');
                try{
                    if (!document.fullscreenElement) {
                        // Prova tutto schermo vero 100% - come prima faceva
                        if (parentContainer.requestFullscreen) {
                            parentContainer.requestFullscreen();
                        } else if (document.documentElement.requestFullscreen) {
                            document.documentElement.requestFullscreen();
                        } else if (mapContainer.requestFullscreen) {
                            mapContainer.requestFullscreen();
                        }
                        // Fallback parent window Streamlit
                        try{
                            var parentDoc = window.parent.document;
                            var iframe = parentDoc.querySelector('iframe[title*="st.components"]');
                            if(iframe && iframe.requestFullscreen) iframe.requestFullscreen();
                            else if(parentDoc.documentElement.requestFullscreen) parentDoc.documentElement.requestFullscreen();
                        }catch(e){}
                        setTimeout(function(){
                            parentContainer.style.width = '100vw';
                            parentContainer.style.height = '100vh';
                            parentContainer.style.position = 'fixed';
                            parentContainer.style.top = '0';
                            parentContainer.style.left = '0';
                            parentContainer.style.zIndex = '9999';
                            parentContainer.style.background = 'white';
                            mapContainer.style.height = '100vh';
                            mapContainer.style.width = '100vw';
                            mapContainer.style.borderRadius = '0';
                            mapContainer.style.border = 'none';
                            map.invalidateSize();
                        }, 100);
                    } else {
                        if (document.exitFullscreen) document.exitFullscreen();
                        else if (document.webkitExitFullscreen) document.webkitExitFullscreen();
                        try{
                            var parentDoc = window.parent.document;
                            if(parentDoc.exitFullscreen) parentDoc.exitFullscreen();
                            else if(parentDoc.webkitExitFullscreen) parentDoc.webkitExitFullscreen();
                        }catch(e){}
                        setTimeout(function(){
                            parentContainer.style.width = '100%';
                            parentContainer.style.height = '';
                            parentContainer.style.position = 'relative';
                            parentContainer.style.top = '';
                            parentContainer.style.left = '';
                            parentContainer.style.zIndex = '';
                            parentContainer.style.background = '';
                            mapContainer.style.height = '650px';
                            mapContainer.style.width = '100%';
                            mapContainer.style.borderRadius = '12px';
                            mapContainer.style.border = '3px solid #1A5D1A';
                            map.invalidateSize();
                        }, 100);
                    }
                }catch(err){ console.log('Fullscreen error', err); }
                setTimeout(function(){ map.invalidateSize(); }, 600);
            };
            return container;
        }
    });
    new FullscreenControl({position: 'bottomleft'}).addTo(map);
    // Listener fullscreen 100% tutto schermo
    document.addEventListener('fullscreenchange', function(){
        var mapContainer = document.getElementById('map');
        var parentContainer = document.getElementById('map-container');
        if (document.fullscreenElement) {
            parentContainer.style.width = '100vw';
            parentContainer.style.height = '100vh';
            parentContainer.style.position = 'fixed';
            parentContainer.style.top = '0';
            parentContainer.style.left = '0';
            parentContainer.style.zIndex = '9999';
            parentContainer.style.background = 'white';
            mapContainer.style.height = '100vh';
            mapContainer.style.width = '100vw';
            mapContainer.style.borderRadius = '0';
            mapContainer.style.border = 'none';
        } else {
            parentContainer.style.width = '100%';
            parentContainer.style.height = '';
            parentContainer.style.position = 'relative';
            parentContainer.style.top = '';
            parentContainer.style.left = '';
            parentContainer.style.zIndex = '';
            parentContainer.style.background = '';
            mapContainer.style.height = '650px';
            mapContainer.style.width = '100%';
            mapContainer.style.borderRadius = '12px';
            mapContainer.style.border = '3px solid #1A5D1A';
        }
        setTimeout(function(){ map.invalidateSize(); }, 600);
    });
    // CSS fullscreen 100% vero tutto schermo
    var styleFs = document.createElement('style');
    styleFs.innerHTML = `
        #map-container:fullscreen {
            width: 100vw !important;
            height: 100vh !important;
            background: white !important;
            padding: 0 !important;
            margin: 0 !important;
            position: fixed !important;
            top: 0 !important;
            left: 0 !important;
            z-index: 9999 !important;
        }
        #map-container:fullscreen #map {
            height: 100vh !important;
            width: 100vw !important;
            border-radius: 0 !important;
            border: none !important;
        }
        #map:fullscreen {
            width: 100vw !important;
            height: 100vh !important;
        }
    `;
    document.head.appendChild(styleFs);
    function getColorCode(c){ var m={'red':'#d32f2f','blue':'#1976d2','green':'#388e3c','orange':'#f57c00','purple':'#7b1fa2'}; return m[c]||'#388e3c'; }
    var allMarkers = [];
    markersData.forEach(function(md){
        // FIX: Vedi tutti i marker nelle mappe - richiesta Ezio - mostra tutti, non solo con file
        var mk;
        if(md.hasFile && md.fileB64){
            var icon = L.icon({iconUrl: "data:image/png;base64," + md.fileB64, iconSize: [40, 40], iconAnchor: [20, 20]});
            mk = L.marker([md.lat, md.lon], {icon: icon}).addTo(map).bindPopup("<b>📍 " + md.nome + "</b><br>📻 Alias Radio: <b>" + (md.alias||"--") + "</b> - Chiamata radio<br>🏙️ Comune: " + md.comune + "<br>📍 Via: " + md.via + "<br>Lat: " + md.lat + " Lon: " + md.lon + "<br><small>Emergenza: " + (md.emergenza||"--") + " | Evento: " + (md.evento||"--") + "</small>");
        } else {
            var colorMapMain = {'red':'#d32f2f','blue':'#1976d2','green':'#388e3c','orange':'#f57c00','purple':'#7b1fa2'};
            var colMain = colorMapMain[md.colore] || '#388e3c';
            var divIconMain = L.divIcon({html:"<div style='background:white;border:2px solid "+colMain+";width:36px;height:36px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:20px;box-shadow:0 2px 6px rgba(0,0,0,0.3);'>"+(md.emoji||"📍")+"</div>",iconSize:[36,36],iconAnchor:[18,18]});
            mk = L.marker([md.lat, md.lon], {icon: divIconMain}).addTo(map).bindPopup("<b>📍 " + md.nome + "</b><br>📻 Alias Radio: <b>" + (md.alias||"--") + "</b> - Chiamata radio<br>🏙️ Comune: " + md.comune + "<br>📍 Via: " + md.via + "<br>Lat: " + md.lat + " Lon: " + md.lon + "<br><small>Emergenza: " + (md.emergenza||"--") + " | Evento: " + (md.evento||"--") + "</small>");
        }
        allMarkers.push(mk);
    });
    if (focusMarker && focusMarker.Lat){
        map.setView([focusMarker.Lat, focusMarker.Lon], 18);
        // Zoom su mappa grande - icona selezionata ingrandita, non cerchio giallo
        if(focusMarker.hasFile && focusMarker.fileB64){
            var focusIconBig = L.icon({iconUrl: "data:image/png;base64," + focusMarker.fileB64, iconSize: [52, 52], iconAnchor: [26, 26]});
            L.marker([focusMarker.Lat, focusMarker.Lon], {icon: focusIconBig}).addTo(map).bindPopup("<b>📍 SELEZIONATA: "+focusMarker.nome+"</b><br>📻 Alias Radio: <b>"+(focusMarker.alias||"--")+"</b><br>🏙️ "+focusMarker.comune+"<br>📍 "+focusMarker.via).openPopup();
        }
        document.getElementById('coords').innerHTML = "📍 Zoom su: "+focusMarker.nome+" - "+focusMarker.comune+" "+focusMarker.via+" - Ingrandita su mappa grande";
    } else {
        if(allMarkers.length>0){
            var g = new L.featureGroup(allMarkers);
            map.fitBounds(g.getBounds().pad(0.3));
        }
    }
    map.on('click', function(e){
        var lat = e.latlng.lat.toFixed(6);
        var lon = e.latlng.lng.toFixed(6);
        // Solo icona scelta caricata da te - niente cerchio giallo, niente emoji - richiesta Ezio
        if(selectedIconHasFile && selectedIconFileB64){
            var tmpIcon = L.icon({iconUrl: "data:image/png;base64," + selectedIconFileB64, iconSize: [40,40], iconAnchor: [20,20]});
            L.marker([lat, lon], {icon: tmpIcon}).addTo(map).bindPopup("Nuova<br>" + lat + "," + lon).openPopup();
            document.getElementById('coords').innerHTML = "📍 Nuovo marker - Lat: " + lat + " Lon: " + lon + " - Compilo maschera lat/lon/comune/via";
        } else {
            // Se non hai caricato icona, avvisa - non mettere cerchio giallo né emoji
            document.getElementById('coords').innerHTML = "⚠️ Carica prima un'icona in Libreria Icone - poi selezionala - niente cerchio giallo";
            return;
        }
        // FIX: Quando metti marker deve compilarmi maschera con lat long comune e via - richiesta Ezio
        try{
            fetch('https://nominatim.openstreetmap.org/reverse?format=json&lat='+lat+'&lon='+lon)
                .then(r=>r.json()).then(d=>{
                    var com = d.address.city||d.address.town||d.address.village||"";
                    var via = d.address.road||"";
                    document.getElementById('coords').innerHTML = "📍 Nuovo Lat: "+lat+" Lon: "+lon+"<br>Comune: "+com+" Via: "+via+" - Maschera compilata";
                    try{
                        var url = new URL(window.parent.location.href);
                        url.searchParams.set('lat', lat);
                        url.searchParams.set('lon', lon);
                        url.searchParams.set('comune', com);
                        url.searchParams.set('via', via);
                        window.parent.history.replaceState(null, '', url.toString());
                    }catch(err){}
                });
        }catch(e){
            try {
                var url = new URL(window.parent.location.href);
                url.searchParams.set('lat', lat);
                url.searchParams.set('lon', lon);
                window.parent.history.replaceState(null, '', url.toString());
            } catch(err) {}
        }
    });
    </script>
    """
    import json as json_lib2
    import base64 as b64lib_main
    sel_e = selected_ico_obj.get('Emoji','👷')
    sel_c = selected_ico_obj.get('Colore','green')
    sel_hasfile_main = selected_ico_obj.get("HasFile", False)
    sel_fileb64_main = ""
    if sel_hasfile_main and selected_ico_obj.get("FileBytes"):
        try:
            fb = selected_ico_obj.get("FileBytes")
            if isinstance(fb, bytes):
                sel_fileb64_main = b64lib_main.b64encode(fb).decode()
        except:
            sel_fileb64_main = ""
    html_code = html_code.replace("MARKERS_JSON_PLACEHOLDER", markers_for_js)
    html_code = html_code.replace("FOCUS_JSON_PLACEHOLDER", focus_for_js)
    html_code = html_code.replace("SELECTED_EMOJI_PLACEHOLDER", json_lib2.dumps(sel_e))
    html_code = html_code.replace("SELECTED_COLOR_PLACEHOLDER", json_lib2.dumps(sel_c))
    html_code = html_code.replace("SELECTED_HASFILE_PLACEHOLDER", json_lib2.dumps(sel_hasfile_main))
    html_code = html_code.replace("SELECTED_FILEB64_PLACEHOLDER", json_lib2.dumps(sel_fileb64_main))
    st.components.v1.html(html_code, height=700)

    st.divider()
    # Filtro per vedere solo alias da form Alias - richiesta Ezio
    if all_markers and (alias_esistenti or alias_form_list):
        col_f1, col_f2 = st.columns([2,1])
        with col_f1:
            filtro_alias = st.selectbox("🔍 Filtra per Alias (da form Alias)", ["-- Tutti --"] + (alias_form_list if alias_form_list else alias_esistenti), key="filtro_alias_tab")
        with col_f2:
            if st.button("🧹 Pulisci filtro Alias", key="btn_pulisci_filtro_alias"):
                st.session_state["filtro_alias_tab"] = "-- Tutti --"
                st.rerun()
        if filtro_alias != "-- Tutti --":
            all_markers_filtered = [m for m in all_markers if m.get('Alias','') == filtro_alias]
            st.success(f"📻 Vedo solo alias: {filtro_alias} - {len(all_markers_filtered)} postazioni")
            all_markers_display = all_markers_filtered
        else:
            all_markers_display = all_markers
    else:
        all_markers_display = all_markers
        filtro_alias = "-- Tutti --"

    st.markdown(f"### 📋 Tabella Postazioni Salvate - {len(all_markers_display)} / {len(all_markers)} - Alias: {filtro_alias} - Clicca Vedi su mappa")
    if all_markers_display:
        for idx, m in enumerate(all_markers_display):
            is_focus = focus_marker and str(focus_marker.get('Lat')) == str(m['Lat']) and str(focus_marker.get('Lon')) == str(m['Lon'])
            bg = "#fffde7" if is_focus else "white"
            border = "#FFD700" if is_focus else "#e0e0e0"
            c1, c2, c3, c4 = st.columns([1,2,2,2])
            with c1:
                with st.container(border=True):
                    # RIGA 4722-4730 - Mostra icona salvata, non quadrato bianco - fix Ezio
                    if m.get("HasFile") and m.get("FileBytes"):
                        try:
                            st.image(m.get("FileBytes"), width=80, caption=m.get("IconaNome",""))
                        except:
                            st.markdown(f"<div style='text-align:center;background:{bg};border:2px solid {border};border-radius:8px;padding:4px;'><div style='font-size:26px;'>{m.get('Emoji','👷')}</div><small>{m.get('IconaNome','')}</small></div>", unsafe_allow_html=True)
                    else:
                        st.markdown(f"<div style='text-align:center;background:{bg};border:2px solid {border};border-radius:8px;padding:4px;'><div style='font-size:26px;'>{m.get('Emoji','👷')}</div><small>{m.get('IconaNome','')}</small></div>", unsafe_allow_html=True)
                    st.caption(f"{m.get('IconaNome','')} - {m.get('FileName','')[:15] if m.get('HasFile') else ''}")
                    if is_focus:
                        st.caption("👆 SELEZIONATA - Ingrandita su mappa")
            with c2:
                st.write(f"**{m['Nome']}**")
                # Alias radio - richiesta Ezio
                if m.get('Alias'):
                    st.markdown(f"<span style='background:#1A5D1A;color:white;padding:2px 8px;border-radius:12px;font-weight:bold;font-size:12px;'>📻 {m.get('Alias')}</span>", unsafe_allow_html=True)
                st.caption(f"Tipo: {m.get('Tipo','')} - {m.get('IconaNome','')}")
                st.caption(f"Emergenza: {m.get('NomeEmergenza','--')}")
            with c3:
                # FIX: Comune e Via associati al marker - sempre visibili - richiesta Ezio
                comune_txt = m.get('Comune','') or '--'
                via_txt = m.get('Via','') or '--'
                st.write(f"**🏙️ {comune_txt}**")
                st.write(f"📍 {via_txt}")
                st.caption(f"Lat: {m.get('Lat','')} Lon: {m.get('Lon','')} - Comune/Via da marker")
            with c4:
                if st.button("📍 Vedi su mappa", key=f"focus_{idx}", use_container_width=True, type="primary" if is_focus else "secondary", help="Ingrandisce postazione su anteprima e mappa grande"):
                    st.session_state.map_focus = m
                    st.session_state.last_clicked_lat = str(m['Lat'])
                    st.session_state.last_clicked_lon = str(m['Lon'])
                    st.rerun()
                col_del, col_dup = st.columns(2)
                with col_del:
                    if st.button("🗑️", key=f"del_{idx}", use_container_width=True):
                        st.session_state.mappa_avanzata_markers.pop(idx)
                        if is_focus:
                            st.session_state.map_focus = None
                        st.rerun()
                with col_dup:
                    if st.button("📋", key=f"dup_{idx}", use_container_width=True):
                        nm = m.copy()
                        nm["Nome"] = m["Nome"] + " copia"
                        st.session_state.mappa_avanzata_markers.append(nm)
                        st.rerun()
            st.divider()

    excel_import_inline("mappa_postazioni", "Mappe Postazioni")


elif cur == "Turni":
    hdr_form("TURNI - Gestione Turni Volontari")
    if "turni" not in st.session_state:
        st.session_state.turni = []
    tab1, tab2 = st.tabs(["➕ Nuovo Turno", "📋 Elenco Turni"])
    if show_tab1:
        c1, c2, c3 = st.columns(3)
        with c1:
            data_turno = st.date_input("Data Turno *", value=date.today(), format="DD/MM/YYYY", key="turno_data")
            ora_inizio = st.time_input("Ora Inizio *", value=time(8,0), key="turno_ora_in")
            ora_fine = st.time_input("Ora Fine *", value=time(12,0), key="turno_ora_fine")
        with c2:
            volontari_list = st.session_state.get("volontari", [])
            nomi_vol = [f"{v.get('Cognome','')} {v.get('Nome','')} - {v.get('Telefono','')}" for v in volontari_list] if volontari_list else ["-- Nessun volontario --"]
            volontario_sel = st.selectbox("Volontario *", nomi_vol, key="turno_volontario")
            tipo_turno = st.selectbox("Tipo Turno *", ["Mattina", "Pomeriggio", "Sera", "Notte", "Reperibilità", "Emergenza", "Evento", "Formazione", "Altro"], key="turno_tipo")
            luogo_turno = combo_comune("Luogo / Comune", "turno_comune", "Varese")
        with c3:
            via_turno = combo_vie("Via", luogo_turno, "turno_via", "")
            stato_turno = st.selectbox("Stato", ["Programmato", "Confermato", "In Corso", "Completato", "Annullato"], key="turno_stato")
            note_turno = st.text_area("Note Turno", key="turno_note")
            bg_t, txt_t, lab_t = get_stato_color(stato_turno)
            st.markdown(f'<div style="background:{bg_t};color:{txt_t};padding:6px;border-radius:6px;text-align:center;">{lab_t}: {stato_turno}</div>', unsafe_allow_html=True)
        if st.button("💾 Salva Turno", type="primary", use_container_width=True, key="btn_salva_turno"):
            if volontario_sel and volontario_sel != "-- Nessun volontario --":
                nuovo_turno = {
                    "Data": str(data_turno),
                    "OraInizio": str(ora_inizio),
                    "OraFine": str(ora_fine),
                    "Volontario": volontario_sel,
                    "Tipo": tipo_turno,
                    "Comune": luogo_turno,
                    "Via": via_turno,
                    "Stato": stato_turno,
                    "Note": note_turno,
                    "DataIns": datetime.now().strftime("%d/%m/%Y %H:%M")
                }
                st.session_state.turni.append(nuovo_turno)
                st.success(f"✅ Turno salvato: {volontario_sel} - {data_turno}")
                st.rerun()
            else:
                st.error("Seleziona volontario")
    if show_tab2:
        if st.session_state.turni:
            df_turni = pd.DataFrame(st.session_state.turni)
            st.dataframe(df_turni, use_container_width=True)
            for idx, row in enumerate(st.session_state.turni):
                c1, c2, c3 = st.columns([4,1,1])
                bg, txt, lab = get_stato_color(row.get("Stato","Programmato"))
                c1.markdown(f"**{row.get('Data','')} {row.get('OraInizio','')}-{row.get('OraFine','')}** - {row.get('Volontario','')} - {row.get('Tipo','')} - {row.get('Comune','')} {row.get('Via','')}")
                c2.markdown(f"<span style='background:{bg};color:{txt};padding:4px 8px;border-radius:4px;'>{row.get('Stato','')}</span>", unsafe_allow_html=True)
                if c3.button("🗑️", key=f"del_turno_{idx}"):
                    st.session_state.turni.pop(idx)
                    st.rerun()
            if REPORTLAB_OK:
                st.download_button("📄 PDF Turni", data=to_pdf(df_turni, "TURNI"), file_name="turni.pdf", mime="application/pdf", use_container_width=True, key="pdf_turni_final")
            st.download_button("📊 Excel Turni", data=to_excel(df_turni), file_name="turni.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="excel_turni_final")
        else:
            pass  # istruzione rimossa
            pass
            # Istruzione rimossa - form pulito


    # IMPORT/EXPORT INLINE - Ezio - TUTTI I FORM - Excel + PDF + Template ODV
    excel_import_inline("turni", "Turni")


elif cur == "Libreria Icone":
    hdr_form("LIBRERIA ICONE")

    st.markdown("""
    <div style="background:#e8f5e9;padding:8px;border-radius:8px;border-left:4px solid #1A5D1A;margin-bottom:12px;">
    <b>Qui crei le icone che poi usi su Mappe Postazioni - Decidi tu che marker usare - Ogni icona ha Emoji + Colore + Nome</b>
    </div>
    """, unsafe_allow_html=True)

    # Icone di default tolte - le carichi tu - richiesta Ezio
    # Nessuna icona predefinita - libreria vuota all'inizio
    if "icone" not in st.session_state:
        st.session_state.icone = []
    # Se vuoi ripristinare, carica da Excel o crea manualmente

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        nome_icona = st.text_input("Nome Icona *", key="ico_nome", placeholder="Es: Postazione 1")
        emoji_icona = st.text_input("Emoji Icona *", value="📍", key="ico_emoji", help="Inserisci emoji: 🚨 📅 🚐 👤 🏥 🔥 💧 📻 👷 🚒 🚑")
    with c2:
        tipo_icona = st.selectbox("Tipo", ["Emergenza", "Evento", "Mezzo", "Volontario", "Postazione", "Punto Interesse", "Altro"], key="ico_tipo")
        colore_icona = st.selectbox("Colore Marker", ["red", "blue", "green", "orange", "purple", "darkred", "darkblue", "cadetblue"], key="ico_colore")
    with c3:
        desc_icona = st.text_input("Descrizione Icona", key="ico_desc", placeholder="Descrizione")
        file_icona = st.file_uploader("File Icona (opzionale)", type=["png", "jpg", "svg"], key="ico_file")
    with c4:
        st.markdown("**Anteprima Icona Scelta**")
        # Anteprima emoji
        preview_emoji = st.session_state.get("ico_emoji", "📍") if "ico_emoji" in st.session_state else emoji_icona
        # Anteprima file caricato - richiesta Ezio: vedere anteprima icona caricata
        if file_icona is not None:
            try:
                file_bytes = file_icona.getvalue()
                if file_icona.type.startswith("image/"):
                    st.image(file_bytes, caption=f"Anteprima file: {file_icona.name}", width=120)
                    st.success(f"File caricato: {file_icona.name} - {len(file_bytes)} bytes")
                else:
                    st.markdown(f"<div style='font-size:40px;text-align:center;background:white;padding:10px;border-radius:8px;border:2px solid #1A5D1A;'>{preview_emoji}</div>", unsafe_allow_html=True)
                    st.caption(f"File: {file_icona.name}")
                # Salva temporaneo per preview
                st.session_state["ico_file_bytes"] = file_bytes
                st.session_state["ico_file_name"] = file_icona.name
            except Exception as e:
                st.error(f"Errore anteprima file: {e}")
                st.markdown(f"<div style='font-size:40px;text-align:center;background:white;padding:10px;border-radius:8px;border:2px solid #1A5D1A;'>{preview_emoji}</div>", unsafe_allow_html=True)
        else:
            # Se esiste file precedente in session
            if "ico_file_bytes" in st.session_state:
                try:
                    st.image(st.session_state["ico_file_bytes"], caption=f"Anteprima file: {st.session_state.get('ico_file_name','')}", width=120)
                except:
                    st.markdown(f"<div style='font-size:40px;text-align:center;background:white;padding:10px;border-radius:8px;border:2px solid #1A5D1A;'>{preview_emoji}</div>", unsafe_allow_html=True)
            else:
                st.markdown(f"<div style='font-size:40px;text-align:center;background:white;padding:10px;border-radius:8px;border:2px solid #1A5D1A;'>{preview_emoji}</div>", unsafe_allow_html=True)
                st.caption(f"Emoji: {preview_emoji} - Colore: {st.session_state.get('ico_colore','red')}")

    if st.button("💾 Salva Icona in Libreria", type="primary", use_container_width=True):
        if nome_icona and emoji_icona:
            # Controlla se esiste già
            exists = False
            for ico in st.session_state.icone:
                if ico.get("Nome") == nome_icona:
                    exists = True
                    break
            if not exists:
                new_ico = {
                    "Nome": nome_icona,
                    "Emoji": emoji_icona,
                    "Tipo": tipo_icona,
                    "Colore": colore_icona,
                    "Descrizione": desc_icona,
                    "Data": datetime.now().strftime("%d/%m/%Y %H:%M")
                }
                # Salva anche file se caricato
                if "ico_file_bytes" in st.session_state:
                    try:
                        new_ico["FileBytes"] = st.session_state["ico_file_bytes"]
                        new_ico["FileName"] = st.session_state.get("ico_file_name","")
                        new_ico["HasFile"] = True
                    except:
                        pass
                st.session_state.icone.append(new_ico)
                # Pulisci temp
                for k in ["ico_file_bytes", "ico_file_name"]:
                    if k in st.session_state:
                        try:
                            del st.session_state[k]
                        except:
                            pass
                st.success(f"Icona {emoji_icona} {nome_icona} salvata - Ora la puoi usare su Mappe Postazioni e Intervento Emergenza")
                st.rerun()
            else:
                st.warning("Nome già esistente - cambia nome")
        else:
            st.error("Nome e Emoji obbligatori")

    st.divider()
    st.markdown(f"### Libreria Icone - {len(st.session_state.icone)} icone disponibili - Le usi su Mappe Postazioni")

    if st.session_state.icone:
        st.markdown("**Anteprima libreria - clicca per vedere dettaglio**")
        cols = st.columns(4)
        for idx, ico in enumerate(st.session_state.icone):
            col = cols[idx % 4]
            with col:
                # Se ha file, mostra file, altrimenti emoji
                if ico.get("HasFile") and ico.get("FileBytes"):
                    try:
                        st.image(ico.get("FileBytes"), caption=f"{ico.get('Emoji','')} {ico.get('Nome','')}", width=100)
                    except:
                        st.markdown(f"""
                        <div style="background:white;padding:8px;border-radius:8px;border:2px solid #1A5D1A;text-align:center;margin-bottom:8px;">
                        <div style="font-size:32px;">{ico.get('Emoji','📍')}</div>
                        <b>{ico.get('Nome','')}</b><br>
                        <small>{ico.get('Tipo','')} - {ico.get('Colore','')}</small><br>
                        <small>{ico.get('Descrizione','')}</small>
                        </div>
                        """, unsafe_allow_html=True)
                else:
                    st.markdown(f"""
                    <div style="background:white;padding:8px;border-radius:8px;border:2px solid #1A5D1A;text-align:center;margin-bottom:8px;">
                    <div style="font-size:32px;">{ico.get('Emoji','📍')}</div>
                    <b>{ico.get('Nome','')}</b><br>
                    <small>{ico.get('Tipo','')} - {ico.get('Colore','')}</small><br>
                    <small>{ico.get('Descrizione','')}</small>
                    </div>
                    """, unsafe_allow_html=True)
                if st.button(f"🗑️ Elimina", key=f"del_ico_{idx}"):
                    st.session_state.icone.pop(idx)
                    st.rerun()
        st.divider()
        df_ico = pd.DataFrame(st.session_state.icone)
        st.dataframe(df_ico, use_container_width=True)
        st.download_button("Excel Libreria Icone", to_excel(df_ico), "libreria_icone.xlsx", use_container_width=True)
    else:
        pass  # istruzione rimossa
        pass
        # Istruzione rimossa - form pulito

# CHAT
    # IMPORT/EXPORT INLINE - Ezio - TUTTI I FORM - Excel + PDF + Template ODV
    excel_import_inline("icone", "Libreria Icone")


elif cur == "Chat":
    hdr_form("CHAT - Comunicazioni Squadra - Chi è collegato")

    # Mostra chi è collegato ora - Richiesta Ezio
    st.markdown("#### 🟢 Chi è collegato ora")
    try:
        presenza = load_presenza()
        if presenza:
            # Aggiorna mia presenza
            aggiorna_presenza(st.session_state.get("username",""), st.session_state.get("nome_utente",""), st.session_state.get("ruolo_utente",""))
            presenza = load_presenza()
            cols = st.columns(min(len(presenza), 4))
            for idx, p in enumerate(presenza[-12:]):  # ultimi 12
                with cols[idx % len(cols)]:
                    ruolo = p.get("ruolo","operatore")
                    colore = {"amministratore":"#d32f2f","operatore":"#1A5D1A","lettore":"#1976d2"}.get(ruolo, "#1A5D1A")
                    st.markdown(f"""
                    <div style="background:{colore};color:white;padding:8px;border-radius:8px;text-align:center;margin-bottom:6px;">
                    <b>🟢 {p.get("nome","")}</b><br>
                    <small>{p.get("username","")} - {ruolo}</small><br>
                    <small>{p.get("ora","")}</small>
                    </div>
                    """, unsafe_allow_html=True)
            st.caption(f"{len(presenza)} utenti collegati negli ultimi 30 minuti - Aggiornamento automatico")
        else:
            st.info("Nessun utente collegato oltre te - File presenza.json vuoto")
    except Exception as e:
        st.warning(f"Presenza non disponibile: {e}")

    st.divider()
    
    # Info multi-utente
    st.markdown("""
    <div style="background:#e3f2fd;padding:10px;border-radius:8px;border-left:4px solid #1976d2;margin-bottom:10px;">
    <b>ℹ️ Multi-utente contemporaneo:</b> Sì, più utenti possono aprire il gestionale insieme!<br>
    - Ogni utente ha il suo login (admin, operatore1, lettore1)<br>
    - Su Streamlit Cloud session_state è separato per utente, ma presenza.json è condiviso<br>
    - Per dati condivisi in tempo reale serve database esterno (Firebase/Supabase) - per ora chat e presenza usano file condiviso<br>
    - Chat salvata in sessione locale (per condivisione reale serve DB)
    </div>
    """, unsafe_allow_html=True)

    st.markdown("#### 💬 Chat operativa volontari")
    
    # Mostra utente corrente
    curr_user = st.session_state.get("nome_utente", "Admin")
    curr_username = st.session_state.get("username", "admin")
    st.caption(f"Stai chattando come: {curr_user} ({curr_username})")

    msg = st.text_input("Messaggio", key="chat_msg", placeholder="Scrivi messaggio e premi Invio o Invia")

    c1, c2, c3 = st.columns([1, 1, 3])
    with c1:
        if st.button("📤 Invia Messaggio", type="primary", use_container_width=True):
            if msg:
                st.session_state.chat.append({
                    "Data": datetime.now().strftime("%d/%m/%Y"),
                    "Ora": datetime.now().strftime("%H:%M:%S"),
                    "Utente": curr_user,
                    "Username": curr_username,
                    "Messaggio": msg,
                    "Ruolo": st.session_state.get("ruolo_utente","operatore")
                })
                # Salva anche su file per condivisione parziale
                try:
                    with open("chat.json","a", encoding="utf-8") as f:
                        f.write(json.dumps({"Data": datetime.now().strftime("%d/%m/%Y"), "Ora": datetime.now().strftime("%H:%M:%S"), "Utente": curr_user, "Messaggio": msg}) + "\n")
                except:
                    pass
                st.rerun()
    with c2:
        if st.button("🔄 Aggiorna", use_container_width=True):
            try:
                aggiorna_presenza(curr_username, curr_user, st.session_state.get("ruolo_utente",""))
            except:
                pass
            st.rerun()

    st.divider()

    if st.session_state.chat:
        st.markdown(f"#### Ultimi {len(st.session_state.chat[-30:])} messaggi")
        for chat_msg in reversed(st.session_state.chat[-30:]):
            ruolo = chat_msg.get('Ruolo','operatore')
            colore = {"amministratore":"#d32f2f","operatore":"#1A5D1A","lettore":"#1976d2"}.get(ruolo, "#1A5D1A")
            st.markdown(
                f"""
                <div style="background:white;padding:10px;border-radius:8px;
                margin-bottom:6px;border-left:4px solid {colore};">
                <strong>{chat_msg.get('Data','')} {chat_msg.get('Ora','')} - {chat_msg.get('Utente','')} ({chat_msg.get('Ruolo','')})</strong><br>
                {chat_msg.get('Messaggio','')}
                </div>
                """,
                unsafe_allow_html=True
            )
    else:
        st.info("Nessun messaggio - Inizia conversazione - I messaggi sono visibili solo nella tua sessione (per chat condivisa serve DB)")

    # Mostra anche chat da file se esiste
    try:
        if os.path.exists("chat.json"):
            st.divider()
            st.markdown("#### 📁 Chat condivisa da file (ultimi 10)")
            with open("chat.json","r", encoding="utf-8") as f:
                lines = f.readlines()[-10:]
                for line in reversed(lines):
                    try:
                        cj = json.loads(line)
                        st.caption(f"{cj.get('Data','')} {cj.get('Ora','')} - {cj.get('Utente','')}: {cj.get('Messaggio','')}")
                    except:
                        pass
    except:
        pass

elif False: # Geolocalizzazione rimossa su richiesta Ezio
    if cur == "Geolocalizzazione Hytera + Anytone RIMOSSA":
        st.session_state.menu = "Geolocalizzazione Hytera + Anytone"
        pass
    hdr_form("GEOLOCALIZZAZIONE HYTERA + ANYTONE - RX AUTOMATICA + PONTE RETE - MAPPA ANTEPRIMA")

    # === MAPPA ANTEPRIMA SEMPRE IN CIMA - FIX EZIO - BOTTONI VISIBILI SUBITO ===
    st.markdown("### 🗺️ MAPPA LIVE - ANTEPRIMA IN CIMA - Sempre visibile")
    if "map_filter" not in st.session_state:
        st.session_state.map_filter = "Tutti"
    if "show_big_map" not in st.session_state:
        st.session_state.show_big_map = False

    c_fs_geo1, c_fs_geo2 = st.columns([1,3])
    with c_fs_geo1:
        if st.button("⛶ Fullscreen 100%", key="btn_fs_geo_top", use_container_width=True, type="primary"):
            st.session_state["fs_geo_active"] = not st.session_state.get("fs_geo_active", False)
            st.rerun()
    with c_fs_geo2:
        st.markdown('<span style="background:red;color:white;padding:6px 12px;border-radius:6px;font-weight:bold;">🔴 FULL = mappa 100% come Mappe Postazioni - ESC per uscire</span>', unsafe_allow_html=True)

    if st.session_state.get("fs_geo_active"):
        st.components.v1.html("""
            <script>
            (function(){
                try {
                    const docEl = window.parent.document.documentElement;
                    if (docEl.requestFullscreen) docEl.requestFullscreen();
                } catch(e){}
                try {
                    if (document.documentElement.requestFullscreen) document.documentElement.requestFullscreen();
                } catch(e){}
            })();
            </script>
            <div style="background:#ff0000;color:white;padding:8px;border-radius:6px;text-align:center;font-weight:bold;">
            🔴 FULLSCREEN 100% ATTIVO - premi ESC per uscire - Mappa Geo 100%
            </div>
            """, height=70)
        if st.button("❌ Esci Fullscreen 100%", key="btn_exit_fs_geo", use_container_width=True):
            st.components.v1.html("<script>try{document.exitFullscreen(); parent.document.exitFullscreen();}catch(e){}</script>", height=0)
            st.session_state["fs_geo_active"] = False
            st.rerun()

    col_f1, col_f2, col_f3, col_f4, col_f5 = st.columns(5)
    with col_f1:
        if st.button("🔴 PD785", key="filter_pd_TOP", use_container_width=True):
            st.session_state.map_filter = "PD785"
            st.rerun()
    with col_f2:
        if st.button("🔵 Anytone", key="filter_any_TOP", use_container_width=True):
            st.session_state.map_filter = "Anytone"
            st.rerun()
    with col_f3:
        if st.button("🟢 APRS.fi", key="filter_aprs_TOP", use_container_width=True):
            st.session_state.map_filter = "APRS"
            st.rerun()
    with col_f4:
        if st.button("🌍 Tutti", key="filter_all_TOP", use_container_width=True):
            st.session_state.map_filter = "Tutti"
            st.rerun()
    with col_f5:
        if st.button("🔍 Ingrandisci", key="btn_big_map_TOP", use_container_width=True):
            st.session_state.show_big_map = not st.session_state.show_big_map
            st.rerun()
    
    st.caption(f"Filtro: {st.session_state.map_filter} | Big: {'ON' if st.session_state.show_big_map else 'OFF'}")

    # Mostra mappa anteprima SEMPRE - demo se vuota
    try:
        all_pos_top = st.session_state.get("posizioni_pd785", []) + st.session_state.get("posizioni_anytone", [])
        if all_pos_top:
            map_data_top = []
            for p in all_pos_top:
                try:
                    lat = float(str(p.get("Lat","0")).replace(",","."))
                    lon = float(str(p.get("Lon","0")).replace(",","."))
                    if -90 <= lat <= 90 and -180 <= lon <= 180 and lat != 0 and lon != 0:
                        modello = str(p.get("Modello","")).lower()
                        fonte = str(p.get("Fonte","")).lower()
                        if "anytone" in modello or "878" in modello:
                            color = "#00FF00" if "aprs" in fonte or "aprs" in modello else "#0000FF"
                        else:
                            color = "#FF0000"
                        filt = st.session_state.map_filter
                        tipo = "PD785" if color=="#FF0000" else "Anytone" if color=="#0000FF" else "APRS"
                        if filt == "Tutti" or filt == tipo:
                            map_data_top.append({"lat": lat, "lon": lon, "color": color, "ID": p.get("ID","")})
                except:
                    continue
            if map_data_top:
                map_df_top = pd.DataFrame(map_data_top)
                st.markdown(f"##### 🗺️ ANTEPRIMA - {len(map_data_top)} posizioni - Filtro {st.session_state.map_filter} - Leaflet OSM")
                show_leaflet_map(map_data_top, height=400, zoom=12)
                if st.session_state.show_big_map:
                    st.markdown(f"### 🔍 MAPPA GRANDE - {len(map_data_top)} posizioni - Ingrandita")
                    show_leaflet_map(map_data_top, height=600, zoom=13)
            else:
                st.warning(f"Nessuna per filtro {st.session_state.map_filter} - Mostro demo")
                show_leaflet_map([], height=400, zoom=11)
        else:
            st.info("Nessuna posizione ancora - Demo Varese + Caronno - Inserisci ID/Lat/Lon sotto o APRS.fi")
            show_leaflet_map([], height=400, zoom=11)
    except Exception as e:
        st.error(f"Errore mappa anteprima: {e}")
        show_leaflet_map([], height=400, zoom=11)

    st.divider()

    # === RX AUTOMATICA VIA URL ===
    try:
        qp = st.query_params
        rx_id = qp.get("rx_id", "")
        rx_lat = qp.get("rx_lat", "")
        rx_lon = qp.get("rx_lon", "")
        rx_model = qp.get("rx_model", "Anytone")
        if rx_id and rx_lat and rx_lon:
            try:
                lat_f = float(str(rx_lat).replace(",", "."))
                lon_f = float(str(rx_lon).replace(",", "."))
                if "any" in rx_model.lower() or "878" in rx_model.lower():
                    st.session_state.posizioni_anytone.append({
                        "ID": rx_id,
                        "Lat": str(lat_f),
                        "Lon": str(lon_f),
                        "Ora": datetime.now().strftime("%H:%M:%S"),
                        "Modello": rx_model,
                        "Fonte": "RX_AUTO_PONTE"
                    })
                else:
                    st.session_state.posizioni_pd785.append({
                        "ID": rx_id,
                        "Lat": str(lat_f),
                        "Lon": str(lon_f),
                        "Ora": datetime.now().strftime("%H:%M:%S"),
                        "Modello": rx_model,
                        "Fonte": "RX_AUTO_PONTE"
                    })
                st.success(f"📡 RX AUTO PONTE: {rx_id} - {lat_f}, {lon_f} - {rx_model}")
                try:
                    st.query_params.clear()
                except:
                    pass
            except Exception as e:
                st.error(f"Errore RX auto ponte: {e}")
    except:
        pass

    # === PARSER NMEA / APRS / Hytera ===
    with st.expander("📡 RX Parser - Incolla stringa GPS da radio (Hytera / Anytone ponte)"):
        gps_string = st.text_area("Stringa GPS", key="gps_parser_input", placeholder="Es: 101,45.8167,8.8333 oppure $GPRMC,... oppure APRS: IU2XYZ>APRS,=4549.00N/00850.00E-")
        c_parse1, c_parse2 = st.columns(2)
        with c_parse1:
            parse_id = st.text_input("ID Radio (se non in stringa)", key="parse_id")
        with c_parse2:
            parse_model = st.selectbox("Modello", ["Anytone 878 - Ponte Rete", "PD785", "Altro"], key="parse_model")
        
        if st.button("📡 Decodifica e Metti su Mappa", type="primary", use_container_width=True, key="btn_parse_gps"):
            if gps_string:
                lat_parsed = None
                lon_parsed = None
                id_parsed = parse_id
                try:
                    parts = gps_string.replace(";", ",").split(",")
                    if len(parts) >= 3:
                        for i in range(len(parts)-1):
                            try:
                                maybe_lat = float(parts[i].replace(",", "."))
                                maybe_lon = float(parts[i+1].replace(",", "."))
                                if -90 <= maybe_lat <= 90 and -180 <= maybe_lon <= 180:
                                    lat_parsed = maybe_lat
                                    lon_parsed = maybe_lon
                                    if not id_parsed:
                                        id_parsed = parts[0].strip()
                                    break
                            except:
                                continue
                    if "$GPRMC" in gps_string or "$GPGGA" in gps_string:
                        try:
                            p = gps_string.split(",")
                            if len(p) >= 6:
                                lat_raw = p[3]
                                lat_dir = p[4]
                                lon_raw = p[5]
                                lon_dir = p[6] if len(p) > 6 else "E"
                                lat_deg = int(float(lat_raw)//100)
                                lat_min = float(lat_raw) - lat_deg*100
                                lat_parsed = lat_deg + lat_min/60
                                if lat_dir == "S":
                                    lat_parsed = -lat_parsed
                                lon_deg = int(float(lon_raw)//100)
                                lon_min = float(lon_raw) - lon_deg*100
                                lon_parsed = lon_deg + lon_min/60
                                if lon_dir == "W":
                                    lon_parsed = -lon_parsed
                        except:
                            pass
                except Exception as e:
                    st.error(f"Errore parsing: {e}")
                
                if lat_parsed and lon_parsed:
                    if not id_parsed:
                        id_parsed = "RX_"+datetime.now().strftime("%H%M%S")
                    if "Anytone" in parse_model or "878" in parse_model:
                        st.session_state.posizioni_anytone.append({
                            "ID": id_parsed,
                            "Lat": str(lat_parsed),
                            "Lon": str(lon_parsed),
                            "Ora": datetime.now().strftime("%H:%M:%S"),
                            "Modello": parse_model,
                            "Fonte": "RX_PARSER"
                        })
                    else:
                        st.session_state.posizioni_pd785.append({
                            "ID": id_parsed,
                            "Lat": str(lat_parsed),
                            "Lon": str(lon_parsed),
                            "Ora": datetime.now().strftime("%H:%M:%S"),
                            "Modello": parse_model,
                            "Fonte": "RX_PARSER"
                        })
                    st.success(f"✅ RX: {id_parsed} -> {lat_parsed}, {lon_parsed}")
                    st.rerun()
                else:
                    st.warning("Non riesco a leggere lat/lon - usa formato ID,lat,lon")

    # === IMPORT CSV + RETE ===
    with st.expander("📂 Import CSV da Hytera / Anytone + da RETE (URL) - Aperto di default"):
        st.markdown("**1) Da file locale (export Hytera SmartDispatch / Anytone CPS)**")
        up_csv = st.file_uploader("CSV con ID,Lat,Lon", type=["csv","txt"], key="up_csv_gps")
        if up_csv:
            try:
                df_csv = pd.read_csv(up_csv)
                st.write(df_csv.head())
                if st.button("Importa tutto su mappa da FILE", key="btn_import_csv_gps"):
                    for _, row in df_csv.iterrows():
                        lat_col = None
                        lon_col = None
                        id_col = None
                        for c in df_csv.columns:
                            if "lat" in c.lower():
                                lat_col = c
                            if "lon" in c.lower() or "lng" in c.lower():
                                lon_col = c
                            if "id" in c.lower():
                                id_col = c
                        if lat_col and lon_col:
                            # Scegli lista in base a modello se presente
                            mod = str(row.get("Modello","") if "Modello" in df_csv.columns else "")
                            target = st.session_state.posizioni_anytone if "any" in mod.lower() or "878" in mod.lower() else st.session_state.posizioni_pd785
                            target.append({
                                "ID": str(row[id_col]) if id_col else "CSV",
                                "Lat": str(row[lat_col]),
                                "Lon": str(row[lon_col]),
                                "Ora": datetime.now().strftime("%H:%M:%S"),
                                "Modello": mod if mod else "CSV_IMPORT"
                            })
                    st.success(f"Importati {len(df_csv)} posizioni")
                    st.rerun()
            except Exception as e:
                st.error(f"Errore CSV: {e}")

        st.divider()
        st.markdown("**2) Da RETE - URL diretto (ponte Anytone / Hytera / Google Sheets CSV)**")
        st.caption("Incolla URL del CSV che sta in rete: es. http://192.168.1.100/posizioni.csv oppure link Google Sheets pubblicato come CSV, oppure http://ponte-anytone.local/gps.csv")
        url_csv = st.text_input("URL CSV da rete", key="url_csv_rete", placeholder="http://192.168.1.50/gps.csv oppure https://docs.google.com/spreadsheets/d/.../export?format=csv")
        c_url1, c_url2 = st.columns(2)
        with c_url1:
            auto_refresh = st.checkbox("🔄 Auto-refresh ogni 30s da rete (ponte)", key="auto_refresh_rete", value=False)
        with c_url2:
            url_model = st.selectbox("Modello per URL rete", ["Anytone 878 - Ponte Rete", "PD785", "Mix da CSV"], key="url_model")

        if st.button("🌐 SCARICA da RETE e metti su mappa", type="primary", use_container_width=True, key="btn_download_rete"):
            if url_csv:
                try:
                    r = requests.get(url_csv, timeout=10)
                    r.raise_for_status()
                    # Prova a leggere CSV da testo
                    from io import StringIO
                    df_rete = pd.read_csv(StringIO(r.text))
                    st.success(f"Scaricato da rete: {len(df_rete)} righe")
                    st.write(df_rete.head(20))
                    # Importa
                    count = 0
                    for _, row in df_rete.iterrows():
                        lat_col = None
                        lon_col = None
                        id_col = None
                        for c in df_rete.columns:
                            if "lat" in c.lower():
                                lat_col = c
                            if "lon" in c.lower() or "lng" in c.lower():
                                lon_col = c
                            if "id" in c.lower() or "callsign" in c.lower() or "alias" in c.lower():
                                id_col = c
                        if lat_col and lon_col:
                            if url_model == "Anytone 878 - Ponte Rete":
                                st.session_state.posizioni_anytone.append({
                                    "ID": str(row[id_col]) if id_col and id_col in row else f"RETE_{count}",
                                    "Lat": str(row[lat_col]),
                                    "Lon": str(row[lon_col]),
                                    "Ora": datetime.now().strftime("%H:%M:%S"),
                                    "Modello": "Anytone 878 - Ponte Rete",
                                    "Fonte": f"RETE_URL:{url_csv[:50]}"
                                })
                            elif url_model == "PD785":
                                st.session_state.posizioni_pd785.append({
                                    "ID": str(row[id_col]) if id_col and id_col in row else f"RETE_{count}",
                                    "Lat": str(row[lat_col]),
                                    "Lon": str(row[lon_col]),
                                    "Ora": datetime.now().strftime("%H:%M:%S"),
                                    "Modello": "PD785",
                                    "Fonte": f"RETE_URL:{url_csv[:50]}"
                                })
                            else:
                                mod_row = str(row.get("Modello",""))
                                if "any" in mod_row.lower() or "878" in mod_row.lower():
                                    st.session_state.posizioni_anytone.append({
                                        "ID": str(row[id_col]) if id_col else f"RETE_{count}",
                                        "Lat": str(row[lat_col]),
                                        "Lon": str(row[lon_col]),
                                        "Ora": datetime.now().strftime("%H:%M:%S"),
                                        "Modello": mod_row,
                                        "Fonte": "RETE_CSV_MIX"
                                    })
                                else:
                                    st.session_state.posizioni_pd785.append({
                                        "ID": str(row[id_col]) if id_col else f"RETE_{count}",
                                        "Lat": str(row[lat_col]),
                                        "Lon": str(row[lon_col]),
                                        "Ora": datetime.now().strftime("%H:%M:%S"),
                                        "Modello": mod_row if mod_row else "PD785",
                                        "Fonte": "RETE_CSV_MIX"
                                    })
                            count += 1
                    st.success(f"✅ {count} posizioni da RETE messe su mappa!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Errore scarico rete: {e} - Controlla che URL sia CSV accessibile senza login")
            else:
                st.warning("Inserisci URL CSV")

        # Auto-refresh rimosso blocco caricamento - ora solo info, non fa request automatica
        if auto_refresh and url_csv:
            st.caption(f"🔄 Auto-refresh attivo per: {url_csv} - verrà aggiornato al prossimo click su SCARICA da RETE")

    # === APRS.FI - ANYTONE POSIZIONE DA APRS === - FIX anti-dashboard - fuori da expander
    st.markdown("### 📡 APRS.fi - Prendi posizione Anytone da APRS.fi")
    with st.container(border=True):
        st.markdown("""
        <div style="background:#e3f2fd;padding:10px;border-radius:8px;border-left:4px solid #1976d2;">
        <b>📡 Anytone 878 su ponte DMR → APRS.fi</b><br>
        L'Anytone invia GPS via APRS su ponte → finisce su aprs.fi → noi lo prendiamo da qui e lo mettiamo su mappa.<br>
        Serve API key gratuita da aprs.fi
        </div>
        """, unsafe_allow_html=True)
        
        st.info("🔑 Come prendi API key gratis: 1) Vai su https://aprs.fi 2) Registrati gratis 3) https://aprs.fi/page/api → copia la tua API key")
        
        # FIX: Usa form per evitare salto su Dashboard quando clicchi API key
        with st.form("form_aprs_fi_fix", clear_on_submit=False):
            c_aprs1, c_aprs2 = st.columns(2)
            with c_aprs1:
                aprs_callsign = st.text_input("Callsign Anytone (es: IU2XYZ-9, IZ2xxx-7)", key="aprs_callsign", placeholder="IU2XYZ-9", help="Il callsign che usi su Anytone APRS - con SSID -9 per portatile")
                aprs_api_key = st.text_input("API Key aprs.fi (gratis)", key="aprs_api_key", type="password", placeholder="Inserisci API key da aprs.fi/page/api", help="Gratis - prendila su aprs.fi/page/api")
            with c_aprs2:
                aprs_filter = st.selectbox("Modello", ["Anytone 878 - Ponte Rete", "PD785 + APRS", "Tutti"], key="aprs_model_filter")
                auto_aprs = st.checkbox("🔄 Auto-refresh APRS ogni 60s", key="auto_aprs_refresh", value=False)
            aprs_multi = st.text_area("Lista callsign squadra (uno per riga, opzionale)", key="aprs_multi", placeholder="IU2XYZ-9\nIZ2ABC-7\nIU2DEF-9", help="Se hai più Anytone, metti uno per riga e prende tutti")
            btn_aprs = st.form_submit_button("📡 PRENDI da aprs.fi ORA - Anytone", type="primary", use_container_width=True)
        # Fine form - mantiene valori senza saltare a Dashboard

        c_btn1, c_btn2 = st.columns(2)
        with c_btn1:
            st.caption("Usa bottone dentro form sopra per prendere da aprs.fi")
        with c_btn2:
            btn_aprs_clear = st.button("🧹 Pulisci lista APRS", use_container_width=True, key="btn_aprs_clear")

        if btn_aprs_clear:
            # non cancella tutto, solo info
            st.info("Lista APRS pulita - usa bottoni pulisci sotto mappa per cancellare posizioni")

        def fetch_aprs_fi(callsign, api_key):
            """Prende posizione da aprs.fi API ufficiale"""
            try:
                # API ufficiale aprs.fi
                url = f"https://api.aprs.fi/api/get?name={callsign.strip()}&what=loc&apikey={api_key.strip()}&format=json"
                r = requests.get(url, timeout=10)
                r.raise_for_status()
                data = r.json()
                if data.get("result") == "ok" and data.get("entries"):
                    entry = data["entries"][0]
                    lat = entry.get("lat")
                    lng = entry.get("lng")
                    lasttime = entry.get("lasttime")
                    comment = entry.get("comment", "")
                    return {"lat": lat, "lon": lng, "time": lasttime, "comment": comment, "found": True}
                else:
                    # Prova senza filtro what
                    url2 = f"https://api.aprs.fi/api/get?name={callsign.strip()}&what=loc&apikey={api_key.strip()}&format=json&lasttime=86400"
                    r2 = requests.get(url2, timeout=10)
                    data2 = r2.json()
                    if data2.get("entries"):
                        e = data2["entries"][0]
                        return {"lat": e.get("lat"), "lon": e.get("lng"), "time": e.get("lasttime"), "comment": e.get("comment",""), "found": True}
                    return {"found": False, "error": data.get("description","Non trovato - verifica callsign e API key")}
            except Exception as e:
                return {"found": False, "error": str(e)}

        def fetch_aprs_no_key(callsign):
            """Fallback senza API key - prova aprs.fi via scraping leggero o api alternativa"""
            try:
                # Prova api alternativa aprsdirect o aprs.is che non richiede key
                # 1) Prova aprs.fi page json non ufficiale
                url = f"https://aprs.fi/json/get?call={callsign.strip()}&time=86400"
                headers = {"User-Agent": "ANA-Varese-App/1.0"}
                r = requests.get(url, headers=headers, timeout=10)
                if r.status_code == 200:
                    try:
                        data = r.json()
                        if data.get("entries"):
                            e = data["entries"][0]
                            return {"lat": e.get("lat"), "lon": e.get("lng"), "found": True}
                    except:
                        pass
                # 2) Prova aprsdirect
                url2 = f"https://api.aprsdirect.de/api/v1/get?c={callsign.strip()}"
                r2 = requests.get(url2, timeout=10)
                if r2.status_code == 200:
                    try:
                        data2 = r2.json()
                        # formato variabile
                        if "entries" in data2 and data2["entries"]:
                            e = data2["entries"][0]
                            return {"lat": e.get("lat") or e.get("latitude"), "lon": e.get("lng") or e.get("longitude"), "found": True}
                    except:
                        pass
                return {"found": False, "error": "Serve API key aprs.fi - prendi gratis su aprs.fi/page/api"}
            except Exception as e:
                return {"found": False, "error": str(e)}

        if btn_aprs:
            callsigns = []
            if aprs_multi.strip():
                callsigns = [c.strip() for c in aprs_multi.strip().split("\n") if c.strip()]
            elif aprs_callsign.strip():
                callsigns = [aprs_callsign.strip()]
            
            if not callsigns:
                st.warning("Inserisci almeno un callsign es: IU2XYZ-9")
            else:
                progress = st.progress(0)
                for idx, cs in enumerate(callsigns):
                    progress.progress((idx+1)/len(callsigns))
                    st.write(f"🔍 Cerco {cs} su aprs.fi...")
                    if aprs_api_key.strip():
                        res = fetch_aprs_fi(cs, aprs_api_key)
                    else:
                        res = fetch_aprs_no_key(cs)
                        if not res.get("found"):
                            st.warning(f"⚠️ {cs}: {res.get('error')} - Inserisci API key gratis da aprs.fi/page/api per funzionare meglio")
                            continue
                    
                    if res.get("found") and res.get("lat") and res.get("lon"):
                        # Aggiungi su mappa
                        target_list = st.session_state.posizioni_anytone if "Anytone" in aprs_filter else st.session_state.posizioni_pd785 if "PD785" in aprs_filter else st.session_state.posizioni_anytone
                        # Per tutti, metti in anytone se callsign contiene -9 -7 ecc, altrimenti PD785
                        if aprs_filter == "Tutti":
                            target_list = st.session_state.posizioni_anytone
                        
                        target_list.append({
                            "ID": cs,
                            "Lat": str(res["lat"]),
                            "Lon": str(res["lon"]),
                            "Ora": datetime.now().strftime("%H:%M:%S"),
                            "Modello": "Anytone 878 - APRS.fi",
                            "Fonte": f"APRS.fi:{cs}",
                            "Comment": res.get("comment",""),
                            "LastTime": res.get("time","")
                        })
                        st.success(f"✅ {cs} -> {res['lat']}, {res['lon']} messo su mappa! {res.get('comment','')}")
                    else:
                        st.error(f"❌ {cs} non trovato su aprs.fi: {res.get('error','Non trovato - verifica che Anytone stia trasmettendo APRS sul ponte')}")
                
                st.balloons()
                st.rerun()

        if auto_aprs:
            st.caption(f"🔄 Auto APRS attivo: {aprs_callsign} - premi PRENDI da aprs.fi per aggiornare (evita blocco pagina)")



    st.divider()
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Posizioni Hytera PD785 - RX Manuale + Auto + Ponte**")
        id_pd = st.text_input("ID Radio PD785", key="pd_id")
        lat_pd = st.text_input("Latitudine PD785", value="45.8167", key="pd_lat")
        lon_pd = st.text_input("Longitudine PD785", value="8.8333", key="pd_lon")
        if st.button("Aggiorna Posizione PD785", use_container_width=True, key="btn_pd785"):
            if id_pd and lat_pd and lon_pd:
                st.session_state.posizioni_pd785.append({
                    "ID": id_pd,
                    "Lat": lat_pd,
                    "Lon": lon_pd,
                    "Ora": datetime.now().strftime("%H:%M:%S"),
                    "Modello": "PD785",
                    "Fonte": "MANUALE"
                })
                st.success("Posizione PD785 aggiornata")
                st.session_state.menu = "Geolocalizzazione Hytera + Anytone"
                st.rerun()

    with c2:
        st.markdown("**Posizioni Anytone 878 - Ponte in Rete - RX Auto**")
        id_any = st.text_input("ID Radio Anytone", key="any_id")
        lat_any = st.text_input("Latitudine Anytone", value="45.82", key="any_lat")
        lon_any = st.text_input("Longitudine Anytone", value="8.84", key="any_lon")
        if st.button("Aggiorna Posizione Anytone - Ponte Rete", use_container_width=True, key="btn_anytone"):
            if id_any and lat_any and lon_any:
                st.session_state.posizioni_anytone.append({
                    "ID": id_any,
                    "Lat": lat_any,
                    "Lon": lon_any,
                    "Ora": datetime.now().strftime("%H:%M:%S"),
                    "Modello": "Anytone 878 - Ponte Rete",
                    "Fonte": "MANUALE_PONTE"
                })
                st.success("Posizione Anytone ponte rete aggiornata")
                st.session_state.menu = "Geolocalizzazione Hytera + Anytone"
                st.rerun()

    st.divider()
    st.markdown("### 🗺️ MAPPA - Tutte le posizioni RX - Hytera + Anytone ponte rete")

    all_pos = st.session_state.posizioni_pd785 + st.session_state.posizioni_anytone
    
    st.markdown("#### 🗺️ MAPPA LIVE - Colori: 🔴 PD785 | 🔵 Anytone Ponte | 🟢 APRS.fi")
    
    # Pulsanti colorati cliccabili - FIX Ezio
    if "map_filter" not in st.session_state:
        st.session_state.map_filter = "Tutti"
    if "show_big_map" not in st.session_state:
        st.session_state.show_big_map = False

    col_f1, col_f2, col_f3, col_f4, col_f5 = st.columns(5)
    with col_f1:
        if st.button("🔴 PD785", key="filter_pd", use_container_width=True):
            st.session_state.map_filter = "PD785"
            st.session_state.menu = "Geolocalizzazione Hytera + Anytone"
            st.rerun()
    with col_f2:
        if st.button("🔵 Anytone", key="filter_any", use_container_width=True):
            st.session_state.map_filter = "Anytone"
            st.session_state.menu = "Geolocalizzazione Hytera + Anytone"
            st.rerun()
    with col_f3:
        if st.button("🟢 APRS.fi", key="filter_aprs", use_container_width=True):
            st.session_state.map_filter = "APRS"
            st.session_state.menu = "Geolocalizzazione Hytera + Anytone"
            st.rerun()
    with col_f4:
        if st.button("🌍 Tutti", key="filter_all", use_container_width=True):
            st.session_state.map_filter = "Tutti"
            st.session_state.menu = "Geolocalizzazione Hytera + Anytone"
            st.rerun()
    with col_f5:
        if st.button("🔍 Ingrandisci", key="btn_big_map", use_container_width=True):
            st.session_state.show_big_map = not st.session_state.show_big_map
            st.session_state.menu = "Geolocalizzazione Hytera + Anytone"
            st.rerun()
    
    st.caption(f"Filtro attivo: {st.session_state.map_filter} | Big: {'ON' if st.session_state.show_big_map else 'OFF'}")
    
    if all_pos:
        df_pos = pd.DataFrame(all_pos)
        st.dataframe(df_pos, use_container_width=True)
        try:
            map_data = []
            for p in all_pos:
                try:
                    lat = float(str(p.get("Lat","0")).replace(",","."))
                    lon = float(str(p.get("Lon","0")).replace(",","."))
                    if -90 <= lat <= 90 and -180 <= lon <= 180 and lat != 0 and lon != 0:
                        modello = str(p.get("Modello","")).lower()
                        fonte = str(p.get("Fonte","")).lower()
                        if "anytone" in modello or "878" in modello:
                            if "aprs" in fonte or "aprs" in modello:
                                color = "#00FF00"
                                tipo = "APRS"
                            else:
                                color = "#0000FF"
                                tipo = "Anytone"
                        else:
                            color = "#FF0000"
                            tipo = "PD785"
                        filt = st.session_state.map_filter
                        if filt == "Tutti" or (filt == "PD785" and tipo == "PD785") or (filt == "Anytone" and tipo == "Anytone") or (filt == "APRS" and tipo == "APRS"):
                            map_data.append({"lat": lat, "lon": lon, "color": color, "ID": p.get("ID",""), "Modello": p.get("Modello",""), "tipo": tipo})
                except:
                    continue
            if map_data:
                map_df = pd.DataFrame(map_data)
                st.markdown("##### Mappa Anteprima - Leaflet OSM - Sempre visibile")
                show_leaflet_map(map_data, height=400, zoom=12)
                
                if st.session_state.get("show_big_map", False):
                    with st.expander("🔍 MAPPA INGRANDITA - Anteprima grande", expanded=True):
                        st.markdown(f"### 🗺️ MAPPA GRANDE - {st.session_state.map_filter} - {len(map_data)} posizioni")
                        show_leaflet_map(map_data, height=500, zoom=13)
                        # Lista grande
                        for p in map_data:
                            emoji = "🔴" if p["color"]=="#FF0000" else "🔵" if p["color"]=="#0000FF" else "🟢"
                            st.markdown(f"{emoji} **{p['ID']}** {p['lat']},{p['lon']}")
                
                st.success(f"📍 {len(map_data)} visibili - Filtro: {st.session_state.map_filter}")
                for p in map_data:
                    emoji = "🔴" if p["color"]=="#FF0000" else "🔵" if p["color"]=="#0000FF" else "🟢"
                    st.markdown(f"{emoji} **{p['ID']}** ({p['Modello']}) -> {p['lat']}, {p['lon']}")
            else:
                st.warning(f"Nessuna posizione per filtro {st.session_state.map_filter} - Demo Leaflet")
                show_leaflet_map([], height=400, zoom=11)
        except Exception as e:
            st.error(f"Errore mappa: {e}")
            show_leaflet_map([], height=400, zoom=11)
        c_exp1, c_exp2, c_exp3 = st.columns(3)
        with c_exp1:
            st.download_button("Excel Posizioni", to_excel(df_pos), "posizioni_hytera_anytone.xlsx", use_container_width=True, key="dl_pos_excel")
        with c_exp2:
            if st.button("🧹 Pulisci PD785", use_container_width=True, key="btn_clear_pd"):
                st.session_state.posizioni_pd785 = []
                st.rerun()
        with c_exp3:
            if st.button("🧹 Pulisci Anytone", use_container_width=True, key="btn_clear_any"):
                st.session_state.posizioni_anytone = []
                st.rerun()
    else:
        st.info("Nessuna posizione ancora - Demo Leaflet OSM - Inserisci manualmente o APRS.fi")
        show_leaflet_map([], height=400, zoom=11)
        st.caption("Mappa demo - Varese + Caronno - Clicca i pulsanti colorati sopra per filtrare quando avrai posizioni")



# GESTIONE UTENTI - Amministratore e utenti view/insert - Ezio richiesta
elif cur == "Gestione Utenti":
    hdr_form("GESTIONE UTENTI - Solo Amministratore - Maschera Configurazione VISIBILE")
    
    # Solo amministratore può accedere - RICHIESTA EZIO - maschera solo per admin
    if st.session_state.get("ruolo_utente") != "amministratore":
        st.error("⛔ Accesso negato - Solo amministratore può gestire utenti e vedere maschera configurazione")
        st.info(f"Il tuo ruolo: {st.session_state.get('ruolo_utente')} - Contatta amministratore (admin / ana2024)")
        st.stop()
    
    st.markdown("""
    <div style="background:#ffebee;padding:12px;border-radius:8px;border-left:4px solid #d32f2f;margin-bottom:12px;">
    <b>🔐 MASCHERA CONFIGURAZIONE UTENTI - Solo Amministratore - SEMPRE VISIBILE</b><br>
    Qui come Amministratore crei i vari utenti con livelli accesso - Maschera configurazione in alto sempre visibile<br>
    <b>Livelli:</b> 🔴 Amministratore=tutto | 🟠 Coordinatore=gestione | 🟢 Operatore=vede+inserisce | 🔵 Volontario=base | ⚪ Lettore=solo vista
    </div>
    """, unsafe_allow_html=True)
    
    utenti = load_utenti()
    
    # === MASCHERA CONFIGURAZIONE SEMPRE VISIBILE IN ALTO - Richiesta Ezio ===
    st.markdown("### ➕ MASCHERA CONFIGURAZIONE - Crea Nuovo Utente con Livello Accesso")
    st.markdown("""
    <div style="background:#e8f5e9;padding:12px;border-radius:8px;border:3px solid #1A5D1A;margin-bottom:12px;">
    <b>📌 Come Amministratore crei utenti (maschera configurazione):</b><br>
    1. Username (senza spazi, es: mario.rossi) - minuscolo<br>
    2. Nome Completo (es: Mario Rossi ODV Varese)<br>
    3. Scegli Livello Accesso / Ruolo<br>
    4. Password + Conferma (min 4 caratteri)<br>
    5. Spunta Permessi extra<br>
    6. Clicca CREA UTENTE - Salva in utenti.json
    </div>
    """, unsafe_allow_html=True)
    
    # Campi maschera SEMPRE VISIBILI - 3 colonne
    col_u1, col_u2, col_u3 = st.columns(3)
    with col_u1:
        mu_username = st.text_input("Username * (minuscolo, senza spazi)", placeholder="es: mario.rossi", key="cfg_username")
        mu_nome = st.text_input("Nome Completo *", placeholder="es: Mario Rossi - Squadra A", key="cfg_nome")
        mu_ruolo = st.selectbox("Livello Accesso / Ruolo *", 
            ["operatore","coordinatore","volontario","lettore","amministratore"], 
            index=0, key="cfg_ruolo")
        desc_ruoli = {
            "amministratore": "🔴 Tutto: utenti, backup, tutti form",
            "coordinatore": "🟠 Squadre, volontari, mezzi, emergenze",
            "operatore": "🟢 Vede+inserisce volontari, mezzi, brogliaccio",
            "volontario": "🔵 Check-in, chat, vista base",
            "lettore": "⚪ Solo vista"
        }
        st.caption(desc_ruoli.get(mu_ruolo, ""))
    
    with col_u2:
        mu_pwd = st.text_input("Password *", type="password", key="cfg_pwd")
        mu_pwd2 = st.text_input("Conferma Password *", type="password", key="cfg_pwd2")
        mu_attivo = st.checkbox("✅ Utente Attivo", value=True, key="cfg_attivo")
        st.caption("Se non attivo, non può fare login")
        st.markdown("---")
        st.caption(f"Stai creando come: {st.session_state.get('username','admin')} - {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    
    with col_u3:
        st.markdown("**Permessi per singolo utente - Spunta TUTTI i form che può usare:**")
        st.caption("Decidi per singolo utente cosa può fare - TUTTI visibili qui")
        is_admin = mu_ruolo == "amministratore"
        is_coord = mu_ruolo in ["amministratore","coordinatore"]
        is_oper = mu_ruolo in ["amministratore","coordinatore","operatore"]
        
        mu_perm_dashboard = st.checkbox("🏠 Dashboard", value=True, key="cfg_perm_dashboard")
        mu_perm_volontari = st.checkbox("👤 Volontari (con foto)", value=is_oper, key="cfg_perm_volontari")
        mu_perm_db_radio = st.checkbox("📻 DB Radio", value=is_oper, key="cfg_perm_db_radio")
        mu_perm_consegna = st.checkbox("🤝 Consegna Radio", value=is_oper, key="cfg_perm_consegna")
        mu_perm_alias = st.checkbox("🔖 Alias Radio", value=is_oper, key="cfg_perm_alias")
        mu_perm_brog = st.checkbox("📓 Brogliaccio", value=is_oper, key="cfg_perm_brog")
        mu_perm_eventi = st.checkbox("📅 Eventi", value=is_oper, key="cfg_perm_eventi")
        mu_perm_emergenze = st.checkbox("🚨 Emergenze", value=is_oper, key="cfg_perm_emergenze")
        mu_perm_tab_em = st.checkbox("📋 Tabella Emergenze", value=is_oper, key="cfg_perm_tab_em")
        mu_perm_checkin = st.checkbox("✅ Check-in", value=True, key="cfg_perm_checkin")
        mu_perm_interv = st.checkbox("🚒 Interventi Emergenza", value=is_oper, key="cfg_perm_interv")
        mu_perm_tab_interv = st.checkbox("📊 Tabella Interventi", value=is_oper, key="cfg_perm_tab_interv")
        mu_perm_mezzi = st.checkbox("🚐 Mezzi", value=is_oper, key="cfg_perm_mezzi")
        mu_perm_attrezz = st.checkbox("🧰 Attrezzature", value=is_oper, key="cfg_perm_attrezz")
        mu_perm_mappe = st.checkbox("🗺️ Mappe Postazioni", value=is_coord, key="cfg_perm_mappe")
        mu_perm_icone = st.checkbox("🎨 Libreria Icone", value=is_coord, key="cfg_perm_icone")
        mu_perm_chat = st.checkbox("💬 Chat", value=True, key="cfg_perm_chat")
        mu_perm_geo = st.checkbox("📍 Geolocalizzazione", value=is_coord, key="cfg_perm_geo")
        mu_perm_backup = st.checkbox("💾 Backup", value=is_admin, key="cfg_perm_backup")
        mu_perm_gest = st.checkbox("👥 Gestione Utenti (solo admin)", value=is_admin, key="cfg_perm_gest")
    
    # Bottone CREA con TUTTI i permessi
    if st.button("✅ CREA UTENTE - SALVA CON PERMESSI SCELTI PER SINGOLO UTENTE", type="primary", use_container_width=True, key="btn_crea_utente_mask_visibile"):
        if not mu_username or not mu_nome or not mu_pwd:
            st.error("❌ Compila Username, Nome, Password *")
        elif mu_pwd != mu_pwd2:
            st.error("❌ Password non coincidono")
        elif len(mu_pwd) < 4:
            st.error("❌ Password minimo 4 caratteri")
        elif any(u.get("username") == mu_username.strip().lower() for u in utenti):
            st.error(f"❌ Username {mu_username} già esistente")
        elif " " in mu_username.strip():
            st.error("❌ Username senza spazi - usa punto: mario.rossi")
        else:
            perm_list = []
            if mu_perm_dashboard: perm_list.append("dashboard")
            if mu_perm_volontari: perm_list.append("volontari")
            if mu_perm_db_radio: perm_list.append("db_radio")
            if mu_perm_consegna: perm_list.append("consegna_radio")
            if mu_perm_alias: perm_list.append("alias_radio")
            if mu_perm_brog: perm_list.append("brogliaccio")
            if mu_perm_eventi: perm_list.append("eventi")
            if mu_perm_emergenze: perm_list.append("emergenze")
            if mu_perm_tab_em: perm_list.append("tabella_emergenze")
            if mu_perm_checkin: perm_list.append("checkin")
            if mu_perm_interv: perm_list.append("interventi_emergenza")
            if mu_perm_tab_interv: perm_list.append("tabella_interventi")
            if mu_perm_mezzi: perm_list.append("mezzi")
            if mu_perm_attrezz: perm_list.append("attrezzature")
            if mu_perm_mappe: perm_list.append("mappe_postazioni")
            if mu_perm_icone: perm_list.append("libreria_icone")
            if mu_perm_chat: perm_list.append("chat")
            if mu_perm_geo: perm_list.append("geolocalizzazione")
            if mu_perm_backup: perm_list.append("backup")
            if mu_perm_gest: perm_list.append("gestione_utenti")
            
            nuovo = {
                "username": mu_username.strip().lower(),
                "password": hash_pwd(mu_pwd.strip()),
                "nome": mu_nome.strip(),
                "ruolo": mu_ruolo,
                "attivo": mu_attivo,
                "permessi": perm_list,
                "creato_da": st.session_state.get("username","admin"),
                "data_creazione": datetime.now().strftime("%d/%m/%Y %H:%M")
            }
            utenti.append(nuovo)
            if save_utenti(utenti):
                st.success(f"✅ Utente {mu_username.strip().lower()} creato! Livello: {mu_ruolo.upper()} - Login: {mu_username.strip().lower()} / {mu_pwd.strip()}")
                st.balloons()
                st.rerun()
            else:
                st.error("Errore salvataggio utenti.json")
    
    st.divider()
    
    # Sotto maschera: elenco e modifica
    tab_list, tab_edit, tab_online = st.tabs(["👥 Elenco Utenti", "✏️ Modifica/Elimina", "📊 Online"])
    
    with tab_list:
        st.markdown(f"#### Utenti configurati: {len(utenti)}")
        if utenti:
            df_show = []
            for u in utenti:
                df_show.append({
                    "Username": u.get("username"),
                    "Nome": u.get("nome"),
                    "Ruolo": u.get("ruolo"),
                    "Livello": {"amministratore":"🔴 Tutto","coordinatore":"🟠 Gestione","operatore":"🟢 Vede+Ins","volontario":"🔵 Base","lettore":"⚪ Vista"}.get(u.get("ruolo"), u.get("ruolo")),
                    "Attivo": "✅" if u.get("attivo",True) else "❌",
                    "Permessi": ", ".join(u.get("permessi", [])[:3]) if u.get("permessi") else "Base"
                })
            st.dataframe(pd.DataFrame(df_show), use_container_width=True)
    
    with tab_edit:
        st.markdown("#### Modifica / Elimina Utente")
        sel = st.selectbox("Seleziona utente", ["--"] + [u.get("username") for u in utenti], key="sel_edit_cfg")
        if sel != "--":
            uo = next((u for u in utenti if u.get("username")==sel), None)
            if uo:
                c1, c2 = st.columns(2)
                with c1:
                    en_nome = st.text_input("Nome", value=uo.get("nome",""), key="en_nome")
                    en_ruolo = st.selectbox("Ruolo", ["amministratore","coordinatore","operatore","volontario","lettore"], index=["amministratore","coordinatore","operatore","volontario","lettore"].index(uo.get("ruolo","operatore")), key="en_ruolo")
                    en_attivo = st.checkbox("Attivo", value=uo.get("attivo",True), key="en_attivo")
                with c2:
                    en_pwd = st.text_input("Nuova Password (vuoto=no cambio)", type="password", key="en_pwd")
                    st.caption(f"Username: {sel}")
                col1, col2, col3 = st.columns(3)
                with col1:
                    if st.button("💾 Salva", type="primary", use_container_width=True, key="btn_en_save"):
                        for u in utenti:
                            if u.get("username")==sel:
                                u["nome"]=en_nome
                                u["ruolo"]=en_ruolo
                                u["attivo"]=en_attivo
                                if en_pwd.strip():
                                    u["password"]=hash_pwd(en_pwd.strip())
                        save_utenti(utenti)
                        st.success("Aggiornato")
                        st.rerun()
                with col2:
                    if st.button("🔄 Reset pwd a 'password'", use_container_width=True, key="btn_en_reset"):
                        for u in utenti:
                            if u.get("username")==sel:
                                u["password"]=hash_pwd("password")
                        save_utenti(utenti)
                        st.success("Reset a 'password'")
                        st.rerun()
                with col3:
                    if st.button("🗑️ Elimina", use_container_width=True, key="btn_en_del"):
                        if sel=="admin":
                            st.error("Non puoi eliminare admin")
                        else:
                            if st.session_state.get("conf_del")==sel:
                                utenti2=[u for u in utenti if u.get("username")!=sel]
                                save_utenti(utenti2)
                                st.success("Eliminato")
                                st.session_state["conf_del"]=None
                                st.rerun()
                            else:
                                st.session_state["conf_del"]=sel
                                st.warning("Clicca di nuovo per confermare")
    
    with tab_online:
        st.markdown("#### Online")
        try:
            pres = load_presenza()
            if pres:
                st.dataframe(pd.DataFrame(pres), use_container_width=True)
            else:
                st.info("Solo tu online")
        except Exception as e:
            st.error(f"{e}")

# VERBALI - Numero progressivo + Responsabile agganciato Volontari + campo libero
elif cur == "Verbali":
    hdr_form("VERBALI - Numero progressivo + PDF con logo PC ANA")

    if "verbali" not in st.session_state:
        st.session_state.verbali = []
    if "archivio_documenti" not in st.session_state:
        st.session_state.archivio_documenti = []

    # Numero progressivo automatico - anno corrente
    anno_corr = datetime.now().year
    num_prog = len(st.session_state.verbali) + 1
    num_verbale_auto = f"{num_prog:03d}/{anno_corr}"

    st.markdown(f"""<div style="background:#C8E6C9;padding:8px;border-radius:8px;border-left:5px solid #1A5D1A;margin-bottom:8px;text-align:center;"><b>📝 Verbale {num_verbale_auto}</b> - Totale: {len(st.session_state.verbali)}</div>""", unsafe_allow_html=True)

    # Maschera colore di fondo come altri form - verde #C8E6C9 - Richiesta Ezio
    st.markdown("""
    <style>
    /* Verde chiaro #C8E6C9 per form verbali come volontari - Riga maschera */
    </style>
    """, unsafe_allow_html=True)

    # Contenitore verde chiaro per maschera verbali - come altri form
    st.markdown('<div style="background:#C8E6C9;padding:15px;border-radius:10px;border:2px solid #1A5D1A;margin-bottom:15px;">', unsafe_allow_html=True)

    # Form verbale - 2 colonne
    c1, c2 = st.columns(2)
    with c1:
        num_verbale = st.text_input("Numero Verbale *", value=num_verbale_auto, key="verb_num")
        data_verbale = st.date_input("Data Verbale", value=date.today(), min_value=date(1950,1,1), max_value=date.today(), format="DD/MM/YYYY", key="verb_data")
        ora_verbale = st.time_input("Ora", value=datetime.now().time(), key="verb_ora")
        luogo_verbale = st.text_input("Luogo *", value="Sede ANA Varese", key="verb_luogo", placeholder="Es: Sede ANA Varese - Via...")

    with c2:
        tipo_verbale = st.selectbox("Tipo Verbale *", ["Consiglio Direttivo", "Assemblea Ordinaria", "Assemblea Straordinaria", "Riunione Squadra", "Riunione Emergenza", "Riunione Formazione", "Verbale Intervento", "Altro"], key="verb_tipo")
        oggetto_verbale = st.text_input("Oggetto *", key="verb_oggetto", placeholder="Es: Approvazione bilancio, Organizzazione esercitazione...")
        # Responsabile agganciato a Volontari - Nome Cognome unificato - Riga richiesta Ezio
        if st.session_state.volontari:
            vol_options = [f"{v.get('Cognome','').strip()} {v.get('Nome','').strip()}" for v in st.session_state.volontari if v.get('Cognome') or v.get('Nome')]
            vol_options = sorted(list(set([o for o in vol_options if o.strip()])))
            # Aggiungi opzione manuale
            vol_options = [""] + vol_options + ["Altro - inserisci manualmente"]
            responsabile_sel = st.selectbox("Responsabile * (da Volontari - Nome Cognome unificato)", vol_options, key="verb_resp_sel")
            if responsabile_sel == "Altro - inserisci manualmente":
                responsabile = st.text_input("Responsabile - inserisci Nome Cognome", key="verb_resp_manual", placeholder="Mario Rossi")
            elif responsabile_sel == "":
                responsabile = st.text_input("Responsabile * - Nome Cognome", key="verb_resp", value=st.session_state.get("nome_utente",""), placeholder="Mario Rossi")
            else:
                responsabile = responsabile_sel
                st.caption(f"✅ Selezionato da DB Volontari: {responsabile}")
        else:
            responsabile = st.text_input("Responsabile * - Nome Cognome (agganciato Volontari)", key="verb_resp", value=st.session_state.get("nome_utente",""), placeholder="Mario Rossi - Quando aggiungi volontari compariranno qui")

    st.divider()
    st.markdown("#### 📋 Contenuto Verbale - Campo libero")
    # Campo verbale mano libera - Richiesta Ezio
    st.markdown("""<div style="background:#e8f5e9;padding:6px;border-radius:6px;border-left:4px solid #1A5D1A;margin-bottom:6px;text-align:center;"><b>✍️ Verbale</b></div>""", unsafe_allow_html=True)
    testo_verbale = st.text_area("Verbale - Scrivi a mano libera * (campo grande)", key="verb_testo_libero", placeholder="Scrivi qui il verbale completo a mano libera...\n\nEs:\nIl giorno ... alle ore ... presso ... si è riunito il Consiglio...\nPresenti: ...\nODG: ...\nSi discute: ...\nSi delibera: ...", height=500)

    st.divider()
    st.markdown("#### 📋 Dettagli strutturati (opzionali - per PDF strutturato)")
    odg = st.text_area("Ordine del Giorno (ODG)", key="verb_odg", placeholder="1. Approvazione verbale precedente\n2. Comunicazioni\n3. Varie", height=80)
    delibere = st.text_area("Delibere / Decisioni prese", key="verb_delibere", placeholder="Il consiglio delibera: ...", height=100)
    incarichi = st.text_area("Incarichi assegnati", key="verb_incarichi", placeholder="Mario Rossi: preparazione mezzi...", height=60)
    note_verb = st.text_area("Note finali", key="verb_note", placeholder="Note, allegati...")

    st.markdown("</div>", unsafe_allow_html=True)  # chiude contenitore verde maschera
    st.divider()
    col_save, col_pdf = st.columns(2)
    with col_save:
        if st.button("💾 SALVA VERBALE", type="primary", use_container_width=True, key="btn_salva_verbale"):
            if num_verbale and luogo_verbale and responsabile and testo_verbale:
                nuovo_verb = {
                    "NumVerbale": num_verbale,
                    "Data": str(data_verbale),
                    "Ora": str(ora_verbale),
                    "Luogo": luogo_verbale,
                    "Tipo": tipo_verbale,
                    "Oggetto": oggetto_verbale,
                    "Responsabile": responsabile,
                    "TestoVerbale": testo_verbale,
                    "ODG": odg,
                    "Delibere": delibere,
                    "Incarichi": incarichi,
                    "Note": note_verb,
                    "DataIns": datetime.now().strftime("%d/%m/%Y %H:%M")
                }
                st.session_state.verbali.append(nuovo_verb)
                # Pulisci maschera
                for k in ["verb_num","verb_data","verb_ora","verb_luogo","verb_tipo","verb_oggetto","verb_resp","verb_resp_sel","verb_resp_manual","verb_testo_libero","verb_odg","verb_delibere","verb_incarichi","verb_note"]:
                    if k in st.session_state:
                        try:
                            del st.session_state[k]
                        except:
                            pass
                st.success(f"✅ Verbale {num_verbale} salvato! Responsabile: {responsabile} - Prossimo: {len(st.session_state.verbali)+1:03d}/{anno_corr}")
                st.balloons()
                st.rerun()
            else:
                st.error("Compila campi obbligatori: Numero, Luogo, Responsabile, Verbale mano libera *")


    with col_pdf:
        if st.session_state.verbali:
            ultimo = st.session_state.verbali[-1]
            try:
                def verbale_to_pdf_logo(verb):
                    buf = io.BytesIO()
                    try:
                        from reportlab.lib.pagesizes import A4
                        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
                        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
                        from reportlab.lib import colors
                        from reportlab.lib.units import cm
                        from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
                        doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=2*cm, rightMargin=2*cm, topMargin=2*cm, bottomMargin=2*cm)
                        styles = getSampleStyleSheet()
                        style_title = ParagraphStyle('TitleCustom', parent=styles['Title'], fontSize=14, alignment=TA_CENTER, spaceAfter=12, textColor=colors.HexColor("#1A5D1A"))
                        style_heading = ParagraphStyle('HeadingCustom', parent=styles['Heading2'], fontSize=11, textColor=colors.HexColor("#1A5D1A"), spaceBefore=10, spaceAfter=6)
                        style_normal = ParagraphStyle('NormalCustom', parent=styles['Normal'], fontSize=10, leading=13, alignment=TA_JUSTIFY)
                        story = []
                        try:
                            if os.path.exists("logo.png"):
                                logo = Image("logo.png", width=70, height=70)
                                story.append(logo)
                        except:
                            pass
                        try:
                            for p in ["gruppo_CPB.jpeg","logo2.png","/mnt/data/gruppo_CPB.jpeg"]:
                                if os.path.exists(p):
                                    logo2 = Image(p, width=60, height=60)
                                    story.append(logo2)
                                    break
                        except:
                            pass
                        story.append(Paragraph(f"<b>VERBALE N. {verb.get('NumVerbale','')} - {verb.get('Tipo','').upper()}</b>", style_title))
                        story.append(Spacer(1, 6))
                        story.append(Paragraph(f"<b>Squadra Volontari di protezione civile - Gruppo Alpini di Caronno Pertusella Bariola<br/>NUCLEO VOLONTARI DI P.C. A.N.A. - SEZIONE DI VARESE<br/>ASSOCIAZIONE NAZIONALE ALPINI</b>", style_normal))
                        story.append(Spacer(1, 8))
                        story.append(Paragraph(f"Data: {verb.get('Data','')} - Ora: {verb.get('Ora','')} - Luogo: {verb.get('Luogo','')} - Responsabile: {verb.get('Responsabile','')}", style_normal))
                        story.append(Spacer(1, 8))
                        story.append(Paragraph(f"<b>Oggetto:</b> {verb.get('Oggetto','')}", style_heading))
                        story.append(Paragraph(verb.get('Oggetto',''), style_normal))
                        story.append(Spacer(1, 10))
                        story.append(Paragraph(f"<b>VERBALE:</b>", style_heading))
                        # Testo libero con a capo
                        testo_libero_html = verb.get('TestoVerbale','').replace('\n','<br/>')
                        story.append(Paragraph(testo_libero_html, style_normal))
                        story.append(Spacer(1, 10))
                        if verb.get('ODG',''):
                            story.append(Paragraph("<b>ODG:</b>", style_heading))
                            story.append(Paragraph(verb.get('ODG','').replace('\n','<br/>'), style_normal))
                            story.append(Spacer(1, 6))
                        if verb.get('Delibere',''):
                            story.append(Paragraph("<b>Delibere:</b>", style_heading))
                            story.append(Paragraph(verb.get('Delibere','').replace('\n','<br/>'), style_normal))
                            story.append(Spacer(1, 6))
                        if verb.get('Incarichi',''):
                            story.append(Paragraph("<b>Incarichi:</b>", style_heading))
                            story.append(Paragraph(verb.get('Incarichi','').replace('\n','<br/>'), style_normal))
                        story.append(Spacer(1, 20))
                        data_firme = [["Il Responsabile", "Data"], [verb.get('Responsabile',''), verb.get('Data','')], ["___________________", "___________________"]]
                        t_firme = Table(data_firme, colWidths=[7.5*cm, 7.5*cm])
                        t_firme.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('FONTSIZE', (0,0), (-1,-1), 9), ('FONTNAME', (0,0), (0,0), 'Helvetica-Bold'),]))
                        story.append(t_firme)
                        doc.build(story)
                        buf.seek(0)
                        return buf.getvalue()
                    except Exception as e:
                        return f"Verbale {verb.get('NumVerbale','')} - Errore {e}".encode()

                pdf_ultimo = verbale_to_pdf_logo(ultimo)
                st.download_button(f"📄 PDF Ultimo Verbale {ultimo.get('NumVerbale','')}", data=pdf_ultimo, file_name=f"Verbale_{ultimo.get('NumVerbale','').replace('/','_')}.pdf", mime="application/pdf", use_container_width=True, key="pdf_ultimo_verbale")
                if st.button("💾 Salva PDF in Archivio Documenti", use_container_width=True, key="salva_pdf_archivio"):
                    nuovo_doc = {
                        "Titolo": f"Verbale {ultimo.get('NumVerbale','')} - {ultimo.get('Oggetto','')}",
                        "Tipo": "Verbale",
                        "Categoria": "Amministrativo",
                        "Descrizione": f"Verbale {ultimo.get('Tipo','')} del {ultimo.get('Data','')} - Resp: {ultimo.get('Responsabile','')} - {ultimo.get('Oggetto','')}",
                        "NomeFile": f"Verbale_{ultimo.get('NumVerbale','').replace('/','_')}.pdf",
                        "TipoFile": "application/pdf",
                        "DimensioneKB": round(len(pdf_ultimo)/1024, 1),
                        "DataDoc": ultimo.get('Data',''),
                        "CaricatoDa": ultimo.get('Responsabile',''),
                        "DataIns": datetime.now().strftime("%d/%m/%Y %H:%M"),
                        "FileBytes": pdf_ultimo
                    }
                    st.session_state.archivio_documenti.append(nuovo_doc)
                    st.success(f"✅ PDF Verbale {ultimo.get('NumVerbale','')} salvato in Archivio Documenti!")
            except Exception as e:
                st.error(f"Errore PDF: {e}")

    st.divider()
    if st.session_state.verbali:
        st.markdown(f"#### 📋 Elenco Verbali ({len(st.session_state.verbali)}) - Numero progressivo + Responsabile agganciato Volontari")
        for idx, verb in enumerate(reversed(st.session_state.verbali)):
            real_idx = len(st.session_state.verbali) - 1 - idx
            c1, c2, c3, c4, c5 = st.columns([1, 1, 2, 1, 0.8])
            with c1:
                st.write(f"**{verb.get('NumVerbale','')}**")
                st.caption(f"{verb.get('Data','')} {verb.get('Ora','')}")
            with c2:
                st.write(f"{verb.get('Tipo','')}")
                st.caption(f"Resp: {verb.get('Responsabile','')[:20]}")
            with c3:
                st.write(f"{verb.get('Oggetto','')[:50]}")
                st.caption(f"{verb.get('TestoVerbale','')[:60]}...")
            with c4:
                if st.button(f"👁️ {verb.get('NumVerbale','')}", key=f"view_verb_{real_idx}"):
                    st.session_state["verb_view_idx"] = real_idx
            with c5:
                if st.button("🗑️", key=f"del_verb_{real_idx}"):
                    st.session_state.verbali.pop(real_idx)
                    st.rerun()

        if "verb_view_idx" in st.session_state:
            try:
                v_idx = st.session_state["verb_view_idx"]
                verb_sel = st.session_state.verbali[v_idx]
                st.divider()
                st.markdown(f"### 📝 Verbale {verb_sel.get('NumVerbale','')} - Dettaglio Responsabile: {verb_sel.get('Responsabile','')}")
                st.write(f"**Oggetto:** {verb_sel.get('Oggetto','')}")
                st.text_area("Testo verbale mano libera", value=verb_sel.get('TestoVerbale',''), height=250, key=f"view_testo_{v_idx}", disabled=True)
                st.write(f"**ODG:** {verb_sel.get('ODG','')}")
                st.write(f"**Delibere:** {verb_sel.get('Delibere','')}")
            except:
                pass

        st.divider()
        df_verb = pd.DataFrame([{k:v for k,v in verb.items() if k not in ["FileBytes"]} for verb in st.session_state.verbali])
        st.dataframe(df_verb[["NumVerbale","Data","Tipo","Oggetto","Responsabile"]].head(20) if not df_verb.empty and "NumVerbale" in df_verb.columns else df_verb, use_container_width=True)
        c_exp1, c_exp2 = st.columns(2)
        with c_exp1:
            try:
                st.download_button("⬇️ EXCEL Verbali", data=to_excel(df_verb), file_name=f"verbali_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="exp_excel_verbali")
            except Exception as e:
                st.error(f"Excel errore: {e}")
        with c_exp2:
            try:
                if REPORTLAB_OK:
                    st.download_button("📄 PDF Lista Verbali", data=to_pdf(df_verb, "LISTA VERBALI"), file_name="lista_verbali.pdf", mime="application/pdf", use_container_width=True, key="exp_pdf_verbali")
            except Exception as e:
                st.error(f"PDF errore: {e}")
    else:
        pass  # istruzione rimossa
        pass
        # Istruzione rimossa - form pulito

    excel_import_inline("verbali", "Verbali")

elif cur == "Archivio Documenti":
    hdr_form("ARCHIVIO DOCUMENTI - Salva PDF, Word, Excel ecc")

    st.markdown("""
    <div style="background:#e8f5e9;padding:12px;border-radius:8px;border-left:4px solid #1A5D1A;margin-bottom:12px;">
    <b>📁 Archivio documenti</b><br>
    Regolamenti, convenzioni, attestati, verbali, circolari<br>
    <small>PDF, DOC, XLS, XLSX, JPG, PNG, ZIP</small>
    </div>
    """, unsafe_allow_html=True)

    # Inizializza sessione se manca
    if "archivio_documenti" not in st.session_state:
        st.session_state.archivio_documenti = []

    c1, c2 = st.columns(2)
    with c1:
        titolo_doc = st.text_input("Titolo Documento *", key="arch_titolo", placeholder="Es: Regolamento ODV 2024")
        tipo_doc = st.selectbox("Tipo Documento", ["Regolamento", "Convenzione", "Attestato", "Verbale", "Circolare", "Manuale Radio", "Modulo", "Autorizzazione", "Altro"], key="arch_tipo")
        categoria_doc = st.selectbox("Categoria", ["Generale", "Volontari", "Radio", "Mezzi", "Emergenze", "Formazione", "Amministrativo", "Sicurezza"], key="arch_cat")
        descrizione_doc = st.text_area("Descrizione", key="arch_desc", placeholder="Descrizione breve del documento")
    
    with c2:
        data_doc = st.date_input("Data Documento", value=date.today(), format="DD/MM/YYYY", key="arch_data")
        uploader_doc = st.text_input("Caricato da", value=st.session_state.get("nome_utente","Admin"), key="arch_uploader")
        file_doc = st.file_uploader("Carica File - PDF, Word, Excel ecc *", type=["pdf", "doc", "docx", "xls", "xlsx", "jpg", "jpeg", "png", "zip", "txt", "ppt", "pptx"], key="arch_file")
        if file_doc:
            # Info file rimossa - form pulito
            # Preview se immagine
            if file_doc.type and "image" in file_doc.type:
                st.image(file_doc.getvalue(), width=200, caption="Anteprima")

    if st.button("💾 SALVA DOCUMENTO IN ARCHIVIO", type="primary", use_container_width=True, key="btn_salva_arch"):
        if titolo_doc and file_doc:
            file_bytes = file_doc.getvalue()
            nuovo_doc = {
                "Titolo": titolo_doc,
                "Tipo": tipo_doc,
                "Categoria": categoria_doc,
                "Descrizione": descrizione_doc,
                "NomeFile": file_doc.name,
                "TipoFile": file_doc.type,
                "DimensioneKB": round(len(file_bytes)/1024, 1),
                "DataDoc": str(data_doc),
                "CaricatoDa": uploader_doc,
                "DataIns": datetime.now().strftime("%d/%m/%Y %H:%M"),
                "FileBytes": file_bytes
            }
            st.session_state.archivio_documenti.append(nuovo_doc)
            # Pulisci campi
            for k in ["arch_titolo", "arch_tipo", "arch_cat", "arch_desc", "arch_data", "arch_file"]:
                if k in st.session_state:
                    try:
                        del st.session_state[k]
                    except:
                        pass
            st.success(f"✅ Documento '{titolo_doc}' salvato! Maschera pulita per nuovo inserimento.")
            st.balloons()
            st.rerun()
        else:
            st.error("Compila Titolo Documento * e carica File *")

    st.divider()
    if st.session_state.archivio_documenti:
        st.markdown(f"#### 📁 Archivio Documenti ({len(st.session_state.archivio_documenti)}) - Tutti i file salvati")
        # Filtro categoria
        cat_filter = st.selectbox("Filtra per Categoria", ["Tutte"] + ["Generale", "Volontari", "Radio", "Mezzi", "Emergenze", "Formazione", "Amministrativo", "Sicurezza"], key="arch_filter_cat")
        docs_to_show = st.session_state.archivio_documenti
        if cat_filter != "Tutte":
            docs_to_show = [d for d in docs_to_show if d.get("Categoria")==cat_filter]
        
        for idx, doc in enumerate(docs_to_show):
            # Trova indice reale in lista completa
            real_idx = st.session_state.archivio_documenti.index(doc)
            c1, c2, c3, c4, c5 = st.columns([2, 1, 1, 1, 0.8])
            with c1:
                st.write(f"**{doc.get('Titolo','')}**")
                st.caption(f"{doc.get('NomeFile','')} - {doc.get('DimensioneKB','')} KB")
            with c2:
                st.write(f"{doc.get('Tipo','')} - {doc.get('Categoria','')}")
                st.caption(f"{doc.get('DataDoc','')}")
            with c3:
                st.write(f"{doc.get('CaricatoDa','')}")
            with c4:
                # Download
                try:
                    st.download_button(
                        f"⬇️ Scarica",
                        data=doc.get("FileBytes", b""),
                        file_name=doc.get("NomeFile", f"doc_{real_idx}.pdf"),
                        mime=doc.get("TipoFile", "application/octet-stream"),
                        use_container_width=True,
                        key=f"dl_arch_{real_idx}"
                    )
                except Exception as e:
                    st.error(f"Err dl: {e}")
            with c5:
                if st.button("🗑️", key=f"del_arch_{real_idx}", help=f"Elimina {doc.get('Titolo','')}"):
                    st.session_state.archivio_documenti.pop(real_idx)
                    st.success("Documento eliminato")
                    st.rerun()
        
        st.divider()
        # Tabella riepilogo
        df_arch = pd.DataFrame([{k:v for k,v in d.items() if "Bytes" not in k} for d in st.session_state.archivio_documenti])
        st.dataframe(df_arch, use_container_width=True)
        
        # Export Excel e PDF - SOLO EXCEL XLSX + PDF per tutti i form - Richiesta Ezio
        c_exp1, c_exp2 = st.columns(2)
        with c_exp1:
            try:
                excel_data = to_excel(df_arch)
                if excel_data and excel_data[:2] == b'PK' and len(excel_data) > 100:
                    st.download_button(
                        "⬇️ EXCEL Archivio Documenti",
                        data=excel_data,
                        file_name=f"archivio_documenti_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="exp_excel_arch",
                        type="primary"
                    )
                else:
                    st.warning("Excel non disponibile - verifica openpyxl")
            except Exception as e:
                st.error(f"Excel errore: {e}")
        with c_exp2:
            try:
                if REPORTLAB_OK:
                    pdf_data = to_pdf(df_arch, "ARCHIVIO DOCUMENTI")
                    st.download_button(
                        "📄 PDF Archivio Documenti",
                        data=pdf_data,
                        file_name=f"archivio_documenti_{datetime.now().strftime('%Y%m%d')}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                        key="exp_pdf_arch",
                        type="primary"
                    )
                else:
                    st.warning("PDF non disponibile - verifica reportlab in requirements.txt + Reboot Cloud")
            except Exception as e:
                st.error(f"PDF errore: {e}")
    else:
        pass  # istruzione rimossa
        pass
        # Istruzione rimossa - form pulito

    # Import/Export inline
    excel_import_inline("archivio_documenti", "Archivio Documenti")



elif cur == "Diplomi Attestati":
    hdr_form("DIPLOMI ATTESTATI - Bordi Tricolore Italiano - Combo Check-in")

    volontari_all = st.session_state.get("volontari", [])
    checkin_all = st.session_state.get("checkin", [])
    nomi_checkin = set()
    for c in checkin_all:
        if isinstance(c, dict):
            for k in ["volontario","nome","nominativo","volontario_nome","cognome"]:
                if k in c and c[k]:
                    val = str(c[k]).strip()
                    if len(val)>2:
                        nomi_checkin.add(val)
                    break
        elif isinstance(c, str):
            if len(c.strip())>2:
                nomi_checkin.add(c.strip())

    volontari_filtrati = []
    if nomi_checkin and volontari_all:
        for v in volontari_all:
            full = (str(v.get("Nome",""))+" "+str(v.get("Cognome",""))).strip().lower()
            nome_solo = str(v.get("Nome","")).strip().lower()
            cogn_solo = str(v.get("Cognome","")).strip().lower()
            for n in nomi_checkin:
                nl = n.lower()
                if full in nl or nl in full or nome_solo in nl or cogn_solo in nl:
                    volontari_filtrati.append(v)
                    break
        if not volontari_filtrati:
            volontari_filtrati = volontari_all
    else:
        volontari_filtrati = volontari_all

    st.info(f"Combo volontari: {len(volontari_filtrati)} filtrati da check-in ({len(checkin_all)} check-in totali) - SOLO chi ha fatto check-in")

    if "diplomi_form_version" not in st.session_state:
        st.session_state["diplomi_form_version"] = 0
    ver_dip = st.session_state.get("diplomi_form_version", 0)
    def k_dip(b): return f"{b}_v{ver_dip}"

    c1, c2 = st.columns([1, 1.3])
    with c1:
        st.markdown("#### Combo Volontari (solo check-in)")
        if volontari_filtrati:
            opzioni = [f"{v.get('Nome','')} {v.get('Cognome','')} - {v.get('Comune','')}".strip() for v in volontari_filtrati]
            sel_idx = st.selectbox("Seleziona Volontario COMBO - solo chi ha fatto check-in", range(len(volontari_filtrati)), format_func=lambda i: opzioni[i], key=k_dip("combo_vol"))
            nome_default = f"{volontari_filtrati[sel_idx].get('Nome','')} {volontari_filtrati[sel_idx].get('Cognome','')}".strip()
            comune_default = volontari_filtrati[sel_idx].get('Comune','')
        else:
            nome_default = "ALBERTO VIGANO'"
            comune_default = "CARONNO PERTUSELLA"
        nome_dip = st.text_input("Nome Volontario *", value=nome_default, key=k_dip("nome"))
        evento_dip = st.text_input("Titolo evento (\n per a capo)", value="CAMPO SCUOLA 2026\nCARONNO PERTUSELLA", key=k_dip("evento"))
        data_dip = st.text_input("Data", value="6 e 7 giugno 2026", key=k_dip("data"))
        ruolo_dip = st.text_input("Ruolo", value="VOLONTARIO DI P.C. ANA VARESE (Campo Scuola)", key=k_dip("ruolo"))
        motto1 = st.text_input("Motto riga 1", value="Insieme con Noi ...", key=k_dip("m1"))
        motto2 = st.text_input("Motto riga 2", value="Addestramento alla Protezione Civile", key=k_dip("m2"))
        capo = st.text_input("Capogruppo", value="Fiscato Stefano", key=k_dip("capo"))
        coord = st.text_input("Coordinatore", value="", key=k_dip("coord"))
        if st.button("Pulisci", key=f"pul_{ver_dip}"):
            st.session_state["diplomi_form_version"] += 1
            st.rerun()

    def get_b64(paths):
        for p in paths:
            if os.path.exists(p):
                try:
                    import base64
                    with open(p, "rb") as fh:
                        return base64.b64encode(fh.read()).decode()
                except:
                    continue
        return ""
    logo_a = get_b64(["logo.png","gruppo_caronno.png","volontario_varese.png"])
    logo_b = get_b64(["logo2.png","logo_pc.png","logo_protezione.png"])
    logo_a_html = f'<img src="data:image/png;base64,{logo_a}" style="width:92px;height:92px;object-fit:contain;">' if logo_a else '<div style="width:88px;height:88px;border:2px solid #1A5D1A;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:7px;font-weight:bold;background:white;">ANA<br>VARESE</div>'
    logo_b_html = f'<img src="data:image/png;base64,{logo_b}" style="width:92px;height:92px;object-fit:contain;">' if logo_b else '<div style="width:88px;height:88px;border:2px solid #CE2B37;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:7px;font-weight:bold;background:white;">P.C.</div>'
    ev_html = "<br>".join([l.strip() for l in evento_dip.split("\n")])

    with c2:
        st.markdown("#### Anteprima con bordi Tricolore Italiano 🇮🇹")
        html = f"""
        <div style="background:#9aa3b0;padding:18px;border-radius:6px;">
            <div style="background:white;width:700px;height:980px;margin:0 auto;box-shadow:0 0 12px rgba(0,0,0,0.4);position:relative;overflow:hidden;">
                <div style="position:absolute;top:0;left:0;right:0;bottom:0;border:10px solid #009246;z-index:10;pointer-events:none;"></div>
                <div style="position:absolute;top:10px;left:10px;right:10px;bottom:10px;border:8px solid white;z-index:11;pointer-events:none;"></div>
                <div style="position:absolute;top:18px;left:18px;right:18px;bottom:18px;border:10px solid #CE2B37;z-index:12;pointer-events:none;"></div>
                <div style="padding:36px 40px 28px 40px;height:100%;box-sizing:border-box;display:flex;flex-direction:column;position:relative;z-index:5;">
                    <div style="display:flex;justify-content:space-between;align-items:flex-start;">
                        <div style="width:92px;text-align:center;">{logo_a_html}</div>
                        <div style="flex:1;text-align:center;padding-top:6px;">
                            <div style="font-family:Arial,sans-serif;font-size:14px;font-weight:800;color:#000;line-height:1.35;">{ev_html}</div>
                            <div style="font-family:Arial,sans-serif;font-size:13px;font-weight:700;color:#000;margin-top:10px;">{data_dip}</div>
                        </div>
                        <div style="width:92px;text-align:center;">{logo_b_html}</div>
                    </div>
                    <div style="text-align:center;margin-top:52px;margin-bottom:52px;">
                        <div style="font-family:Arial Black,Arial,sans-serif;font-size:40px;font-weight:900;color:#000;letter-spacing:1px;">ATTESTATO</div>
                    </div>
                    <div style="text-align:center;margin:36px 0 54px 0;">
                        <div style="font-family:'Comic Sans MS','Segoe Script','Brush Script MT',cursive;font-size:32px;color:#000;font-style:italic;transform:rotate(-0.8deg);">{nome_dip}</div>
                    </div>
                    <div style="text-align:center;margin:20px 0 44px 0;">
                        <div style="font-family:Arial,sans-serif;font-size:18px;font-weight:700;color:#000;font-style:italic;line-height:1.5;">“{motto1} ...<br>{motto2}”</div>
                    </div>
                    <div style="text-align:center;margin:30px 0 20px 0;flex-grow:1;">
                        <div style="font-family:Arial,sans-serif;font-size:15px;color:#000;line-height:1.6;">E' stato operativo per le attività di<br><span style="font-weight:700;">{ruolo_dip}</span></div>
                    </div>
                    <div style="display:flex;justify-content:space-between;align-items:flex-end;margin-top:auto;padding:0 12px 16px 12px;">
                        <div style="text-align:left;">
                            <div style="font-family:Arial,sans-serif;font-size:10px;color:#000;">Il Capogruppo</div>
                            <div style="font-family:Arial,sans-serif;font-size:11px;font-weight:700;color:#000;font-style:italic;margin-top:2px;">{capo}</div>
                            <div style="font-family:'Segoe Script','Comic Sans MS',cursive;font-size:20px;color:#001a8a;margin-top:6px;transform:rotate(-3deg);">{capo.split()[0] if capo else 'Firma'}</div>
                        </div>
                        <div style="text-align:right;">
                            <div style="font-family:Arial,sans-serif;font-size:10px;color:#000;">Firma del Coordinatore di P.C.</div>
                            <div style="font-family:'Segoe Script','Comic Sans MS',cursive;font-size:18px;color:#444;margin-top:22px;min-height:26px;">{coord}</div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        """
        st.components.v1.html(html, height=1080, scrolling=True)

    st.divider()
    if REPORTLAB_OK:
        col1, col2 = st.columns(2)
        with col1:
            if st.button("📄 PDF Singolo Tricolore", type="primary", use_container_width=True, key=f"pdf_tri_{ver_dip}"):
                from reportlab.lib.pagesizes import A4
                from reportlab.lib.units import mm
                from reportlab.lib import colors
                from reportlab.pdfgen import canvas
                import io
                buf = io.BytesIO()
                c = canvas.Canvas(buf, pagesize=A4)
                w, h = A4
                c.setStrokeColor(colors.HexColor("#009246"))
                c.setLineWidth(7)
                c.rect(9*mm, 9*mm, w-18*mm, h-18*mm, stroke=1, fill=0)
                c.setStrokeColor(colors.white)
                c.setLineWidth(4)
                c.rect(12.5*mm, 12.5*mm, w-25*mm, h-25*mm, stroke=1, fill=0)
                c.setStrokeColor(colors.HexColor("#CE2B37"))
                c.setLineWidth(7)
                c.rect(14.5*mm, 14.5*mm, w-29*mm, h-29*mm, stroke=1, fill=0)
                c.setFillColor(colors.black)
                c.setFont("Helvetica-Bold", 11)
                y = h - 22*mm
                for line in evento_dip.split("\n"):
                    c.drawCentredString(w/2, y, line.strip())
                    y -= 5.5*mm
                y -= 3*mm
                c.setFont("Helvetica-Bold", 10)
                c.drawCentredString(w/2, y, data_dip)
                c.setFont("Helvetica-Bold", 32)
                c.drawCentredString(w/2, h/2 + 70*mm, "ATTESTATO")
                c.setFont("Helvetica-Oblique", 22)
                c.drawCentredString(w/2, h/2 + 32*mm, nome_dip)
                c.setFont("Helvetica-BoldOblique", 13)
                c.drawCentredString(w/2, h/2 + 2*mm, f'"{motto1} ...')
                c.drawCentredString(w/2, h/2 - 6*mm, f'{motto2}"')
                c.setFont("Helvetica", 11)
                c.drawCentredString(w/2, h/2 - 24*mm, "E' stato operativo per le attività di")
                c.setFont("Helvetica-Bold", 11)
                c.drawCentredString(w/2, h/2 - 31*mm, ruolo_dip)
                c.setFont("Helvetica", 8)
                c.drawString(20*mm, 36*mm, "Il Capogruppo")
                c.setFont("Helvetica-BoldOblique", 9)
                c.drawString(20*mm, 32*mm, capo)
                c.setFont("Helvetica-Oblique", 13)
                c.setFillColor(colors.HexColor("#001a8a"))
                c.drawString(20*mm, 22*mm, capo.split()[0] if capo else "Firma")
                c.setFillColor(colors.black)
                c.setFont("Helvetica", 8)
                c.drawRightString(w-20*mm, 36*mm, "Firma del Coordinatore di P.C.")
                if coord:
                    c.setFont("Helvetica-Oblique", 11)
                    c.drawRightString(w-20*mm, 22*mm, coord)
                c.showPage()
                c.save()
                buf.seek(0)
                st.download_button("⬇️ Scarica PDF Tricolore", data=buf, file_name=f"Attestato_{nome_dip.replace(' ', '_')}_Tricolore.pdf", mime="application/pdf", use_container_width=True, key=f"dl_tri_{ver_dip}")
        with col2:
            if st.button(f"📚 PDF Tutti Check-in Tricolore ({len(volontari_filtrati)})", use_container_width=True, key=f"multi_tri_{ver_dip}"):
                from reportlab.lib.pagesizes import A4
                from reportlab.lib.units import mm
                from reportlab.lib import colors
                from reportlab.pdfgen import canvas
                import io
                buf = io.BytesIO()
                c = canvas.Canvas(buf, pagesize=A4)
                w, h = A4
                for vol in volontari_filtrati:
                    nome_vol = f"{vol.get('Nome','')} {vol.get('Cognome','')}".strip()
                    if not nome_vol:
                        nome_vol = f"{vol.get('nome','')} {vol.get('cognome','')}".strip()
                    c.setStrokeColor(colors.HexColor("#009246"))
                    c.setLineWidth(7)
                    c.rect(9*mm, 9*mm, w-18*mm, h-18*mm, stroke=1, fill=0)
                    c.setStrokeColor(colors.white)
                    c.setLineWidth(4)
                    c.rect(12.5*mm, 12.5*mm, w-25*mm, h-25*mm, stroke=1, fill=0)
                    c.setStrokeColor(colors.HexColor("#CE2B37"))
                    c.setLineWidth(7)
                    c.rect(14.5*mm, 14.5*mm, w-29*mm, h-29*mm, stroke=1, fill=0)
                    c.setFillColor(colors.black)
                    c.setFont("Helvetica-Bold", 11)
                    y = h - 22*mm
                    for line in evento_dip.split("\n"):
                        c.drawCentredString(w/2, y, line.strip())
                        y -= 5.5*mm
                    y -= 3*mm
                    c.setFont("Helvetica-Bold", 10)
                    c.drawCentredString(w/2, y, data_dip)
                    c.setFont("Helvetica-Bold", 32)
                    c.drawCentredString(w/2, h/2 + 70*mm, "ATTESTATO")
                    c.setFont("Helvetica-Oblique", 22)
                    c.drawCentredString(w/2, h/2 + 32*mm, nome_vol)
                    c.setFont("Helvetica-BoldOblique", 13)
                    c.drawCentredString(w/2, h/2 + 2*mm, f'"{motto1} ...')
                    c.drawCentredString(w/2, h/2 - 6*mm, f'{motto2}"')
                    c.setFont("Helvetica", 11)
                    c.drawCentredString(w/2, h/2 - 24*mm, "E' stato operativo per le attività di")
                    c.setFont("Helvetica-Bold", 11)
                    c.drawCentredString(w/2, h/2 - 31*mm, ruolo_dip)
                    c.showPage()
                c.save()
                buf.seek(0)
                st.download_button("⬇️ Scarica Tutti Tricolore", data=buf, file_name="Attestati_Checkin_Tricolore.pdf", mime="application/pdf", use_container_width=True, key=f"dl_all_{ver_dip}")

elif cur == "Report Filtro":

    hdr_form("REPORT FILTRO - Verde ANA - Font Nero Times Bold - Tasto Filtro OK")
    st.markdown("""
    <div style="background:linear-gradient(135deg,#1A5D1A 0%,#2e7d32 100%);color:white;padding:12px;border-radius:10px;text-align:center;margin-bottom:14px;border:3px solid #FFD700;">
    <b>📊 REPORT FILTRATI - Verde ANA - Font Nero Times Bold</b>
    </div>
    """, unsafe_allow_html=True)
    FORM_KEYS_REPORT = {
        "Volontari (con foto)": "volontari",
        "Ospiti": "ospiti",
        "DB Radio": "radio_db",
        "Consegna Radio": "consegna_radio",
        "Alias Radio": "alias_radio",
        "Brogliaccio": "brogliaccio",
        "Eventi": "eventi",
        "Emergenze": "emergenze",
        "Check-in": "checkin",
        "Interventi Emergenza": "interventi",
        "Mezzi": "mezzi",
        "Attrezzature": "attrezzature",
        "Turni": "turni",
        "Verbali": "verbali",
        "Diplomi Attestati": "diplomi"
    }
    c1, c2 = st.columns([1, 1.6])
    with c1:
        sel_form_label = st.selectbox("Form per report", list(FORM_KEYS_REPORT.keys()), key="report_form_sel_final")
        sel_form_key = FORM_KEYS_REPORT[sel_form_label]
        data_list = st.session_state.get(sel_form_key, [])
        if not data_list:
            st.warning(f"Nessun dato in {sel_form_label}")
            all_fields = []; selected_fields = []; filtro_testo=""; campo_filtro="-- Nessuno --"; valore_filtro=""
        else:
            st.success(f"{len(data_list)} record")
            all_fields = set()
            for item in data_list[:20]:
                if isinstance(item, dict):
                    all_fields.update([k for k in item.keys() if "Bytes" not in k and "Foto" not in k])
            all_fields = sorted(list(all_fields))
            selected_fields = st.multiselect("Campi da includere", all_fields, default=all_fields[:8] if len(all_fields)>8 else all_fields, key="report_fields_sel_final")
            filtro_testo = st.text_input("🔍 Filtro testo libero", key="report_filtro_testo_final", placeholder="Mario, Varese...")
            campo_filtro="-- Nessuno --"; valore_filtro=""
            if selected_fields:
                campo_filtro = st.selectbox("Filtra per campo", ["-- Nessuno --"] + selected_fields, key="report_campo_filtro_final")
                if campo_filtro != "-- Nessuno --":
                    valore_filtro = st.text_input(f"Valore per '{campo_filtro}'", key="report_valore_filtro_final")
    with c2:
        if not data_list:
            st.info("👈 Seleziona form con dati")
        else:
            # Tasto filtro esplicito - fix Ezio
            if st.button("🔍 Applica Filtro", key="btn_applica_filtro_report", type="primary", use_container_width=True):
                st.session_state["filtro_attivo"] = True
            filtered = data_list
            if filtro_testo:
                filtered = [r for r in filtered if isinstance(r, dict) and any(filtro_testo.lower() in str(v).lower() for v in r.values())]
            if campo_filtro != "-- Nessuno --" and valore_filtro:
                filtered = [r for r in filtered if isinstance(r, dict) and valore_filtro.lower() in str(r.get(campo_filtro,"")).lower()]
            st.success(f"Filtrati: {len(filtered)} su {len(data_list)}")
            filtered_display = [{k: r.get(k,"") for k in selected_fields} if isinstance(r, dict) else {"Valore": r} for r in filtered] if selected_fields else filtered
            st.markdown(f"#### 📄 Anteprima - {len(filtered_display)} record")
            if filtered_display:
                df_report = pd.DataFrame(filtered_display)
                st.dataframe(df_report, use_container_width=True, height=400)
                col1, col2, col3 = st.columns(3)
                with col1:
                    try:
                        excel_bytes = to_excel_bytes({sel_form_label[:25]: df_report})
                        st.download_button(f"⬇️ Excel ({len(df_report)})", data=excel_bytes, file_name=f"report_{sel_form_key}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="dl_excel_report_final")
                    except Exception as e: st.error(f"Excel: {e}")
                with col2:
                    try:
                        csv_data = df_report.to_csv(index=False).encode('utf-8')
                        st.download_button("⬇️ CSV", data=csv_data, file_name=f"report_{sel_form_key}.csv", mime="text/csv", use_container_width=True, key="dl_csv_report_final")
                    except Exception as e: st.error(f"CSV: {e}")
                with col3:
                    if REPORTLAB_OK:
                        try:
                            from reportlab.lib.pagesizes import landscape, A4
                            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
                            from reportlab.lib import colors
                            import io
                            buf = io.BytesIO()
                            doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=1*cm, rightMargin=1*cm)
                            df_pdf = df_report.copy()
                            for col in df_pdf.columns: df_pdf[col] = df_pdf[col].astype(str).apply(lambda x: x[:35])
                            cols = list(df_pdf.columns)[:12]
                            data = [cols] + df_pdf[cols].values.tolist()[:60]
                            t = Table(data, repeatRows=1)
                            t.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1A5D1A")), ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke), ('FONTSIZE', (0,0), (-1,-1), 6), ('GRID', (0,0), (-1,-1), 0.4, colors.grey)]))
                            from reportlab.platypus import Paragraph, Spacer
                            from reportlab.lib.styles import getSampleStyleSheet
                            styles = getSampleStyleSheet()
                            story = [Paragraph(f"Report {sel_form_label} - {len(filtered_display)} record", styles['Normal']), Spacer(1,12), t]
                            doc.build(story); buf.seek(0)
                            st.download_button(f"⬇️ PDF ({len(df_report)})", data=buf.getvalue(), file_name=f"report_{sel_form_key}.pdf", mime="application/pdf", use_container_width=True, key="dl_pdf_report_final")
                        except Exception as e: st.error(f"PDF: {e}")



elif cur == "Spese ODV per Evento":
    hdr_form("💰 SPESE ODV PER EVENTO / EMERGENZA - Registro Spese Sostenute")

    edit_mode_spese = False
    edit_data_spese = {}
    if st.session_state.get("spese_edit_index") is not None:
        try:
            edit_data_spese = st.session_state.spese_odv[st.session_state.spese_edit_index]
            edit_mode_spese = True
        except:
            edit_data_spese = {}
            edit_mode_spese = False

    if edit_mode_spese:
        st.warning(f"✏️ Modifica Spesa: {edit_data_spese.get('Evento','')} - {edit_data_spese.get('ODV','')} - € {edit_data_spese.get('Importo','')}")

    lista_eventi = []
    try:
        if st.session_state.get("eventi"):
            for e in st.session_state.eventi:
                nome = e.get('Nome','') or e.get('Titolo','') or e.get('Descrizione','') or 'Evento senza nome'
                data = e.get('Data','')
                lista_eventi.append(f"EVENTO: {nome} - {data}")
        if st.session_state.get("emergenze"):
            for e in st.session_state.emergenze:
                nome = e.get('Nome','') or e.get('Titolo','') or e.get('Descrizione','') or 'Emergenza senza nome'
                data = e.get('Data','')
                lista_eventi.append(f"EMERGENZA: {nome} - {data}")
        if st.session_state.get("interventi"):
            for e in st.session_state.interventi:
                nome = e.get('Descrizione','')[:50] or e.get('Nome','') or 'Intervento'
                data = e.get('Data','')
                lista_eventi.append(f"INTERVENTO: {nome} - {data}")
    except:
        pass
    if not lista_eventi:
        lista_eventi = ["Evento Generico", "Emergenza Generica", "Esercitazione", "Manutenzione", "Altro"]

    c1, c2, c3 = st.columns(3)
    with c1:
        evento_combo = st.selectbox("Evento / Emergenza * (combo)", lista_eventi + ["Altro (manuale)"], key="spese_evento_sel")
        if evento_combo == "Altro (manuale)":
            evento_nome_manuale = st.text_input("Nome Evento Manuale *", key="spese_evento_manuale")
            evento_sel = evento_nome_manuale
        else:
            evento_sel = evento_combo
        st.text_input("Nome Evento/Emergenza (blindato - spesa attribuita a questo)", value=evento_sel if evento_sel else "", key="spese_evento_blindato", disabled=True, help="Campo blindato - attribuisce la spesa all'evento/emergenza scelto nella combo sopra")
        data_spesa = st.date_input("Data Spesa *", value=date.today(), format="DD/MM/YYYY", key="spese_data")
        odv_list = ["ANA Varese", "ANA Sez. Varese", "Protezione Civile Varese", "Croce Rossa", "Alpini", "Altro"]
        odv_sel = st.selectbox("ODV che ha sostenuto spesa *", odv_list, key="spese_odv_sel")
        if odv_sel == "Altro":
            odv_sel = st.text_input("Altra ODV", key="spese_odv_altro")
    with c2:
        tipo_spesa = st.selectbox("Tipo Spesa *", ["Carburante", "Pasti / Vitto", "Alloggio", "Trasporto", "Materiali", "Attrezzature", "Noleggio Mezzi", "Comunicazioni", "Altro"], key="spese_tipo")
        importo = st.number_input("Importo € *", min_value=0.0, step=0.5, format="%.2f", key="spese_importo")
        modalita_pag = st.selectbox("Modalità Pagamento", ["Contanti", "Carta", "Bonifico", "Fattura da pagare"], key="spese_modalita")
    with c3:
        pagato_da = st.text_input("Pagato da (Nome)", value=edit_data_spese.get("PagatoDa",""), key="spese_pagato_da")
        rimborsato = st.selectbox("Rimborsato?", ["No", "Sì", "In attesa", "Parziale"], key="spese_rimborsato")
        note_spesa = st.text_area("Note / Descrizione", value=edit_data_spese.get("Note",""), key="spese_note")
        scontrino_file = st.file_uploader("Allega Scontrino / Fattura (jpg, png, pdf)", type=["jpg","jpeg","png","pdf"], key="spese_scontrino")
        scontrino_bytes = None
        if scontrino_file:
            scontrino_bytes = scontrino_file.getvalue()
            if scontrino_file.type != "application/pdf":
                st.image(scontrino_bytes, width=150)
            else:
                st.success("PDF allegato")

    st.divider()
    if st.button("💾 Salva Spesa ODV", type="primary", use_container_width=True, key="btn_save_spese"):
        if not evento_sel or importo <= 0:
            st.error("Evento e Importo obbligatori")
        else:
            nuova_spesa = {
                "Evento": evento_sel,
                "DataSpesa": str(data_spesa),
                "ODV": odv_sel,
                "TipoSpesa": tipo_spesa,
                "Importo": importo,
                "ModalitaPagamento": modalita_pag,
                "PagatoDa": pagato_da,
                "Rimborsato": rimborsato,
                "Note": note_spesa,
                "ScontrinoBytes": scontrino_bytes,
                "Timestamp": datetime.now().strftime("%d/%m/%Y %H:%M")
            }
            if edit_mode_spese:
                st.session_state.spese_odv[st.session_state.spese_edit_index] = nuova_spesa
                st.session_state.spese_edit_index = None
                st.success("✅ Spesa modificata!")
            else:
                st.session_state.spese_odv.append(nuova_spesa)
                st.success("✅ Spesa aggiunta!")
            st.rerun()

    st.divider()
    if st.session_state.get("spese_odv"):
        totale = sum([float(s.get("Importo",0)) for s in st.session_state.spese_odv])
        st.markdown(f"#### 📋 Elenco Spese ODV ({len(st.session_state.spese_odv)}) - Totale: € {totale:.2f}")
        c_tot1, c_tot3 = st.columns(2)
        with c_tot1:
            st.metric("💰 Totale Spese", f"€ {totale:.2f}")
        with c_tot3:
            st.info(f"Numero spese: {len(st.session_state.spese_odv)}")
    else:
        st.markdown("#### 📋 Elenco Spese ODV (0) - Totale: € 0.00")
        c_tot1, c_tot3 = st.columns(2)
        with c_tot1:
            st.metric("💰 Totale Spese", "€ 0.00")
        with c_tot3:
            st.info("Nessuna spesa registrata")

    if st.session_state.get("spese_odv"):
        import pandas as pd
        df_spese = pd.DataFrame([{k:v for k,v in s.items() if "Bytes" not in k} for s in st.session_state.spese_odv])
        st.dataframe(df_spese, use_container_width=True)
        try:
            df_group_evento = df_spese.groupby("Evento")["Importo"].sum().reset_index()
            st.markdown("##### 📊 Totale per Evento")
            st.dataframe(df_group_evento, use_container_width=True)
            df_group_odv = df_spese.groupby("ODV")["Importo"].sum().reset_index()
            st.markdown("##### 📊 Totale per ODV")
            st.dataframe(df_group_odv, use_container_width=True)
        except:
            pass
        if st.button("📥 Esporta Spese in Excel", key="btn_export_spese"):
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_spese.to_excel(writer, index=False, sheet_name="Spese ODV")
            st.download_button("⬇️ Scarica Excel Spese", data=output.getvalue(), file_name="spese_odv.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        # PDF SPECIALE con totale in cella + filtro per evento blindato + casella in basso
        if REPORTLAB_OK and st.session_state.get("spese_odv"):
            try:
                from reportlab.lib.pagesizes import landscape, A4
                from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
                from reportlab.lib.styles import getSampleStyleSheet
                from reportlab.lib import colors
                from reportlab.lib.units import cm
                
                # Determina evento filtrato dal campo blindato (combo)
                try:
                    evento_filtro = evento_sel
                except:
                    evento_filtro = ""
                
                # Prepara due PDF: uno completo e uno filtrato per evento blindato
                for tipo_report in ["completo", "filtrado"]:
                    if tipo_report == "filtrado":
                        if not evento_filtro or evento_filtro == "Altro (manuale)":
                            continue
                        df_report = df_spese[df_spese["Evento"] == evento_filtro].copy() if "Evento" in df_spese.columns else df_spese.copy()
                        if df_report.empty:
                            continue
                        totale_report = df_report["Importo"].sum() if "Importo" in df_report.columns else totale
                        titolo_report = f"SPESE ODV - {evento_filtro}"
                        filename_report = f"spese_{evento_filtro[:20].replace(' ', '_')}_totale.pdf"
                        label_btn = f"📄 PDF Spese per: {evento_filtro[:40]} - con casella Totale"
                    else:
                        df_report = df_spese
                        totale_report = totale
                        titolo_report = "SPESE ODV PER EVENTO - ANA Varese - TUTTE"
                        filename_report = "spese_odv_tutte_con_totale.pdf"
                        label_btn = "📄 PDF Spese TUTTE con Totale (casella in basso)"
                    
                    buf_spese = BytesIO()
                    doc_spese = SimpleDocTemplate(buf_spese, pagesize=landscape(A4), leftMargin=1*cm, rightMargin=1*cm, topMargin=1.5*cm, bottomMargin=1*cm)
                    styles_spese = getSampleStyleSheet()
                    story_spese = []
                    try:
                        if os.path.exists("logo.png"):
                            logo_img_spese = Image("logo.png", width=70, height=70)
                        else:
                            logo_img_spese = Paragraph("", styles_spese["Normal"])
                    except:
                        logo_img_spese = Paragraph("", styles_spese["Normal"])
                    title_spese = Paragraph(f"<b>{titolo_report}</b><br/><font size=9>Generato il {datetime.now().strftime('%d/%m/%Y %H:%M')} - Totale: € {totale_report:.2f}</font>", styles_spese["Title"])
                    header_spese = Table([[logo_img_spese, title_spese]], colWidths=[3*cm, landscape(A4)[0] - 5*cm])
                    header_spese.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"MIDDLE")]))
                    story_spese.append(header_spese)
                    story_spese.append(Spacer(1,16))
                    cols_spese = list(df_report.columns)[:12]
                    data_spese_pdf = [cols_spese]
                    for _, r in df_report.iterrows():
                        data_spese_pdf.append([str(r.get(c,""))[:120] for c in cols_spese])
                    # Calcola larghezza per far occupare tutto il foglio
                    available_width_spese = landscape(A4)[0] - 2*cm
                    col_width_spese = available_width_spese / len(cols_spese) if cols_spese else available_width_spese
                    col_widths_spese = [col_width_spese] * len(cols_spese)
                    # Riga totale
                    totale_row = [""] * len(cols_spese)
                    try:
                        idx_importo = cols_spese.index("Importo")
                        totale_row[idx_importo] = f"TOTALE: € {totale_report:.2f}"
                        totale_row[0] = "TOTALE SPESE"
                    except:
                        totale_row[-1] = f"TOTALE: € {totale_report:.2f}"
                    data_spese_pdf.append(totale_row)
                    try:
                        t_spese = Table(data_spese_pdf, colWidths=col_widths_spese, repeatRows=1)
                    except:
                        t_spese = Table(data_spese_pdf, repeatRows=1)
                    style_spese = [
                        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1A5D1A")),
                        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                        ("FONTSIZE", (0, 0), (-1, -1), 7),
                        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                        ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, colors.HexColor("#e8f5e9")]),
                        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#FFD700")),
                        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
                        ("FONTSIZE", (0, -1), (-1, -1), 9),
                    ]
                    t_spese.setStyle(TableStyle(style_spese))
                    story_spese.append(t_spese)
                    story_spese.append(Spacer(1,20))
                    # CASELLA IN BASSO CON TOTALE - richiesta Ezio
                    casella_totale = [
                        [Paragraph(f"<b>CASELLA TOTALE - {evento_filtro if tipo_report=='filtrado' else 'TUTTE LE SPESE'}</b>", styles_spese["Normal"]), 
                         Paragraph(f"<b>€ {totale_report:.2f}</b>", styles_spese["Title"])],
                        [Paragraph(f"Numero spese: {len(df_report)}", styles_spese["Normal"]), 
                         Paragraph(f"Evento: {evento_filtro if tipo_report=='filtrado' else 'Tutti'}", styles_spese["Normal"])]
                    ]
                    box_totale = Table(casella_totale, colWidths=[landscape(A4)[0]/2 - 1*cm, landscape(A4)[0]/2 - 1*cm])
                    box_totale.setStyle(TableStyle([
                        ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#FFF9C4")),
                        ("BOX", (0,0), (-1,-1), 2, colors.HexColor("#1A5D1A")),
                        ("GRID", (0,0), (-1,-1), 1, colors.HexColor("#FFD700")),
                        ("FONTSIZE", (0,0), (-1,-1), 10),
                        ("ALIGN", (1,0), (1,0), "RIGHT"),
                        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                        ("TOPPADDING", (0,0), (-1,-1), 8),
                        ("BOTTOMPADDING", (0,0), (-1,-1), 8),
                    ]))
                    story_spese.append(box_totale)
                    doc_spese.build(story_spese)
                    buf_spese.seek(0)
                    st.download_button(label_btn, data=buf_spese.getvalue(), file_name=filename_report, mime="application/pdf", use_container_width=True, key=f"btn_pdf_{tipo_report}_{evento_filtro[:10]}")
            except Exception as e:
                st.error(f"Errore PDF spese: {e}")
                import traceback
                st.code(traceback.format_exc())

    excel_import_inline("spese_odv", "Spese ODV per Evento")

el
if cur == "Splash":
    hdr_form("🎬 SPLASH - Presentazione ANA - Come primo file")
    st.markdown("""
    <div style="text-align:center;padding:40px;background:linear-gradient(135deg,#1A5D1A,#2e7d32);border-radius:16px;color:white;border:3px solid #FFD700;">
    <div style="font-size:80px;">🛡️</div>
    <h1>ANA Varese</h1>
    <h2>Squadra Volontari di Protezione Civile</h2>
    <h3>Gruppo Alpini di Caronno Pertusella Bariola</h3>
    <p>Gestionale ODV - Versione 2026 - UNICO</p>
    </div>
    """, unsafe_allow_html=True)
    st.write("")
    col1,col2,col3 = st.columns([1,2,1])
    with col2:
        try:
            if os.path.exists("copertina.png"):
                st.image("copertina.png", width=400)
            elif os.path.exists("logo.png"):
                st.image("logo.png", width=250)
        except:
            pass
        tempo = st.slider("Durata splash (secondi)", 3, 15, 5, key="splash_dur")
        if st.button("🎬 MOSTRA SPLASH POPUP (come primo file)", type="primary", use_container_width=True):
            try:
                inject_popout_splash()
                st.balloons()
                st.success(f"Splash {tempo}s - ESC per chiudere")
            except Exception as e:
                st.error(f"Errore splash: {e}")
        if st.button("🏠 Torna alla Dashboard", use_container_width=True):
            st.session_state.menu = "Dashboard"
            st.rerun()

if cur == "Eventi":
    hdr_form("📅 EVENTI - Form ripristinato - Era mancante")
    with st.form("eventi_form_unico"):
        c1,c2 = st.columns(2)
        with c1:
            nome_ev = st.text_input("Nome Evento")
            data_ev = st.date_input("Data Evento")
            luogo_ev = st.text_input("Luogo")
        with c2:
            tipo_ev = st.selectbox("Tipo Evento", ["Esercitazione","Emergenza","Manifestazione","Formazione","Altro"])
            note_ev = st.text_area("Note Evento")
        if st.form_submit_button("💾 Salva Evento", type="primary", use_container_width=True):
            st.session_state.setdefault("eventi", []).append({
                "nome": nome_ev, "data": str(data_ev), "luogo": luogo_ev,
                "tipo": tipo_ev, "note": note_ev, "ts": str(datetime.now())
            })
            st.success("Evento salvato!")
            st.rerun()
    if st.session_state.get("eventi"):
        st.dataframe(pd.DataFrame(st.session_state.eventi), use_container_width=True)

if cur == "Statistiche":
    hdr_form("📊 STATISTICHE - Form aggiunto")
    spese = st.session_state.get("spese_odv", [])
    note = st.session_state.get("note_spese", [])
    eventi = st.session_state.get("eventi", [])
    try:
        tot_spese = sum([float(x.get("importo",0)) if isinstance(x, dict) else 0 for x in spese]) if spese else 0
    except:
        tot_spese = 0
    try:
        tot_note = sum([float(x.get("importo",0)) if isinstance(x, dict) else 0 for x in note]) if note else 0
    except:
        tot_note = 0
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Tot Spese ODV", f"€ {tot_spese:.2f}")
    c2.metric("Tot Note Spese", f"€ {tot_note:.2f}")
    c3.metric("Totale", f"€ {tot_spese+tot_note:.2f}")
    c4.metric("Eventi", len(eventi))
    if spese or note:
        try:
            df_stat = pd.DataFrame([{"tipo":"Spese ODV","importo":tot_spese},{"tipo":"Note Spese","importo":tot_note}])
            st.bar_chart(df_stat.set_index("tipo"))
        except:
            pass

if cur == "Note Spese":
    hdr_form("🧾 NOTE SPESE - Rimborso Volontari")
    with st.form("note_spese_form_unico"):
        c1,c2,c3 = st.columns(3)
        with c1:
            data_ns = st.date_input("Data Spesa")
            volontario_ns = st.text_input("Volontario")
            evento_ns = st.text_input("Evento")
        with c2:
            tipo_ns = st.selectbox("Tipo Spesa", ["Carburante","Pedaggio","Vitto","Materiale","Parcheggio","Altro"])
            importo_ns = st.number_input("Importo €", min_value=0.0, step=0.10)
            pagato_ns = st.selectbox("Pagato da", ["Volontario","ODV","Anticipo"])
        with c3:
            scontrino_ns = st.file_uploader("Scontrino", type=["png","jpg","pdf"])
            note_ns = st.text_area("Note")
        if st.form_submit_button("💾 Salva Nota Spesa", type="primary", use_container_width=True):
            st.session_state.setdefault("note_spese", []).append({
                "data": str(data_ns), "volontario": volontario_ns, "evento": evento_ns,
                "tipo": tipo_ns, "importo": importo_ns, "pagato_da": pagato_ns, "note": note_ns, "ts": str(datetime.now())
            })
            st.success("Salvata!")
            st.rerun()
    if st.session_state.get("note_spese"):
        st.dataframe(pd.DataFrame(st.session_state.note_spese), use_container_width=True)


if cur == "Backup":
    hdr_form("BACKUP")

    FORM_KEYS = {
        "Volontari (con foto)": "volontari",
        "DB Radio": "radio_db",
        "Consegna Radio": "consegna_radio",
        "Alias Radio": "alias_radio",
        "Brogliaccio": "brogliaccio",
        "Eventi": "eventi",
        "Emergenze": "emergenze",
        "Check-in": "checkin",
        "Interventi Emergenza": "interventi",
        "Tabella Interventi Emergenza": "tabella_interventi",
        "Mezzi": "mezzi",
        "Attrezzature": "attrezzature",
        "Mappe Postazioni": "mappa_avanzata_markers",
        "Libreria Icone": "icone",
        "Turni": "turni",
        "Chat": "chat",
        "Posizioni PD785": "posizioni_pd785",
        "Posizioni Anytone": "posizioni_anytone",
        "Archivio Documenti": "archivio_documenti",
        "Spese ODV per Evento": "spese_odv"
    }

    st.markdown("""
    <div style="background:#e8f5e9;padding:10px;border-radius:8px;border-left:4px solid #1A5D1A;margin-bottom:12px;">
    <b>Template Excel per ODV</b> - Invia file vuoto, ODV compila, importi in Volontari<br>
    <small>Backup Totale | Template ODV | Import Multi-Foglio | Solo Excel</small>
    </div>
    """, unsafe_allow_html=True)

    # Riepilogo
    cols = st.columns(4)
    tot_records = 0
    for i, (label, key) in enumerate(FORM_KEYS.items()):
        cnt = len(st.session_state.get(key, []))
        tot_records += cnt
        cols[i % 4].metric(label[:18], cnt)
    st.metric("Totale Record", tot_records)

    st.divider()

    # === SEZIONE 1: TEMPLATE EXCEL PER ODV - VOLONTARI ===
    st.markdown("### 📋 TEMPLATE EXCEL PER ODV - Volontari")
    # Istruzione rimossa - form pulito

    def get_volontari_template_df():
        # Template OFFICE 2016 COMPATIBILE - solo header, no righe esempio che danno errore formato
        columns = [
            "Nome", "Cognome", "Comune", "Via", "CapoODV", "ODVAppartenenza",
            "DataNascita", "CodFisc", "Cellulare", "Email", "TelEmergenza",
            "Ruolo", "Squadra", "RadioID", "Documento", "ScadDoc", "Note"
        ]
        # Office 2016 FIX: DataFrame vuoto solo con colonne, niente righe esempio
        # Office 2016 da errore "formato non valido" se ci sono righe con tipi misti esempio
        return pd.DataFrame(columns=columns)

    c_t1, c_t2 = st.columns(2)
    with c_t1:
        df_template_vol = get_volontari_template_df()
        st.dataframe(df_template_vol, use_container_width=True)
        st.download_button(
            "📥 Scarica TEMPLATE Volontari per ODV (Excel vuoto + esempio)",
            data=to_excel(df_template_vol),
            file_name="TEMPLATE_Volontari_ODV_da_compilare.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
            key="download_template_vol_odv"
        )
    with c_t2:
        st.markdown("""
        **Istruzioni per ODV:**
        1. Scarica template a sinistra
        2. Compila righe: Nome, Cognome, Comune, CapoODV, Cellulare...
        3. Data formato: gg/mm/aaaa (es: 15/06/1985)
        4. Lascia prima riga esempio o cancellala
        5. Salva e rimanda file a te
        6. Tu carichi file sotto in "Import Template ODV"
        
        **Campi obbligatori:** Nome, Cognome, CapoODV
        """)
        # Template anche per altri form
        sel_template_other = st.selectbox("Scarica Template altro Form", ["--"] + list(FORM_KEYS.keys()), key="sel_template_other")
        if sel_template_other != "--":
            key_other = FORM_KEYS[sel_template_other]
            data_other = st.session_state.get(key_other, [])
            if data_other:
                df_other = pd.DataFrame([{k:v for k,v in r.items() if "Bytes" not in k and "Foto" not in k and "File" not in k} for r in data_other[:1]])
                if df_other.empty:
                    df_other = pd.DataFrame(columns=["Col1","Col2"])
            else:
                # Template vuoto con colonne generiche
                df_other = pd.DataFrame(columns=["Campo1","Campo2","Note"])
            st.download_button(
                f"📥 Template {sel_template_other}",
                data=to_excel(df_other),
                file_name=f"TEMPLATE_{key_other}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"tpl_{key_other}"
            )

    st.divider()

    # Backup Totale Excel
    st.markdown("### 💾 Backup Totale Excel")
    c1, c2, c3 = st.columns(3)
    with c1:
        try:
            datasets = {}
            for label, key in FORM_KEYS.items():
                data = st.session_state.get(key, [])
                if data:
                    clean = [{kk: vv for kk, vv in r.items() if "Bytes" not in kk and "Foto" not in kk and "File" not in kk} for r in data if isinstance(r, dict)]
                    if clean:
                        datasets[label[:31]] = pd.DataFrame(clean)
            if datasets:
                st.download_button("⬇️ Backup Totale Excel", data=to_excel_multi(datasets), file_name=f"backup_totale_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary", key="backup_tot_excel")
            else:
                st.info("Nessun dato")
        except Exception as e:
            st.error(f"Excel: {e}")
    with c2:
        if REPORTLAB_OK:
            try:
                df_summary = pd.DataFrame([{"Form": label, "Record": len(st.session_state.get(key, []))} for label, key in FORM_KEYS.items()])
                st.download_button("⬇️ PDF Riepilogo", data=to_pdf(df_summary, "BACKUP TOTALE"), file_name="backup_riepilogo.pdf", mime="application/pdf", use_container_width=True, key="backup_pdf")
            except:
                pass
    with c3:
        if st.button("🗑️ Azzera Tutto", use_container_width=True, key="azzera_backup_unico"):
            for k in FORM_KEYS.values():
                st.session_state[k] = []
            st.success("Azzerati")
            st.rerun()

    st.divider()

    # Export Singolo Excel
    st.markdown("### 📄 Export Singolo Form - Solo Excel")
    sel_label = st.selectbox("Seleziona Form per Export Excel", list(FORM_KEYS.keys()), key="backup_sel_form_unico")
    sel_key = FORM_KEYS[sel_label]
    sel_data = st.session_state.get(sel_key, [])
    st.metric(f"Record in {sel_label}", len(sel_data))
    if sel_data:
        df_sel = pd.DataFrame([{k:v for k,v in r.items() if "Bytes" not in k and "Foto" not in k and "File" not in k} for r in sel_data if isinstance(r, dict)])
        st.dataframe(df_sel.head(20), use_container_width=True)
        c1, c2 = st.columns(2)
        with c1:
            st.download_button(f"⬇️ Excel {sel_label}", data=to_excel(df_sel), file_name=f"{sel_key}_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key=f"exp_excel_{sel_key}_unico")
        with c2:
            if REPORTLAB_OK:
                st.download_button(f"📄 PDF {sel_label}", data=to_pdf(df_sel, sel_label.upper()), file_name=f"{sel_key}.pdf", mime="application/pdf", use_container_width=True, key=f"exp_pdf_{sel_key}_unico")
    else:
        st.warning(f"{sel_label} vuoto")

    st.divider()

    # === IMPORT TEMPLATE ODV - NUOVO - PER VOLONTARI ===
    st.markdown("### 📥 IMPORT TEMPLATE ODV - Volontari da Excel")
    st.success("Carica qui il file Excel compilato dalle ODV - Importa tutti i volontari in un click senza inserire uno per uno!")

    sel_label_imp_odv = st.selectbox("Form destinazione", ["Volontari (con foto)"] + list(FORM_KEYS.keys()), index=0, key="import_odv_dest")
    sel_key_imp_odv = FORM_KEYS.get(sel_label_imp_odv, "volontari")
    up_mode_odv = st.radio("Modalità", ["Aggiungi a esistenti", "Sostituisci tutto"], key="up_mode_odv", horizontal=True)
    up_file_odv = st.file_uploader(f"Carica Excel ODV compilato per {sel_label_imp_odv}", type=["xlsx", "xls"], key="up_odv_excel")

    if up_file_odv:
        try:
            df_odv = None
            last_err = ""
            # Prova lettura con engine automatico
            for eng in [None, "openpyxl", "xlrd"]:
                try:
                    up_file_odv.seek(0)
                    if eng is None:
                        df_odv = pd.read_excel(up_file_odv)
                    else:
                        df_odv = pd.read_excel(up_file_odv, engine=eng)
                    if df_odv is not None and not df_odv.empty:
                        break
                except Exception as e:
                    last_err = str(e)
                    continue

            if df_odv is not None and not df_odv.empty:
                # Pulisci colonne vuote e righe vuote
                df_odv = df_odv.dropna(how='all')
                # Rimuovi righe dove Nome e Cognome vuoti
                if "Nome" in df_odv.columns and "Cognome" in df_odv.columns:
                    df_odv = df_odv[~(df_odv["Nome"].astype(str).str.strip().isin(["", "nan", "None"]) & df_odv["Cognome"].astype(str).str.strip().isin(["", "nan", "None"]))]
                # Rimuovi riga esempio se c'è
                df_odv = df_odv[~((df_odv.astype(str).apply(lambda x: x.str.contains("Esempio", na=False)).any(axis=1)) | (df_odv.astype(str).apply(lambda x: x.str.contains("gg/mm/aaaa", na=False)).any(axis=1)))]

                st.success(f"✅ {len(df_odv)} volontari trovati nel file Excel ODV")
                st.dataframe(df_odv.head(30), use_container_width=True)

                # Mappatura colonne -> campi volontari
                # Converte DataNascita in formato gg/mm/aaaa se necessario
                if st.button(f"✅ IMPORTA {len(df_odv)} VOLONTARI IN {sel_label_imp_odv}", type="primary", use_container_width=True, key="btn_import_odv_vol"):
                    imported_list = []
                    for _, row in df_odv.iterrows():
                        rec = {}
                        for col in df_odv.columns:
                            val = row[col]
                            # Salta NaN
                            if pd.isna(val):
                                continue
                            # Converte date
                            if "Data" in col or "Scad" in col:
                                try:
                                    if isinstance(val, (pd.Timestamp, datetime, date)):
                                        rec[col] = val.strftime("%d/%m/%Y")
                                    else:
                                        rec[col] = str(val).strip()
                                except:
                                    rec[col] = str(val)
                            else:
                                rec[col] = str(val).strip() if isinstance(val, str) else val
                        # Normalizza chiavi comuni
                        # Mappa CodFisc varianti
                        if "CodFisc" not in rec:
                            for k in ["CodiceFiscale","CF","Codice Fiscale"]:
                                if k in rec:
                                    rec["CodFisc"] = rec.pop(k)
                        # Aggiungi campi default se mancano
                        if "ODVAppartenenza" not in rec:
                            rec["ODVAppartenenza"] = "ANA Varese"
                        if "Ruolo" not in rec:
                            rec["Ruolo"] = "Volontario"
                        if "Squadra" not in rec:
                            rec["Squadra"] = "Squadra A"
                        if "Comune" not in rec:
                            rec["Comune"] = "Varese"
                        # Solo se ha Nome o Cognome
                        if rec.get("Nome") or rec.get("Cognome"):
                            imported_list.append(rec)

                    if up_mode_odv.startswith("Sostituisci"):
                        st.session_state[sel_key_imp_odv] = imported_list
                    else:
                        st.session_state[sel_key_imp_odv] = st.session_state.get(sel_key_imp_odv, []) + imported_list

                    st.success(f"🎉 Importati {len(imported_list)} volontari in {sel_label_imp_odv}!")
                    st.balloons()
                    st.rerun()
            elif df_odv is not None:
                st.warning("File Excel vuoto o solo intestazioni")
            else:
                st.error(f"Errore lettura Excel: {last_err}")
                st.error("Verifica che file sia .xlsx valido e che requirements.txt contenga openpyxl, xlrd")
        except Exception as e:
            st.error(f"Errore import ODV: {e}")
            import traceback
            st.code(traceback.format_exc())

    st.divider()

    # Import Singolo - Solo Excel - FIX DEFINITIVO - Mantenuto per compatibilità
    st.markdown("### 📥 Import Singolo Form - Solo Excel xlsx/xls (Generico)")
    sel_label_imp = st.selectbox("Seleziona Form per Import Excel", list(FORM_KEYS.keys()), key="import_sel_form_excel")
    sel_key_imp = FORM_KEYS[sel_label_imp]
    up_mode_single = st.radio("Modalità Import Singolo Excel", ["Aggiungi", "Sostituisci"], key="up_mode_single_excel", horizontal=True)
    up_file_single = st.file_uploader(f"Carica Excel per {sel_label_imp} - Solo xlsx/xls", type=["xlsx", "xls"], key="up_single_excel")
    if up_file_single:
        try:
            df_imp = None
            last_err = ""
            for eng in [None, "openpyxl", "xlrd"]:
                try:
                    up_file_single.seek(0)
                    if eng is None:
                        df_imp = pd.read_excel(up_file_single)
                    else:
                        df_imp = pd.read_excel(up_file_single, engine=eng)
                    if df_imp is not None and len(df_imp.columns) > 0:
                        break
                except Exception as e:
                    last_err = str(e)
                    continue

            if df_imp is not None and not df_imp.empty:
                imported = df_imp.to_dict(orient="records")
                st.success(f"{len(imported)} record letti da Excel - OK")
                st.dataframe(pd.DataFrame(imported).head(10), use_container_width=True)
                if st.button(f"✅ Importa Excel in {sel_label_imp}", type="primary", use_container_width=True, key=f"btn_import_excel_{sel_key_imp}"):
                    if up_mode_single == "Sostituisci":
                        st.session_state[sel_key_imp] = imported
                    else:
                        st.session_state[sel_key_imp] = st.session_state.get(sel_key_imp, []) + imported
                    st.success(f"Importato {len(imported)} record in {sel_label_imp}")
                    st.rerun()
            elif df_imp is not None:
                st.warning("File Excel vuoto")
            else:
                st.error(f"Errore lettura Excel: {last_err}")
                if "openpyxl" in last_err.lower():
                    st.error("⚠️ openpyxl non installato - Controlla requirements.txt e Reboot Cloud")
        except Exception as e:
            st.error(f"Errore import Excel: {e}")

    st.divider()

    # Import Totale - Solo Excel Multi-foglio - Un file con tanti fogli
    st.markdown("### 📥 Import Backup Totale - Solo Excel xlsx/xls Multi-fogli")
    # Istruzione rimossa - form pulito
    up_total_excel = st.file_uploader("Carica Backup Totale Excel - Solo xlsx/xls", type=["xlsx", "xls"], key="up_total_excel")
    if up_total_excel:
        try:
            xls = None
            last_err = ""
            for eng in [None, "openpyxl", "xlrd"]:
                try:
                    up_total_excel.seek(0)
                    if eng is None:
                        xls = pd.ExcelFile(up_total_excel)
                    else:
                        xls = pd.ExcelFile(up_total_excel, engine=eng)
                    if xls is not None:
                        break
                except Exception as e:
                    last_err = str(e)
                    continue

            if xls is not None:
                st.write(f"Fogli trovati: {xls.sheet_names}")
                for sh in xls.sheet_names:
                    try:
                        df_preview = pd.read_excel(xls, sheet_name=sh)
                        st.write(f"**{sh}**: {len(df_preview)} righe")
                    except:
                        pass
                mode_total = st.radio("Modalità Import Totale Excel", ["Aggiungi", "Sostituisci"], key="mode_total_excel", horizontal=True)
                if st.button("✅ CONFERMA IMPORT TOTALE EXCEL MULTI-FOGLIO", type="primary", use_container_width=True, key="btn_import_tot_excel"):
                    for sheet in xls.sheet_names:
                        for label, key in FORM_KEYS.items():
                            if label[:31].lower() in sheet.lower() or key.lower() in sheet.lower() or label.lower() in sheet.lower():
                                try:
                                    df_sheet = pd.read_excel(xls, sheet_name=sheet)
                                    df_sheet = df_sheet.dropna(how='all')
                                    imported_sheet = df_sheet.to_dict(orient="records")
                                    if imported_sheet:
                                        if mode_total.startswith("Sostituisci"):
                                            st.session_state[key] = imported_sheet
                                        else:
                                            st.session_state[key] = st.session_state.get(key, []) + imported_sheet
                                        st.success(f"Importato {len(imported_sheet)} in {label}")
                                except Exception as e:
                                    st.error(f"Errore foglio {sheet}: {e}")
                                break
                    st.success("Import totale Excel completato!")
                    st.balloons()
                    st.rerun()
            else:
                st.error(f"Errore apertura Excel: {last_err}")
        except Exception as e:
            st.error(f"Errore import totale Excel: {e}")

    st.divider()

    with st.expander("⚠️ Azzera Singolo Form"):
        sel_zero = st.selectbox("Form da azzerare", ["--"] + list(FORM_KEYS.keys()), key="zero_sel_unico")
        if sel_zero != "--":
            if st.button(f"🗑️ Azzera {sel_zero}", key=f"btn_zero_{sel_zero}"):
                st.session_state[FORM_KEYS[sel_zero]] = []
                st.success(f"{sel_zero} azzerato")
                st.rerun()




# Footer
st.divider()
st.markdown(
    """
    <div style="text-align:center;padding:8px;background:linear-gradient(135deg,#1A5D1A,#2e7d32);border-radius:8px;color:white;font-size:12px;">
    ANA Varese - Dashboard | Volontari linguette + ODV | Date gg/mm/aaaa | Backup Import/Export<br>
    Sviluppato per Ezio
    </div>
    """,
    unsafe_allow_html=True
)
