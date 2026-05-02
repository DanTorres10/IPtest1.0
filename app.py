import streamlit as st
import random
import sqlite3
import pandas as pd
from datetime import datetime
import os
import json

st.set_page_config(
    page_title="Validador IPWeb 1.3 · Interbank",
    page_icon="🔐",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ══════════════════════════════════════════════════════════════════
#  RUTAS
# ══════════════════════════════════════════════════════════════════
BASE_DIR            = os.path.dirname(os.path.abspath(__file__))
RUTA_FALLIDOS       = os.path.join(BASE_DIR, "Historial_Fallidos.xlsx")
RUTA_DB             = os.path.join(BASE_DIR, "clientes.db")
RUTA_PREGUNTAS_JSON = os.path.join(BASE_DIR, "preguntas.json")

PROCESOS = [
    "Desbloqueo Temporal UPF",
    "Actualización de Datos Personales",
    "Validación de VCAS",
]

BLOQUES_DNI = ["DNI_BLOQUE_1", "DNI_BLOQUE_2", "DNI_BLOQUE_3"]
BLOQUES_CE  = ["CE_BLOQUE_1",  "CE_BLOQUE_2",  "CE_BLOQUE_3"]

# ── CONTROL DE ACCESO AL PANEL DE ADMINISTRACIÓN ─────────────────
# Solo DanteTorres puede acceder a la gestión de usuarios y
# otorgar/revocar ese permiso a otros usuarios del sistema.
ADMIN_MAESTRO = "DanteTorres"   # único con acceso permanente garantizado

# ══════════════════════════════════════════════════════════════════
#  SESSION STATE
# ══════════════════════════════════════════════════════════════════
defaults = {
    "autenticado": False, "login_usuario": "", "login_error": "",
    "dark_mode": False,
    "vista": "validador",
    "preguntas": None, "finalizado": False, "resultado": None,
    "datos_bloqueo": None, "doc_sesion": "", "tipo_sesion": "DNI",
    "proceso_sesion": "", "user_sesion": "", "nombre_sesion": "",
    "admin_msg": "",
    "usuarios_extra": {}, "especialistas_extra": {},
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

dark = st.session_state.dark_mode

# ══════════════════════════════════════════════════════════════════
#  ESTILOS — LIGHT (verde Excel #217346) / DARK (verde neon)
# ══════════════════════════════════════════════════════════════════
if dark:
    PRIMARY     = "#4ade80"   # verde neon dark
    PRIMARY_DK  = "#22c55e"
    BG_APP      = "#0f172a"
    BG_CARD     = "#1e293b"
    BG_SIDEBAR  = "#0f172a"
    BORDER      = "#334155"
    TEXT_MAIN   = "#f1f5f9"
    TEXT_MUTED  = "#94a3b8"
    GOLD        = "#fbbf24"
    GREEN_OK    = "#4ade80"
    GREEN_BG    = "#052e16"
    RED_ERR     = "#f87171"
    RED_BG      = "#2d0a0a"
    AMBER       = "#fbbf24"
    AMBER_BG    = "#1c1400"
    SHADOW      = "0 2px 16px rgba(0,0,0,.5)"
else:
    PRIMARY     = "#217346"   # verde Excel
    PRIMARY_DK  = "#185c38"
    BG_APP      = "#f0f4f0"
    BG_CARD     = "#ffffff"
    BG_SIDEBAR  = "#1a3a2a"
    BORDER      = "#c8ddd0"
    TEXT_MAIN   = "#1a2e1f"
    TEXT_MUTED  = "#5a7a63"
    GOLD        = "#d4a017"
    GREEN_OK    = "#217346"
    GREEN_BG    = "#e6f2ec"
    RED_ERR     = "#b91c1c"
    RED_BG      = "#fef2f2"
    AMBER       = "#b45309"
    AMBER_BG    = "#fffbeb"
    SHADOW      = "0 2px 12px rgba(0,0,0,.08)"

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500&family=IBM+Plex+Sans:wght@300;400;500;600;700&display=swap');

:root {{
    --primary:    {PRIMARY};
    --primary-dk: {PRIMARY_DK};
    --gold:       {GOLD};
    --bg-app:     {BG_APP};
    --bg-card:    {BG_CARD};
    --bg-sidebar: {BG_SIDEBAR};
    --border:     {BORDER};
    --text:       {TEXT_MAIN};
    --muted:      {TEXT_MUTED};
    --green:      {GREEN_OK};
    --green-bg:   {GREEN_BG};
    --red:        {RED_ERR};
    --red-bg:     {RED_BG};
    --amber:      {AMBER};
    --amber-bg:   {AMBER_BG};
    --shadow:     {SHADOW};
    --radius:     8px;
}}

/* ─ BASE ─────────────────────────────────────────────────────── */
.stApp {{ background: var(--bg-app) !important; font-family:'IBM Plex Sans',sans-serif; color:var(--text); }}
.stApp * {{ color: var(--text); }}
.stTabs [data-baseweb="tab-list"] {{ background: var(--bg-card); border-radius:8px; padding:.25rem; }}
.stTabs [data-baseweb="tab"] {{ color: var(--muted) !important; font-weight:600; font-size:.82rem; }}
.stTabs [aria-selected="true"] {{ background: var(--primary) !important; color:white !important; border-radius:6px; }}
.stTextInput input, .stSelectbox select, .stTextArea textarea {{
    background: var(--bg-card) !important; border:1px solid var(--border) !important;
    color: var(--text) !important; border-radius:6px !important;
}}
.stRadio label {{ color: var(--text) !important; }}
hr {{ border-color: var(--border) !important; }}

/* ─ SIDEBAR ─────────────────────────────────────────────────── */
[data-testid="stSidebar"] {{ background: var(--bg-sidebar) !important; }}
[data-testid="stSidebar"] * {{ color: white !important; }}
[data-testid="stSidebar"] .stTextInput input,
[data-testid="stSidebar"] .stSelectbox select {{
    background: rgba(255,255,255,.1) !important;
    border: 1px solid rgba(255,255,255,.2) !important;
    color: white !important; border-radius:6px !important;
}}
[data-testid="stSidebar"] label {{
    color: rgba(255,255,255,.7) !important; font-size:.72rem !important;
    font-weight:600 !important; letter-spacing:.07em !important; text-transform:uppercase !important;
}}
[data-testid="stSidebar"] .stButton button {{
    background: var(--gold) !important; color: #1a1a1a !important;
    border:none !important; border-radius:6px !important; font-weight:700 !important;
    font-size:.83rem !important; padding:.6rem 1rem !important;
    width:100% !important; text-transform:uppercase !important; transition:all .2s !important;
}}
[data-testid="stSidebar"] .stButton button:hover:not(:disabled) {{
    opacity:.9 !important; transform:translateY(-1px) !important;
}}
[data-testid="stSidebar"] .stButton button:disabled {{
    background: rgba(255,255,255,.12) !important; color:rgba(255,255,255,.3) !important;
}}

/* ─ SIDEBAR ELEMENTS ─────────────────────────────────────────── */
.sb-logo {{ text-align:center; padding:1.3rem 1rem 1rem; border-bottom:1px solid rgba(255,255,255,.12); margin-bottom:.8rem; }}
.sb-logo .brand {{ font-size:1.15rem; font-weight:700; letter-spacing:.04em; }}
.sb-logo .sub   {{ font-size:.62rem; color:rgba(255,255,255,.5); text-transform:uppercase; letter-spacing:.15em; margin-top:.15rem; }}
.sb-logo .ver   {{ display:inline-block; background:var(--gold); color:#1a1a1a; font-size:.58rem; font-weight:700; padding:.1rem .4rem; border-radius:3px; margin-top:.4rem; }}
.chip {{ background:rgba(255,255,255,.08); border:1px solid rgba(255,255,255,.15); border-radius:6px; padding:.5rem .8rem; margin:.3rem 0; }}
.chip .lbl {{ font-size:.6rem; color:rgba(255,255,255,.5); text-transform:uppercase; letter-spacing:.1em; }}
.chip .val {{ font-size:.85rem; font-weight:600; margin-top:.08rem; }}
.bb {{ display:inline-block; padding:.15rem .5rem; border-radius:20px; font-size:.66rem; font-weight:700; letter-spacing:.05em; margin:.1rem .1rem .2rem 0; }}
.bb1 {{ background:#cfe2ff; color:#003F8A; }}
.bb2 {{ background:#d1f0e0; color:#1A6B3C; }}
.bb3 {{ background:#ffe5c0; color:#8B4513; }}
.logout-wrap .stButton button {{
    background:rgba(255,255,255,.08) !important; border:1px solid rgba(255,255,255,.2) !important;
    font-size:.76rem !important; text-transform:none !important; padding:.38rem .8rem !important;
}}

/* ─ LOGIN ─────────────────────────────────────────────────────── */
.login-card {{
    background:var(--bg-card); border-radius:16px; border-top:5px solid var(--primary);
    box-shadow:var(--shadow); padding:2.6rem 2.2rem 2.2rem; max-width:420px; margin:0 auto;
}}
.login-logo {{ text-align:center; margin-bottom:1.8rem; }}
.login-logo .ico   {{ font-size:2.4rem; }}
.login-logo .brand {{ font-size:1.25rem; font-weight:700; color:var(--primary); margin-top:.3rem; }}
.login-logo .sub   {{ font-size:.68rem; color:var(--muted); text-transform:uppercase; letter-spacing:.15em; margin-top:.18rem; }}
.login-logo .ver   {{ display:inline-block; background:var(--gold); color:#1a1a1a; font-size:.58rem; font-weight:700; padding:.1rem .4rem; border-radius:3px; margin-top:.4rem; }}
.login-err {{ background:var(--red-bg); border:1px solid var(--red); border-radius:6px; color:var(--red); font-size:.82rem; font-weight:600; padding:.6rem 1rem; margin-bottom:1rem; text-align:center; }}

/* ─ CARDS ─────────────────────────────────────────────────────── */
.vheader {{ background:var(--bg-card); border:1px solid var(--border); border-radius:var(--radius); padding:.85rem 1.1rem; box-shadow:var(--shadow); }}
.vheader .lbl {{ font-size:.66rem; color:var(--muted); font-weight:600; text-transform:uppercase; letter-spacing:.07em; }}
.vheader .val {{ font-family:'IBM Plex Mono',monospace; font-size:.88rem; color:var(--primary); font-weight:500; margin-top:.08rem; }}
.q-card {{ background:var(--bg-card); border:1px solid var(--border); border-left:4px solid var(--primary); border-radius:var(--radius); padding:1rem 1.2rem; margin-bottom:.8rem; box-shadow:var(--shadow); transition:all .2s; }}
.q-num  {{ font-family:'IBM Plex Mono',monospace; font-size:.65rem; font-weight:500; color:var(--primary); text-transform:uppercase; letter-spacing:.1em; margin-bottom:.28rem; }}
.q-text {{ font-size:.92rem; color:var(--text); line-height:1.55; }}
.admin-card {{ background:var(--bg-card); border:1px solid var(--border); border-radius:10px; padding:1.3rem 1.5rem; margin-bottom:.8rem; box-shadow:var(--shadow); }}
.admin-title {{ font-size:1rem; font-weight:700; color:var(--primary); margin-bottom:.8rem; border-bottom:2px solid var(--green-bg); padding-bottom:.45rem; }}
.user-row {{ display:flex; align-items:center; padding:.5rem .75rem; border-radius:6px; margin-bottom:.28rem; background:var(--bg-app); border:1px solid var(--border); }}
.user-name {{ font-weight:600; font-size:.86rem; color:var(--text); }}
.user-meta {{ font-size:.7rem; color:var(--muted); }}

/* ─ PROGRESS ─────────────────────────────────────────────────── */
.prog-bg   {{ background:var(--border); border-radius:4px; height:5px; margin:.4rem 0 1.2rem; }}
.prog-fill {{ background:var(--primary); border-radius:4px; height:5px; transition:width .4s ease; }}

/* ─ RESULT PANELS ────────────────────────────────────────────── */
.r-ok   {{ background:linear-gradient(135deg,{GREEN_OK} 0%,{PRIMARY_DK} 100%); color:white; border-radius:14px; padding:2.2rem 2rem; text-align:center; margin:1.5rem 0; box-shadow:0 8px 32px rgba(0,0,0,.18); }}
.r-fail {{ background:linear-gradient(135deg,#b91c1c 0%,#dc2626 100%); color:white; border-radius:14px; padding:2.2rem 2rem; text-align:center; margin:1.5rem 0; box-shadow:0 8px 32px rgba(0,0,0,.18); }}
.r-block{{ background:var(--bg-card); border:2px solid var(--red); border-radius:14px; padding:2.2rem 2rem; text-align:center; margin:1.5rem 0; }}
.r-mono  {{ font-family:'IBM Plex Mono',monospace; font-size:1.7rem; font-weight:500; letter-spacing:.06em; margin:.5rem 0; }}
.r-title {{ font-size:1.35rem; font-weight:700; margin:0 0 .35rem; }}
.r-sub   {{ font-size:.86rem; opacity:.88; line-height:1.5; }}
.r-table {{ width:100%; border-collapse:collapse; max-width:420px; margin:1.2rem auto 0; text-align:left; font-size:.82rem; }}
.r-table td {{ padding:.5rem .85rem; border-bottom:1px solid rgba(255,255,255,.15); }}
.r-table td:first-child {{ opacity:.65; font-size:.7rem; font-weight:600; text-transform:uppercase; letter-spacing:.06em; width:40%; }}
.block-alert {{ background:var(--amber-bg); border:1px solid var(--amber); border-radius:8px; padding:.85rem 1rem; margin-top:1rem; color:var(--amber); font-weight:600; font-size:.86rem; }}

/* ─ BADGES ───────────────────────────────────────────────────── */
.bdg {{ display:inline-flex; align-items:center; gap:.25rem; padding:.2rem .55rem; border-radius:20px; font-size:.68rem; font-weight:600; letter-spacing:.04em; }}
.bdg-ok   {{ background:var(--green-bg);  color:var(--green); }}
.bdg-fail {{ background:var(--red-bg);    color:var(--red); }}
.bdg-warn {{ background:var(--amber-bg);  color:var(--amber); }}
.bdg-info {{ background:var(--green-bg);  color:var(--primary); }}

/* ─ BUTTONS ──────────────────────────────────────────────────── */
.main-btn .stButton button {{
    background:var(--primary) !important; color:white !important;
    border:none !important; border-radius:6px !important; font-weight:700 !important;
    font-size:.86rem !important; padding:.6rem 1.6rem !important; transition:all .2s !important;
}}
.fin-btn .stButton button {{
    background:var(--primary) !important; color:white !important;
    font-weight:700 !important; font-size:.88rem !important; padding:.7rem 2rem !important;
    border-radius:6px !important; border:none !important; width:100% !important; margin-top:.4rem !important;
}}
.add-btn .stButton button {{
    background:var(--primary) !important; color:white !important;
    font-weight:700 !important; border:none !important; border-radius:6px !important;
    font-size:.82rem !important;
}}

/* ─ IDLE ─────────────────────────────────────────────────────── */
.idle {{ display:flex; flex-direction:column; align-items:center; justify-content:center; min-height:55vh; text-align:center; color:var(--muted); }}
.idle-icon {{ font-size:3rem; opacity:.3; margin-bottom:.8rem; }}

/* ─ CONFETTI CANVAS ──────────────────────────────────────────── */
#confetti-canvas {{
    position:fixed; top:0; left:0; width:100%; height:100%;
    pointer-events:none; z-index:9999;
}}

/* ─ DARK TOGGLE ──────────────────────────────────────────────── */
.dark-toggle .stButton button {{
    background:rgba(255,255,255,.1) !important; border:1px solid rgba(255,255,255,.2) !important;
    font-size:.75rem !important; padding:.35rem .7rem !important;
    text-transform:none !important; border-radius:20px !important; font-weight:600 !important;
}}
</style>

<!-- CONFETTI SCRIPT -->
<canvas id="confetti-canvas"></canvas>
<script>
(function(){{
  var canvas = document.getElementById('confetti-canvas');
  if(!canvas) return;
  var ctx = canvas.getContext('2d');
  var W = canvas.width  = window.innerWidth;
  var H = canvas.height = window.innerHeight;
  var pieces = [];
  var running = false;
  var colors = ['#217346','#4ade80','#fbbf24','#60a5fa','#f87171','#a78bfa','#34d399'];

  function Piece(){{
    this.x  = Math.random()*W;
    this.y  = Math.random()*H - H;
    this.w  = Math.random()*10+6;
    this.h  = Math.random()*5+4;
    this.color = colors[Math.floor(Math.random()*colors.length)];
    this.rot   = Math.random()*360;
    this.drot  = (Math.random()-0.5)*4;
    this.vy    = Math.random()*4+3;
    this.vx    = (Math.random()-0.5)*3;
    this.opacity = 1;
  }}
  Piece.prototype.draw = function(){{
    ctx.save();
    ctx.globalAlpha = this.opacity;
    ctx.translate(this.x,this.y);
    ctx.rotate(this.rot*Math.PI/180);
    ctx.fillStyle = this.color;
    ctx.fillRect(-this.w/2,-this.h/2,this.w,this.h);
    ctx.restore();
  }};
  Piece.prototype.update = function(){{
    this.x   += this.vx;
    this.y   += this.vy;
    this.rot += this.drot;
    if(this.y > H+20) this.opacity -= 0.02;
  }};

  function launch(){{
    if(running) return;
    running = true;
    pieces = [];
    for(var i=0;i<180;i++){{ pieces.push(new Piece()); }}
    loop();
  }}
  function loop(){{
    if(!running) return;
    ctx.clearRect(0,0,W,H);
    var alive = false;
    for(var i=0;i<pieces.length;i++){{
      pieces[i].update();
      if(pieces[i].opacity>0){{ pieces[i].draw(); alive=true; }}
    }}
    if(alive) requestAnimationFrame(loop);
    else {{ running=false; ctx.clearRect(0,0,W,H); }}
  }}

  window.launchConfetti = launch;
  window.addEventListener('resize',function(){{
    W = canvas.width  = window.innerWidth;
    H = canvas.height = window.innerHeight;
  }});
}})();
</script>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════
#  PREGUNTAS — archivo JSON local con fallback hardcodeado
# ══════════════════════════════════════════════════════════════════
PREGUNTAS_DEFAULT = {
    "DNI_BLOQUE_1": [
        "¿Cuál es su fecha de nacimiento?",
        "¿Cuál es su nombre completo tal como aparece en su DNI?",
        "¿En qué distrito figura su domicilio en el DNI?",
    ],
    "DNI_BLOQUE_2": [
        "¿Cuál es el número de su cuenta principal?",
        "¿Cuál fue el monto de su último depósito o retiro?",
        "¿Tiene alguna tarjeta de crédito o débito activa con nosotros?",
    ],
    "DNI_BLOQUE_3": [
        "¿Cuál es el correo electrónico registrado en su cuenta?",
        "¿Cuál es el número de celular registrado?",
        "¿Cuál es su dirección de correspondencia actual?",
    ],
    "CE_BLOQUE_1": [
        "¿Cuál es su nombre completo tal como figura en su carné de extranjería?",
        "¿Cuál es su nacionalidad?",
        "¿Cuál es su fecha de nacimiento?",
    ],
    "CE_BLOQUE_2": [
        "¿Cuál es el número de su cuenta principal en Interbank?",
        "¿Recuerda la fecha en que abrió su cuenta?",
        "¿Tiene algún producto adicional con nosotros (tarjeta, préstamo)?",
    ],
    "CE_BLOQUE_3": [
        "¿Cuál es el correo electrónico registrado en su cuenta?",
        "¿Cuál es el número de celular o teléfono registrado?",
        "¿Tiene habilitada la banca por internet?",
    ],
}

def cargar_preguntas() -> dict:
    try:
        if os.path.exists(RUTA_PREGUNTAS_JSON):
            with open(RUTA_PREGUNTAS_JSON, "r", encoding="utf-8") as f:
                data = json.load(f)
            # Asegurar que todos los bloques existan
            for k, v in PREGUNTAS_DEFAULT.items():
                if k not in data or not data[k]:
                    data[k] = v
            return data
    except Exception:
        pass
    return dict(PREGUNTAS_DEFAULT)

def guardar_preguntas(data: dict) -> bool:
    try:
        with open(RUTA_PREGUNTAS_JSON, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        st.error(f"Error al guardar preguntas: {e}")
        return False

def sortear_preguntas(tipo_doc: str) -> list:
    bd = cargar_preguntas()
    bloques = BLOQUES_DNI if tipo_doc == "DNI" else BLOQUES_CE
    seleccion = []
    for bloque in bloques:
        pool = bd.get(bloque, [])
        if pool:
            seleccion.append({
                "texto":  random.choice(pool),
                "bloque": bloque,
            })
    return seleccion   # 3 preguntas: 1 por bloque


# ══════════════════════════════════════════════════════════════════
#  USUARIOS DEL SISTEMA
# ══════════════════════════════════════════════════════════════════
def cargar_usuarios_sistema() -> dict:
    try:
        esp = st.secrets.get("usuarios", {})
        if esp:
            u = dict(esp)
            u.update(st.session_state.get("usuarios_extra", {}))
            return u
    except Exception:
        pass
    u = {"RaizaGrados": "Ibk2026", "DanteTorres": "Ibk2026",
         "LuceroCori": "Ibk2026",  "BryanRamirez": "Ibk2026"}
    u.update(st.session_state.get("usuarios_extra", {}))
    return u

def agregar_usuario_sistema(usr, pwd):
    st.session_state.usuarios_extra[usr] = pwd; return True

def puede_acceder_admin() -> bool:
    """Devuelve True si el usuario activo tiene permiso para el panel de admin."""
    usuario = st.session_state.get("login_usuario", "")
    # El maestro siempre puede
    if usuario == ADMIN_MAESTRO:
        return True
    # Otros usuarios solo si DanteTorres les otorgó permiso
    try:
        admins_extra = list(st.secrets.get("admins_autorizados", []))
        if usuario in admins_extra:
            return True
    except Exception:
        pass
    return usuario in st.session_state.get("admins_sesion", [])

def eliminar_usuario_sistema(usr):
    try:
        if usr in dict(st.secrets.get("usuarios", {})):
            st.session_state.admin_msg = f"⚠️ '{usr}' está en Secrets. Edítalo en Streamlit → Settings → Secrets."
            return False
    except Exception:
        pass
    st.session_state.usuarios_extra.pop(usr, None); return True


# ══════════════════════════════════════════════════════════════════
#  ESPECIALISTAS
# ══════════════════════════════════════════════════════════════════
def cargar_especialistas() -> dict:
    try:
        esp = st.secrets.get("especialistas", {})
        if esp:
            d = {k.upper(): v for k, v in dict(esp).items()}
            d.update({k.upper(): v for k, v in st.session_state.get("especialistas_extra", {}).items()})
            return d
    except Exception:
        pass
    d = {
        "B46453":"Especialista 1","B46450":"Especialista 2","B46449":"Especialista 3",
        "B46419":"Especialista 4","B46447":"Especialista 5","B46446":"Especialista 6",
        "B46444":"Especialista 7","B46287":"Especialista 8","B45325":"Especialista 9",
        "B45320":"Especialista 10",
    }
    d.update({k.upper(): v for k, v in st.session_state.get("especialistas_extra", {}).items()})
    return d


# ══════════════════════════════════════════════════════════════════
#  BASE DE DATOS
# ══════════════════════════════════════════════════════════════════
def iniciar_db():
    conn = sqlite3.connect(RUTA_DB)
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS fallidos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            dni TEXT NOT NULL, proceso TEXT NOT NULL,
            trabajador TEXT NOT NULL, fecha TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS idx_dni ON fallidos(dni);
    """)
    conn.commit(); conn.close()

def guardar_fallo(dni, proceso, trabajador):
    fecha = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    conn  = sqlite3.connect(RUTA_DB)
    conn.execute("INSERT INTO fallidos(dni,proceso,trabajador,fecha) VALUES(?,?,?,?)",
                 (dni, proceso, trabajador, fecha))
    conn.commit(); conn.close()
    try:
        nuevo = pd.DataFrame([[dni,proceso,trabajador,fecha]],
                             columns=["DNI/CE","PROCESO","ASESOR","FECHA_HORA"])
        if os.path.exists(RUTA_FALLIDOS):
            df = pd.concat([pd.read_excel(RUTA_FALLIDOS, engine="openpyxl"), nuevo], ignore_index=True)
        else:
            df = nuevo
        df.to_excel(RUTA_FALLIDOS, index=False, engine="openpyxl")
    except Exception:
        pass

def buscar_historial_fallido(dni):
    conn = sqlite3.connect(RUTA_DB)
    row  = conn.execute(
        "SELECT trabajador,fecha FROM fallidos WHERE dni=? ORDER BY id DESC LIMIT 1",(dni,)
    ).fetchone()
    conn.close()
    return {"trabajador":row[0],"fecha":row[1]} if row else None


# ══════════════════════════════════════════════════════════════════
#  INIT
# ══════════════════════════════════════════════════════════════════
iniciar_db()


# ══════════════════════════════════════════════════════════════════
#  LOGIN
# ══════════════════════════════════════════════════════════════════
if not st.session_state.autenticado:
    _, col_c, _ = st.columns([1, 1.1, 1])
    with col_c:
        st.markdown("""
        <div class="login-card">
          <div class="login-logo">
            <div class="ico">🔐</div>
            <div class="brand">Validador IPWeb</div>
            <div class="sub">Sistema de Autenticación · Interbank</div>
            <span class="ver">v1.3</span>
          </div>
        </div>""", unsafe_allow_html=True)

        if st.session_state.login_error:
            st.markdown(f'<div class="login-err">⚠️ {st.session_state.login_error}</div>',
                        unsafe_allow_html=True)

        u_in = st.text_input("Usuario",    placeholder="Ingrese su usuario")
        p_in = st.text_input("Contraseña", placeholder="Ingrese su contraseña", type="password")
        st.markdown("<br>", unsafe_allow_html=True)

        if st.button("INGRESAR AL SISTEMA", use_container_width=True):
            u, p = u_in.strip(), p_in.strip()
            if not u and not p:
                st.session_state.login_error = "Debe completar el usuario y la contraseña."
            elif not u:
                st.session_state.login_error = "El campo Usuario es obligatorio."
            elif not p:
                st.session_state.login_error = "El campo Contraseña es obligatorio."
            elif cargar_usuarios_sistema().get(u) == p:
                st.session_state.autenticado   = True
                st.session_state.login_usuario = u
                st.session_state.login_error   = ""
            else:
                st.session_state.login_error = "Usuario o contraseña incorrectos."
            st.rerun()

        st.markdown('<div style="text-align:center;margin-top:1.2rem;font-size:.7rem;color:var(--muted);">Acceso restringido · Solo personal autorizado</div>',
                    unsafe_allow_html=True)
    st.stop()


# ══════════════════════════════════════════════════════════════════
#  SIDEBAR (autenticado)
# ══════════════════════════════════════════════════════════════════
dict_esp = cargar_especialistas()

with st.sidebar:
    st.markdown("""
    <div class="sb-logo">
      <div class="brand">🔐 Validador IPWeb</div>
      <div class="sub">Sistema de Autenticación</div>
      <span class="ver">v1.3</span>
    </div>""", unsafe_allow_html=True)

    # Sesión + dark mode
    col_s, col_d = st.columns([2, 1])
    with col_s:
        st.markdown(f'<div class="chip"><div class="lbl">🟢 Sesión activa</div><div class="val">{st.session_state.login_usuario}</div></div>',
                    unsafe_allow_html=True)
    with col_d:
        st.markdown('<div class="dark-toggle">', unsafe_allow_html=True)
        icono = "☀️" if dark else "🌙"
        if st.button(icono, use_container_width=True, key="toggle_dark"):
            st.session_state.dark_mode = not dark; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    # Navegación — ⚙️ solo visible para usuarios con permiso admin
    _puede_admin = puede_acceder_admin()
    if _puede_admin:
        c1, c2, c3 = st.columns(3)
        with c1:
            if st.button("🛡️", use_container_width=True, help="Validador"):
                st.session_state.vista = "validador"; st.rerun()
        with c2:
            if st.button("⚙️", use_container_width=True, help="Gestión de usuarios"):
                st.session_state.vista = "admin"; st.rerun()
        with c3:
            if st.button("❓", use_container_width=True, help="Preguntas"):
                st.session_state.vista = "preguntas"; st.rerun()
    else:
        c1, c2 = st.columns(2)
        with c1:
            if st.button("🛡️", use_container_width=True, help="Validador"):
                st.session_state.vista = "validador"; st.rerun()
        with c2:
            if st.button("❓", use_container_width=True, help="Preguntas"):
                st.session_state.vista = "preguntas"; st.rerun()
        # Si estaba en admin sin permiso, redirigir
        if st.session_state.vista == "admin":
            st.session_state.vista = "validador"; st.rerun()

    st.markdown('<div class="logout-wrap">', unsafe_allow_html=True)
    if st.button("⏻ Cerrar sesión", use_container_width=True):
        for k in list(st.session_state.keys()): del st.session_state[k]
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("---")

    # ── Controles del validador ───────────────────────────────────
    if st.session_state.vista == "validador":
        st.markdown("**IDENTIFICACIÓN DEL ASESOR**")
        user_b = st.text_input("Código Usuario B", placeholder="B46453",
                               label_visibility="collapsed").strip().upper()
        es_valido   = user_b in dict_esp
        nombre_ases = dict_esp.get(user_b, "")

        if user_b:
            if es_valido:
                st.markdown(f'<div class="chip"><div class="lbl">✓ Asesor autorizado</div><div class="val">{nombre_ases}</div></div>',
                            unsafe_allow_html=True)
            else:
                st.markdown('<div style="color:#f87171;font-size:.78rem;margin-top:.3rem;">✗ Código no autorizado</div>',
                            unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("**DATOS DEL CLIENTE**")

        tipo_doc = st.selectbox("Tipo de documento", ["DNI","CE"],
                                disabled=not es_valido, label_visibility="collapsed")
        max_d    = 8 if tipo_doc == "DNI" else 9
        doc_num  = st.text_input(f"N° de {tipo_doc}", max_chars=max_d,
                                 disabled=not es_valido,
                                 placeholder="12345678" if tipo_doc=="DNI" else "123456789",
                                 label_visibility="collapsed")

        # Proceso solo habilitado si tipo+doc están completos
        doc_completo = es_valido and len(doc_num) == max_d
        proceso_sel  = st.selectbox("Proceso a realizar", PROCESOS,
                                    disabled=not doc_completo,
                                    label_visibility="collapsed")
        if not doc_completo and es_valido:
            st.markdown('<div style="color:#fbbf24;font-size:.72rem;margin-top:.2rem;">⚠ Ingrese el número de documento para seleccionar el proceso.</div>',
                        unsafe_allow_html=True)

        # Bloques info
        st.markdown("---")
        bd = cargar_preguntas()
        bloques_activos = BLOQUES_DNI if tipo_doc == "DNI" else BLOQUES_CE
        etiquetas = {
            "DNI_BLOQUE_1":"Bloque 1 (DNI)","DNI_BLOQUE_2":"Bloque 2 (DNI)","DNI_BLOQUE_3":"Bloque 3 (DNI)",
            "CE_BLOQUE_1":"Bloque 1 (CE)", "CE_BLOQUE_2":"Bloque 2 (CE)", "CE_BLOQUE_3":"Bloque 3 (CE)",
        }
        cls_map = {"DNI_BLOQUE_1":"bb1","DNI_BLOQUE_2":"bb2","DNI_BLOQUE_3":"bb3",
                   "CE_BLOQUE_1":"bb1","CE_BLOQUE_2":"bb2","CE_BLOQUE_3":"bb3"}
        st.markdown('<div style="font-size:.62rem;color:rgba(255,255,255,.45);text-transform:uppercase;letter-spacing:.08em;margin-bottom:.35rem;">Preguntas disponibles</div>',
                    unsafe_allow_html=True)
        for b in bloques_activos:
            n   = len(bd.get(b, []))
            cls = cls_map.get(b, "bb1")
            st.markdown(f'<span class="bb {cls}">{etiquetas[b]}</span>'
                        f'<span style="font-size:.66rem;color:rgba(255,255,255,.4);margin-left:.35rem;">{n} preguntas</span><br>',
                        unsafe_allow_html=True)

        st.markdown("---")
        if st.button("▶ INICIAR VALIDACIÓN", disabled=not doc_completo, use_container_width=True):
            historial = buscar_historial_fallido(doc_num)
            if historial:
                st.session_state.resultado     = "bloqueado"
                st.session_state.datos_bloqueo = historial
                st.session_state.finalizado    = True
                st.session_state.preguntas     = None
            else:
                st.session_state.preguntas     = sortear_preguntas(tipo_doc)
                st.session_state.finalizado    = False
                st.session_state.resultado     = None
                st.session_state.datos_bloqueo = None
            st.session_state.doc_sesion     = doc_num
            st.session_state.tipo_sesion    = tipo_doc
            st.session_state.proceso_sesion = proceso_sel
            st.session_state.user_sesion    = user_b
            st.session_state.nombre_sesion  = nombre_ases
            st.rerun()

    else:
        lbl = {"admin":"⚙️ Gestión de usuarios y especialistas",
               "preguntas":"❓ Editor de preguntas por bloque"}.get(st.session_state.vista,"")
        st.markdown(f'<div style="font-size:.75rem;color:rgba(255,255,255,.55);margin-top:.4rem;">{lbl}</div>',
                    unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════
#  VISTA: EDITOR DE PREGUNTAS
# ══════════════════════════════════════════════════════════════════
if st.session_state.vista == "preguntas":
    st.markdown("## ❓ Editor de Preguntas")
    st.markdown("Administra las preguntas de validación por tipo de documento y bloque.")

    if st.session_state.admin_msg:
        es_ok = "✅" in st.session_state.admin_msg
        bg = f"{'var(--green-bg)' if es_ok else 'var(--red-bg)'}"
        cl = f"{'var(--green)' if es_ok else 'var(--red)'}"
        st.markdown(f'<div style="background:{bg};border:1px solid {cl};border-radius:6px;padding:.75rem 1rem;margin-bottom:1rem;color:{cl};font-weight:600;">{st.session_state.admin_msg}</div>',
                    unsafe_allow_html=True)
        st.session_state.admin_msg = ""

    tab_dni, tab_ce = st.tabs(["🪪  Preguntas DNI", "🌍  Preguntas CE (Extranjeros)"])
    bd = cargar_preguntas()

    def render_bloque_editor(bloque_key: str, titulo: str):
        st.markdown(f'<div class="admin-card"><div class="admin-title">{titulo}</div>',
                    unsafe_allow_html=True)
        preguntas_bloque = bd.get(bloque_key, [])

        # Mostrar preguntas existentes
        to_delete = None
        for idx, preg in enumerate(preguntas_bloque):
            col_p, col_e, col_d = st.columns([6, 1, 1])
            with col_p:
                st.markdown(f'<div style="padding:.4rem .6rem;background:var(--bg-app);border:1px solid var(--border);border-radius:5px;font-size:.85rem;color:var(--text);">{idx+1}. {preg}</div>',
                            unsafe_allow_html=True)
            with col_e:
                if st.button("✏️", key=f"edit_{bloque_key}_{idx}", help="Editar"):
                    st.session_state[f"editing_{bloque_key}_{idx}"] = True
            with col_d:
                if st.button("🗑️", key=f"del_{bloque_key}_{idx}", help="Eliminar"):
                    to_delete = idx

            # Modo edición inline
            if st.session_state.get(f"editing_{bloque_key}_{idx}"):
                nueva = st.text_area("Editar pregunta:", value=preg,
                                     key=f"ea_{bloque_key}_{idx}", height=80)
                col_s, col_c2 = st.columns(2)
                with col_s:
                    if st.button("💾 Guardar", key=f"save_{bloque_key}_{idx}"):
                        bd[bloque_key][idx] = nueva.strip()
                        guardar_preguntas(bd)
                        del st.session_state[f"editing_{bloque_key}_{idx}"]
                        st.session_state.admin_msg = "✅ Pregunta actualizada."
                        st.rerun()
                with col_c2:
                    if st.button("✖ Cancelar", key=f"can_{bloque_key}_{idx}"):
                        del st.session_state[f"editing_{bloque_key}_{idx}"]
                        st.rerun()

        if to_delete is not None:
            bd[bloque_key].pop(to_delete)
            guardar_preguntas(bd)
            st.session_state.admin_msg = "✅ Pregunta eliminada."
            st.rerun()

        # Agregar nueva pregunta
        st.markdown("<br>", unsafe_allow_html=True)
        nueva_p = st.text_area(f"Nueva pregunta para {titulo}:",
                               placeholder="Escribe la nueva pregunta aquí...",
                               key=f"new_{bloque_key}", height=70)
        st.markdown('<div class="add-btn">', unsafe_allow_html=True)
        if st.button(f"➕ Agregar", key=f"add_{bloque_key}", use_container_width=True):
            np = nueva_p.strip()
            if not np:
                st.session_state.admin_msg = "⚠️ Escribe la pregunta antes de agregar."
            else:
                if bloque_key not in bd:
                    bd[bloque_key] = []
                bd[bloque_key].append(np)
                guardar_preguntas(bd)
                st.session_state.admin_msg = f"✅ Pregunta agregada a {titulo}."
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with tab_dni:
        st.markdown("Preguntas para clientes con **Documento Nacional de Identidad (DNI)**.")
        col1, col2, col3 = st.columns(3)
        with col1: render_bloque_editor("DNI_BLOQUE_1", "🔵 Bloque 1 — Datos Personales")
        with col2: render_bloque_editor("DNI_BLOQUE_2", "🟢 Bloque 2 — Datos Bancarios")
        with col3: render_bloque_editor("DNI_BLOQUE_3", "🟠 Bloque 3 — Datos de Contacto")

    with tab_ce:
        st.markdown("Preguntas para clientes **Extranjeros** con Carné de Extranjería (CE).")
        col4, col5, col6 = st.columns(3)
        with col4: render_bloque_editor("CE_BLOQUE_1", "🔵 Bloque 1 — Datos Personales")
        with col5: render_bloque_editor("CE_BLOQUE_2", "🟢 Bloque 2 — Datos Bancarios")
        with col6: render_bloque_editor("CE_BLOQUE_3", "🟠 Bloque 3 — Datos de Contacto")

    st.stop()


# ══════════════════════════════════════════════════════════════════
#  VISTA: ADMINISTRACIÓN DE USUARIOS Y ESPECIALISTAS
# ══════════════════════════════════════════════════════════════════
if st.session_state.vista == "admin":
    # Doble verificación de seguridad — bloquear acceso directo
    if not puede_acceder_admin():
        st.error("⛔ No tienes permiso para acceder a esta sección.")
        st.stop()
    st.markdown("## ⚙️ Panel de Administración")

    if st.session_state.admin_msg:
        es_ok = "✅" in st.session_state.admin_msg
        bg = "var(--green-bg)" if es_ok else "var(--red-bg)"
        cl = "var(--green)"    if es_ok else "var(--red)"
        st.markdown(f'<div style="background:{bg};border:1px solid {cl};border-radius:6px;padding:.75rem 1rem;margin-bottom:1rem;color:{cl};font-weight:600;">{st.session_state.admin_msg}</div>',
                    unsafe_allow_html=True)
        st.session_state.admin_msg = ""

    tab1, tab2, tab3 = st.tabs(["👥  Usuarios del sistema", "🪪  Códigos de especialistas", "🔑  Permisos de administrador"])

    # ── TAB 1: USUARIOS ──────────────────────────────────────────
    with tab1:
        col1, col2 = st.columns([1.4, 1])
        with col1:
            st.markdown('<div class="admin-card"><div class="admin-title">👥 Usuarios activos</div>', unsafe_allow_html=True)
            for i, (usr, pwd) in enumerate(cargar_usuarios_sistema().items()):
                try:    en_sec = usr in dict(st.secrets.get("usuarios", {}))
                except: en_sec = False
                cn, cd = st.columns([3, 1])
                with cn:
                    etq = (' <span style="font-size:.58rem;background:var(--green-bg);color:var(--green);padding:.08rem .35rem;border-radius:3px;font-weight:700;">SECRETS</span>'
                           if en_sec else
                           ' <span style="font-size:.58rem;background:var(--amber-bg);color:var(--amber);padding:.08rem .35rem;border-radius:3px;font-weight:700;">SESIÓN</span>')
                    st.markdown(f'<div class="user-row"><div class="user-name">👤 {usr}{etq}</div><div class="user-meta">{"•"*min(len(pwd),8)}</div></div>',
                                unsafe_allow_html=True)
                with cd:
                    if not en_sec and st.button("🗑️", key=f"delu_{i}_{usr}"):
                        if usr == st.session_state.login_usuario:
                            st.session_state.admin_msg = "⚠️ No puedes eliminar tu usuario activo."
                        elif eliminar_usuario_sistema(usr):
                            st.session_state.admin_msg = f"✅ Usuario '{usr}' eliminado."
                        st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

        with col2:
            st.markdown('<div class="admin-card"><div class="admin-title">➕ Nuevo usuario</div>', unsafe_allow_html=True)
            nu = st.text_input("Nombre de usuario", placeholder="Ej: CarlosMendez", key="nu_usr")
            np = st.text_input("Contraseña", placeholder="Ej: Ibk2026", type="password", key="nu_pwd")
            st.markdown('<div class="add-btn">', unsafe_allow_html=True)
            if st.button("➕ AGREGAR USUARIO", use_container_width=True, key="btn_add_usr"):
                nu_, np_ = nu.strip(), np.strip()
                if not nu_ or not np_: st.session_state.admin_msg = "⚠️ Completa usuario y contraseña."
                elif nu_ in cargar_usuarios_sistema(): st.session_state.admin_msg = f"⚠️ '{nu_}' ya existe."
                elif agregar_usuario_sistema(nu_, np_): st.session_state.admin_msg = f"✅ Usuario '{nu_}' agregado (sesión)."
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)
            st.markdown('<div style="margin-top:.9rem;background:var(--green-bg);border-radius:6px;padding:.7rem;font-size:.73rem;color:var(--primary);"><b>🔒 Para permanentes:</b> Streamlit Cloud → Settings → Secrets → sección <code>[usuarios]</code></div>',
                        unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

    # ── TAB 2: ESPECIALISTAS ─────────────────────────────────────
    with tab2:
        col3, col4 = st.columns([1.4, 1])
        with col3:
            st.markdown('<div class="admin-card"><div class="admin-title">🪪 Especialistas autorizados</div>', unsafe_allow_html=True)
            for i, (cod, nom) in enumerate(cargar_especialistas().items()):
                try:    en_sec2 = cod in {k.upper() for k in dict(st.secrets.get("especialistas", {}))}
                except: en_sec2 = False
                cc, cd2 = st.columns([3, 1])
                with cc:
                    etq2 = (' <span style="font-size:.58rem;background:var(--green-bg);color:var(--green);padding:.08rem .35rem;border-radius:3px;font-weight:700;">SECRETS</span>'
                            if en_sec2 else
                            ' <span style="font-size:.58rem;background:var(--amber-bg);color:var(--amber);padding:.08rem .35rem;border-radius:3px;font-weight:700;">SESIÓN</span>')
                    st.markdown(f'<div class="user-row"><div class="user-name">🪪 {cod}{etq2}</div><div class="user-meta">{nom}</div></div>',
                                unsafe_allow_html=True)
                with cd2:
                    if not en_sec2 and st.button("🗑️", key=f"dele_{i}_{cod}"):
                        st.session_state.especialistas_extra.pop(cod, None)
                        st.session_state.admin_msg = f"✅ '{cod}' eliminado de sesión."
                        st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

        with col4:
            st.markdown('<div class="admin-card"><div class="admin-title">➕ Nuevo especialista</div>', unsafe_allow_html=True)
            nc = st.text_input("Código (Usuario B)", placeholder="Ej: B46999", key="ne_cod")
            nn = st.text_input("Nombre completo",   placeholder="Ej: Carlos Mendez", key="ne_nom")
            st.markdown('<div class="add-btn">', unsafe_allow_html=True)
            if st.button("➕ AGREGAR ESPECIALISTA", use_container_width=True, key="btn_add_esp"):
                nc_, nn_ = nc.strip().upper(), nn.strip()
                if not nc_ or not nn_: st.session_state.admin_msg = "⚠️ Completa código y nombre."
                elif nc_ in cargar_especialistas(): st.session_state.admin_msg = f"⚠️ '{nc_}' ya existe."
                else:
                    st.session_state.especialistas_extra[nc_] = nn_
                    st.session_state.admin_msg = f"✅ '{nc_}' agregado (sesión)."
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)
            st.markdown('<div style="margin-top:.9rem;background:var(--green-bg);border-radius:6px;padding:.7rem;font-size:.73rem;color:var(--primary);"><b>🔒 Para permanentes:</b> Streamlit Cloud → Settings → Secrets → sección <code>[especialistas]</code></div>',
                        unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

    # ── TAB 3: PERMISOS DE ADMIN ─────────────────────────────────
    with tab3:
        if st.session_state.login_usuario != ADMIN_MAESTRO:
            st.warning("⚠️ Solo **DanteTorres** puede gestionar los permisos de administrador.")
        else:
            st.markdown("Controla qué usuarios pueden acceder al **Panel de Administración** (⚙️).")
            st.markdown(f'<div style="background:var(--green-bg);border:1px solid var(--green);border-radius:6px;padding:.75rem 1rem;font-size:.82rem;color:var(--primary);margin-bottom:1rem;"><b>👑 Administrador maestro:</b> DanteTorres — acceso permanente garantizado.</div>', unsafe_allow_html=True)

            # Admins actuales con permiso
            todos_usuarios = list(cargar_usuarios_sistema().keys())
            admins_sesion  = st.session_state.get("admins_sesion", [])
            try:
                admins_secrets = list(st.secrets.get("admins_autorizados", []))
            except Exception:
                admins_secrets = []
            admins_actuales = list(set(admins_secrets + admins_sesion))
            admins_actuales = [a for a in admins_actuales if a != ADMIN_MAESTRO]

            col_left, col_right = st.columns([1.4, 1])

            with col_left:
                st.markdown('<div class="admin-card"><div class="admin-title">🔑 Usuarios con permiso de admin</div>', unsafe_allow_html=True)
                if not admins_actuales:
                    st.info("Ningún usuario adicional tiene permiso de administrador actualmente.")
                for adm in admins_actuales:
                    en_sec = adm in admins_secrets
                    ca, cb = st.columns([3, 1])
                    with ca:
                        etq = (' <span style="font-size:.58rem;background:var(--green-bg);color:var(--green);padding:.08rem .35rem;border-radius:3px;font-weight:700;">SECRETS</span>'
                               if en_sec else
                               ' <span style="font-size:.58rem;background:var(--amber-bg);color:var(--amber);padding:.08rem .35rem;border-radius:3px;font-weight:700;">SESIÓN</span>')
                        st.markdown(f'<div class="user-row"><div class="user-name">🔑 {adm}{etq}</div><div class="user-meta">Acceso al panel de administración</div></div>',
                                    unsafe_allow_html=True)
                    with cb:
                        if not en_sec and st.button("🗑️", key=f"rev_{adm}", help=f"Revocar permiso a {adm}"):
                            lista = st.session_state.get("admins_sesion", [])
                            if adm in lista: lista.remove(adm)
                            st.session_state.admins_sesion = lista
                            st.session_state.admin_msg = f"✅ Permiso de '{adm}' revocado."
                            st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)

            with col_right:
                st.markdown('<div class="admin-card"><div class="admin-title">➕ Otorgar permiso</div>', unsafe_allow_html=True)
                usuarios_sin_permiso = [u for u in todos_usuarios if u != ADMIN_MAESTRO and u not in admins_actuales]
                if not usuarios_sin_permiso:
                    st.info("Todos los usuarios ya tienen permiso.")
                else:
                    nuevo_admin = st.selectbox("Seleccionar usuario", usuarios_sin_permiso, key="sel_new_admin")
                    st.markdown('<div class="add-btn">', unsafe_allow_html=True)
                    if st.button("🔑 OTORGAR PERMISO", use_container_width=True, key="btn_grant"):
                        if "admins_sesion" not in st.session_state:
                            st.session_state.admins_sesion = []
                        if nuevo_admin not in st.session_state.admins_sesion:
                            st.session_state.admins_sesion.append(nuevo_admin)
                        st.session_state.admin_msg = f"✅ Permiso otorgado a '{nuevo_admin}' (activo en esta sesión)."
                        st.rerun()
                    st.markdown('</div>', unsafe_allow_html=True)
                st.markdown('<div style="margin-top:.9rem;background:var(--green-bg);border-radius:6px;padding:.7rem;font-size:.73rem;color:var(--primary);"><b>🔒 Para permisos permanentes:</b> Streamlit Cloud → Settings → Secrets → agrega sección:<br><code>[admins_autorizados]<br>usuario1 = true</code></div>',
                            unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)

    st.stop()


# ══════════════════════════════════════════════════════════════════
#  VISTA: VALIDADOR
# ══════════════════════════════════════════════════════════════════

# ── IDLE ──────────────────────────────────────────────────────────
if not st.session_state.preguntas and not st.session_state.finalizado:
    st.markdown("""
    <div class="idle">
      <div class="idle-icon">🛡️</div>
      <div style="font-size:.95rem;">Ingrese los datos en el panel lateral<br>para iniciar la validación de identidad.</div>
    </div>""", unsafe_allow_html=True)

# ── CUESTIONARIO ──────────────────────────────────────────────────
elif st.session_state.preguntas and not st.session_state.finalizado:
    doc     = st.session_state.doc_sesion
    tipo    = st.session_state.tipo_sesion
    proceso = st.session_state.proceso_sesion
    preg    = st.session_state.preguntas

    c1, c2, c3 = st.columns(3)
    now = datetime.now().strftime("%d/%m/%Y  %H:%M")
    tipo_label = f"{'🪪' if tipo=='DNI' else '🌍'} {tipo}"
    for col, lbl, val in [
        (c1, "Documento",  f"{tipo_label} · {doc}"),
        (c2, "Proceso",    proceso[:38] + ("…" if len(proceso)>38 else "")),
        (c3, "Fecha/Hora", now),
    ]:
        with col:
            st.markdown(f'<div class="vheader"><div class="lbl">{lbl}</div><div class="val">{val}</div></div>',
                        unsafe_allow_html=True)

    st.markdown(f'<p style="font-size:.82rem;color:var(--muted);margin:1rem 0 .4rem;">Formule cada pregunta al cliente y registre si la respuesta es <b>Correcta</b> o <b>Incorrecta</b>. · Tipo de cliente: <b>{tipo}</b></p>',
                unsafe_allow_html=True)

    respuestas = {i: st.session_state.get(f"p_{i}", "Pendiente") for i in range(len(preg))}
    respondidas = sum(1 for v in respuestas.values() if v != "Pendiente")
    pct = int(respondidas / len(preg) * 100)
    st.markdown(f'<div class="prog-bg"><div class="prog-fill" style="width:{pct}%;"></div></div>',
                unsafe_allow_html=True)

    etq_bloque = {
        "DNI_BLOQUE_1":"Bloque 1","DNI_BLOQUE_2":"Bloque 2","DNI_BLOQUE_3":"Bloque 3",
        "CE_BLOQUE_1":"Bloque 1", "CE_BLOQUE_2":"Bloque 2", "CE_BLOQUE_3":"Bloque 3",
    }
    cls_bloque = {
        "DNI_BLOQUE_1":"bb1","DNI_BLOQUE_2":"bb2","DNI_BLOQUE_3":"bb3",
        "CE_BLOQUE_1":"bb1","CE_BLOQUE_2":"bb2","CE_BLOQUE_3":"bb3",
    }

    for i, item in enumerate(preg):
        estado = st.session_state.get(f"p_{i}", "Pendiente")
        bcolor = {"Correcto":PRIMARY,"Incorrecto":"#dc2626","Pendiente":PRIMARY}.get(estado, PRIMARY)
        if estado == "Incorrecto": bcolor = "#dc2626"
        nb  = etq_bloque.get(item["bloque"], item["bloque"])
        cls = cls_bloque.get(item["bloque"], "bb1")
        st.markdown(
            f'<div class="q-card" style="border-left-color:{bcolor};">'
            f'<div class="q-num">Pregunta {i+1} de {len(preg)}'
            f' &nbsp;<span class="bb {cls}">{nb}</span></div>'
            f'<div class="q-text">{item["texto"]}</div></div>',
            unsafe_allow_html=True)
        cr, _ = st.columns([2, 3])
        with cr:
            respuestas[i] = st.radio(f"_r{i}", ["Pendiente","Correcto","Incorrecto"],
                                     key=f"p_{i}", horizontal=True, label_visibility="collapsed")

    st.markdown("---")
    ok_c  = sum(1 for v in respuestas.values() if v == "Correcto")
    err_c = sum(1 for v in respuestas.values() if v == "Incorrecto")
    pend  = sum(1 for v in respuestas.values() if v == "Pendiente")
    cols  = st.columns(4)
    cols[0].markdown(f'<span class="bdg bdg-info">📋 Total: {len(preg)}</span>',  unsafe_allow_html=True)
    cols[1].markdown(f'<span class="bdg bdg-ok">✓ Correctas: {ok_c}</span>',     unsafe_allow_html=True)
    cols[2].markdown(f'<span class="bdg bdg-fail">✗ Incorrectas: {err_c}</span>',unsafe_allow_html=True)
    cols[3].markdown(f'<span class="bdg bdg-warn">⏳ Pendientes: {pend}</span>',  unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="fin-btn">', unsafe_allow_html=True)
    if st.button("REGISTRAR RESULTADO", use_container_width=True):
        if any(v == "Pendiente" for v in respuestas.values()):
            st.warning("⚠️ Debe calificar todas las preguntas antes de finalizar.")
        elif any(v == "Incorrecto" for v in respuestas.values()):
            guardar_fallo(doc, proceso, f"{st.session_state.user_sesion} - {st.session_state.nombre_sesion}")
            st.session_state.resultado  = "fail"
            st.session_state.finalizado = True
            st.rerun()
        else:
            st.session_state.resultado  = "ok"
            st.session_state.finalizado = True
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)


# ── RESULTADOS ────────────────────────────────────────────────────
elif st.session_state.finalizado:
    resultado = st.session_state.resultado
    doc       = st.session_state.doc_sesion
    tipo      = st.session_state.tipo_sesion
    proceso   = st.session_state.proceso_sesion
    nombre    = st.session_state.nombre_sesion
    now       = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    if resultado == "ok":
        # Lanzar confetti via JS
        st.markdown("""
        <script>
        setTimeout(function(){
            if(typeof launchConfetti === 'function') launchConfetti();
        }, 300);
        </script>""", unsafe_allow_html=True)

        st.markdown(f"""
        <div class="r-ok">
          <div style="font-size:3.5rem;margin-bottom:.4rem;">🎉</div>
          <div class="r-mono">✓ VALIDACIÓN EXITOSA</div>
          <div class="r-title">Identidad Verificada</div>
          <div class="r-sub" style="max-width:460px;margin:.6rem auto 0;">
            ¡Perfecto! La validación fue exitosa.<br>
            <b>Proceder con el proceso según corresponda.</b>
          </div>
          <table class="r-table">
            <tr><td>Documento</td><td style="font-family:monospace;">{tipo} · {doc}</td></tr>
            <tr><td>Proceso</td><td>{proceso}</td></tr>
            <tr><td>Asesor</td><td>{nombre}</td></tr>
            <tr><td>Fecha / Hora</td><td>{now}</td></tr>
          </table>
        </div>""", unsafe_allow_html=True)

    elif resultado == "bloqueado":
        datos = st.session_state.datos_bloqueo or {}
        st.markdown(f"""
        <div class="r-block">
          <div style="font-size:2.5rem;margin-bottom:.4rem;">⛔</div>
          <div class="r-title" style="color:var(--red);">CLIENTE CON FALLO PREVIO REGISTRADO</div>
          <p style="color:var(--muted);font-size:.86rem;margin:.5rem 0 1rem;">
            Este DNI/CE ya cuenta con un intento de validación fallido en el sistema.
          </p>
          <table class="r-table" style="max-width:460px;margin:0 auto;">
            <tr><td style="color:var(--muted);">Documento</td>
                <td style="font-family:monospace;font-weight:600;color:var(--text);">{doc}</td></tr>
            <tr><td style="color:var(--muted);">Asesor responsable</td>
                <td style="color:var(--text);">{datos.get('trabajador','N/D')}</td></tr>
            <tr><td style="color:var(--muted);">Fecha del fallo</td>
                <td style="color:var(--text);">{datos.get('fecha','N/D')}</td></tr>
          </table>
          <div class="block-alert">📍 ACCIÓN REQUERIDA: Derivar al cliente a una tienda física Interbank.</div>
        </div>""", unsafe_allow_html=True)

    else:
        st.markdown(f"""
        <div class="r-fail">
          <div style="font-size:3.5rem;margin-bottom:.4rem;">😔</div>
          <div class="r-mono">✗ VALIDACIÓN NO SUPERADA</div>
          <div class="r-title">No fue posible verificar la identidad</div>
          <div class="r-sub" style="max-width:520px;margin:.6rem auto 0;line-height:1.65;">
            Estimado asesor, lamentablemente el cliente no ha podido superar el proceso
            de verificación de identidad en esta ocasión. Por favor, de manera amable y
            cordial, <b>invita al cliente a acercarse a la agencia Interbank más cercana</b>,
            donde nuestro equipo presencial podrá ayudarlo con el proceso de forma segura.
            ¡Gracias por tu gestión! 🙏
          </div>
          <table class="r-table">
            <tr><td>Documento</td><td style="font-family:monospace;">{tipo} · {doc}</td></tr>
            <tr><td>Proceso</td><td>{proceso}</td></tr>
            <tr><td>Asesor</td><td>{nombre}</td></tr>
            <tr><td>Fecha / Hora</td><td>{now}</td></tr>
          </table>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="main-btn">', unsafe_allow_html=True)
    if st.button("↩  Nueva Consulta"):
        for k in ["preguntas","finalizado","resultado","datos_bloqueo",
                  "doc_sesion","tipo_sesion","proceso_sesion","user_sesion","nombre_sesion"]:
            st.session_state[k] = None if k in ("preguntas","datos_bloqueo") else \
                                   False if k=="finalizado" else ""
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
