import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import re
import io
from pathlib import Path
import base64

st.set_page_config(page_title="DON VALENTIN", page_icon="🧀", layout="wide", initial_sidebar_state="expanded")

APP_NAME = "DON VALENTIN"
DEMO_USER = "demo"
DEMO_PASS = "demo123"
DEFAULT_EXCEL = Path(__file__).parent / "lista_don_valentin.xlsx"
MOZZARELLA_IMG = Path(__file__).parent / "assets" / "mozzarella_hero.png"
try:
    MOZZARELLA_B64 = base64.b64encode(MOZZARELLA_IMG.read_bytes()).decode("utf-8") if MOZZARELLA_IMG.exists() else ""
except Exception:
    MOZZARELLA_B64 = ""
WHATSAPP_LINK = "https://wa.me/TUNUMERO?text=Hola,%20quiero%20solicitar%20acceso%20completo%20a%20DON%20VALENTIN"

# =========================
# ESTADO
# =========================
if "logged" not in st.session_state:
    st.session_state.logged = False
if "page" not in st.session_state:
    st.session_state.page = "Dashboard"
if "ventas_demo" not in st.session_state:
    st.session_state.ventas_demo = []
if "last_ticket" not in st.session_state:
    st.session_state.last_ticket = None
if "clientes_demo" not in st.session_state:
    st.session_state.clientes_demo = [
        "Pizzería La Esquina", "Rotisería Avenida", "Almacén Don Luis", "Casa de Comidas Norte", "Panadería Centro"
    ]

# =========================
# ESTILO NEGRO + DORADO
# =========================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');
html, body, [class*="css"] {font-family:'Inter',sans-serif;}
.stApp {background: radial-gradient(circle at top left, rgba(212,175,55,.18), transparent 30%), linear-gradient(135deg,#050505 0%,#111 55%,#050505 100%); color:#F8F1D7;}
section[data-testid="stSidebar"]{background:linear-gradient(180deg,#040404,#111);border-right:1px solid rgba(212,175,55,.30);}
section[data-testid="stSidebar"] *{color:#F8F1D7!important;}
[data-testid="stHeader"]{background:rgba(5,5,5,0);}
.block-container{padding-top:1.4rem; max-width:1320px;}
.hero{background:linear-gradient(135deg,rgba(20,20,20,.98),rgba(5,5,5,.98));border:1px solid rgba(245,213,106,.34);border-radius:30px;padding:30px 34px;margin-bottom:22px;box-shadow:0 20px 60px rgba(0,0,0,.55);}
.hero h1{color:#F5D56A;font-size:42px;font-weight:900;margin:0;letter-spacing:-.6px;}
.hero p{color:#d8c982;font-size:16px;font-weight:700;margin:8px 0 0 0;}
.demo-banner{background:linear-gradient(90deg,rgba(212,175,55,.24),rgba(245,213,106,.09));border:1px solid rgba(245,213,106,.45);color:#FFE89A;padding:14px 18px;border-radius:18px;font-weight:800;margin-bottom:18px;box-shadow:0 12px 32px rgba(212,175,55,.10);}
.card{background:linear-gradient(180deg,rgba(22,22,22,.96),rgba(10,10,10,.96));border:1px solid rgba(212,175,55,.22);border-radius:24px;padding:22px;box-shadow:0 18px 44px rgba(0,0,0,.44);transition:.22s ease;margin-bottom:16px;}
.card:hover{transform:translateY(-2px);border-color:rgba(245,213,106,.44);}
.kpi-label{color:#bba762;font-size:12px;font-weight:800;text-transform:uppercase;letter-spacing:.4px;}
.kpi-value{color:#F5D56A;font-size:30px;font-weight:900;margin-top:4px;}
.kpi-note{color:#e7dca8;font-size:13px;margin-top:6px;}
.locked-box{background:rgba(245,213,106,.08);border:1px dashed rgba(245,213,106,.45);color:#FFE89A;border-radius:18px;padding:14px 16px;margin:12px 0;font-weight:700;}
.success-box{background:rgba(34,197,94,.10);border:1px solid rgba(34,197,94,.35);color:#bbf7d0;border-radius:18px;padding:14px 16px;margin:12px 0;font-weight:700;}
.ticket-box{background:#fff;color:#111;border-radius:18px;padding:18px;border:2px dashed #B98A1E;font-family:monospace;box-shadow:0 14px 38px rgba(0,0,0,.35);}
.ticket-box h3{color:#111;text-align:center;margin:0 0 8px 0;}
.ticket-line{border-top:1px dashed #555;margin:8px 0;}
.stButton>button, .stDownloadButton>button{background:linear-gradient(135deg,#B98A1E,#F5D56A)!important;color:#111!important;border:0!important;border-radius:14px!important;font-weight:900!important;padding:12px 18px!important;box-shadow:0 12px 28px rgba(212,175,55,.18)!important;}
.stDownloadButton>button:hover, .stButton>button:hover{filter:brightness(1.05);transform:translateY(-1px);}
.stButton>button:disabled{background:#2b2b2b!important;color:#807244!important;border:1px solid rgba(245,213,106,.18)!important;}
.login-shell{min-height:88vh;display:flex;align-items:center;justify-content:center;}
.login-card{background:radial-gradient(circle at 50% 7%,rgba(245,213,106,.20),transparent 36%),linear-gradient(180deg,rgba(22,22,22,.99),rgba(4,4,4,.99));border:1px solid rgba(245,213,106,.38);border-radius:34px;padding:38px 42px 34px 42px;text-align:center;box-shadow:0 30px 90px rgba(0,0,0,.78), 0 0 80px rgba(212,175,55,.08);margin-bottom:18px;}
.mozzarella-wrap{height:185px;display:flex;align-items:center;justify-content:center;margin-bottom:2px;filter:drop-shadow(0 20px 36px rgba(212,175,55,.28));}
.mozzarella{width:210px;height:130px;border-radius:56% 44% 52% 48% / 58% 48% 52% 42%;background:radial-gradient(circle at 34% 27%,#fffef9 0%,#fff4cc 30%,#e7d09a 66%,#b98a1e 100%);box-shadow:0 24px 60px rgba(245,213,106,.24), inset -20px -15px 30px rgba(73,51,12,.28), inset 22px 18px 30px rgba(255,255,255,.86);position:relative;transform:rotate(-3deg);}
.mozzarella:before{content:"";position:absolute;width:54px;height:32px;border-radius:50%;background:linear-gradient(135deg,#1f8f3a,#8ce99a);left:-22px;bottom:22px;transform:rotate(-28deg);box-shadow:48px -18px 0 -9px #37b24d, 30px 12px 20px rgba(0,0,0,.18);}
.mozzarella:after{content:"• • •";position:absolute;top:25px;left:68px;color:#755719;font-size:22px;letter-spacing:9px;opacity:.48;}
.brand-word{font-size:48px;font-weight:900;color:#F5D56A;letter-spacing:.8px;text-shadow:0 8px 28px rgba(245,213,106,.18);line-height:1;}
.brand-subline{display:flex;align-items:center;gap:16px;justify-content:center;margin-top:18px;color:#F8F1D7;font-size:18px;font-weight:900;letter-spacing:10px;}
.brand-subline:before,.brand-subline:after{content:"";height:1px;width:74px;background:linear-gradient(90deg,transparent,#F5D56A);opacity:.8;}
.brand-subline:after{background:linear-gradient(90deg,#F5D56A,transparent);}
.stTextInput input,.stNumberInput input,.stSelectbox div[data-baseweb="select"],.stTextArea textarea{background:#111!important;border:1px solid rgba(245,213,106,.35)!important;color:#F8F1D7!important;border-radius:14px!important;min-height:50px!important;}
.stTextInput input:focus{box-shadow:0 0 0 1px rgba(245,213,106,.50),0 0 24px rgba(245,213,106,.10)!important;}
div[data-testid="stLinkButton"] a{background:linear-gradient(135deg,#B98A1E,#F5D56A)!important;color:#111!important;border:0!important;border-radius:14px!important;font-weight:900!important;padding:12px 18px!important;box-shadow:0 12px 28px rgba(212,175,55,.18)!important;text-decoration:none!important;width:100%!important;text-align:center!important;}

[data-testid="stMetric"]{background:linear-gradient(180deg,rgba(22,22,22,.95),rgba(10,10,10,.95));border:1px solid rgba(212,175,55,.22);border-radius:20px;padding:18px;}
hr{border:0;border-top:1px solid rgba(245,213,106,.18);}

/* Ajuste pedido por cliente: textos grises/secondary en dorado legible */
p, span, label, .stMarkdown, .stCaption, [data-testid="stCaptionContainer"], [data-testid="stCaptionContainer"] p,
[data-testid="stMarkdownContainer"] p, .stSelectbox label, .stTextInput label, .stNumberInput label, .stTextArea label,
[data-testid="stWidgetLabel"], [data-testid="stWidgetLabel"] p, .css-1dp5vir, .css-10trblm {
    color:#E7DCA8 !important;
}
small, .caption, .st-emotion-cache-1vbkxwb, .st-emotion-cache-16idsys p {
    color:#F5D56A !important;
}
.stInfo, .stWarning, .stSuccess, .stAlert { color:#F8F1D7 !important; }

</style>
""", unsafe_allow_html=True)


# =========================
# AJUSTE LOGIN PREMIUM COMPACTO
# =========================
st.markdown(f"""
<style>
/* Login compacto: todo visible sin usar rueda */
.stApp {{
    background:
        linear-gradient(180deg, rgba(0,0,0,.82), rgba(0,0,0,.90)),
        radial-gradient(circle at 18% 82%, rgba(64,120,45,.28), transparent 20%),
        radial-gradient(circle at 84% 78%, rgba(64,120,45,.24), transparent 18%),
        radial-gradient(circle at 50% 15%, rgba(245,213,106,.17), transparent 30%),
        url("data:image/png;base64,{MOZZARELLA_B64}");
    background-size: cover;
    background-position: center;
    background-attachment: fixed;
}}
.block-container {{
    padding-top: .45rem !important;
    padding-bottom: .25rem !important;
}}
.login-shell {{
    min-height: auto !important;
    display: flex !important;
    align-items: center !important;
    justify-content: center !important;
    padding: 0 !important;
    margin: 0 auto !important;
}}
.premium-login-card {{
    width: min(620px, 92vw) !important;
    border-radius: 34px !important;
    padding: 18px 24px 22px !important;
    margin: 0 auto 10px !important;
    background: linear-gradient(180deg, rgba(10,10,10,.78), rgba(4,4,4,.94)) !important;
    backdrop-filter: blur(8px) !important;
    border: 1px solid rgba(245,213,106,.58) !important;
    box-shadow: 0 28px 95px rgba(0,0,0,.78), 0 0 95px rgba(212,175,55,.18) !important;
}}
.food-hero {{
    width: 100% !important;
    height: 205px !important;
    object-fit: cover !important;
    object-position: center 42% !important;
    border-radius: 26px !important;
    margin-bottom: 14px !important;
}}
.brand-word {{
    font-size: clamp(42px, 5vw, 58px) !important;
    line-height: .9 !important;
}}
.brand-subline {{
    font-size: 18px !important;
    letter-spacing: 12px !important;
    margin-top: 14px !important;
}}
.brand-subline:before,.brand-subline:after {{
    width: 74px !important;
}}
.login-form-wrap {{
    width: min(620px, 92vw) !important;
    margin: 0 auto !important;
}}
.login-form-wrap .stTextInput {{
    margin-bottom: .35rem !important;
}}
.login-form-wrap .stTextInput input {{
    height: 48px !important;
    min-height: 48px !important;
    border-radius: 14px !important;
    font-size: 16px !important;
}}
.login-form-wrap .stButton>button {{
    height: 54px !important;
    border-radius: 16px !important;
    font-size: 16px !important;
    margin-top: .15rem !important;
}}
.login-form-wrap a {{
    height: 50px !important;
    min-height: 50px !important;
    border-radius: 15px !important;
}}
/* Oculta espacios extra del login */
div[data-testid="stVerticalBlock"]:has(.login-shell) {{
    gap: .35rem !important;
}}
@media (max-height: 720px) {{
    .food-hero {{ height: 170px !important; }}
    .premium-login-card {{ padding: 14px 22px 16px !important; }}
    .brand-word {{ font-size: 42px !important; }}
    .brand-subline {{ font-size: 14px !important; margin-top: 10px !important; }}
    .login-form-wrap .stTextInput input {{ height: 44px !important; min-height: 44px !important; }}
    .login-form-wrap .stButton>button {{ height: 48px !important; }}
}}
</style>
""", unsafe_allow_html=True)

# =========================
# FUNCIONES DE DATOS
# =========================
def money(x):
    try:
        return "$ {:,.0f}".format(float(x)).replace(",", ".")
    except Exception:
        return "$ 0"

def to_num(x):
    if pd.isna(x):
        return None
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip().replace("$", "").replace(".", "").replace(",", ".")
    if not s or "NO STOCK" in s.upper() or "XXXX" in s.upper():
        return None
    try:
        return float(s)
    except Exception:
        return None

def guess_category_row(name, unit_price, kg_price):
    if pd.isna(name):
        return False
    txt = str(name).strip()
    if not txt:
        return False
    return unit_price is None and kg_price is None and len(txt) < 35

def parse_price_excel(file_obj):
    if isinstance(file_obj, (bytes, bytearray)):
        file_obj = io.BytesIO(file_obj)
    else:
        try:
            file_obj.seek(0)
        except Exception:
            pass
    raw = pd.read_excel(file_obj, sheet_name=0, header=None, engine="openpyxl")
    productos = []
    blocks = [(0, 6, 8), (9, 15, 17)]
    for name_col, unit_col, kg_col in blocks:
        categoria = "General"
        for _, row in raw.iterrows():
            name = row.get(name_col, None)
            if pd.isna(name):
                continue
            nombre = str(name).strip()
            if not nombre or "whatsapp" in nombre.lower() or "lista de precios" in nombre.lower():
                continue
            precio_und = to_num(row.get(unit_col, None))
            precio_kg = to_num(row.get(kg_col, None))
            if guess_category_row(nombre, precio_und, precio_kg):
                categoria = nombre.title()
                continue
            estado = "Sin stock" if any("NO STOCK" in str(v).upper() for v in row.values if not pd.isna(v)) else "Activo"
            if precio_und is None and precio_kg is None and estado != "Sin stock":
                continue
            productos.append({
                "Código": f"DV-{len(productos)+1:04d}",
                "Producto": nombre.title(),
                "Categoría": categoria,
                "Precio unidad": precio_und if precio_und is not None else 0,
                "Precio kg": precio_kg if precio_kg is not None else 0,
                "Permite fraccionar": "Sí" if precio_kg is not None or any(w in nombre.lower() for w in ["kg", "gr", "grs", "panceta", "jamon", "paleta", "queso"]) else "No",
                "Stock demo": 0 if estado == "Sin stock" else 25 + (len(productos) * 7) % 95,
                "Estado": estado
            })
    df = pd.DataFrame(productos).drop_duplicates(subset=["Producto", "Categoría"], keep="first")
    return df.reset_index(drop=True)

@st.cache_data(show_spinner=False)
def load_default_products():
    if DEFAULT_EXCEL.exists():
        return parse_price_excel(DEFAULT_EXCEL)
    return pd.DataFrame({
        "Código": ["DV-0001"], "Producto": ["Mozzarella X 10kg"], "Categoría": ["Mozzarellas"],
        "Precio unidad": [65000], "Precio kg": [0], "Permite fraccionar": ["No"], "Stock demo": [30], "Estado": ["Activo"]
    })

def get_products():
    uploaded_bytes = st.session_state.get("uploaded_prices_bytes", None)
    if uploaded_bytes is not None:
        try:
            return parse_price_excel(uploaded_bytes)
        except Exception as e:
            st.error("No se pudo leer el Excel cargado. Revisá que sea un .xlsx válido y que no esté protegido con contraseña.")
            st.caption(f"Detalle técnico: {e}")
            return load_default_products()
    return load_default_products()

def price_for_sale(row, grams, units):
    precio_kg = float(row.get("Precio kg", 0) or 0)
    precio_und = float(row.get("Precio unidad", 0) or 0)
    if grams and grams > 0:
        if precio_kg > 0:
            return precio_kg * grams / 1000
        return precio_und * grams / 1000
    return precio_und * units

def format_catalog(df):
    out = df.copy()
    if "Precio unidad" in out.columns:
        out["Precio unidad"] = out["Precio unidad"].apply(money)
    if "Precio kg" in out.columns:
        out["Precio kg"] = out["Precio kg"].apply(money)
    return out

def excel_catalog_download(df):
    out = format_catalog(df)
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="Lista DON VALENTIN")
        ws = writer.book["Lista DON VALENTIN"]
        for cell in ws[1]:
            cell.font = cell.font.copy(bold=True)
        widths = {"A":12, "B":36, "C":22, "D":16, "E":16, "F":18, "G":12, "H":14}
        for col, width in widths.items():
            ws.column_dimensions[col].width = width
    bio.seek(0)
    return bio.getvalue()

def make_ticket_html(ticket):
    items_rows = "".join([
        f"<tr><td>{i['Producto']}<br><small>{i['Cantidad']}</small></td><td style='text-align:right'>{money(i['Total'])}</td></tr>"
        for i in ticket.get("Items", [])
    ])
    return f"""
<!doctype html>
<html>
<head><meta charset='utf-8'><title>Ticket {ticket.get('Número')}</title>
<style>
body{{font-family:Arial, sans-serif;background:#f3f3f3;padding:20px;}}
.ticket{{width:300px;margin:auto;background:white;padding:16px;border:1px dashed #111;color:#111;}}
h2,h3{{text-align:center;margin:4px 0;}}
.line{{border-top:1px dashed #111;margin:10px 0;}}
table{{width:100%;border-collapse:collapse;font-size:12px;}}
td{{padding:5px 0;border-bottom:1px dotted #bbb;vertical-align:top;}}
.total{{font-size:20px;font-weight:bold;text-align:right;margin-top:10px;}}
.btn{{display:block;text-align:center;margin:20px auto;background:#111;color:white;padding:10px;border-radius:6px;text-decoration:none;width:160px;}}
@media print{{.btn{{display:none}} body{{background:white;padding:0}} .ticket{{border:none;width:280px}}}}
</style></head>
<body>
<a href='javascript:window.print()' class='btn'>Imprimir en HP / Ctrl+P</a>
<div class='ticket'>
<h2>DON VALENTIN</h2>
<h3>Ticket interno</h3>
<div class='line'></div>
<b>N°:</b> {ticket.get('Número')}<br>
<b>Fecha:</b> {ticket.get('Fecha')}<br>
<b>Cliente:</b> {ticket.get('Cliente')}<br>
<b>Pago:</b> {ticket.get('Método de pago')}<br>
<div class='line'></div>
<table>{items_rows}</table>
<div class='line'></div>
<div class='total'>TOTAL {money(ticket.get('Total',0))}</div>
<div class='line'></div>
<p style='text-align:center;font-size:11px;'>Comprobante interno no fiscal.<br>Para factura ARCA se requiere módulo fiscal.</p>
</div>
</body></html>
"""

def banner():
    st.markdown('<div class="demo-banner">✨ Sistema comercial para distribuidora gastronómica · Productos · Precios · Fraccionamiento · Tickets.</div>', unsafe_allow_html=True)

def header(title, subtitle):
    st.markdown(f'<div class="hero"><h1>{title}</h1><p>{subtitle}</p></div>', unsafe_allow_html=True)

def kpi(label, value, note):
    st.markdown(f'<div class="card"><div class="kpi-label">{label}</div><div class="kpi-value">{value}</div><div class="kpi-note">{note}</div></div>', unsafe_allow_html=True)

def styled_fig(fig, height=390):
    fig.update_layout(template="plotly_dark", height=height, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#F8F1D7"), title_font=dict(color="#F5D56A", size=20), margin=dict(l=20,r=20,t=55,b=25))
    return fig

# =========================
# LOGIN / SIDEBAR
# =========================
def login():
    st.markdown("""
    <style>
    .blocked-wrapper {
        min-height: 92vh;
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 24px;
    }
    .blocked-card {
        width: min(760px, 94vw);
        text-align: center;
        background:
            radial-gradient(circle at 50% 0%, rgba(245,213,106,.20), transparent 34%),
            linear-gradient(180deg, rgba(18,18,18,.96), rgba(4,4,4,.98));
        border: 1px solid rgba(245,213,106,.42);
        border-radius: 34px;
        padding: 56px 46px;
        box-shadow: 0 34px 100px rgba(0,0,0,.78), 0 0 80px rgba(212,175,55,.10);
    }
    .blocked-logo {
        font-size: 56px;
        font-weight: 900;
        color: #F5D56A;
        letter-spacing: .8px;
        line-height: 1;
        text-shadow: 0 10px 32px rgba(245,213,106,.18);
    }
    .blocked-subtitle {
        margin-top: 18px;
        color: #F8F1D7;
        font-size: 18px;
        font-weight: 900;
        letter-spacing: 10px;
    }
    .blocked-line {
        width: 170px;
        height: 1px;
        margin: 32px auto;
        background: linear-gradient(90deg, transparent, #F5D56A, transparent);
    }
    .blocked-alert {
        display: inline-block;
        background: rgba(245,213,106,.10);
        border: 1px solid rgba(245,213,106,.40);
        color: #FFE89A;
        padding: 13px 20px;
        border-radius: 999px;
        font-size: 18px;
        font-weight: 900;
        margin-bottom: 22px;
    }
    .blocked-text {
        color: #E7DCA8;
        font-size: 17px;
        line-height: 1.75;
        max-width: 580px;
        margin: 0 auto;
        font-weight: 600;
    }
    .blocked-footer {
        margin-top: 30px;
        color: #BBA762;
        font-size: 14px;
        font-weight: 700;
    }
    </style>

    <div class="blocked-wrapper">
        <div class="blocked-card">
            <div class="blocked-logo">DON VALENTIN</div>
            <div class="blocked-subtitle">DISTRIBUIDORA</div>
            <div class="blocked-line"></div>
            <div class="blocked-alert">🔒 Acceso temporalmente suspendido</div>
            <div class="blocked-text">
                Esta demostración se encuentra actualmente en proceso de revisión e implementación comercial.<br><br>
                Para solicitar una presentación, capacitación o activación del sistema, comuníquese con administración.
            </div>
            <div class="blocked-footer">
                Sistema protegido · Acceso privado · Versión comercial
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

def sidebar():
    with st.sidebar:
        st.markdown("## 🧀 DON VALENTIN")
        st.markdown("**DISTRIBUIDORA**")
        st.caption("Sistema comercial premium")
        st.markdown("---")
        pages=["Dashboard","Lista de precios","Productos","Venta fraccionada","Ticket / Cobro","Clientes","Logística","Reportes","Configuración"]
        icons={"Dashboard":"📊","Lista de precios":"📄","Productos":"📦","Venta fraccionada":"⚖️","Ticket / Cobro":"🧾","Clientes":"👥","Logística":"🚚","Reportes":"📑","Configuración":"⚙️"}
        for p in pages:
            if st.button(f"{icons[p]} {p}", use_container_width=True):
                st.session_state.page=p
                st.rerun()
        st.markdown("---")
        st.success("Sistema visual activo")
        st.caption("Lista cargable · Fraccionamiento · Ticket")
        if st.button("Cerrar demo", use_container_width=True):
            st.session_state.logged=False
            st.rerun()

# =========================
# PÁGINAS
# =========================
def dashboard():
    banner()
    df = get_products()
    header("DON VALENTIN", "Sistema comercial para visualizar precios, productos, ventas por gramos, cobro y ticket simple.")
    activos = int((df["Estado"] == "Activo").sum()) if not df.empty else 0
    fracc = int((df["Permite fraccionar"] == "Sí").sum()) if not df.empty else 0
    valor_lista = df[["Precio unidad", "Precio kg"]].max(axis=1).sum() if not df.empty else 0
    c1,c2,c3,c4=st.columns(4)
    with c1: kpi("Productos cargados", f"{len(df)}", "Desde lista de precios Excel")
    with c2: kpi("Productos activos", f"{activos}", "Disponibles para vender")
    with c3: kpi("Fraccionables", f"{fracc}", "Ventas desde 100g hasta 50kg")
    with c4: kpi("Valor de lista", money(valor_lista), "Precios con signo $")
    col1,col2=st.columns([2,1])
    with col1:
        cat = df.groupby("Categoría", as_index=False).size().rename(columns={"size":"Productos"}).sort_values("Productos", ascending=False).head(12)
        st.plotly_chart(styled_fig(px.bar(cat, x="Categoría", y="Productos", text_auto=True, title="Productos por categoría"), 420), use_container_width=True)
    with col2:
        status = df.groupby("Estado", as_index=False).size().rename(columns={"size":"Cantidad"})
        st.plotly_chart(styled_fig(px.pie(status, names="Estado", values="Cantidad", hole=.55, title="Estado de lista"), 420), use_container_width=True)
    st.markdown('<div class="success-box">💡 Precios visibles con $, descarga de lista formateada, venta fraccionada y ticket simple imprimible. El ticket NO es factura ARCA.</div>', unsafe_allow_html=True)

def lista_precios():
    banner(); header("Lista de precios", "Subida de Excel de productos y precios para actualizar el catálogo comercial.")
    st.markdown('<div class="card"><b>📄 Opción para el cliente:</b> subir su Excel de precios y productos. La app interpreta productos, precios por unidad y precios por kilo cuando existen.</div>', unsafe_allow_html=True)
    uploaded = st.file_uploader("Subir Excel de precios Don Valentin", type=["xlsx", "xls"])
    if uploaded is not None:
        st.session_state.uploaded_prices_bytes = uploaded.getvalue()
        st.success("Excel cargado correctamente. El catálogo se actualizó para esta sesión.")
    df = get_products()
    c1,c2,c3=st.columns(3)
    with c1: st.metric("Total productos", len(df))
    with c2: st.metric("Categorías", df["Categoría"].nunique() if not df.empty else 0)
    with c3: st.metric("Fraccionables", int((df["Permite fraccionar"] == "Sí").sum()) if not df.empty else 0)
    st.dataframe(format_catalog(df), use_container_width=True, hide_index=True, height=430)
    st.download_button("⬇️ Descargar lista con precios en $", data=excel_catalog_download(df), file_name="lista_don_valentin_con_pesos.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

def productos_page():
    banner(); header("Productos", "Catálogo profesional con categorías, precios con $, stock demo y estado comercial.")
    df = get_products()
    c1,c2,c3=st.columns(3)
    with c1: busqueda=st.text_input("Buscar producto")
    with c2: categoria=st.selectbox("Categoría", ["Todas"] + sorted(df["Categoría"].dropna().unique().tolist()))
    with c3: fraccion=st.selectbox("Fraccionamiento", ["Todos", "Sí", "No"])
    view=df.copy()
    if busqueda:
        view=view[view["Producto"].str.contains(busqueda, case=False, na=False)]
    if categoria!="Todas":
        view=view[view["Categoría"]==categoria]
    if fraccion!="Todos":
        view=view[view["Permite fraccionar"]==fraccion]
    st.dataframe(format_catalog(view), use_container_width=True, hide_index=True, height=430)
    st.subheader("➕ Carga manual de producto")
    st.markdown('<div class="card">Formulario para cargar productos nuevos en el sistema comercial.</div>', unsafe_allow_html=True)
    a,b,c,d=st.columns(4)
    with a: st.text_input("Producto nuevo", placeholder="Ej: Queso cremoso x kg")
    with b: st.selectbox("Categoría nueva", sorted(df["Categoría"].dropna().unique().tolist()) + ["Nueva categoría"])
    with c: st.number_input("Precio unidad", min_value=0, step=100, format="%d")
    with d: st.number_input("Precio kg", min_value=0, step=100, format="%d")
    st.caption("Los precios se mostrarán con signo $ en tablas, venta y ticket.")
    st.button("Guardar producto", use_container_width=True)

def venta_fraccionada():
    banner(); header("Venta fraccionada", "Carga visual de compras por unidad o por fracción: desde 100 gramos hasta 50 kilos.")
    df = get_products()
    if df.empty:
        st.warning("No hay productos cargados.")
        return
    col1,col2=st.columns([1,1])
    with col1:
        st.subheader("⚖️ Nueva venta / simulador")
        cliente = st.selectbox("Cliente", st.session_state.clientes_demo)
        productos_lista = df["Producto"].tolist()
        prod = st.selectbox("Producto", productos_lista)
        row = df[df["Producto"] == prod].iloc[0]
        modo = st.radio("Modo de venta", ["Por gramos", "Por unidad"], horizontal=True)
        grams = 0
        units = 1
        if modo == "Por gramos":
            grams = st.selectbox(
                "Cantidad / fracción",
                [100, 200, 300, 400, 500, 600, 700, 800, 900,
                 1000, 1500, 2000, 2500, 3000, 4000, 5000,
                 10000, 15000, 20000, 25000, 30000, 40000, 50000],
                format_func=lambda x: f"{x} g" if x < 1000 else f"{x/1000:g} kg"
            )
        else:
            units = st.number_input("Unidades", min_value=1, value=1, step=1)
        metodo = st.selectbox("Método de pago", ["Efectivo", "Transferencia", "Mercado Pago", "Cuenta corriente"])
        total = price_for_sale(row, grams, units)
        st.markdown(f'<div class="card"><div class="kpi-label">Total estimado</div><div class="kpi-value">{money(total)}</div><div class="kpi-note">Producto: {prod}</div></div>', unsafe_allow_html=True)
        st.caption("Al confirmar, la app te lleva automáticamente al módulo 🧾 Ticket / Cobro para ver la vista previa antes de imprimir.")
        if st.button("Aplicar compra y generar ticket", use_container_width=True):
            numero = "T-" + datetime.now().strftime("%Y%m%d%H%M%S")
            venta = {
                "Fecha": datetime.now().strftime("%d/%m/%Y %H:%M"),
                "Cliente": cliente,
                "Producto": prod,
                "Modo": modo,
                "Cantidad": f"{grams} g" if modo == "Por gramos" else f"{units} u.",
                "Total": round(total, 2),
                "Total $": money(total),
                "Método de pago": metodo,
                "Ticket": numero,
            }
            st.session_state.ventas_demo.append(venta)
            st.session_state.last_ticket = {
                "Número": numero,
                "Fecha": venta["Fecha"],
                "Cliente": cliente,
                "Método de pago": metodo,
                "Items": [venta],
                "Total": round(total, 2),
            }
            st.success("Compra aplicada y ticket generado.")
            st.session_state.page = "Ticket / Cobro"
            st.rerun()
    with col2:
        st.subheader("📋 Últimas compras aplicadas")
        if st.session_state.ventas_demo:
            ventas = pd.DataFrame(st.session_state.ventas_demo)
        else:
            ventas = pd.DataFrame([
                {"Fecha":"Hoy 09:20","Cliente":"Pizzería La Esquina","Producto":"Mozzarella Doña Emilse X10Kg","Modo":"Por gramos","Cantidad":"500 g","Total":3250,"Total $":money(3250),"Método de pago":"Efectivo","Ticket":"T-DEMO1"},
                {"Fecha":"Hoy 10:05","Cliente":"Rotisería Avenida","Producto":"Panceta Luvianka Ahumada","Modo":"Por gramos","Cantidad":"300 g","Total":3300,"Total $":money(3300),"Método de pago":"Transferencia","Ticket":"T-DEMO2"},
                {"Fecha":"Hoy 11:12","Cliente":"Almacén Don Luis","Producto":"Aceituna Verde 1 X5Kg Garrafa","Modo":"Por unidad","Cantidad":"1 u.","Total":24000,"Total $":money(24000),"Método de pago":"Mercado Pago","Ticket":"T-DEMO3"},
            ])
        show_cols = [c for c in ["Fecha","Cliente","Producto","Cantidad","Método de pago","Total $","Ticket"] if c in ventas.columns]
        st.dataframe(ventas[show_cols], use_container_width=True, hide_index=True)
        total_dia = ventas["Total"].sum() if "Total" in ventas else 0
        st.metric("Total aplicado", money(total_dia))

def ticket_page():
    banner(); header("Ticket / Cobro", "Vista previa de ticket interno, listo para imprimir en HP Smart Tank 750 desde el navegador.")
    ticket = st.session_state.get("last_ticket")
    if ticket is None:
        st.info("Todavía no generaste un ticket. Entrá en Venta fraccionada, aplicá una compra y se generará acá.")
        ticket = {
            "Número":"T-DEMO",
            "Fecha":datetime.now().strftime("%d/%m/%Y %H:%M"),
            "Cliente":"Cliente",
            "Método de pago":"Efectivo",
            "Items":[{"Producto":"Producto fraccionado","Cantidad":"300 g","Total":1350}],
            "Total":1350,
        }
    items_html = "".join([f"<div>{i['Producto']} · {i['Cantidad']} <span style='float:right'>{money(i['Total'])}</span></div>" for i in ticket.get("Items", [])])
    st.markdown(f"""
    <div class='ticket-box'>
        <h3>DON VALENTIN</h3>
        <div style='text-align:center;'>Ticket interno</div>
        <div class='ticket-line'></div>
        <b>N°:</b> {ticket.get('Número')}<br>
        <b>Fecha:</b> {ticket.get('Fecha')}<br>
        <b>Cliente:</b> {ticket.get('Cliente')}<br>
        <b>Pago:</b> {ticket.get('Método de pago')}
        <div class='ticket-line'></div>
        {items_html}
        <div class='ticket-line'></div>
        <div style='font-size:22px;font-weight:900;text-align:right;'>TOTAL {money(ticket.get('Total',0))}</div>
        <div class='ticket-line'></div>
        <div style='text-align:center;font-size:12px;'>Comprobante interno no fiscal.<br>Para factura ARCA se requiere módulo fiscal.</div>
    </div>
    """, unsafe_allow_html=True)
    html = make_ticket_html(ticket)
    st.download_button("⬇️ Descargar ticket HTML para imprimir", data=html.encode("utf-8"), file_name=f"ticket_{ticket.get('Número','demo')}.html", mime="text/html", use_container_width=True)
    st.markdown('<div class="success-box">🖨️ Impresión real con HP Smart Tank 750: descargá el ticket, abrilo en Chrome/Edge y tocá <b>Imprimir</b> o <b>Ctrl + P</b>. Elegí la impresora HP Smart Tank 750 en Windows.</div>', unsafe_allow_html=True)
    st.caption("Ticket interno no fiscal. Para factura ARCA se requiere módulo fiscal adicional.")

def clientes_page():
    banner(); header("Clientes", "Alta de clientes, cartera comercial y ejemplo de compras fraccionadas por negocio.")

    if "clientes_registros" not in st.session_state:
        st.session_state.clientes_registros = [
            {"Cliente":"Pizzería La Esquina", "Tipo":"Pizzería", "Zona":"Centro", "Teléfono":"11 2345-6789", "Estado":"Activo"},
            {"Cliente":"Rotisería Avenida", "Tipo":"Rotisería", "Zona":"Norte", "Teléfono":"11 3456-7890", "Estado":"Activo"},
            {"Cliente":"Almacén Don Luis", "Tipo":"Almacén", "Zona":"Oeste", "Teléfono":"11 4567-8901", "Estado":"Activo"},
            {"Cliente":"Casa de Comidas Norte", "Tipo":"Comidas", "Zona":"Norte", "Teléfono":"11 5678-9012", "Estado":"Activo"},
            {"Cliente":"Panadería Centro", "Tipo":"Panadería", "Zona":"Centro", "Teléfono":"11 6789-0123", "Estado":"Activo"},
        ]

    st.subheader("➕ Cargar cliente nuevo")
    st.markdown('<div class="card">Formulario para dar de alta compradores, comercios o cuentas corrientes.</div>', unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        nuevo_cliente = st.text_input("Nombre / Razón social", placeholder="Ej: Pizzería San Martín")
        telefono = st.text_input("Teléfono", placeholder="Ej: 11 1234-5678")
    with c2:
        tipo = st.selectbox("Tipo de cliente", ["Pizzería", "Rotisería", "Panadería", "Almacén", "Restaurant", "Mayorista", "Otro"])
        zona = st.text_input("Zona / Localidad", placeholder="Ej: Centro")
    with c3:
        estado = st.selectbox("Estado", ["Activo", "Inactivo", "Cuenta corriente"])
        st.text_area("Observaciones", placeholder="Ej: compra mozzarella por semana")

    if st.button("Guardar cliente", use_container_width=True):
        if nuevo_cliente.strip():
            st.session_state.clientes_registros.append({
                "Cliente": nuevo_cliente.strip(),
                "Tipo": tipo,
                "Zona": zona.strip() or "Sin zona",
                "Teléfono": telefono.strip() or "-",
                "Estado": estado,
            })
            if nuevo_cliente.strip() not in st.session_state.clientes_demo:
                st.session_state.clientes_demo.append(nuevo_cliente.strip())
            st.success("Cliente cargado correctamente.")
        else:
            st.warning("Ingresá el nombre del cliente.")

    st.subheader("👥 Clientes cargados")
    clientes = pd.DataFrame(st.session_state.clientes_registros)
    st.dataframe(clientes, use_container_width=True, hide_index=True)

    st.subheader("📊 Vista comercial")
    resumen = clientes.groupby("Tipo", as_index=False).size().rename(columns={"size":"Cantidad"})
    if not resumen.empty:
        st.plotly_chart(styled_fig(px.bar(resumen, x="Tipo", y="Cantidad", text_auto=True, title="Clientes por tipo de negocio"), 360), use_container_width=True)

def logistica_page():
    banner(); header("Logística", "Vista comercial de preparación, despacho y entrega de pedidos.")
    rutas = pd.DataFrame({"Ruta":["Centro","Zona Norte","Zona Oeste","Zona Sur"],"Chofer":["Martín","Lucas","Diego","Sergio"],"Pedidos":[12,8,10,7],"Estado":["En reparto","Preparando","En reparto","Pendiente"]})
    st.dataframe(rutas, use_container_width=True, hide_index=True)
    st.plotly_chart(styled_fig(px.bar(rutas, x="Ruta", y="Pedidos", color="Estado", title="Pedidos por ruta")), use_container_width=True)

def reportes_page():
    banner(); header("Reportes", "Resumen comercial de lista, fraccionamiento, categorías, ventas demo y tickets.")
    df=get_products()
    resumen = df.groupby("Categoría", as_index=False).agg(Productos=("Producto", "count"), Fraccionables=("Permite fraccionar", lambda x: (x=="Sí").sum()))
    st.dataframe(resumen, use_container_width=True, hide_index=True)
    st.plotly_chart(styled_fig(px.bar(resumen.sort_values("Productos", ascending=False).head(12), x="Categoría", y="Productos", title="Top categorías")), use_container_width=True)
    if st.session_state.ventas_demo:
        ventas = pd.DataFrame(st.session_state.ventas_demo)
        st.subheader("Ventas registradas")
        st.dataframe(ventas[["Fecha","Cliente","Producto","Cantidad","Método de pago","Total $","Ticket"]], use_container_width=True, hide_index=True)

def config_page():
    banner(); header("Configuración", "Personalización general del sistema.")
    st.text_input("Nombre de la app", value="DON VALENTIN")
    st.selectbox("Tema visual", ["Negro con retoques dorados"])
    st.checkbox("Permitir carga de Excel", value=True)
    st.checkbox("Mostrar precios con signo $", value=True)
    st.checkbox("Permitir venta fraccionada por gramos", value=True)
    st.checkbox("Generar ticket simple", value=True)
    st.checkbox("Activar logística", value=True)
    st.button("Guardar configuración", use_container_width=True)

if not st.session_state.logged:
    login()
else:
    sidebar()
    pages = {
        "Dashboard": dashboard,
        "Lista de precios": lista_precios,
        "Productos": productos_page,
        "Venta fraccionada": venta_fraccionada,
        "Ticket / Cobro": ticket_page,
        "Clientes": clientes_page,
        "Logística": logistica_page,
        "Reportes": reportes_page,
        "Configuración": config_page,
    }
    pages[st.session_state.page]()