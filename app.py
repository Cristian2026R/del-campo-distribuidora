import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import re
from pathlib import Path

st.set_page_config(page_title="DON VALENTIN", page_icon="🏢", layout="wide", initial_sidebar_state="expanded")

APP_NAME = "DON VALENTIN"
DEMO_USER = "demo"
DEMO_PASS = "demo123"
DEFAULT_EXCEL = Path(__file__).parent / "lista_don_valentin.xlsx"
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
.stButton>button{background:linear-gradient(135deg,#B98A1E,#F5D56A);color:#111!important;border:0;border-radius:14px;font-weight:900;padding:12px 18px;}
.stButton>button:disabled{background:#2b2b2b!important;color:#807244!important;border:1px solid rgba(245,213,106,.18);}
.stTextInput input,.stNumberInput input,.stSelectbox div[data-baseweb="select"],.stTextArea textarea{background:#111!important;border:1px solid rgba(245,213,106,.25)!important;color:#F8F1D7!important;}
[data-testid="stMetric"]{background:linear-gradient(180deg,rgba(22,22,22,.95),rgba(10,10,10,.95));border:1px solid rgba(212,175,55,.22);border-radius:20px;padding:18px;}
hr{border:0;border-top:1px solid rgba(245,213,106,.18);}
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
    raw = pd.read_excel(file_obj, sheet_name=0, header=None)
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
    uploaded = st.session_state.get("uploaded_prices", None)
    if uploaded is not None:
        return parse_price_excel(uploaded)
    return load_default_products()

def price_for_sale(row, grams, units):
    precio_kg = float(row.get("Precio kg", 0) or 0)
    precio_und = float(row.get("Precio unidad", 0) or 0)
    if grams and grams > 0:
        if precio_kg > 0:
            return precio_kg * grams / 1000
        # Si no tiene precio por kg, se toma precio de unidad como referencia de 1kg solo para demo visual.
        return precio_und * grams / 1000
    return precio_und * units

def banner():
    st.markdown('<div class="demo-banner">✨ DON VALENTIN · DEMO COMERCIAL PREMIUM · Negro con retoques dorados · Lista de precios, productos fraccionables y ventas por gramos.</div>', unsafe_allow_html=True)

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
    st.markdown("<br><br>", unsafe_allow_html=True)
    c1,c2,c3=st.columns([1,1.15,1])
    with c2:
        st.markdown('<div class="card" style="text-align:center;"><div style="font-size:54px;">🥖</div><div style="font-size:38px;font-weight:900;color:#F5D56A;">DON VALENTIN</div><div style="font-size:17px;font-weight:800;color:#F8F1D7;">Sistema comercial para distribuidora gastronómica</div><div style="color:#c8b46a;margin-top:10px;">Demo premium · Productos · Precios · Fraccionamiento · Ventas</div></div>', unsafe_allow_html=True)
        st.info("Acceso demo: usuario **demo** · contraseña **demo123**")
        user=st.text_input("Usuario", value="demo")
        pwd=st.text_input("Contraseña", type="password", value="demo123")
        if st.button("Ingresar a la demo", use_container_width=True):
            if user==DEMO_USER and pwd==DEMO_PASS:
                st.session_state.logged=True
                st.rerun()
            else:
                st.error("Acceso no autorizado.")
        st.link_button("📞 Solicitar implementación completa", WHATSAPP_LINK, use_container_width=True)

def sidebar():
    with st.sidebar:
        st.markdown("## 🥖 DON VALENTIN")
        st.markdown("**Distribuidora gastronómica**")
        st.caption("Demo premium de venta")
        st.markdown("---")
        pages=["Dashboard","Lista de precios","Productos","Venta fraccionada","Clientes","Logística","Reportes","Configuración"]
        icons={"Dashboard":"📊","Lista de precios":"📄","Productos":"📦","Venta fraccionada":"⚖️","Clientes":"👥","Logística":"🚚","Reportes":"📑","Configuración":"⚙️"}
        for p in pages:
            if st.button(f"{icons[p]} {p}", use_container_width=True):
                st.session_state.page=p
                st.rerun()
        st.markdown("---")
        st.success("Demo visual activa")
        st.caption("Lista cargable · Fraccionamiento por gramos")
        if st.button("Cerrar demo", use_container_width=True):
            st.session_state.logged=False
            st.rerun()

# =========================
# PÁGINAS
# =========================
def dashboard():
    banner()
    df = get_products()
    header("DON VALENTIN", "Boceto comercial para que el cliente visualice el sistema por fuera y por dentro: precios, productos, ventas por gramos y operación diaria.")
    activos = int((df["Estado"] == "Activo").sum()) if not df.empty else 0
    fracc = int((df["Permite fraccionar"] == "Sí").sum()) if not df.empty else 0
    valor_lista = df[["Precio unidad", "Precio kg"]].max(axis=1).sum() if not df.empty else 0
    c1,c2,c3,c4=st.columns(4)
    with c1: kpi("Productos cargados", f"{len(df)}", "Desde lista de precios Excel")
    with c2: kpi("Productos activos", f"{activos}", "Disponibles para vender")
    with c3: kpi("Fraccionables", f"{fracc}", "Ventas por 100g, 200g, 300g...")
    with c4: kpi("Valor de lista", money(valor_lista), "Referencia comercial")
    col1,col2=st.columns([2,1])
    with col1:
        cat = df.groupby("Categoría", as_index=False).size().rename(columns={"size":"Productos"}).sort_values("Productos", ascending=False).head(12)
        st.plotly_chart(styled_fig(px.bar(cat, x="Categoría", y="Productos", text_auto=True, title="Productos por categoría"), 420), use_container_width=True)
    with col2:
        status = df.groupby("Estado", as_index=False).size().rename(columns={"size":"Cantidad"})
        st.plotly_chart(styled_fig(px.pie(status, names="Estado", values="Cantidad", hole=.55, title="Estado de lista"), 420), use_container_width=True)
    st.markdown('<div class="success-box">💡 Boceto de venta: el cliente puede subir su lista de precios, ver productos ordenados por categoría y cargar compras fraccionadas por gramos para atención diaria.</div>', unsafe_allow_html=True)

def lista_precios():
    banner(); header("Lista de precios", "Subida de Excel de productos y precios para actualizar el catálogo comercial.")
    st.markdown('<div class="card"><b>📄 Opción para el cliente:</b> subir su Excel de precios y productos. La app interpreta columnas de productos, precio por unidad y precio por kilo cuando existen.</div>', unsafe_allow_html=True)
    uploaded = st.file_uploader("Subir Excel de precios Don Valentin", type=["xlsx", "xls"])
    if uploaded is not None:
        st.session_state.uploaded_prices = uploaded
        st.success("Excel cargado correctamente. El catálogo se actualizó para esta sesión.")
    df = get_products()
    c1,c2,c3=st.columns(3)
    with c1: st.metric("Total productos", len(df))
    with c2: st.metric("Categorías", df["Categoría"].nunique() if not df.empty else 0)
    with c3: st.metric("Fraccionables", int((df["Permite fraccionar"] == "Sí").sum()) if not df.empty else 0)
    st.dataframe(df, use_container_width=True, hide_index=True, height=430)

def productos_page():
    banner(); header("Productos", "Catálogo profesional con categorías, precios, stock demo y estado comercial.")
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
    st.dataframe(view, use_container_width=True, hide_index=True, height=430)
    st.subheader("➕ Carga manual de producto")
    st.markdown('<div class="card">En la versión real, este formulario guardaría productos nuevos en base de datos. En la demo queda como boceto visual para presentar al cliente.</div>', unsafe_allow_html=True)
    a,b,c,d=st.columns(4)
    with a: st.text_input("Producto nuevo", placeholder="Ej: Queso cremoso x kg")
    with b: st.selectbox("Categoría nueva", sorted(df["Categoría"].dropna().unique().tolist()) + ["Nueva categoría"])
    with c: st.number_input("Precio unidad", min_value=0, step=100)
    with d: st.number_input("Precio kg", min_value=0, step=100)
    st.button("Guardar producto demo", use_container_width=True)

def venta_fraccionada():
    banner(); header("Venta fraccionada", "Carga visual de compras por unidad o por gramos: 100g, 200g, 300g, 500g, 1kg, etc.")
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
            grams = st.selectbox("Cantidad", [100, 200, 250, 300, 500, 750, 1000, 1500, 2000])
        else:
            units = st.number_input("Unidades", min_value=1, value=1, step=1)
        total = price_for_sale(row, grams, units)
        st.markdown(f'<div class="card"><div class="kpi-label">Total estimado</div><div class="kpi-value">{money(total)}</div><div class="kpi-note">Producto: {prod}</div></div>', unsafe_allow_html=True)
        if st.button("Aplicar compra demo", use_container_width=True):
            st.session_state.ventas_demo.append({
                "Fecha": datetime.now().strftime("%d/%m/%Y %H:%M"),
                "Cliente": cliente,
                "Producto": prod,
                "Modo": modo,
                "Cantidad": f"{grams} g" if modo == "Por gramos" else f"{units} u.",
                "Total": round(total, 2)
            })
            st.success("Compra aplicada en historial demo.")
    with col2:
        st.subheader("📋 Últimas compras aplicadas")
        if st.session_state.ventas_demo:
            ventas = pd.DataFrame(st.session_state.ventas_demo)
        else:
            ventas = pd.DataFrame([
                {"Fecha":"Hoy 09:20","Cliente":"Pizzería La Esquina","Producto":"Mozzarella Doña Emilse X10Kg","Modo":"Por gramos","Cantidad":"500 g","Total":3250},
                {"Fecha":"Hoy 10:05","Cliente":"Rotisería Avenida","Producto":"Panceta Luvianka Ahumada","Modo":"Por gramos","Cantidad":"300 g","Total":3300},
                {"Fecha":"Hoy 11:12","Cliente":"Almacén Don Luis","Producto":"Aceituna Verde 1 X5Kg Garrafa","Modo":"Por unidad","Cantidad":"1 u.","Total":24000},
            ])
        st.dataframe(ventas, use_container_width=True, hide_index=True)
        total_dia = ventas["Total"].sum() if "Total" in ventas else 0
        st.metric("Total aplicado demo", money(total_dia))

def clientes_page():
    banner(); header("Clientes", "Clientes comerciales y ejemplo de compras fraccionadas por negocio.")
    clientes = pd.DataFrame({
        "Cliente": st.session_state.clientes_demo,
        "Tipo": ["Pizzería", "Rotisería", "Almacén", "Comidas", "Panadería"],
        "Zona": ["Centro", "Norte", "Oeste", "Sur", "Centro"],
        "Estado": ["Activo", "Activo", "Activo", "Activo", "Activo"]
    })
    st.dataframe(clientes, use_container_width=True, hide_index=True)

def logistica_page():
    banner(); header("Logística", "Vista comercial de preparación, despacho y entrega de pedidos.")
    rutas = pd.DataFrame({"Ruta":["Centro","Zona Norte","Zona Oeste","Zona Sur"],"Chofer":["Martín","Lucas","Diego","Sergio"],"Pedidos":[12,8,10,7],"Estado":["En reparto","Preparando","En reparto","Pendiente"]})
    st.dataframe(rutas, use_container_width=True, hide_index=True)
    st.plotly_chart(styled_fig(px.bar(rutas, x="Ruta", y="Pedidos", color="Estado", title="Pedidos por ruta")), use_container_width=True)

def reportes_page():
    banner(); header("Reportes", "Resumen comercial de lista, fraccionamiento, categorías y ventas demo.")
    df=get_products()
    resumen = df.groupby("Categoría", as_index=False).agg(Productos=("Producto", "count"), Fraccionables=("Permite fraccionar", lambda x: (x=="Sí").sum()))
    st.dataframe(resumen, use_container_width=True, hide_index=True)
    st.plotly_chart(styled_fig(px.bar(resumen.sort_values("Productos", ascending=False).head(12), x="Categoría", y="Productos", title="Top categorías")), use_container_width=True)

def config_page():
    banner(); header("Configuración", "Boceto de personalización para versión real.")
    st.text_input("Nombre de la app", value="DON VALENTIN")
    st.selectbox("Tema visual", ["Negro con retoques dorados"])
    st.checkbox("Permitir carga de Excel", value=True)
    st.checkbox("Permitir venta fraccionada por gramos", value=True)
    st.checkbox("Activar logística", value=True)
    st.button("Guardar configuración demo", use_container_width=True)

if not st.session_state.logged:
    login()
else:
    sidebar()
    pages = {
        "Dashboard": dashboard,
        "Lista de precios": lista_precios,
        "Productos": productos_page,
        "Venta fraccionada": venta_fraccionada,
        "Clientes": clientes_page,
        "Logística": logistica_page,
        "Reportes": reportes_page,
        "Configuración": config_page,
    }
    pages[st.session_state.page]()
