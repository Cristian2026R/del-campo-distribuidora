# app.py
# ============================================================
# DEL CAMPO DISTRIBUIDORA - DEMO PREMIUM
# Streamlit + Pandas + Plotly + SQLite básico
# ============================================================

import streamlit as st
import pandas as pd
import plotly.express as px
import sqlite3
from datetime import datetime, date, timedelta
from pathlib import Path
import random

APP_NAME = "DEL CAMPO DISTRIBUIDORA"
DB_PATH = "del_campo_demo.db"

st.set_page_config(
    page_title=APP_NAME,
    page_icon="DC",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# CSS PREMIUM NEGRO + DORADO
# ============================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.stApp { background: radial-gradient(circle at top left, #2b230f 0%, #0b0b0d 35%, #050506 100%); color: #f5f0df; }
.block-container { padding-top: 1.4rem; padding-bottom: 3rem; }
section[data-testid="stSidebar"] { background: linear-gradient(180deg, #060607 0%, #111111 55%, #1b1608 100%); border-right: 1px solid rgba(212,175,55,.28); }
section[data-testid="stSidebar"] * { color: #f7e7aa !important; }

.main-hero { background: linear-gradient(135deg, rgba(212,175,55,.20), rgba(10,10,10,.96)); border: 1px solid rgba(212,175,55,.35); border-radius: 28px; padding: 30px; box-shadow: 0 24px 60px rgba(0,0,0,.45); margin-bottom: 24px; }
.hero-title { font-size: 38px; font-weight: 900; color: #f7d774; letter-spacing: -1px; margin: 0; }
.hero-subtitle { font-size: 16px; color: #d9d1b3; margin-top: 8px; }
.logo-box { width: 64px; height: 64px; border-radius: 20px; display:flex; align-items:center; justify-content:center; background: linear-gradient(135deg, #f7d774, #9f7928); color:#050506; font-size: 26px; font-weight: 900; box-shadow: 0 0 35px rgba(212,175,55,.35); }

.premium-card { background: linear-gradient(180deg, rgba(25,25,25,.98), rgba(12,12,12,.98)); border: 1px solid rgba(212,175,55,.22); border-radius: 24px; padding: 22px; box-shadow: 0 18px 45px rgba(0,0,0,.35); transition: all .25s ease; min-height: 126px; }
.premium-card:hover { transform: translateY(-4px); border-color: rgba(247,215,116,.65); box-shadow: 0 22px 55px rgba(212,175,55,.16); }
.kpi-label { color:#b8ae8a; font-size:13px; font-weight:700; text-transform:uppercase; letter-spacing:.5px; }
.kpi-value { color:#fff7d6; font-size:31px; font-weight:900; margin-top:8px; }
.kpi-note { color:#d4af37; font-size:13px; font-weight:700; margin-top:8px; }

.section-title { color:#f7d774; font-size:24px; font-weight:900; margin: 18px 0 12px; }
.small-muted { color:#b8ae8a; font-size:14px; }
.module-header { background: linear-gradient(135deg, rgba(212,175,55,.18), rgba(15,15,15,.95)); border:1px solid rgba(212,175,55,.28); border-radius:24px; padding:24px; margin-bottom:20px; box-shadow: 0 16px 40px rgba(0,0,0,.34); }
.module-header h2 { color:#f7d774; font-size:30px; font-weight:900; margin:0; }
.module-header p { color:#d9d1b3; margin:8px 0 0; }

.stButton > button { background: linear-gradient(135deg, #d4af37, #8d6d1f) !important; color: #080808 !important; border: none !important; border-radius: 14px !important; font-weight: 900 !important; padding: 0.65rem 1rem !important; box-shadow: 0 10px 25px rgba(212,175,55,.18); }
.stButton > button:hover { transform: translateY(-2px); box-shadow: 0 14px 35px rgba(212,175,55,.28); }
[data-testid="stMetric"] { background: linear-gradient(180deg, rgba(25,25,25,.95), rgba(10,10,10,.95)); border: 1px solid rgba(212,175,55,.22); border-radius: 18px; padding: 18px; }
[data-testid="stDataFrame"] { border: 1px solid rgba(212,175,55,.18); border-radius: 18px; overflow: hidden; }
hr { border: none; height: 1px; background: rgba(212,175,55,.22); }
.login-wrap { max-width: 460px; margin: 70px auto 20px; background: linear-gradient(180deg, rgba(25,25,25,.98), rgba(8,8,8,.98)); border: 1px solid rgba(212,175,55,.35); border-radius: 30px; padding: 34px; box-shadow: 0 24px 70px rgba(0,0,0,.50); text-align:center; }
.login-title { color:#f7d774; font-size:34px; font-weight:900; margin-top:16px; }
.login-sub { color:#d9d1b3; margin-bottom:18px; }
.badge { display:inline-block; padding:7px 12px; border-radius:999px; background:rgba(212,175,55,.15); border:1px solid rgba(212,175,55,.35); color:#f7d774; font-weight:800; font-size:12px; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# BASE DE DATOS SIMPLE PARA PRODUCTOS CARGABLES
# ============================================================
def con():
    return sqlite3.connect(DB_PATH, check_same_thread=False)

def init_db():
    c = con()
    c.execute('''CREATE TABLE IF NOT EXISTS productos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        codigo TEXT UNIQUE,
        nombre TEXT,
        categoria TEXT,
        marca TEXT,
        precio_compra REAL,
        precio_venta REAL,
        stock INTEGER,
        stock_minimo INTEGER,
        proveedor TEXT,
        estado TEXT,
        fecha TEXT
    )''')
    c.commit()
    if pd.read_sql("SELECT COUNT(*) as n FROM productos", c)["n"].iloc[0] == 0:
        demo = [
            ("P001","Harina 000 x 25kg","Harinas","Molino Norte",9800,12800,180,40,"Molino Norte","Activo"),
            ("P002","Aceite girasol 900ml","Almacén","Solcampo",1150,1690,84,25,"Aceitera Sur","Activo"),
            ("P003","Yerba mate 1kg","Infusiones","La Estancia",3100,4550,42,20,"Mayorista Verde","Activo"),
            ("P004","Azúcar 1kg","Almacén","Dulce Campo",780,1190,36,30,"Distribuidora Dulce","Activo"),
            ("P005","Gaseosa cola 2.25L","Bebidas","Fresca",1450,2450,22,30,"Bebidas Centro","Activo"),
            ("P006","Lavandina 1L","Limpieza","Brillo",520,890,65,20,"Limpieza Total","Activo"),
        ]
        for x in demo:
            c.execute("""INSERT OR IGNORE INTO productos 
            (codigo,nombre,categoria,marca,precio_compra,precio_venta,stock,stock_minimo,proveedor,estado,fecha)
            VALUES (?,?,?,?,?,?,?,?,?,?,?)""", (*x, datetime.now().strftime("%Y-%m-%d")))
        c.commit()
    c.close()

init_db()

def get_productos():
    c = con(); df = pd.read_sql("SELECT * FROM productos ORDER BY id DESC", c); c.close(); return df

def add_producto(data):
    c = con()
    c.execute("""INSERT OR REPLACE INTO productos
    (codigo,nombre,categoria,marca,precio_compra,precio_venta,stock,stock_minimo,proveedor,estado,fecha)
    VALUES (?,?,?,?,?,?,?,?,?,?,?)""", data)
    c.commit(); c.close()

# ============================================================
# DATOS DEMO VIVOS
# ============================================================
random.seed(8)
clientes = pd.DataFrame({
    "Cliente": ["Autoservicio Los Pinos", "Despensa Don Mario", "Supermercado Avenida", "Kiosco Central", "Mayorista San Telmo", "Almacén La Esquina"],
    "Zona": ["CABA", "GBA Oeste", "GBA Sur", "CABA", "GBA Norte", "GBA Oeste"],
    "Estado": ["Activo", "Activo", "Activo", "Activo", "Activo", "Con deuda"],
    "Deuda": [180000, 45000, 0, 82000, 315000, 126000],
    "Última compra": ["Hoy", "Ayer", "Hoy", "Hace 2 días", "Hoy", "Hace 5 días"]
})
ventas = pd.DataFrame({
    "Día": list(range(1, 31)),
    "Ventas": [random.randint(180000, 650000) for _ in range(30)],
    "Ganancia": [random.randint(52000, 210000) for _ in range(30)]
})
logistica = pd.DataFrame({
    "Pedido": ["PED-2041", "PED-2042", "PED-2043", "PED-2044", "PED-2045"],
    "Cliente": ["Autoservicio Los Pinos", "Mayorista San Telmo", "Kiosco Central", "Despensa Don Mario", "Supermercado Avenida"],
    "Chofer": ["Martín", "Lucas", "Diego", "Sergio", "Nicolás"],
    "Zona": ["CABA", "GBA Norte", "CABA", "GBA Oeste", "GBA Sur"],
    "Estado": ["En reparto", "Preparación", "Entregado", "Pendiente", "En reparto"],
    "Importe": [245000, 520000, 78000, 132000, 410000]
})
movimientos = pd.DataFrame({
    "Hora": ["08:40", "09:15", "10:05", "11:30", "13:20", "15:10"],
    "Movimiento": ["Venta contado", "Pago cliente", "Egreso combustible", "Venta transferencia", "Pago proveedor", "Venta cuenta corriente"],
    "Tipo": ["Ingreso", "Ingreso", "Egreso", "Ingreso", "Egreso", "Ingreso"],
    "Importe": [185000, 90000, -35000, 240000, -120000, 310000]
})

# ============================================================
# SESIÓN
# ============================================================
if "logged" not in st.session_state: st.session_state.logged = False
if "page" not in st.session_state: st.session_state.page = "Inicio"

# ============================================================
# COMPONENTES
# ============================================================
def logo():
    st.markdown('<div class="logo-box">DC</div>', unsafe_allow_html=True)

def kpi(label, value, note):
    st.markdown(f"""
    <div class="premium-card">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value}</div>
        <div class="kpi-note">{note}</div>
    </div>
    """, unsafe_allow_html=True)

def header(title, subtitle):
    st.markdown(f"""
    <div class="module-header">
        <h2>{title}</h2>
        <p>{subtitle}</p>
    </div>
    """, unsafe_allow_html=True)

def sidebar():
    with st.sidebar:
        logo()
        st.markdown("### DEL CAMPO")
        st.markdown("**DISTRIBUIDORA**")
        st.markdown('<span class="badge">DEMO COMERCIAL</span>', unsafe_allow_html=True)
        st.markdown("---")
        pages = {
            "Inicio":"🏠", "Productos":"📦", "Ventas":"🧾", "Clientes":"👥", "Caja":"💰", "Logística":"🚚", "Reportes":"📊", "Inteligencia":"🧠", "Configuración":"⚙️"
        }
        for p, icon in pages.items():
            if st.button(f"{icon} {p}", use_container_width=True):
                st.session_state.page = p; st.rerun()
        st.markdown("---")
        st.caption("Usuario demo: Administración")
        if st.button("Cerrar sesión", use_container_width=True):
            st.session_state.logged = False; st.rerun()

# ============================================================
# LOGIN
# ============================================================
def login():
    st.markdown('<div class="login-wrap"><div style="display:flex; justify-content:center;"><div class="logo-box">DC</div></div><div class="login-title">DEL CAMPO</div><div class="login-sub">Distribuidora · Sistema comercial premium</div><span class="badge">Acceso demo</span></div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1,1.15,1])
    with c2:
        user = st.text_input("Usuario", value="admin")
        pwd = st.text_input("Contraseña", value="admin123", type="password")
        if st.button("Ingresar", use_container_width=True):
            st.session_state.logged = True; st.rerun()
        st.caption("Demo: admin / admin123")

# ============================================================
# PÁGINAS
# ============================================================
def page_inicio():
    st.markdown(f"""
    <div class="main-hero">
        <div style="display:flex; gap:18px; align-items:center;">
            <div class="logo-box">DC</div>
            <div>
                <h1 class="hero-title">DEL CAMPO DISTRIBUIDORA</h1>
                <div class="hero-subtitle">Demo premium para gestión diaria: ventas, stock, caja, clientes y logística.</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    dfp = get_productos()
    c1,c2,c3,c4,c5 = st.columns(5)
    with c1: kpi("Ventas del día", "$ 1.245.000", "+18% vs ayer")
    with c2: kpi("Caja actual", "$ 785.000", "Operación positiva")
    with c3: kpi("Pedidos activos", "24", "8 en reparto")
    with c4: kpi("Productos", str(len(dfp)), "Catálogo activo")
    with c5: kpi("Stock crítico", str((dfp['stock'] <= dfp['stock_minimo']).sum()), "Revisión necesaria")

    st.markdown('<div class="section-title">Panel comercial</div>', unsafe_allow_html=True)
    a,b = st.columns([2,1])
    with a:
        fig = px.line(ventas, x="Día", y=["Ventas","Ganancia"], markers=True, title="Evolución mensual de ventas y ganancia")
        fig.update_layout(template="plotly_dark", paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', height=410, font_color="#f5f0df")
        st.plotly_chart(fig, use_container_width=True)
    with b:
        mix = pd.DataFrame({"Tipo":["Contado","Transferencia","Cuenta corriente","Mercado Pago"], "Importe":[42,28,20,10]})
        fig2 = px.pie(mix, names="Tipo", values="Importe", hole=.55, title="Métodos de pago")
        fig2.update_layout(template="plotly_dark", paper_bgcolor='rgba(0,0,0,0)', height=410, font_color="#f5f0df")
        st.plotly_chart(fig2, use_container_width=True)

    c,d = st.columns(2)
    with c:
        st.markdown('<div class="section-title">🚚 Repartos en movimiento</div>', unsafe_allow_html=True)
        st.dataframe(logistica, use_container_width=True, hide_index=True)
    with d:
        st.markdown('<div class="section-title">💰 Movimientos de caja</div>', unsafe_allow_html=True)
        st.dataframe(movimientos, use_container_width=True, hide_index=True)

def page_productos():
    header("Productos", "Carga producto por producto, control de stock y alertas comerciales.")
    df = get_productos()
    with st.expander("➕ Cargar nuevo producto", expanded=True):
        c1,c2,c3 = st.columns(3)
        with c1:
            codigo = st.text_input("Código")
            nombre = st.text_input("Nombre del producto")
            categoria = st.selectbox("Categoría", ["Almacén","Bebidas","Harinas","Limpieza","Lácteos","Golosinas","Otros"])
        with c2:
            marca = st.text_input("Marca")
            proveedor = st.text_input("Proveedor")
            estado = st.selectbox("Estado", ["Activo","Inactivo"])
        with c3:
            pc = st.number_input("Precio compra", min_value=0.0, step=100.0)
            pv = st.number_input("Precio venta", min_value=0.0, step=100.0)
            stock = st.number_input("Stock", min_value=0, step=1)
            minimo = st.number_input("Stock mínimo", min_value=0, step=1)
        if st.button("Guardar producto", use_container_width=True):
            if codigo and nombre:
                add_producto((codigo,nombre,categoria,marca,pc,pv,stock,minimo,proveedor,estado,datetime.now().strftime("%Y-%m-%d")))
                st.success("Producto guardado correctamente."); st.rerun()
            else:
                st.error("Completá código y nombre.")
    st.markdown('<div class="section-title">Inventario demo</div>', unsafe_allow_html=True)
    df["Alerta"] = df.apply(lambda r: "🔴 Bajo stock" if r["stock"] <= r["stock_minimo"] else "🟢 OK", axis=1)
    st.dataframe(df[["codigo","nombre","categoria","marca","precio_venta","stock","stock_minimo","proveedor","Alerta"]], use_container_width=True, hide_index=True)

def page_ventas():
    header("Ventas", "Simulación comercial para que el cliente visualice el uso diario.")
    dfp = get_productos()
    c1,c2 = st.columns([1,1])
    with c1:
        cliente = st.selectbox("Cliente", clientes["Cliente"])
        prod = st.selectbox("Producto", dfp["nombre"] if not dfp.empty else [])
        cant = st.number_input("Cantidad", min_value=1, value=1)
        tipo = st.selectbox("Tipo de venta", ["Blanco", "Negro"])
        pago = st.selectbox("Método de pago", ["Efectivo","Transferencia","Cuenta corriente","Mercado Pago","Cheque"])
        st.button("Registrar venta demo", use_container_width=True)
    with c2:
        st.info("Esta pantalla permite mostrar cómo un empleado cargaría ventas todos los días.")
        st.metric("Total estimado", "$ 128.500", "+ margen incluido")
        st.metric("Stock posterior", "Disponible", "validación demo")
    st.markdown('<div class="section-title">Últimas ventas demo</div>', unsafe_allow_html=True)
    ult = pd.DataFrame({"Fecha":["Hoy","Hoy","Ayer","Ayer"],"Cliente":clientes["Cliente"].head(4),"Importe":[245000,78000,132000,410000],"Pago":["Transferencia","Efectivo","Cuenta corriente","Mercado Pago"]})
    st.dataframe(ult, use_container_width=True, hide_index=True)

def page_clientes():
    header("Clientes", "Base comercial, deuda, historial y oportunidades de venta.")
    c1,c2,c3 = st.columns(3)
    with c1: st.metric("Clientes activos", "148", "+12 este mes")
    with c2: st.metric("Clientes con deuda", "31", "seguimiento")
    with c3: st.metric("Deuda total", "$ 2.180.000", "+5%")
    st.dataframe(clientes, use_container_width=True, hide_index=True)
    fig = px.bar(clientes, x="Cliente", y="Deuda", color="Zona", title="Deuda por cliente")
    fig.update_layout(template="plotly_dark", paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', height=390, font_color="#f5f0df")
    st.plotly_chart(fig, use_container_width=True)

def page_caja():
    header("Caja", "Ingresos, egresos y control diario para administración.")
    c1,c2,c3,c4 = st.columns(4)
    with c1: st.metric("Ingresos", "$ 1.245.000")
    with c2: st.metric("Egresos", "$ 340.000")
    with c3: st.metric("Saldo", "$ 905.000")
    with c4: st.metric("Pendiente cobro", "$ 520.000")
    st.dataframe(movimientos, use_container_width=True, hide_index=True)

def page_logistica():
    header("Logística", "Vista visual de reparto, choferes, zonas y estado de pedidos.")
    st.dataframe(logistica, use_container_width=True, hide_index=True)
    fig = px.bar(logistica, x="Zona", y="Importe", color="Estado", title="Repartos por zona")
    fig.update_layout(template="plotly_dark", paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', height=410, font_color="#f5f0df")
    st.plotly_chart(fig, use_container_width=True)
    st.success("Demo visual: seguimiento de pedidos, zonas, choferes y confirmación de entrega.")

def page_reportes():
    header("Reportes", "Informes exportables para toma de decisiones.")
    option = st.selectbox("Reporte", ["Ventas mensuales", "Stock", "Clientes", "Caja", "Logística"])
    st.download_button("Descargar reporte demo CSV", ventas.to_csv(index=False).encode("utf-8"), "reporte_del_campo.csv", "text/csv", use_container_width=True)
    st.dataframe(ventas, use_container_width=True, hide_index=True)

def page_inteligencia():
    header("Inteligencia comercial", "Lectura ejecutiva para detectar oportunidades y riesgos.")
    c1,c2 = st.columns(2)
    with c1:
        top = pd.DataFrame({"Producto":["Harina 000", "Aceite", "Yerba", "Gaseosa", "Azúcar"], "Ventas":[120,98,75,62,58]})
        fig = px.bar(top, x="Ventas", y="Producto", orientation="h", title="Productos más vendidos")
        fig.update_layout(template="plotly_dark", paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', height=410, font_color="#f5f0df")
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        st.warning("Cliente Mayorista San Telmo superó el umbral de deuda.")
        st.success("Harinas y bebidas muestran alta rotación esta semana.")
        st.info("Sugerencia demo: aumentar stock mínimo en productos de consumo masivo.")

def page_config():
    header("Configuración", "Parámetros visuales y comerciales de la demo.")
    st.text_input("Nombre comercial", APP_NAME)
    st.selectbox("Tema", ["Negro y dorado premium"])
    st.checkbox("Activar módulo logística", True)
    st.checkbox("Activar ventas blanco/negro", True)
    st.checkbox("Activar alertas visuales", True)
    st.button("Guardar configuración demo", use_container_width=True)

# ============================================================
# MAIN
# ============================================================
if not st.session_state.logged:
    login()
else:
    sidebar()
    page = st.session_state.page
    if page == "Inicio": page_inicio()
    elif page == "Productos": page_productos()
    elif page == "Ventas": page_ventas()
    elif page == "Clientes": page_clientes()
    elif page == "Caja": page_caja()
    elif page == "Logística": page_logistica()
    elif page == "Reportes": page_reportes()
    elif page == "Inteligencia": page_inteligencia()
    elif page == "Configuración": page_config()
