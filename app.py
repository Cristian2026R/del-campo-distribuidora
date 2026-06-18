import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, date, timedelta
import sqlite3
import io
import base64
from pathlib import Path

st.set_page_config(page_title="DON VALENTIN", page_icon="🧀", layout="wide", initial_sidebar_state="expanded")

APP_NAME = "DON VALENTIN"
DB_PATH = Path(__file__).parent / "don_valentin.db"
DEFAULT_EXCEL = Path(__file__).parent / "lista_don_valentin.xlsx"
MOZZARELLA_IMG = Path(__file__).parent / "assets" / "mozzarella_hero.png"

try:
    MOZZARELLA_B64 = base64.b64encode(MOZZARELLA_IMG.read_bytes()).decode("utf-8") if MOZZARELLA_IMG.exists() else ""
except Exception:
    MOZZARELLA_B64 = ""

# =========================
# ESTILO
# =========================
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');
html, body, [class*="css"] {{font-family:'Inter',sans-serif;}}
.stApp {{
    background:
        linear-gradient(180deg, rgba(0,0,0,.86), rgba(0,0,0,.92)),
        radial-gradient(circle at 18% 82%, rgba(64,120,45,.26), transparent 20%),
        radial-gradient(circle at 84% 78%, rgba(64,120,45,.22), transparent 18%),
        radial-gradient(circle at 50% 15%, rgba(245,213,106,.15), transparent 30%),
        url("data:image/png;base64,{MOZZARELLA_B64}");
    background-size: cover;
    background-position: center;
    background-attachment: fixed;
    color:#F8F1D7;
}}
section[data-testid="stSidebar"]{{background:linear-gradient(180deg,#040404,#111);border-right:1px solid rgba(212,175,55,.30);}}
section[data-testid="stSidebar"] *{{color:#F8F1D7!important;}}
[data-testid="stHeader"]{{background:rgba(5,5,5,0);}}
.block-container{{padding-top:1rem; max-width:1380px;}}
.hero{{background:linear-gradient(135deg,rgba(20,20,20,.98),rgba(5,5,5,.98));border:1px solid rgba(245,213,106,.34);border-radius:28px;padding:26px 30px;margin-bottom:18px;box-shadow:0 20px 60px rgba(0,0,0,.55);}}
.hero h1{{color:#F5D56A;font-size:38px;font-weight:900;margin:0;letter-spacing:-.6px;}}
.hero p{{color:#E7DCA8;font-size:15px;font-weight:700;margin:8px 0 0 0;}}
.card{{background:linear-gradient(180deg,rgba(22,22,22,.96),rgba(10,10,10,.96));border:1px solid rgba(212,175,55,.22);border-radius:22px;padding:20px;box-shadow:0 16px 38px rgba(0,0,0,.42);margin-bottom:16px;}}
.card:hover{{border-color:rgba(245,213,106,.45);}}
.kpi-label{{color:#bba762;font-size:12px;font-weight:800;text-transform:uppercase;letter-spacing:.4px;}}
.kpi-value{{color:#F5D56A;font-size:28px;font-weight:900;margin-top:4px;}}
.kpi-note{{color:#E7DCA8;font-size:13px;margin-top:6px;}}
.success-box{{background:rgba(34,197,94,.10);border:1px solid rgba(34,197,94,.35);color:#bbf7d0;border-radius:16px;padding:14px 16px;margin:12px 0;font-weight:700;}}
.warn-box{{background:rgba(245,213,106,.08);border:1px solid rgba(245,213,106,.35);color:#FFE89A;border-radius:16px;padding:14px 16px;margin:12px 0;font-weight:700;}}
.ticket-box{{background:#fff;color:#111;border-radius:18px;padding:18px;border:2px dashed #B98A1E;font-family:monospace;box-shadow:0 14px 38px rgba(0,0,0,.35);}}
.ticket-box h3{{color:#111;text-align:center;margin:0 0 8px 0;}}
.ticket-line{{border-top:1px dashed #555;margin:8px 0;}}
.stButton>button, .stDownloadButton>button{{background:linear-gradient(135deg,#B98A1E,#F5D56A)!important;color:#111!important;border:0!important;border-radius:14px!important;font-weight:900!important;padding:12px 18px!important;box-shadow:0 12px 28px rgba(212,175,55,.18)!important;}}
.stButton>button:disabled{{background:#2b2b2b!important;color:#807244!important;border:1px solid rgba(245,213,106,.18)!important;}}
.stTextInput input,.stNumberInput input,.stSelectbox div[data-baseweb="select"],.stTextArea textarea,.stDateInput input{{background:#111!important;border:1px solid rgba(245,213,106,.35)!important;color:#F8F1D7!important;border-radius:14px!important;min-height:44px!important;}}
p, span, label, .stMarkdown, .stCaption, [data-testid="stCaptionContainer"], [data-testid="stMarkdownContainer"] p, [data-testid="stWidgetLabel"] p {{color:#E7DCA8 !important;}}
[data-testid="stMetric"]{{background:linear-gradient(180deg,rgba(22,22,22,.95),rgba(10,10,10,.95));border:1px solid rgba(212,175,55,.22);border-radius:18px;padding:16px;}}
hr{{border:0;border-top:1px solid rgba(245,213,106,.18);}}
.premium-login-card{{width:min(620px,92vw);border-radius:34px;padding:18px 24px 22px;margin:0 auto 10px;background:linear-gradient(180deg,rgba(10,10,10,.78),rgba(4,4,4,.94));backdrop-filter:blur(8px);border:1px solid rgba(245,213,106,.58);box-shadow:0 28px 95px rgba(0,0,0,.78),0 0 95px rgba(212,175,55,.18);text-align:center;}}
.food-hero{{width:100%;height:190px;object-fit:cover;object-position:center 42%;border-radius:26px;margin-bottom:14px;}}
.brand-word{{font-size:50px;font-weight:900;color:#F5D56A;letter-spacing:.8px;text-shadow:0 8px 28px rgba(245,213,106,.18);line-height:.95;}}
.brand-subline{{display:flex;align-items:center;gap:16px;justify-content:center;margin-top:14px;color:#F8F1D7;font-size:18px;font-weight:900;letter-spacing:12px;}}
.brand-subline:before,.brand-subline:after{{content:"";height:1px;width:74px;background:linear-gradient(90deg,transparent,#F5D56A);opacity:.8;}}
.brand-subline:after{{background:linear-gradient(90deg,#F5D56A,transparent);}}
</style>
""", unsafe_allow_html=True)

# =========================
# BASE DE DATOS
# =========================
def conn():
    return sqlite3.connect(DB_PATH, check_same_thread=False)

def q(sql, params=(), fetch=False):
    with conn() as c:
        cur = c.cursor()
        cur.execute(sql, params)
        c.commit()
        if fetch:
            return cur.fetchall()

def df_query(sql, params=()):
    with conn() as c:
        return pd.read_sql_query(sql, c, params=params)

def money(x):
    try:
        return "$ {:,.0f}".format(float(x)).replace(",", ".")
    except Exception:
        return "$ 0"

def to_num(x):
    if pd.isna(x): return None
    if isinstance(x, (int, float)): return float(x)
    s = str(x).strip().replace("$", "").replace(".", "").replace(",", ".")
    if not s or "NO STOCK" in s.upper() or "XXXX" in s.upper(): return None
    try: return float(s)
    except Exception: return None

def parse_price_excel(path):
    try:
        raw = pd.read_excel(path, sheet_name=0, header=None, engine="openpyxl")
    except Exception:
        return pd.DataFrame()
    productos = []
    blocks = [(0, 6, 8), (9, 15, 17)]
    for name_col, unit_col, kg_col in blocks:
        categoria = "General"
        for _, row in raw.iterrows():
            name = row.get(name_col, None)
            if pd.isna(name): continue
            nombre = str(name).strip()
            if not nombre or "whatsapp" in nombre.lower() or "lista de precios" in nombre.lower(): continue
            precio_und = to_num(row.get(unit_col, None))
            precio_kg = to_num(row.get(kg_col, None))
            if precio_und is None and precio_kg is None and len(nombre) < 35:
                categoria = nombre.title(); continue
            estado = "Sin stock" if any("NO STOCK" in str(v).upper() for v in row.values if not pd.isna(v)) else "Activo"
            if precio_und is None and precio_kg is None and estado != "Sin stock": continue
            productos.append({
                "name": nombre.title(), "category": categoria, "sale_unit": precio_und or 0,
                "sale_kg": precio_kg or 0, "cost_unit": 0, "cost_kg": 0,
                "stock_qty": 0, "stock_unit": "unidad", "active": 1
            })
    return pd.DataFrame(productos).drop_duplicates(subset=["name", "category"], keep="first") if productos else pd.DataFrame()

def init_db():
    q("""CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY, username TEXT UNIQUE, password TEXT)""")
    q("""CREATE TABLE IF NOT EXISTS products(
        id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, category TEXT, sale_unit REAL DEFAULT 0,
        sale_kg REAL DEFAULT 0, cost_unit REAL DEFAULT 0, cost_kg REAL DEFAULT 0,
        stock_qty REAL DEFAULT 0, stock_unit TEXT DEFAULT 'unidad', active INTEGER DEFAULT 1,
        created_at TEXT, updated_at TEXT)""")
    q("""CREATE TABLE IF NOT EXISTS clients(
        id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, phone TEXT, zone TEXT, type TEXT,
        notes TEXT, created_at TEXT, updated_at TEXT)""")
    q("""CREATE TABLE IF NOT EXISTS suppliers(
        id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, phone TEXT, notes TEXT, created_at TEXT, updated_at TEXT)""")
    q("""CREATE TABLE IF NOT EXISTS purchases(
        id INTEGER PRIMARY KEY AUTOINCREMENT, supplier_id INTEGER, product_text TEXT, qty REAL,
        unit_cost REAL, total REAL, paid REAL DEFAULT 0, date TEXT, notes TEXT)""")
    q("""CREATE TABLE IF NOT EXISTS sales(
        id INTEGER PRIMARY KEY AUTOINCREMENT, client_id INTEGER, client_name TEXT, date TEXT,
        payment_method TEXT, total REAL, paid REAL DEFAULT 0, status TEXT, ticket_no TEXT, created_at TEXT)""")
    q("""CREATE TABLE IF NOT EXISTS sale_items(
        id INTEGER PRIMARY KEY AUTOINCREMENT, sale_id INTEGER, product_id INTEGER, product_name TEXT,
        qty REAL, unit_type TEXT, price REAL, cost REAL, total REAL, profit REAL)""")
    q("""CREATE TABLE IF NOT EXISTS payments(
        id INTEGER PRIMARY KEY AUTOINCREMENT, client_id INTEGER, sale_id INTEGER, amount REAL, date TEXT, notes TEXT)""")
    q("""CREATE TABLE IF NOT EXISTS cash_movements(
        id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT, kind TEXT, concept TEXT, amount REAL, notes TEXT, ref TEXT)""")
    q("""CREATE TABLE IF NOT EXISTS logistics(
        id INTEGER PRIMARY KEY AUTOINCREMENT, visit_date TEXT, client_id INTEGER, client_name TEXT,
        route TEXT, status TEXT, notes TEXT)""")
    if not q("SELECT id FROM users LIMIT 1", fetch=True):
        q("INSERT INTO users(username,password) VALUES(?,?)", ("admin", "admin123"))
    if not q("SELECT id FROM products LIMIT 1", fetch=True):
        df = parse_price_excel(DEFAULT_EXCEL)
        now = datetime.now().isoformat(timespec="seconds")
        if df.empty:
            df = pd.DataFrame([
                {"name":"Mozzarella X 10kg", "category":"Mozzarellas", "sale_unit":65000, "sale_kg":6500, "cost_unit":52000, "cost_kg":5200, "stock_qty":25, "stock_unit":"kg", "active":1},
                {"name":"Jamón Cocido", "category":"Fiambres", "sale_unit":0, "sale_kg":9000, "cost_unit":0, "cost_kg":6500, "stock_qty":20, "stock_unit":"kg", "active":1},
                {"name":"Harina 0000 Bolsa 25kg", "category":"Harinas", "sale_unit":15000, "sale_kg":600, "cost_unit":10000, "cost_kg":400, "stock_qty":300, "stock_unit":"unidad", "active":1},
            ])
        for _, r in df.iterrows():
            q("""INSERT INTO products(name,category,sale_unit,sale_kg,cost_unit,cost_kg,stock_qty,stock_unit,active,created_at,updated_at)
                 VALUES(?,?,?,?,?,?,?,?,?,?,?)""", (r["name"], r["category"], r["sale_unit"], r["sale_kg"], r.get("cost_unit",0), r.get("cost_kg",0), r.get("stock_qty",0), r.get("stock_unit","unidad"), 1, now, now))
    if not q("SELECT id FROM clients LIMIT 1", fetch=True):
        now = datetime.now().isoformat(timespec="seconds")
        for name, typ in [("Pizzería La Esquina","Pizzería"),("Rotisería Avenida","Rotisería"),("Almacén Don Luis","Almacén")]:
            q("INSERT INTO clients(name,phone,zone,type,notes,created_at,updated_at) VALUES(?,?,?,?,?,?,?)", (name,"","",typ,"",now,now))

init_db()

# =========================
# HELPERS
# =========================
def header(title, subtitle):
    st.markdown(f'<div class="hero"><h1>{title}</h1><p>{subtitle}</p></div>', unsafe_allow_html=True)

def kpi(label, value, note=""):
    st.markdown(f'<div class="card"><div class="kpi-label">{label}</div><div class="kpi-value">{value}</div><div class="kpi-note">{note}</div></div>', unsafe_allow_html=True)

def styled_fig(fig, height=390):
    fig.update_layout(template="plotly_dark", height=height, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#F8F1D7"), title_font=dict(color="#F5D56A", size=20), margin=dict(l=20,r=20,t=55,b=25))
    return fig

def products_df(active_only=False):
    where = "WHERE active=1" if active_only else ""
    return df_query(f"SELECT * FROM products {where} ORDER BY category,name")

def clients_df(): return df_query("SELECT * FROM clients ORDER BY name")
def suppliers_df(): return df_query("SELECT * FROM suppliers ORDER BY name")

def get_user():
    res = q("SELECT username,password FROM users LIMIT 1", fetch=True)
    return res[0] if res else ("admin","admin123")

def set_cash(kind, concept, amount, notes="", ref=""):
    q("INSERT INTO cash_movements(date,kind,concept,amount,notes,ref) VALUES(?,?,?,?,?,?)", (datetime.now().strftime("%Y-%m-%d %H:%M"), kind, concept, float(amount or 0), notes, ref))

def ticket_html(ticket):
    rows = "".join([f"<tr><td>{i['product_name']}<br><small>{i['qty']} {i['unit_type']}</small></td><td style='text-align:right'>{money(i['total'])}</td></tr>" for i in ticket["items"]])
    return f"""<!doctype html><html><head><meta charset='utf-8'><title>Ticket {ticket['ticket_no']}</title>
<style>body{{font-family:Arial;background:#f4f4f4;padding:20px}}.ticket{{width:300px;margin:auto;background:white;padding:16px;border:1px dashed #111;color:#111}}h2,h3{{text-align:center;margin:4px 0}}.line{{border-top:1px dashed #111;margin:10px 0}}table{{width:100%;font-size:12px}}td{{padding:5px 0;border-bottom:1px dotted #bbb;vertical-align:top}}.total{{font-size:20px;font-weight:bold;text-align:right;margin-top:10px}}.btn{{display:block;text-align:center;margin:20px auto;background:#111;color:white;padding:10px;border-radius:6px;text-decoration:none;width:160px}}@media print{{.btn{{display:none}} body{{background:white;padding:0}} .ticket{{border:none;width:280px}}}}</style></head>
<body><a href='javascript:window.print()' class='btn'>Imprimir</a><div class='ticket'><h2>DON VALENTIN</h2><h3>Ticket</h3><div class='line'></div><b>N°:</b> {ticket['ticket_no']}<br><b>Fecha:</b> {ticket['date']}<br><b>Cliente:</b> {ticket['client_name']}<br><b>Pago:</b> {ticket['payment_method']}<br><div class='line'></div><table>{rows}</table><div class='line'></div><div class='total'>TOTAL {money(ticket['total'])}</div></div></body></html>"""

# =========================
# LOGIN / NAV
# =========================
def login():
    image_html = f'<img class="food-hero" src="data:image/png;base64,{MOZZARELLA_B64}" />' if MOZZARELLA_B64 else ''
    st.markdown(f"""<div class="premium-login-card">{image_html}<div class="brand-word">DON VALENTIN</div><div class="brand-subline">DISTRIBUIDORA</div></div>""", unsafe_allow_html=True)
    user0, pass0 = get_user()
    c1,c2,c3 = st.columns([1,0.82,1])
    with c2:
        user = st.text_input("Usuario", placeholder="Usuario", label_visibility="collapsed")
        pwd = st.text_input("Contraseña", type="password", placeholder="Contraseña", label_visibility="collapsed")
        if st.button("🔒 Ingresar", use_container_width=True):
            if user == user0 and pwd == pass0:
                st.session_state.logged=True; st.rerun()
            else:
                st.error("Acceso no autorizado.")

def sidebar():
    with st.sidebar:
        st.markdown("## 🧀 DON VALENTIN")
        st.markdown("**DISTRIBUIDORA**")
        st.markdown("---")
        pages = ["Dashboard","Productos","Venta / Ticket","Clientes","Proveedores","Caja","Ganancias y deudas","Logística","Reportes","Configuración"]
        icons = {"Dashboard":"📊","Productos":"📦","Venta / Ticket":"🧾","Clientes":"👥","Proveedores":"🏭","Caja":"💰","Ganancias y deudas":"📈","Logística":"🚚","Reportes":"📑","Configuración":"⚙️"}
        for p in pages:
            if st.button(f"{icons[p]} {p}", use_container_width=True): st.session_state.page=p; st.rerun()
        st.markdown("---")
        if st.button("Cerrar sesión", use_container_width=True): st.session_state.logged=False; st.rerun()

# =========================
# PAGES
# =========================
def dashboard():
    header("DON VALENTIN", "Sistema operativo con stock real, ventas, clientes, caja, proveedores, ganancias y logística histórica.")
    today = date.today().strftime("%Y-%m-%d")
    ventas = df_query("SELECT COALESCE(SUM(total),0) total FROM sales WHERE substr(date,1,10)=?", (today,)).iloc[0,0]
    ingresos = df_query("SELECT COALESCE(SUM(amount),0) total FROM cash_movements WHERE kind='Ingreso' AND substr(date,1,10)=?", (today,)).iloc[0,0]
    egresos = df_query("SELECT COALESCE(SUM(amount),0) total FROM cash_movements WHERE kind='Egreso' AND substr(date,1,10)=?", (today,)).iloc[0,0]
    ganancia = df_query("SELECT COALESCE(SUM(profit),0) total FROM sale_items si JOIN sales s ON s.id=si.sale_id WHERE substr(s.date,1,10)=?", (today,)).iloc[0,0]
    deuda = df_query("SELECT COALESCE(SUM(total-paid),0) total FROM sales WHERE total>paid").iloc[0,0]
    cols = st.columns(5)
    with cols[0]: kpi("Ventas día", money(ventas), "Facturación diaria")
    with cols[1]: kpi("Ingresos día", money(ingresos), "Caja")
    with cols[2]: kpi("Egresos día", money(egresos), "Caja")
    with cols[3]: kpi("Ganancia día", money(ganancia), "Venta - costo")
    with cols[4]: kpi("Deuda clientes", money(deuda), "Pendiente")
    p = products_df()
    low = p[p["stock_qty"] <= 0]
    if not low.empty: st.warning(f"⚠️ Hay {len(low)} productos sin stock o con stock en cero.")
    c1,c2 = st.columns(2)
    with c1:
        sales_days = df_query("SELECT substr(date,1,10) dia, SUM(total) total FROM sales GROUP BY dia ORDER BY dia DESC LIMIT 30")
        if not sales_days.empty: st.plotly_chart(styled_fig(px.line(sales_days.sort_values('dia'), x='dia', y='total', markers=True, title='Ventas históricas')), use_container_width=True)
    with c2:
        profit_prod = df_query("SELECT product_name Producto, SUM(profit) Ganancia FROM sale_items GROUP BY product_name ORDER BY Ganancia DESC LIMIT 10")
        if not profit_prod.empty: st.plotly_chart(styled_fig(px.bar(profit_prod, x='Producto', y='Ganancia', title='Ganancia por producto')), use_container_width=True)

def productos_page():
    header("Productos", "Alta, edición, eliminación, costos, precios y stock real manual.")
    df = products_df()
    with st.expander("➕ Agregar producto nuevo", expanded=False):
        c1,c2,c3 = st.columns(3)
        with c1: name=st.text_input("Nombre producto") ; cat=st.text_input("Categoría")
        with c2: sale_unit=st.number_input("Precio venta unidad",0.0,step=100.0); sale_kg=st.number_input("Precio venta kilo",0.0,step=100.0)
        with c3: cost_unit=st.number_input("Costo unidad",0.0,step=100.0); cost_kg=st.number_input("Costo kilo",0.0,step=100.0); stock=st.number_input("Stock real",0.0,step=1.0)
        unit=st.selectbox("Unidad de stock", ["unidad","kg","bolsa","caja","pack","litro"])
        if st.button("Guardar producto", use_container_width=True):
            if name.strip():
                now=datetime.now().isoformat(timespec="seconds")
                q("""INSERT INTO products(name,category,sale_unit,sale_kg,cost_unit,cost_kg,stock_qty,stock_unit,active,created_at,updated_at)
                     VALUES(?,?,?,?,?,?,?,?,?,?,?)""", (name.strip(),cat.strip() or "General",sale_unit,sale_kg,cost_unit,cost_kg,stock,unit,1,now,now))
                st.success("Producto agregado."); st.rerun()
    st.subheader("📦 Editar productos existentes")
    if df.empty: st.info("No hay productos."); return
    prod_id = st.selectbox("Seleccionar producto", df["id"].tolist(), format_func=lambda x: df[df.id==x].iloc[0]["name"])
    r = df[df.id==prod_id].iloc[0]
    c1,c2,c3 = st.columns(3)
    with c1: n=st.text_input("Nombre", value=r["name"]); cat=st.text_input("Categoría", value=r["category"])
    with c2: su=st.number_input("Precio venta unidad", value=float(r["sale_unit"]), step=100.0); sk=st.number_input("Precio venta kilo", value=float(r["sale_kg"]), step=100.0)
    with c3: cu=st.number_input("Costo unidad", value=float(r["cost_unit"]), step=100.0); ck=st.number_input("Costo kilo", value=float(r["cost_kg"]), step=100.0); stq=st.number_input("Stock real", value=float(r["stock_qty"]), step=1.0)
    unit=st.selectbox("Unidad stock", ["unidad","kg","bolsa","caja","pack","litro"], index=["unidad","kg","bolsa","caja","pack","litro"].index(r["stock_unit"]) if r["stock_unit"] in ["unidad","kg","bolsa","caja","pack","litro"] else 0)
    active=st.checkbox("Producto activo", value=bool(r["active"]))
    a,b=st.columns(2)
    with a:
        if st.button("💾 Guardar cambios", use_container_width=True):
            q("""UPDATE products SET name=?,category=?,sale_unit=?,sale_kg=?,cost_unit=?,cost_kg=?,stock_qty=?,stock_unit=?,active=?,updated_at=? WHERE id=?""", (n,cat,su,sk,cu,ck,stq,unit,1 if active else 0,datetime.now().isoformat(timespec='seconds'),prod_id))
            st.success("Producto actualizado."); st.rerun()
    with b:
        if st.button("🗑️ Eliminar producto", use_container_width=True):
            q("DELETE FROM products WHERE id=?", (prod_id,)); st.success("Producto eliminado."); st.rerun()
    view=df.copy()
    for col in ["sale_unit","sale_kg","cost_unit","cost_kg"]: view[col]=view[col].apply(money)
    st.dataframe(view.rename(columns={"name":"Producto","category":"Categoría","sale_unit":"Venta unidad","sale_kg":"Venta kg","cost_unit":"Costo unidad","cost_kg":"Costo kg","stock_qty":"Stock","stock_unit":"Unidad","active":"Activo"}), use_container_width=True, hide_index=True)

def venta_page():
    header("Venta / Ticket", "Venta con peso manual exacto, varios productos por ticket y descuento automático de stock al confirmar.")
    if "cart" not in st.session_state: st.session_state.cart=[]
    p=products_df(active_only=True); c=clients_df()
    if p.empty or c.empty: st.warning("Necesitás productos y clientes cargados."); return
    client_id=st.selectbox("Cliente", c["id"].tolist(), format_func=lambda x: c[c.id==x].iloc[0]["name"])
    method=st.selectbox("Método de pago", ["Efectivo","Transferencia","Mercado Pago","Cuenta corriente","Parcial"])
    paid=st.number_input("Monto pagado", min_value=0.0, step=100.0)
    st.subheader("Agregar producto al ticket")
    col1,col2,col3,col4=st.columns(4)
    with col1: pid=st.selectbox("Producto", p["id"].tolist(), format_func=lambda x: p[p.id==x].iloc[0]["name"])
    pr=p[p.id==pid].iloc[0]
    with col2: mode=st.radio("Modo", ["Gramos", "Unidad"], horizontal=True)
    with col3:
        if mode=="Gramos": qty=st.number_input("Gramos exactos", min_value=1.0, value=100.0, step=1.0)
        else: qty=st.number_input("Cantidad", min_value=1.0, value=1.0, step=1.0)
    with col4:
        price=(float(pr.sale_kg or 0)*qty/1000) if mode=="Gramos" else float(pr.sale_unit or 0)*qty
        cost=(float(pr.cost_kg or 0)*qty/1000) if mode=="Gramos" else float(pr.cost_unit or 0)*qty
        st.metric("Total", money(price))
    if st.button("➕ Agregar al ticket", use_container_width=True):
        stock_needed = qty/1000 if mode=="Gramos" else qty
        st.session_state.cart.append({"product_id":int(pid),"product_name":pr["name"],"qty":qty,"unit_type":"g" if mode=="Gramos" else "u.","stock_needed":stock_needed,"price":price,"cost":cost,"total":price,"profit":price-cost})
        st.rerun()
    if st.session_state.cart:
        st.subheader("🧾 Ticket actual")
        cartdf=pd.DataFrame(st.session_state.cart); cartdf["Total $"]=cartdf["total"].apply(money)
        st.dataframe(cartdf[["product_name","qty","unit_type","Total $"]].rename(columns={"product_name":"Producto","qty":"Cantidad","unit_type":"Unidad"}), use_container_width=True, hide_index=True)
        total=sum(i["total"] for i in st.session_state.cart)
        if paid==0 and method!="Cuenta corriente": paid=total
        st.metric("Total ticket", money(total))
        a,b=st.columns(2)
        with a:
            if st.button("✅ Confirmar venta, descontar stock y generar ticket", use_container_width=True):
                # validar stock
                current=products_df().set_index("id")
                for it in st.session_state.cart:
                    if float(current.loc[it["product_id"],"stock_qty"]) < float(it["stock_needed"]):
                        st.error(f"Stock insuficiente: {it['product_name']}"); return
                ticket_no="T-"+datetime.now().strftime("%Y%m%d%H%M%S")
                client_name=c[c.id==client_id].iloc[0]["name"]
                status="Pagado" if paid>=total else "Pendiente"
                now=datetime.now().strftime("%Y-%m-%d %H:%M")
                q("INSERT INTO sales(client_id,client_name,date,payment_method,total,paid,status,ticket_no,created_at) VALUES(?,?,?,?,?,?,?,?,?)", (int(client_id),client_name,now,method,total,paid,status,ticket_no,now))
                sale_id=q("SELECT last_insert_rowid()", fetch=True)[0][0]
                for it in st.session_state.cart:
                    q("INSERT INTO sale_items(sale_id,product_id,product_name,qty,unit_type,price,cost,total,profit) VALUES(?,?,?,?,?,?,?,?,?)", (sale_id,it["product_id"],it["product_name"],it["qty"],it["unit_type"],it["price"],it["cost"],it["total"],it["profit"]))
                    q("UPDATE products SET stock_qty=stock_qty-?, updated_at=? WHERE id=?", (it["stock_needed"], datetime.now().isoformat(timespec='seconds'), it["product_id"]))
                if paid>0:
                    q("INSERT INTO payments(client_id,sale_id,amount,date,notes) VALUES(?,?,?,?,?)", (int(client_id), sale_id, paid, now, "Pago al generar ticket"))
                    set_cash("Ingreso", f"Cobro ticket {ticket_no}", paid, client_name, ticket_no)
                st.session_state.last_ticket={"ticket_no":ticket_no,"date":now,"client_name":client_name,"payment_method":method,"items":st.session_state.cart.copy(),"total":total}
                st.session_state.cart=[]
                st.success("Venta guardada, stock actualizado y ticket generado."); st.rerun()
        with b:
            if st.button("🧹 Limpiar ticket", use_container_width=True): st.session_state.cart=[]; st.rerun()
    if st.session_state.get("last_ticket"):
        t=st.session_state.last_ticket
        st.markdown(f"""<div class='ticket-box'><h3>DON VALENTIN</h3><div style='text-align:center;'>Ticket</div><div class='ticket-line'></div><b>N°:</b> {t['ticket_no']}<br><b>Fecha:</b> {t['date']}<br><b>Cliente:</b> {t['client_name']}<br><b>Pago:</b> {t['payment_method']}<div class='ticket-line'></div>""", unsafe_allow_html=True)
        for it in t["items"]: st.markdown(f"<div>{it['product_name']} · {it['qty']} {it['unit_type']} <span style='float:right'>{money(it['total'])}</span></div>", unsafe_allow_html=True)
        st.markdown(f"<div class='ticket-line'></div><div style='font-size:22px;font-weight:900;text-align:right;'>TOTAL {money(t['total'])}</div></div>", unsafe_allow_html=True)
        st.download_button("⬇️ Descargar ticket HTML para imprimir", data=ticket_html(t).encode('utf-8'), file_name=f"ticket_{t['ticket_no']}.html", mime="text/html", use_container_width=True)

def clientes_page():
    header("Clientes", "Alta, edición, eliminación, pagos parciales, deuda y alertas de más de 7 días.")
    with st.expander("➕ Agregar cliente", expanded=False):
        c1,c2,c3=st.columns(3)
        with c1: name=st.text_input("Nombre/Razón social"); phone=st.text_input("Teléfono")
        with c2: zone=st.text_input("Zona"); typ=st.text_input("Tipo")
        with c3: notes=st.text_area("Notas")
        if st.button("Guardar cliente") and name.strip():
            now=datetime.now().isoformat(timespec='seconds')
            q("INSERT INTO clients(name,phone,zone,type,notes,created_at,updated_at) VALUES(?,?,?,?,?,?,?)", (name,phone,zone,typ,notes,now,now)); st.rerun()
    clients=clients_df()
    if not clients.empty:
        cid=st.selectbox("Editar cliente", clients.id.tolist(), format_func=lambda x: clients[clients.id==x].iloc[0].name)
        r=clients[clients.id==cid].iloc[0]
        c1,c2,c3=st.columns(3)
        with c1: n=st.text_input("Nombre", value=r["name"]); ph=st.text_input("Teléfono", value=r["phone"] or "")
        with c2: z=st.text_input("Zona", value=r["zone"] or ""); ty=st.text_input("Tipo", value=r["type"] or "")
        with c3: no=st.text_area("Notas", value=r["notes"] or "")
        a,b=st.columns(2)
        with a:
            if st.button("💾 Actualizar cliente", use_container_width=True): q("UPDATE clients SET name=?,phone=?,zone=?,type=?,notes=?,updated_at=? WHERE id=?", (n,ph,z,ty,no,datetime.now().isoformat(timespec='seconds'),cid)); st.rerun()
        with b:
            if st.button("🗑️ Eliminar cliente", use_container_width=True): q("DELETE FROM clients WHERE id=?", (cid,)); st.rerun()
        deuda=df_query("SELECT id,ticket_no,date,total,paid,(total-paid) deuda FROM sales WHERE client_id=? AND total>paid ORDER BY date", (cid,))
        st.subheader("Cuenta corriente")
        if not deuda.empty:
            st.dataframe(deuda.assign(total=deuda.total.apply(money), paid=deuda.paid.apply(money), deuda=deuda.deuda.apply(money)), use_container_width=True, hide_index=True)
            sid=st.selectbox("Venta a pagar", deuda.id.tolist(), format_func=lambda x: deuda[deuda.id==x].iloc[0].ticket_no)
            monto=st.number_input("Pago parcial", min_value=0.0, step=100.0)
            if st.button("Registrar pago parcial") and monto>0:
                now=datetime.now().strftime("%Y-%m-%d %H:%M")
                q("UPDATE sales SET paid=paid+?, status=CASE WHEN paid+?>=total THEN 'Pagado' ELSE 'Pendiente' END WHERE id=?", (monto,monto,sid))
                q("INSERT INTO payments(client_id,sale_id,amount,date,notes) VALUES(?,?,?,?,?)", (cid,sid,monto,now,"Pago parcial"))
                set_cash("Ingreso","Pago parcial cliente",monto,n,sid); st.rerun()
        old=df_query("SELECT client_name,ticket_no,date,total-paid deuda FROM sales WHERE total>paid AND julianday('now')-julianday(substr(date,1,10))>7")
        if not old.empty: st.error(f"⚠️ Alertas: {len(old)} deudas con más de 7 días sin pagar."); st.dataframe(old, use_container_width=True, hide_index=True)
        st.subheader("Clientes")
        st.dataframe(clients, use_container_width=True, hide_index=True)

def proveedores_page():
    header("Proveedores", "Proveedor, producto comprado manual, pagos/deudas y egresos.")
    with st.expander("➕ Agregar proveedor", expanded=False):
        name=st.text_input("Proveedor"); phone=st.text_input("Teléfono"); notes=st.text_area("Notas")
        if st.button("Guardar proveedor") and name.strip():
            now=datetime.now().isoformat(timespec='seconds'); q("INSERT INTO suppliers(name,phone,notes,created_at,updated_at) VALUES(?,?,?,?,?)", (name,phone,notes,now,now)); st.rerun()
    sups=suppliers_df()
    if not sups.empty:
        st.subheader("Registrar compra a proveedor")
        c1,c2,c3,c4=st.columns(4)
        with c1: sid=st.selectbox("Proveedor", sups.id.tolist(), format_func=lambda x: sups[sups.id==x].iloc[0].name)
        with c2: product_text=st.text_input("Producto comprado manual", placeholder="Ej: Harina 0000 bolsa 25kg")
        with c3: qty=st.number_input("Cantidad", min_value=0.0, step=1.0); unit_cost=st.number_input("Costo unitario", min_value=0.0, step=100.0)
        with c4: paid=st.number_input("Monto pagado", min_value=0.0, step=100.0)
        total=qty*unit_cost; st.metric("Total compra", money(total))
        if st.button("Guardar compra a proveedor") and product_text.strip():
            now=datetime.now().strftime("%Y-%m-%d %H:%M")
            q("INSERT INTO purchases(supplier_id,product_text,qty,unit_cost,total,paid,date,notes) VALUES(?,?,?,?,?,?,?,?)", (sid,product_text,qty,unit_cost,total,paid,now,""))
            if paid>0: set_cash("Egreso", f"Pago proveedor {sups[sups.id==sid].iloc[0].name}", paid, product_text, "Compra")
            st.success("Compra guardada."); st.rerun()
        st.subheader("Compras / deudas proveedores")
        pur=df_query("SELECT p.id,s.name Proveedor,p.product_text Producto,p.qty Cantidad,p.unit_cost Costo,p.total Total,p.paid Pagado,(p.total-p.paid) Deuda,p.date Fecha FROM purchases p LEFT JOIN suppliers s ON s.id=p.supplier_id ORDER BY p.id DESC")
        st.dataframe(pur, use_container_width=True, hide_index=True)
        if not pur.empty:
            delid=st.selectbox("Eliminar compra/deuda", pur.id.tolist())
            if st.button("Eliminar compra/deuda proveedor"):
                q("DELETE FROM purchases WHERE id=?", (delid,)); st.rerun()

def caja_page():
    header("Caja", "Ingresos, egresos, edición y eliminación manual de movimientos.")
    c1,c2,c3=st.columns(3)
    with c1: kind=st.selectbox("Tipo", ["Ingreso","Egreso"])
    with c2: concept=st.text_input("Concepto")
    with c3: amount=st.number_input("Monto", min_value=0.0, step=100.0)
    notes=st.text_area("Notas")
    if st.button("Agregar movimiento de caja") and concept.strip() and amount>0:
        set_cash(kind, concept, amount, notes); st.rerun()
    mov=df_query("SELECT * FROM cash_movements ORDER BY id DESC")
    st.dataframe(mov, use_container_width=True, hide_index=True)
    if not mov.empty:
        mid=st.selectbox("Editar/eliminar movimiento", mov.id.tolist())
        r=mov[mov.id==mid].iloc[0]
        c1,c2,c3=st.columns(3)
        with c1: k=st.selectbox("Tipo edit", ["Ingreso","Egreso"], index=0 if r.kind=="Ingreso" else 1)
        with c2: co=st.text_input("Concepto edit", value=r.concept)
        with c3: am=st.number_input("Monto edit", value=float(r.amount), step=100.0)
        no=st.text_area("Notas edit", value=r.notes or "")
        a,b=st.columns(2)
        with a:
            if st.button("Actualizar movimiento", use_container_width=True): q("UPDATE cash_movements SET kind=?,concept=?,amount=?,notes=? WHERE id=?", (k,co,am,no,mid)); st.rerun()
        with b:
            if st.button("Eliminar movimiento", use_container_width=True): q("DELETE FROM cash_movements WHERE id=?", (mid,)); st.rerun()

def ganancias_page():
    header("Ganancias y deudas", "Relación completa entre ganancias, deudas de clientes y deudas a proveedores.")
    sales_profit=df_query("SELECT COALESCE(SUM(profit),0) ganancia FROM sale_items").iloc[0,0]
    client_debt=df_query("SELECT COALESCE(SUM(total-paid),0) deuda FROM sales WHERE total>paid").iloc[0,0]
    supp_debt=df_query("SELECT COALESCE(SUM(total-paid),0) deuda FROM purchases WHERE total>paid").iloc[0,0]
    c1,c2,c3=st.columns(3)
    with c1: kpi("Ganancia acumulada", money(sales_profit))
    with c2: kpi("Deuda de clientes", money(client_debt))
    with c3: kpi("Deuda a proveedores", money(supp_debt))
    gp=df_query("SELECT product_name Producto, SUM(profit) Ganancia, SUM(total) Facturación FROM sale_items GROUP BY product_name ORDER BY Ganancia DESC")
    if not gp.empty: st.dataframe(gp.assign(Ganancia=gp.Ganancia.apply(money), Facturación=gp.Facturación.apply(money)), use_container_width=True, hide_index=True)

def logistica_page():
    header("Logística", "Organizar recorridos, clientes a visitar, estado de visita y planificación semanal.")
    c=clients_df()
    if c.empty: st.warning("Primero cargá clientes."); return
    with st.expander("➕ Agendar visita", expanded=False):
        d=st.date_input("Fecha de visita", value=date.today())
        cid=st.selectbox("Cliente", c.id.tolist(), format_func=lambda x: c[c.id==x].iloc[0].name)
        route=st.text_input("Recorrido / zona", placeholder="Ej: Centro mañana")
        notes=st.text_area("Notas")
        if st.button("Guardar visita"):
            cname=c[c.id==cid].iloc[0].name
            q("INSERT INTO logistics(visit_date,client_id,client_name,route,status,notes) VALUES(?,?,?,?,?,?)", (str(d),cid,cname,route,"Pendiente",notes)); st.rerun()
    log=df_query("SELECT * FROM logistics ORDER BY visit_date, id")
    st.dataframe(log, use_container_width=True, hide_index=True)
    if not log.empty:
        lid=st.selectbox("Modificar visita", log.id.tolist(), format_func=lambda x: f"{log[log.id==x].iloc[0].visit_date} - {log[log.id==x].iloc[0].client_name}")
        r=log[log.id==lid].iloc[0]
        status=st.selectbox("Estado", ["Pendiente","Visitado","No visitado","Reprogramar"], index=["Pendiente","Visitado","No visitado","Reprogramar"].index(r.status) if r.status in ["Pendiente","Visitado","No visitado","Reprogramar"] else 0)
        notes=st.text_area("Notas visita", value=r.notes or "")
        a,b=st.columns(2)
        with a:
            if st.button("Actualizar visita", use_container_width=True): q("UPDATE logistics SET status=?,notes=? WHERE id=?", (status,notes,lid)); st.rerun()
        with b:
            if st.button("Eliminar visita", use_container_width=True): q("DELETE FROM logistics WHERE id=?", (lid,)); st.rerun()

def reportes_page():
    header("Reportes y estadísticas", "Facturación diaria, caja, ingresos, egresos, ganancias día/semana/mes/año e históricos.")
    sales=df_query("SELECT * FROM sales ORDER BY date DESC")
    cash=df_query("SELECT * FROM cash_movements ORDER BY date DESC")
    items=df_query("SELECT si.*, s.date FROM sale_items si JOIN sales s ON s.id=si.sale_id")
    today=date.today()
    periods={"Día": today.strftime("%Y-%m-%d"), "Semana": (today-timedelta(days=7)).strftime("%Y-%m-%d"), "Mes": today.replace(day=1).strftime("%Y-%m-%d"), "Año": today.replace(month=1,day=1).strftime("%Y-%m-%d")}
    cols=st.columns(4)
    for col,(name,start) in zip(cols,periods.items()):
        gain=df_query("SELECT COALESCE(SUM(si.profit),0) total FROM sale_items si JOIN sales s ON s.id=si.sale_id WHERE substr(s.date,1,10)>=?", (start,)).iloc[0,0]
        with col: kpi(f"Ganancia {name}", money(gain))
    if not sales.empty:
        diario=df_query("SELECT substr(date,1,10) Fecha, SUM(total) Facturación, SUM(paid) Cobrado FROM sales GROUP BY Fecha ORDER BY Fecha DESC")
        st.subheader("Facturación diaria")
        st.dataframe(diario.assign(Facturación=diario.Facturación.apply(money), Cobrado=diario.Cobrado.apply(money)), use_container_width=True, hide_index=True)
    if not cash.empty:
        caja=df_query("SELECT substr(date,1,10) Fecha, kind Tipo, SUM(amount) Total FROM cash_movements GROUP BY Fecha,Tipo ORDER BY Fecha DESC")
        st.subheader("Caja diaria: ingresos y egresos")
        st.dataframe(caja.assign(Total=caja.Total.apply(money)), use_container_width=True, hide_index=True)
    if not items.empty:
        prod=df_query("SELECT product_name Producto, SUM(total) Facturación, SUM(profit) Ganancia FROM sale_items GROUP BY product_name ORDER BY Ganancia DESC")
        st.subheader("Ganancia por producto")
        st.dataframe(prod.assign(Facturación=prod.Facturación.apply(money), Ganancia=prod.Ganancia.apply(money)), use_container_width=True, hide_index=True)

def config_page():
    header("Configuración", "Cambiar usuario y contraseña.")
    user0, pass0=get_user()
    st.caption(f"Usuario actual: {user0}")
    current=st.text_input("Contraseña actual", type="password")
    new_user=st.text_input("Nuevo usuario", value=user0)
    new_pass=st.text_input("Nueva contraseña", type="password")
    confirm=st.text_input("Confirmar contraseña", type="password")
    if st.button("Guardar usuario y contraseña", use_container_width=True):
        if current!=pass0: st.error("La contraseña actual no coincide.")
        elif not new_user.strip() or not new_pass: st.warning("Completá nuevo usuario y contraseña.")
        elif new_pass!=confirm: st.error("Las contraseñas no coinciden.")
        else:
            q("UPDATE users SET username=?, password=? WHERE id=(SELECT id FROM users LIMIT 1)", (new_user.strip(), new_pass)); st.success("Credenciales actualizadas.")

# =========================
# MAIN
# =========================
if "logged" not in st.session_state: st.session_state.logged=False
if "page" not in st.session_state: st.session_state.page="Dashboard"

if not st.session_state.logged:
    login()
else:
    sidebar()
    pages={
        "Dashboard":dashboard,
        "Productos":productos_page,
        "Venta / Ticket":venta_page,
        "Clientes":clientes_page,
        "Proveedores":proveedores_page,
        "Caja":caja_page,
        "Ganancias y deudas":ganancias_page,
        "Logística":logistica_page,
        "Reportes":reportes_page,
        "Configuración":config_page,
    }
    pages[st.session_state.page]()
