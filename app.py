import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta, date
from pathlib import Path
import sqlite3, hashlib, io, base64, re

BASE = Path(__file__).parent
DB = BASE / "don_valentin_data.sqlite3"
AUTH_DB = BASE / "don_valentin_auth.sqlite3"
DEFAULT_EXCEL = BASE / "lista_don_valentin.xlsx"
MOZZARELLA_IMG = BASE / "assets" / "mozzarella_hero.png"

st.set_page_config(page_title="DON VALENTIN", page_icon="🧀", layout="wide", initial_sidebar_state="expanded")

try:
    MOZZ_B64 = base64.b64encode(MOZZARELLA_IMG.read_bytes()).decode("utf-8") if MOZZARELLA_IMG.exists() else ""
except Exception:
    MOZZ_B64 = ""

# =========================
# ESTILOS
# =========================
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800;900&display=swap');
html, body, [class*="css"] {{font-family:'Inter',sans-serif;}}
.stApp {{background:linear-gradient(180deg,rgba(0,0,0,.86),rgba(0,0,0,.92)), url("data:image/png;base64,{MOZZ_B64}"); background-size:cover; background-attachment:fixed; color:#F8F1D7;}}
[data-testid="stHeader"]{{background:rgba(0,0,0,0);}}
section[data-testid="stSidebar"]{{background:linear-gradient(180deg,#050505,#111);border-right:1px solid rgba(245,213,106,.25);}}
section[data-testid="stSidebar"] *{{color:#F8F1D7!important;}}
.block-container{{padding-top:1rem;max-width:1320px;}}
.hero{{background:linear-gradient(135deg,rgba(18,18,18,.96),rgba(5,5,5,.96));border:1px solid rgba(245,213,106,.34);border-radius:28px;padding:28px 34px;margin-bottom:18px;box-shadow:0 20px 60px rgba(0,0,0,.5);}}
.hero h1{{color:#F5D56A;font-size:42px;font-weight:900;margin:0;}}
.hero p{{color:#E7DCA8;font-size:16px;font-weight:700;margin:8px 0 0 0;}}
.card{{background:linear-gradient(180deg,rgba(22,22,22,.96),rgba(8,8,8,.96));border:1px solid rgba(212,175,55,.22);border-radius:22px;padding:18px;box-shadow:0 14px 38px rgba(0,0,0,.38);margin-bottom:14px;}}
.demo-banner{{background:linear-gradient(90deg,rgba(212,175,55,.25),rgba(245,213,106,.08));border:1px solid rgba(245,213,106,.45);border-radius:18px;padding:14px 18px;color:#FFE89A;font-weight:900;margin-bottom:14px;}}
.kpi-label{{color:#CDBB78;font-size:12px;font-weight:900;text-transform:uppercase;letter-spacing:.5px;}}
.kpi-value{{color:#F5D56A;font-size:30px;font-weight:900;margin-top:3px;}}
.kpi-note{{color:#E7DCA8;font-size:13px;margin-top:5px;}}
.success-box{{background:rgba(34,197,94,.12);border:1px solid rgba(34,197,94,.35);color:#bbf7d0;border-radius:16px;padding:12px 14px;margin:12px 0;font-weight:700;}}
.warn-box{{background:rgba(245,158,11,.12);border:1px solid rgba(245,213,106,.35);color:#FFE89A;border-radius:16px;padding:12px 14px;margin:12px 0;font-weight:700;}}
.ticket-box{{background:#fff;color:#111;border-radius:14px;padding:18px;border:2px dashed #111;font-family:Arial, sans-serif;max-width:420px;margin:auto;}}
.ticket-box *{{color:#111!important;}}
.ticket-line{{border-top:1px dashed #111;margin:10px 0;}}
.stButton>button,.stDownloadButton>button{{background:linear-gradient(135deg,#B98A1E,#F5D56A)!important;color:#111!important;border:0!important;border-radius:14px!important;font-weight:900!important;padding:12px 18px!important;}}
.stTextInput input,.stNumberInput input,.stSelectbox div[data-baseweb="select"],.stTextArea textarea,.stDateInput input{{background:#111!important;border:1px solid rgba(245,213,106,.35)!important;color:#F8F1D7!important;border-radius:12px!important;}}
p,span,label,.stMarkdown,[data-testid="stMarkdownContainer"] p,[data-testid="stWidgetLabel"] p{{color:#E7DCA8!important;}}
[data-testid="stMetric"]{{background:linear-gradient(180deg,rgba(22,22,22,.95),rgba(8,8,8,.95));border:1px solid rgba(212,175,55,.22);border-radius:18px;padding:16px;}}
/* login */
.login-shell{{min-height:80vh;display:flex;align-items:center;justify-content:center;}}
.login-card{{width:min(620px,92vw);background:linear-gradient(180deg,rgba(10,10,10,.78),rgba(4,4,4,.94));border:1px solid rgba(245,213,106,.58);border-radius:34px;padding:18px 24px 22px;text-align:center;box-shadow:0 28px 95px rgba(0,0,0,.78),0 0 95px rgba(212,175,55,.18);}}
.food-hero{{width:100%;height:210px;object-fit:cover;object-position:center 42%;border-radius:26px;margin-bottom:14px;}}
.brand-word{{font-size:50px;font-weight:900;color:#F5D56A;letter-spacing:.8px;line-height:1;text-shadow:0 8px 28px rgba(245,213,106,.18);}}
.brand-sub{{color:#fff;font-size:16px;font-weight:900;letter-spacing:12px;margin-top:14px;}}
</style>
""", unsafe_allow_html=True)

# =========================
# UTILIDADES
# =========================
def money(x):
    try:
        return "$ {:,.0f}".format(float(x or 0)).replace(",", ".")
    except Exception:
        return "$ 0"

def num(x):
    try:
        if pd.isna(x): return 0.0
        if isinstance(x,(int,float)): return float(x)
        s=str(x).replace("$","").replace(".","").replace(",",".").strip()
        return float(s) if s else 0.0
    except Exception:
        return 0.0

def now_str(): return datetime.now().strftime("%d/%m/%Y %H:%M")
def today_str(): return datetime.now().strftime("%d/%m/%Y")

def parse_dt(s):
    for fmt in ("%d/%m/%Y %H:%M", "%d/%m/%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S"):
        try: return datetime.strptime(str(s), fmt)
        except Exception: pass
    return None

def df_query(sql, params=()):
    init_db()
    with sqlite3.connect(DB) as conn:
        return pd.read_sql_query(sql, conn, params=params)

def exec_sql(sql, params=()):
    init_db()
    with sqlite3.connect(DB) as conn:
        conn.execute(sql, params)
        conn.commit()

def hash_password(p): return hashlib.sha256(str(p).encode()).hexdigest()

# =========================
# DB
# =========================
def init_auth():
    with sqlite3.connect(AUTH_DB) as conn:
        cur=conn.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS auth(id INTEGER PRIMARY KEY CHECK(id=1), username TEXT, password_hash TEXT)")
        cur.execute("SELECT COUNT(*) FROM auth")
        if cur.fetchone()[0]==0:
            cur.execute("INSERT INTO auth(id,username,password_hash) VALUES(1,?,?)", ("admin", hash_password("admin123")))
        conn.commit()

def get_auth():
    init_auth()
    with sqlite3.connect(AUTH_DB) as conn:
        row=conn.execute("SELECT username,password_hash FROM auth WHERE id=1").fetchone()
    return row or ("admin", hash_password("admin123"))

def verify_login(u,p):
    su,sh=get_auth()
    return str(u).strip()==su and hash_password(p)==sh

def update_auth(u,p):
    init_auth()
    with sqlite3.connect(AUTH_DB) as conn:
        conn.execute("UPDATE auth SET username=?, password_hash=? WHERE id=1", (u.strip(), hash_password(p)))
        conn.commit()

def init_db():
    with sqlite3.connect(DB) as conn:
        c=conn.cursor()
        c.execute("""CREATE TABLE IF NOT EXISTS productos(
            id INTEGER PRIMARY KEY AUTOINCREMENT, nombre TEXT UNIQUE NOT NULL, categoria TEXT DEFAULT 'General',
            precio_unidad REAL DEFAULT 0, precio_kg REAL DEFAULT 0, costo_unidad REAL DEFAULT 0, costo_kg REAL DEFAULT 0,
            stock REAL DEFAULT 0, unidad_stock TEXT DEFAULT 'kg', activo INTEGER DEFAULT 1, actualizado TEXT DEFAULT CURRENT_TIMESTAMP)""")
        c.execute("""CREATE TABLE IF NOT EXISTS clientes(
            id INTEGER PRIMARY KEY AUTOINCREMENT, nombre TEXT UNIQUE NOT NULL, tipo TEXT DEFAULT '', zona TEXT DEFAULT '', telefono TEXT DEFAULT '', estado TEXT DEFAULT 'Activo', observaciones TEXT DEFAULT '', creado TEXT DEFAULT CURRENT_TIMESTAMP)""")
        c.execute("""CREATE TABLE IF NOT EXISTS cliente_pagos(
            id INTEGER PRIMARY KEY AUTOINCREMENT, fecha TEXT, cliente_id INTEGER, monto REAL, medio TEXT, observaciones TEXT)""")
        c.execute("""CREATE TABLE IF NOT EXISTS proveedores(
            id INTEGER PRIMARY KEY AUTOINCREMENT, nombre TEXT UNIQUE NOT NULL, telefono TEXT DEFAULT '', zona TEXT DEFAULT '', observaciones TEXT DEFAULT '')""")
        c.execute("""CREATE TABLE IF NOT EXISTS compras(
            id INTEGER PRIMARY KEY AUTOINCREMENT, fecha TEXT, proveedor_id INTEGER, producto TEXT, cantidad REAL, unidad TEXT, costo_total REAL, pagado REAL DEFAULT 0, detalle TEXT DEFAULT '')""")
        c.execute("""CREATE TABLE IF NOT EXISTS proveedor_pagos(
            id INTEGER PRIMARY KEY AUTOINCREMENT, fecha TEXT, proveedor_id INTEGER, monto REAL, medio TEXT, observaciones TEXT)""")
        c.execute("""CREATE TABLE IF NOT EXISTS ventas(
            id INTEGER PRIMARY KEY AUTOINCREMENT, ticket TEXT UNIQUE, fecha TEXT, cliente_id INTEGER, metodo_pago TEXT, total REAL DEFAULT 0, pagado REAL DEFAULT 0, observaciones TEXT DEFAULT '')""")
        c.execute("""CREATE TABLE IF NOT EXISTS venta_items(
            id INTEGER PRIMARY KEY AUTOINCREMENT, venta_id INTEGER, producto_id INTEGER, producto_nombre TEXT, modo TEXT, cantidad_texto TEXT, cantidad_base REAL DEFAULT 0, precio_unitario REAL DEFAULT 0, costo_unitario REAL DEFAULT 0, total REAL DEFAULT 0, costo_total REAL DEFAULT 0)""")
        c.execute("""CREATE TABLE IF NOT EXISTS caja(
            id INTEGER PRIMARY KEY AUTOINCREMENT, fecha TEXT, tipo TEXT, concepto TEXT, monto REAL, medio TEXT, observaciones TEXT)""")
        c.execute("""CREATE TABLE IF NOT EXISTS logistica(
            id INTEGER PRIMARY KEY AUTOINCREMENT, fecha TEXT, cliente_id INTEGER, cliente_manual TEXT, zona TEXT, direccion TEXT, detalle TEXT, estado TEXT DEFAULT 'Pendiente')""")
        conn.commit()
        # seed once
        c.execute("SELECT COUNT(*) FROM clientes")
        if c.fetchone()[0]==0:
            for n,t,z,tel in [("Pizzería La Esquina","Pizzería","Centro","11 2345-6789"),("Rotisería Avenida","Rotisería","Norte","11 3456-7890"),("Almacén Don Luis","Almacén","Oeste","11 4567-8901")]:
                c.execute("INSERT OR IGNORE INTO clientes(nombre,tipo,zona,telefono) VALUES(?,?,?,?)", (n,t,z,tel))
        c.execute("SELECT COUNT(*) FROM productos")
        if c.fetchone()[0]==0:
            import_excel_seed(c)
        conn.commit()

def import_excel_seed(cur):
    if DEFAULT_EXCEL.exists():
        try:
            raw=pd.read_excel(DEFAULT_EXCEL, header=None, engine="openpyxl")
            blocks=[(0,6,8),(9,15,17)]
            categoria="General"; count=0
            for name_col, unit_col, kg_col in blocks:
                categoria="General"
                for _,row in raw.iterrows():
                    name=row.get(name_col,None)
                    if pd.isna(name): continue
                    nombre=str(name).strip().title()
                    if not nombre or "whatsapp" in nombre.lower() or "lista" in nombre.lower(): continue
                    pu=num(row.get(unit_col,0)); pk=num(row.get(kg_col,0))
                    if pu==0 and pk==0 and len(nombre)<35:
                        categoria=nombre; continue
                    if pu==0 and pk==0: continue
                    cur.execute("INSERT OR IGNORE INTO productos(nombre,categoria,precio_unidad,precio_kg,costo_unidad,costo_kg,stock,unidad_stock,activo) VALUES(?,?,?,?,?,?,?,?,1)",
                                (nombre,categoria,pu,pk,0,0,0,"kg"))
                    count+=1
            if count: return
        except Exception:
            pass
    for p,cat,pu,pk in [("Mozzarella X10Kg","Mozzarellas",65000,6500),("Jamón Cocido","Fiambres",0,9000),("Harina 000 25Kg","Harinas",15000,600),("Aceitunas Verdes","Conservas",24000,4800)]:
        cur.execute("INSERT OR IGNORE INTO productos(nombre,categoria,precio_unidad,precio_kg,stock,unidad_stock,activo) VALUES(?,?,?,?,?,?,1)", (p,cat,pu,pk,0,"kg"))

init_db()

# =========================
# LOADERS
# =========================
def productos_df(active_only=False):
    where="WHERE activo=1" if active_only else ""
    df=df_query(f"SELECT * FROM productos {where} ORDER BY nombre")
    if not df.empty:
        df["Precio unidad $"]=df["precio_unidad"].apply(money); df["Precio kg $"]=df["precio_kg"].apply(money)
        df["Costo unidad $"]=df["costo_unidad"].apply(money); df["Costo kg $"]=df["costo_kg"].apply(money)
    return df

def clientes_df(): return df_query("SELECT * FROM clientes ORDER BY nombre")
def proveedores_df(): return df_query("SELECT * FROM proveedores ORDER BY nombre")
def caja_df(): return df_query("SELECT * FROM caja ORDER BY id DESC")
def compras_df(): return df_query("SELECT c.*, p.nombre proveedor FROM compras c LEFT JOIN proveedores p ON p.id=c.proveedor_id ORDER BY c.id DESC")
def ventas_df(): return df_query("SELECT v.*, c.nombre cliente FROM ventas v LEFT JOIN clientes c ON c.id=v.cliente_id ORDER BY v.id DESC")
def items_df(): return df_query("SELECT vi.*, v.fecha, v.ticket, c.nombre cliente FROM venta_items vi LEFT JOIN ventas v ON v.id=vi.venta_id LEFT JOIN clientes c ON c.id=v.cliente_id ORDER BY vi.id DESC")
def pagos_clientes_df(): return df_query("SELECT p.*, c.nombre cliente FROM cliente_pagos p LEFT JOIN clientes c ON c.id=p.cliente_id ORDER BY p.id DESC")
def pagos_proveedores_df(): return df_query("SELECT p.*, pr.nombre proveedor FROM proveedor_pagos p LEFT JOIN proveedores pr ON pr.id=p.proveedor_id ORDER BY p.id DESC")
def logistica_df(): return df_query("SELECT l.*, c.nombre cliente FROM logistica l LEFT JOIN clientes c ON c.id=l.cliente_id ORDER BY fecha DESC, id DESC")

# =========================
# COMPONENTES
# =========================
def banner(): st.markdown('<div class="demo-banner">✨ DON VALENTIN · Sistema comercial operativo · Stock · Caja · Clientes · Proveedores · Logística · Reportes históricos</div>', unsafe_allow_html=True)
def header(t,s): st.markdown(f'<div class="hero"><h1>{t}</h1><p>{s}</p></div>', unsafe_allow_html=True)
def kpi(l,v,n=""): st.markdown(f'<div class="card"><div class="kpi-label">{l}</div><div class="kpi-value">{v}</div><div class="kpi-note">{n}</div></div>', unsafe_allow_html=True)
def fig_style(fig,h=380):
    fig.update_layout(template="plotly_dark",height=h,paper_bgcolor="rgba(0,0,0,0)",plot_bgcolor="rgba(0,0,0,0)",font=dict(color="#F8F1D7"),title_font=dict(color="#F5D56A"),margin=dict(l=20,r=20,t=50,b=20))
    return fig

def id_label(df, col="nombre"):
    return {int(r["id"]): str(r[col]) for _,r in df.iterrows()}

# =========================
# LOGIN/SIDEBAR
# =========================
def login():
    img = f'<img class="food-hero" src="data:image/png;base64,{MOZZ_B64}">' if MOZZ_B64 else ""
    st.markdown(f'<div class="login-shell"><div class="login-card">{img}<div class="brand-word">DON VALENTIN</div><div class="brand-sub">DISTRIBUIDORA</div></div></div>', unsafe_allow_html=True)
    c1,c2,c3=st.columns([1,0.8,1])
    with c2:
        u=st.text_input("Usuario", placeholder="Usuario", label_visibility="collapsed")
        p=st.text_input("Contraseña", type="password", placeholder="Contraseña", label_visibility="collapsed")
        if st.button("🔒 Ingresar", use_container_width=True):
            if verify_login(u,p):
                st.session_state.logged=True; st.session_state.user=u; st.rerun()
            else: st.error("Acceso no autorizado.")

def sidebar():
    with st.sidebar:
        st.markdown("## 🧀 DON VALENTIN")
        st.markdown("**DISTRIBUIDORA**")
        st.markdown("---")
        pages=["Dashboard","Productos","Venta / Ticket","Clientes","Proveedores","Caja","Ganancias y deudas","Logística","Reportes","Configuración"]
        icons={"Dashboard":"📊","Productos":"📦","Venta / Ticket":"🧾","Clientes":"👥","Proveedores":"🏭","Caja":"💰","Ganancias y deudas":"📈","Logística":"🚚","Reportes":"📑","Configuración":"⚙️"}
        for p in pages:
            if st.button(f"{icons[p]} {p}", use_container_width=True): st.session_state.page=p; st.rerun()
        st.markdown("---")
        if st.button("Cerrar sesión", use_container_width=True): st.session_state.logged=False; st.rerun()

# =========================
# PÁGINAS
# =========================
def dashboard():
    banner(); header("DON VALENTIN", "Panel general con stock, ventas, caja, ganancias y alertas.")
    v=ventas_df(); it=items_df(); c=caja_df(); p=productos_df(True); cc=clientes_resumen()
    hoy=today_str()
    ventas_hoy=v[v["fecha"].str.startswith(hoy)] if not v.empty else pd.DataFrame()
    gan_hoy=profit_between(datetime.now().date(), datetime.now().date())
    c1,c2,c3,c4=st.columns(4)
    with c1: kpi("Venta del día", money(ventas_hoy["total"].sum() if not ventas_hoy.empty else 0), "Facturación diaria")
    with c2: kpi("Ganancia del día", money(gan_hoy), "Venta - costo")
    with c3: kpi("Productos activos", str(len(p)), "Catálogo editable")
    with c4: kpi("Clientes con deuda", str(len(cc[cc["Debe"]>0]) if not cc.empty else 0), "Cuenta corriente")
    if not p.empty:
        low=p[p["stock"]<=0]
        if not low.empty: st.warning(f"⚠️ Hay {len(low)} productos con stock cero o negativo.")
    if not cc.empty and not cc[cc["Alerta"].str.contains("⚠️")].empty:
        st.error("⚠️ Hay clientes con deuda de más de 7 días.")
    col1,col2=st.columns(2)
    with col1:
        if not p.empty: st.plotly_chart(fig_style(px.bar(p.head(15), x="nombre", y="stock", title="Stock real por producto")), use_container_width=True)
    with col2:
        if not it.empty:
            top=it.groupby("producto_nombre",as_index=False)["total"].sum().sort_values("total",ascending=False).head(10)
            st.plotly_chart(fig_style(px.bar(top,x="producto_nombre",y="total",title="Ventas por producto")), use_container_width=True)

def productos_page():
    banner(); header("Productos", "Agregar, editar, eliminar, cambiar nombre, precios, costos y stock real.")
    df=productos_df(False)
    with st.expander("➕ Agregar producto", expanded=False):
        c1,c2,c3=st.columns(3)
        with c1: nombre=st.text_input("Nombre producto nuevo"); categoria=st.text_input("Categoría", value="General")
        with c2: pu=st.number_input("Precio unidad",0.0,step=100.0); pk=st.number_input("Precio kg",0.0,step=100.0)
        with c3: cu=st.number_input("Costo unidad",0.0,step=100.0); ck=st.number_input("Costo kg",0.0,step=100.0); stock=st.number_input("Stock real inicial",0.0,step=1.0)
        unidad=st.selectbox("Unidad stock",["kg","unidad","bolsa","caja","bulto"])
        if st.button("Guardar producto", use_container_width=True):
            if nombre.strip():
                exec_sql("INSERT OR IGNORE INTO productos(nombre,categoria,precio_unidad,precio_kg,costo_unidad,costo_kg,stock,unidad_stock,activo,actualizado) VALUES(?,?,?,?,?,?,?,?,1,?)", (nombre.strip(),categoria,pu,pk,cu,ck,stock,unidad,now_str()))
                st.success("Producto guardado."); st.rerun()
    st.subheader("📋 Lista editable")
    if df.empty: st.info("No hay productos."); return
    show=df[["id","nombre","categoria","Precio unidad $","Precio kg $","Costo unidad $","Costo kg $","stock","unidad_stock","activo"]].copy()
    st.dataframe(show, use_container_width=True, hide_index=True, height=300)
    ids=id_label(df)
    pid=st.selectbox("Editar / eliminar producto", list(ids.keys()), format_func=lambda x: ids.get(x,str(x)))
    row=df[df.id==pid].iloc[0]
    c1,c2,c3=st.columns(3)
    with c1: en=st.text_input("Nombre", value=row.nombre); ec=st.text_input("Categoría", value=row.categoria)
    with c2: epu=st.number_input("Precio unidad editable", value=float(row.precio_unidad), step=100.0); epk=st.number_input("Precio kg editable", value=float(row.precio_kg), step=100.0)
    with c3: ecu=st.number_input("Costo unidad editable", value=float(row.costo_unidad), step=100.0); eck=st.number_input("Costo kg editable", value=float(row.costo_kg), step=100.0); estock=st.number_input("Stock real editable", value=float(row.stock), step=1.0)
    eun=st.selectbox("Unidad", ["kg","unidad","bolsa","caja","bulto"], index=["kg","unidad","bolsa","caja","bulto"].index(row.unidad_stock) if row.unidad_stock in ["kg","unidad","bolsa","caja","bulto"] else 0)
    active=st.checkbox("Producto activo", value=bool(row.activo))
    c1,c2=st.columns(2)
    if c1.button("💾 Guardar cambios producto", use_container_width=True):
        exec_sql("UPDATE productos SET nombre=?,categoria=?,precio_unidad=?,precio_kg=?,costo_unidad=?,costo_kg=?,stock=?,unidad_stock=?,activo=?,actualizado=? WHERE id=?", (en,ec,epu,epk,ecu,eck,estock,eun,1 if active else 0,now_str(),int(pid)))
        st.success("Producto actualizado."); st.rerun()
    if c2.button("🗑️ Eliminar producto", use_container_width=True):
        exec_sql("DELETE FROM productos WHERE id=?", (int(pid),)); st.success("Producto eliminado."); st.rerun()

def venta_page():
    banner(); header("Venta / Ticket", "Ticket con varios productos, gramos exactos y descuento automático de stock.")
    if "cart" not in st.session_state: st.session_state.cart=[]
    prods=productos_df(True); cl=clientes_df()
    if prods.empty or cl.empty: st.warning("Cargá productos y clientes primero."); return
    labels_c=id_label(cl); labels_p=id_label(prods)
    c1,c2=st.columns([1,1])
    with c1:
        cid=st.selectbox("Cliente", list(labels_c.keys()), format_func=lambda x: labels_c.get(x,str(x)))
        metodo=st.selectbox("Método de pago",["Efectivo","Transferencia","Mercado Pago","Cuenta corriente"])
        pagado=0.0
        if metodo=="Cuenta corriente": pagado=st.number_input("Monto pagado ahora (parcial)",0.0,step=100.0)
        pid=st.selectbox("Producto", list(labels_p.keys()), format_func=lambda x: labels_p.get(x,str(x)))
        pr=prods[prods.id==pid].iloc[0]
        modo=st.radio("Modo",["Por gramos","Por unidad"],horizontal=True)
        if modo=="Por gramos":
            gramos=st.number_input("Peso exacto en gramos", min_value=1.0, max_value=50000.0, value=100.0, step=1.0)
            cantidad_base=gramos/1000; cantidad_txt=f"{gramos:g} g" if gramos<1000 else f"{gramos/1000:g} kg"; precio_unit=float(pr.precio_kg or 0); costo_unit=float(pr.costo_kg or 0)
            total=precio_unit*cantidad_base; costo=costo_unit*cantidad_base
        else:
            unidades=st.number_input("Unidades", min_value=1.0, value=1.0, step=1.0)
            cantidad_base=unidades; cantidad_txt=f"{unidades:g} u."; precio_unit=float(pr.precio_unidad or 0); costo_unit=float(pr.costo_unidad or 0)
            total=precio_unit*unidades; costo=costo_unit*unidades
        kpi("Subtotal", money(total), f"{pr.nombre} · {cantidad_txt}")
        if st.button("➕ Agregar al ticket", use_container_width=True):
            st.session_state.cart.append({"producto_id":int(pid),"producto_nombre":pr.nombre,"modo":modo,"cantidad_texto":cantidad_txt,"cantidad_base":float(cantidad_base),"precio_unitario":precio_unit,"costo_unitario":costo_unit,"total":round(total,2),"costo_total":round(costo,2)})
            st.rerun()
    with c2:
        st.subheader("🧾 Ticket actual")
        if st.session_state.cart:
            cart=pd.DataFrame(st.session_state.cart); cart["Total $"]=cart.total.apply(money)
            st.dataframe(cart[["producto_nombre","cantidad_texto","Total $"]],use_container_width=True,hide_index=True)
            total=sum(i["total"] for i in st.session_state.cart)
            st.metric("Total",money(total))
            a,b=st.columns(2)
            if a.button("Vaciar", use_container_width=True): st.session_state.cart=[]; st.rerun()
            if b.button("Confirmar venta", use_container_width=True):
                ticket="T-"+datetime.now().strftime("%Y%m%d%H%M%S")
                paid=total if metodo!="Cuenta corriente" else pagado
                with sqlite3.connect(DB) as conn:
                    cur=conn.cursor(); cur.execute("INSERT INTO ventas(ticket,fecha,cliente_id,metodo_pago,total,pagado) VALUES(?,?,?,?,?,?)", (ticket,now_str(),int(cid),metodo,total,paid)); vid=cur.lastrowid
                    for item in st.session_state.cart:
                        cur.execute("INSERT INTO venta_items(venta_id,producto_id,producto_nombre,modo,cantidad_texto,cantidad_base,precio_unitario,costo_unitario,total,costo_total) VALUES(?,?,?,?,?,?,?,?,?,?)", (vid,item["producto_id"],item["producto_nombre"],item["modo"],item["cantidad_texto"],item["cantidad_base"],item["precio_unitario"],item["costo_unitario"],item["total"],item["costo_total"]))
                        cur.execute("UPDATE productos SET stock=stock-?, actualizado=? WHERE id=?", (item["cantidad_base"],now_str(),item["producto_id"]))
                    if paid>0: cur.execute("INSERT INTO cliente_pagos(fecha,cliente_id,monto,medio,observaciones) VALUES(?,?,?,?,?)", (now_str(),int(cid),paid,metodo,f"Pago ticket {ticket}"))
                    if paid>0: cur.execute("INSERT INTO caja(fecha,tipo,concepto,monto,medio,observaciones) VALUES(?,?,?,?,?,?)", (now_str(),"Ingreso",f"Cobro ticket {ticket}",paid,metodo,labels_c.get(cid,"")))
                    conn.commit()
                st.session_state.last_ticket=ticket; st.session_state.cart=[]; st.success("Venta guardada y stock descontado."); st.rerun()
        else: st.info("Agregá productos al ticket.")
    render_last_ticket()

def render_last_ticket():
    ticket=st.session_state.get("last_ticket")
    if not ticket: return
    v=ventas_df(); venta=v[v.ticket==ticket]
    if venta.empty: return
    row=venta.iloc[0]; items=items_df(); its=items[items.ticket==ticket]
    html_items="".join([f"<div>{r.producto_nombre}<span style='float:right'>{money(r.total)}</span><br><small>{r.cantidad_texto}</small></div><div class='ticket-line'></div>" for _,r in its.iterrows()])
    html=f"""<div class='ticket-box'><h2 style='text-align:center'>DON VALENTIN</h2><h3 style='text-align:center'>Ticket</h3><div class='ticket-line'></div><b>N°:</b> {ticket}<br><b>Fecha:</b> {row.fecha}<br><b>Cliente:</b> {row.cliente}<br><b>Pago:</b> {row.metodo_pago}<div class='ticket-line'></div>{html_items}<div style='font-size:22px;font-weight:900;text-align:right'>TOTAL {money(row.total)}</div></div>"""
    st.markdown(html, unsafe_allow_html=True)
    full=f"""<html><head><meta charset='utf-8'><style>body{{font-family:Arial;padding:20px}}.ticket{{width:320px;margin:auto;border:1px dashed #111;padding:16px}}.line{{border-top:1px dashed #111;margin:10px 0}}@media print{{button{{display:none}}}}</style></head><body><button onclick='window.print()'>Imprimir</button><div class='ticket'><h2 style='text-align:center'>DON VALENTIN</h2><h3 style='text-align:center'>Ticket</h3><div class='line'></div><b>N°:</b> {ticket}<br><b>Fecha:</b> {row.fecha}<br><b>Cliente:</b> {row.cliente}<br><b>Pago:</b> {row.metodo_pago}<div class='line'></div>{'<div class="line"></div>'.join([f'{r.producto_nombre}<span style="float:right">{money(r.total)}</span><br><small>{r.cantidad_texto}</small>' for _,r in its.iterrows()])}<div class='line'></div><h2 style='text-align:right'>TOTAL {money(row.total)}</h2></div></body></html>"""
    st.download_button("⬇️ Descargar ticket para imprimir", data=full.encode("utf-8"), file_name=f"ticket_{ticket}.html", mime="text/html", use_container_width=True)

def clientes_resumen():
    cl=clientes_df(); v=ventas_df(); pagos=pagos_clientes_df(); rows=[]
    for _,c in cl.iterrows():
        vv=v[v.cliente_id==c.id] if not v.empty else pd.DataFrame(); pp=pagos[pagos.cliente_id==c.id] if not pagos.empty else pd.DataFrame()
        fact=float(vv.total.sum()) if not vv.empty else 0; pag=float(pp.monto.sum()) if not pp.empty else 0; debe=max(fact-pag,0)
        debt_dates=[]
        if not vv.empty:
            for _,r in vv[vv.pagado < vv.total].iterrows():
                dt=parse_dt(r.fecha)
                if dt: debt_dates.append(dt)
        dias=max([(datetime.now()-d).days for d in debt_dates], default=0)
        rows.append({"id":c.id,"Cliente":c.nombre,"Facturado":fact,"Pagado":pag,"Debe":debe,"Facturado $":money(fact),"Pagado $":money(pag),"Debe $":money(debe),"Días deuda":dias,"Alerta":"⚠️ Más de 7 días sin pagar" if debe>0 and dias>=7 else "OK"})
    return pd.DataFrame(rows)

def clientes_page():
    banner(); header("Clientes", "Alta, edición, eliminación, pagos parciales, deuda y alertas.")
    with st.expander("➕ Agregar cliente", expanded=False):
        c1,c2,c3=st.columns(3)
        with c1: n=st.text_input("Nombre cliente"); tel=st.text_input("Teléfono")
        with c2: tipo=st.text_input("Tipo", value="Cliente"); zona=st.text_input("Zona")
        with c3: estado=st.selectbox("Estado",["Activo","Inactivo","Cuenta corriente"]); obs=st.text_area("Observaciones")
        if st.button("Guardar cliente", use_container_width=True):
            if n.strip(): exec_sql("INSERT OR IGNORE INTO clientes(nombre,tipo,zona,telefono,estado,observaciones) VALUES(?,?,?,?,?,?)", (n,tipo,zona,tel,estado,obs)); st.success("Cliente guardado."); st.rerun()
    cl=clientes_df(); res=clientes_resumen()
    if not res.empty: st.dataframe(res[["Cliente","Facturado $","Pagado $","Debe $","Días deuda","Alerta"]],use_container_width=True,hide_index=True)
    if cl.empty: return
    labels=id_label(cl)
    cid=st.selectbox("Editar / eliminar cliente", list(labels.keys()), format_func=lambda x: labels.get(x,str(x)))
    row=cl[cl.id==cid].iloc[0]
    c1,c2,c3=st.columns(3)
    with c1: en=st.text_input("Nombre",value=row.nombre); etel=st.text_input("Teléfono",value=row.telefono)
    with c2: etipo=st.text_input("Tipo",value=row.tipo); ezona=st.text_input("Zona",value=row.zona)
    with c3: eest=st.selectbox("Estado cliente",["Activo","Inactivo","Cuenta corriente"], index=["Activo","Inactivo","Cuenta corriente"].index(row.estado) if row.estado in ["Activo","Inactivo","Cuenta corriente"] else 0); eobs=st.text_area("Observaciones cliente", value=row.observaciones)
    a,b=st.columns(2)
    if a.button("💾 Guardar cambios cliente", use_container_width=True): exec_sql("UPDATE clientes SET nombre=?,tipo=?,zona=?,telefono=?,estado=?,observaciones=? WHERE id=?", (en,etipo,ezona,etel,eest,eobs,int(cid))); st.success("Cliente actualizado."); st.rerun()
    if b.button("🗑️ Eliminar cliente", use_container_width=True): exec_sql("DELETE FROM clientes WHERE id=?", (int(cid),)); st.success("Cliente eliminado."); st.rerun()
    st.subheader("💵 Registrar pago parcial")
    c1,c2,c3=st.columns(3)
    with c1: monto=st.number_input("Monto pago",0.0,step=100.0)
    with c2: medio=st.selectbox("Medio pago",["Efectivo","Transferencia","Mercado Pago","Cheque","Otro"])
    with c3: obsp=st.text_input("Observación pago")
    if st.button("Guardar pago cliente", use_container_width=True):
        exec_sql("INSERT INTO cliente_pagos(fecha,cliente_id,monto,medio,observaciones) VALUES(?,?,?,?,?)", (now_str(),int(cid),monto,medio,obsp)); exec_sql("INSERT INTO caja(fecha,tipo,concepto,monto,medio,observaciones) VALUES(?,?,?,?,?,?)", (now_str(),"Ingreso",f"Pago cliente {labels[cid]}",monto,medio,obsp)); st.success("Pago guardado."); st.rerun()
    pagos=pagos_clientes_df()
    if not pagos.empty:
        st.subheader("Pagos cargados")
        st.dataframe(pagos[["id","fecha","cliente","monto","medio","observaciones"]],use_container_width=True,hide_index=True)
        did=st.number_input("ID de pago a eliminar",0,step=1)
        if st.button("Eliminar pago cliente", use_container_width=True) and did>0: exec_sql("DELETE FROM cliente_pagos WHERE id=?",(int(did),)); st.rerun()

def proveedores_page():
    banner(); header("Proveedores", "Agregar, editar, eliminar proveedores, compras manuales, pagos y deudas.")
    with st.expander("➕ Agregar proveedor", expanded=False):
        n=st.text_input("Nombre proveedor"); tel=st.text_input("Teléfono proveedor"); zona=st.text_input("Zona proveedor"); obs=st.text_area("Observaciones proveedor")
        if st.button("Guardar proveedor") and n.strip(): exec_sql("INSERT OR IGNORE INTO proveedores(nombre,telefono,zona,observaciones) VALUES(?,?,?,?)", (n,tel,zona,obs)); st.rerun()
    pr=proveedores_df(); labels=id_label(pr) if not pr.empty else {}
    if not pr.empty: st.dataframe(pr,use_container_width=True,hide_index=True)
    if labels:
        pid=st.selectbox("Proveedor",list(labels.keys()),format_func=lambda x: labels.get(x,str(x)))
        row=pr[pr.id==pid].iloc[0]
        c1,c2=st.columns(2)
        with c1: en=st.text_input("Editar nombre proveedor",value=row.nombre); etel=st.text_input("Editar teléfono",value=row.telefono)
        with c2: ez=st.text_input("Editar zona",value=row.zona); eo=st.text_area("Editar observaciones",value=row.observaciones)
        a,b=st.columns(2)
        if a.button("Guardar proveedor",use_container_width=True): exec_sql("UPDATE proveedores SET nombre=?,telefono=?,zona=?,observaciones=? WHERE id=?",(en,etel,ez,eo,int(pid))); st.rerun()
        if b.button("Eliminar proveedor",use_container_width=True): exec_sql("DELETE FROM proveedores WHERE id=?",(int(pid),)); st.rerun()
        st.subheader("🧾 Cargar compra manual")
        c1,c2,c3=st.columns(3)
        with c1: prod=st.text_input("Producto comprado manual", placeholder="Ej: Harina 000 25kg")
        with c2: cant=st.number_input("Cantidad comprada",0.0,step=1.0); unidad=st.selectbox("Unidad compra",["kg","unidad","bolsa","caja","bulto"])
        with c3: costo=st.number_input("Costo total",0.0,step=100.0); pag=st.number_input("Pagado al proveedor",0.0,step=100.0)
        det=st.text_input("Detalle compra")
        if st.button("Guardar compra proveedor",use_container_width=True):
            exec_sql("INSERT INTO compras(fecha,proveedor_id,producto,cantidad,unidad,costo_total,pagado,detalle) VALUES(?,?,?,?,?,?,?,?)",(now_str(),int(pid),prod,cant,unidad,costo,pag,det))
            if pag>0: exec_sql("INSERT INTO proveedor_pagos(fecha,proveedor_id,monto,medio,observaciones) VALUES(?,?,?,?,?)",(now_str(),int(pid),pag,"Transferencia",f"Pago compra {prod}")); exec_sql("INSERT INTO caja(fecha,tipo,concepto,monto,medio,observaciones) VALUES(?,?,?,?,?,?)",(now_str(),"Egreso",f"Pago proveedor {labels[pid]}",pag,"Transferencia",prod))
            st.success("Compra guardada."); st.rerun()
        compras=compras_df(); pagos=pagos_proveedores_df()
        if not compras.empty: st.dataframe(compras,use_container_width=True,hide_index=True)
        delc=st.number_input("ID compra a eliminar",0,step=1)
        if st.button("Eliminar compra",use_container_width=True) and delc>0: exec_sql("DELETE FROM compras WHERE id=?",(int(delc),)); st.rerun()

def caja_page():
    banner(); header("Caja", "Ingresos y egresos editables y eliminables.")
    c1,c2,c3=st.columns(3)
    with c1: tipo=st.selectbox("Tipo movimiento",["Ingreso","Egreso"]); monto=st.number_input("Monto",0.0,step=100.0)
    with c2: concepto=st.text_input("Concepto"); medio=st.selectbox("Medio",["Efectivo","Transferencia","Mercado Pago","Cheque","Otro"])
    with c3: obs=st.text_area("Observaciones caja")
    if st.button("Guardar movimiento caja",use_container_width=True): exec_sql("INSERT INTO caja(fecha,tipo,concepto,monto,medio,observaciones) VALUES(?,?,?,?,?,?)",(now_str(),tipo,concepto,monto,medio,obs)); st.rerun()
    caja=caja_df()
    if not caja.empty:
        caja["Monto $"]=caja.monto.apply(money); st.dataframe(caja,use_container_width=True,hide_index=True)
        cid=st.number_input("ID movimiento a editar/eliminar",0,step=1)
        if cid>0:
            row=caja[caja.id==cid]
            if not row.empty:
                r=row.iloc[0]; c1,c2=st.columns(2)
                with c1: em=st.number_input("Nuevo monto",value=float(r.monto),step=100.0); ec=st.text_input("Nuevo concepto",value=r.concepto)
                with c2: et=st.selectbox("Nuevo tipo",["Ingreso","Egreso"],index=0 if r.tipo=="Ingreso" else 1); eo=st.text_input("Nueva observación",value=r.observaciones)
                a,b=st.columns(2)
                if a.button("Guardar edición caja",use_container_width=True): exec_sql("UPDATE caja SET tipo=?,concepto=?,monto=?,observaciones=? WHERE id=?",(et,ec,em,eo,int(cid))); st.rerun()
                if b.button("Eliminar movimiento caja",use_container_width=True): exec_sql("DELETE FROM caja WHERE id=?",(int(cid),)); st.rerun()

def profit_between(start,end):
    it=items_df()
    if it.empty: return 0.0
    total=0.0
    for _,r in it.iterrows():
        dt=parse_dt(r.fecha)
        if dt and start<=dt.date()<=end: total += float(r.total or 0)-float(r.costo_total or 0)
    return total

def ganancias_page():
    banner(); header("Ganancias y deudas", "Ganancias por período y relación con deudas de clientes/proveedores.")
    hoy=datetime.now().date()
    periods={"Día":(hoy,hoy),"Semana":(hoy-timedelta(days=7),hoy),"Mes":(hoy-timedelta(days=30),hoy),"Año":(hoy-timedelta(days=365),hoy)}
    cols=st.columns(4)
    for col,(name,(s,e)) in zip(cols,periods.items()):
        with col: st.metric(f"Ganancia {name.lower()}", money(profit_between(s,e)))
    res=clientes_resumen()
    if not res.empty: st.dataframe(res[["Cliente","Facturado $","Pagado $","Debe $","Días deuda","Alerta"]],use_container_width=True,hide_index=True)
    its=items_df()
    if not its.empty:
        g=its.groupby("producto_nombre",as_index=False).agg(Vendido=("total","sum"),Costo=("costo_total","sum")); g["Ganancia"]=g.Vendido-g.Costo; g["Ganancia $"]=g.Ganancia.apply(money)
        st.subheader("Ganancia por producto")
        st.dataframe(g,use_container_width=True,hide_index=True)
        st.plotly_chart(fig_style(px.bar(g,x="producto_nombre",y="Ganancia",title="Ganancia histórica por producto")),use_container_width=True)

def logistica_page():
    banner(); header("Logística", "Organizar recorridos, entregas, visitas diarias y agenda semanal.")
    cl=clientes_df(); labels=id_label(cl) if not cl.empty else {}
    with st.expander("➕ Agendar visita / entrega", expanded=False):
        c1,c2,c3=st.columns(3)
        with c1: fecha=st.date_input("Fecha", value=datetime.now().date()); cid=st.selectbox("Cliente", [0]+list(labels.keys()), format_func=lambda x: "Cliente manual" if x==0 else labels.get(x,str(x)))
        with c2: manual=st.text_input("Cliente manual"); zona=st.text_input("Zona")
        with c3: direccion=st.text_input("Dirección"); estado=st.selectbox("Estado",["Pendiente","Visitado","No visitado","Reprogramar","Entregado"])
        detalle=st.text_area("Detalle / pedido / recorrido")
        if st.button("Guardar visita",use_container_width=True): exec_sql("INSERT INTO logistica(fecha,cliente_id,cliente_manual,zona,direccion,detalle,estado) VALUES(?,?,?,?,?,?,?)",(fecha.strftime("%d/%m/%Y"),None if cid==0 else int(cid),manual,zona,direccion,detalle,estado)); st.rerun()
    l=logistica_df();
    if not l.empty:
        l["Cliente final"]=l.apply(lambda r: r.cliente if pd.notna(r.cliente) and r.cliente else r.cliente_manual,axis=1)
        st.dataframe(l[["id","fecha","Cliente final","zona","direccion","detalle","estado"]],use_container_width=True,hide_index=True)
        lid=st.number_input("ID visita para editar/eliminar",0,step=1)
        if lid>0:
            row=l[l.id==lid]
            if not row.empty:
                r=row.iloc[0]; c1,c2=st.columns(2)
                with c1: ef=st.text_input("Fecha visita",value=r.fecha); ez=st.text_input("Zona visita",value=r.zona); ed=st.text_input("Dirección visita",value=r.direccion)
                with c2: ee=st.selectbox("Estado visita",["Pendiente","Visitado","No visitado","Reprogramar","Entregado"], index=["Pendiente","Visitado","No visitado","Reprogramar","Entregado"].index(r.estado) if r.estado in ["Pendiente","Visitado","No visitado","Reprogramar","Entregado"] else 0); edet=st.text_area("Detalle visita",value=r.detalle)
                a,b=st.columns(2)
                if a.button("Guardar edición logística",use_container_width=True): exec_sql("UPDATE logistica SET fecha=?,zona=?,direccion=?,detalle=?,estado=? WHERE id=?",(ef,ez,ed,edet,ee,int(lid))); st.rerun()
                if b.button("Eliminar visita",use_container_width=True): exec_sql("DELETE FROM logistica WHERE id=?",(int(lid),)); st.rerun()
    else: st.info("No hay visitas agendadas.")

def reportes_page():
    banner(); header("Reportes", "Estadísticas históricas: ventas, caja, ganancias, stock y deudas.")
    v=ventas_df(); c=caja_df(); its=items_df(); p=productos_df(False)
    hoy=datetime.now().date()
    c1,c2,c3,c4=st.columns(4)
    with c1: st.metric("Venta día", money(v[v.fecha.apply(lambda x: (parse_dt(x) or datetime.min).date()==hoy)].total.sum() if not v.empty else 0))
    with c2: st.metric("Ganancia día", money(profit_between(hoy,hoy)))
    with c3: st.metric("Ingresos caja día", money(c[(c.tipo=="Ingreso") & (c.fecha.apply(lambda x: (parse_dt(x) or datetime.min).date()==hoy))].monto.sum() if not c.empty else 0))
    with c4: st.metric("Egresos caja día", money(c[(c.tipo=="Egreso") & (c.fecha.apply(lambda x: (parse_dt(x) or datetime.min).date()==hoy))].monto.sum() if not c.empty else 0))
    if not v.empty:
        tmp=v.copy(); tmp["Día"]=tmp.fecha.apply(lambda x: (parse_dt(x) or datetime.now()).strftime("%d/%m/%Y")); d=tmp.groupby("Día",as_index=False).agg(Facturación=("total","sum"),Tickets=("ticket","count")); d["Facturación $"]=d.Facturación.apply(money)
        st.subheader("Facturación diaria histórica"); st.dataframe(d,use_container_width=True,hide_index=True); st.plotly_chart(fig_style(px.bar(d,x="Día",y="Facturación",title="Facturación diaria")),use_container_width=True)
    if not its.empty:
        st.subheader("Histórico de ventas por producto"); st.dataframe(its,use_container_width=True,hide_index=True)
    if not p.empty:
        st.subheader("Stock real actual"); st.dataframe(p[["nombre","categoria","stock","unidad_stock","Precio unidad $","Precio kg $","Costo unidad $","Costo kg $"]],use_container_width=True,hide_index=True)

def config_page():
    banner(); header("Configuración", "Cambiar usuario y contraseña.")
    user,_=get_auth()
    nu=st.text_input("Nuevo usuario",value=user); ca=st.text_input("Contraseña actual",type="password"); nc=st.text_input("Nueva contraseña",type="password"); cf=st.text_input("Confirmar contraseña",type="password")
    if st.button("Guardar usuario y contraseña",use_container_width=True):
        if not verify_login(user,ca): st.error("Contraseña actual incorrecta.")
        elif len(nc)<4: st.warning("La contraseña debe tener al menos 4 caracteres.")
        elif nc!=cf: st.error("Las contraseñas no coinciden.")
        else: update_auth(nu,nc); st.success("Usuario y contraseña actualizados.")

# =========================
# ROUTER
# =========================
if "logged" not in st.session_state: st.session_state.logged=False
if "page" not in st.session_state: st.session_state.page="Dashboard"

if not st.session_state.logged:
    login()
else:
    sidebar()
    pages={"Dashboard":dashboard,"Productos":productos_page,"Venta / Ticket":venta_page,"Clientes":clientes_page,"Proveedores":proveedores_page,"Caja":caja_page,"Ganancias y deudas":ganancias_page,"Logística":logistica_page,"Reportes":reportes_page,"Configuración":config_page}
    pages[st.session_state.page]()
