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

def log_event(modulo, accion, detalle=""):
    try:
        with sqlite3.connect(DB) as conn:
            conn.execute("CREATE TABLE IF NOT EXISTS actividad(id INTEGER PRIMARY KEY AUTOINCREMENT, fecha TEXT, modulo TEXT, accion TEXT, detalle TEXT)")
            conn.execute("INSERT INTO actividad(fecha,modulo,accion,detalle) VALUES(?,?,?,?)", (now_str(), str(modulo), str(accion), str(detalle)))
            conn.commit()
    except Exception:
        pass

def actividad_df():
    return df_query("SELECT * FROM actividad ORDER BY id DESC")

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
            id INTEGER PRIMARY KEY AUTOINCREMENT, nombre TEXT UNIQUE NOT NULL, tipo TEXT DEFAULT '', zona TEXT DEFAULT '', telefono TEXT DEFAULT '', estado TEXT DEFAULT 'Activo', observaciones TEXT DEFAULT '', facturado_inicial REAL DEFAULT 0, pagado_inicial REAL DEFAULT 0, debe_inicial REAL DEFAULT 0, dias_deuda_inicial INTEGER DEFAULT 0, creado TEXT DEFAULT CURRENT_TIMESTAMP)""")
        
        # Compatibilidad: si la base ya existía, agregamos campos manuales de cuenta corriente
        for col, typ in [("facturado_inicial","REAL DEFAULT 0"),("pagado_inicial","REAL DEFAULT 0"),("debe_inicial","REAL DEFAULT 0"),("dias_deuda_inicial","INTEGER DEFAULT 0")]:
            try:
                c.execute(f"ALTER TABLE clientes ADD COLUMN {col} {typ}")
            except Exception:
                pass
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
        c.execute("""CREATE TABLE IF NOT EXISTS actividad(
            id INTEGER PRIMARY KEY AUTOINCREMENT, fecha TEXT, modulo TEXT, accion TEXT, detalle TEXT)""")
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
    return {str(int(r["id"])): str(r[col]) for _,r in df.iterrows()}

def selected_int(value):
    try:
        return int(value)
    except Exception:
        return 0

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
        unidad=st.selectbox("Unidad stock",["kg","litro","unidad","bolsa","caja","bulto"])
        if st.button("Guardar producto", use_container_width=True):
            if nombre.strip():
                exec_sql("INSERT OR IGNORE INTO productos(nombre,categoria,precio_unidad,precio_kg,costo_unidad,costo_kg,stock,unidad_stock,activo,actualizado) VALUES(?,?,?,?,?,?,?,?,1,?)", (nombre.strip(),categoria,pu,pk,cu,ck,stock,unidad,now_str()))
                log_event("Productos", "Alta producto", f"{nombre.strip()} · Stock {stock} {unidad} · Venta unidad {money(pu)} · Venta kg {money(pk)}")
                st.success("Producto guardado."); st.rerun()
    st.subheader("📋 Lista editable")
    if df.empty: st.info("No hay productos."); return
    show=df[["id","nombre","categoria","Precio unidad $","Precio kg $","Costo unidad $","Costo kg $","stock","unidad_stock","activo"]].copy()
    st.dataframe(show, use_container_width=True, hide_index=True, height=300)
    ids=id_label(df)
    pid=st.selectbox("Editar / eliminar producto", list(ids.keys()), format_func=lambda x: ids.get(str(x),str(x)))
    pid_int=selected_int(pid)
    row=df[df.id==pid_int].iloc[0]
    c1,c2,c3=st.columns(3)
    with c1: en=st.text_input("Nombre", value=row.nombre); ec=st.text_input("Categoría", value=row.categoria)
    with c2: epu=st.number_input("Precio unidad editable", value=float(row.precio_unidad), step=100.0); epk=st.number_input("Precio kg editable", value=float(row.precio_kg), step=100.0)
    with c3: ecu=st.number_input("Costo unidad editable", value=float(row.costo_unidad), step=100.0); eck=st.number_input("Costo kg editable", value=float(row.costo_kg), step=100.0); estock=st.number_input("Stock real editable", value=float(row.stock), step=1.0)
    eun=st.selectbox("Unidad", ["kg","litro","unidad","bolsa","caja","bulto"], index=["kg","litro","unidad","bolsa","caja","bulto"].index(row.unidad_stock) if row.unidad_stock in ["kg","litro","unidad","bolsa","caja","bulto"] else 0)
    active=st.checkbox("Producto activo", value=bool(row.activo))
    c1,c2=st.columns(2)
    if c1.button("💾 Guardar cambios producto", use_container_width=True):
        exec_sql("UPDATE productos SET nombre=?,categoria=?,precio_unidad=?,precio_kg=?,costo_unidad=?,costo_kg=?,stock=?,unidad_stock=?,activo=?,actualizado=? WHERE id=?", (en,ec,epu,epk,ecu,eck,estock,eun,1 if active else 0,now_str(),int(pid_int)))
        log_event("Productos", "Edición producto", f"{en} · Stock {estock} {eun} · Venta unidad {money(epu)} · Venta kg {money(epk)}")
        st.success("Producto actualizado."); st.rerun()
    if c2.button("🗑️ Eliminar producto", use_container_width=True):
        exec_sql("DELETE FROM productos WHERE id=?", (int(pid_int),)); log_event("Productos", "Eliminación producto", str(row.nombre)); st.success("Producto eliminado."); st.rerun()

def venta_page():
    banner(); header("Venta / Ticket", "Ticket editable con varios productos, gramos/litros exactos, precios automáticos editables y descuento automático de stock.")
    if "cart" not in st.session_state:
        st.session_state.cart=[]
    prods=productos_df(True); cl=clientes_df()
    if prods.empty or cl.empty:
        st.warning("Cargá productos y clientes primero."); return
    labels_c=id_label(cl); labels_p=id_label(prods)
    c1,c2=st.columns([1,1])
    with c1:
        cid_str=st.selectbox("Cliente", list(labels_c.keys()), format_func=lambda x: labels_c.get(str(x),str(x)))
        cid=selected_int(cid_str)
        metodo=st.selectbox("Método de pago",["Efectivo","Transferencia","Mercado Pago","Cuenta corriente"])
        pagado=0.0
        if metodo=="Cuenta corriente":
            pagado=st.number_input("Monto pagado ahora (parcial)",0.0,step=100.0)
        pid_str=st.selectbox("Producto", list(labels_p.keys()), format_func=lambda x: labels_p.get(str(x),str(x)))
        pid=selected_int(pid_str)
        pr=prods[prods.id==pid].iloc[0]
        modo=st.radio("Modo",["Por gramos","Por litros","Por unidad"],horizontal=True)
        if modo=="Por gramos":
            gramos=st.number_input("Peso exacto en gramos", min_value=1.0, max_value=50000.0, value=100.0, step=1.0, key=f"gramos_{pid}")
            cantidad_base=gramos/1000
            cantidad_txt=f"{gramos:g} g" if gramos<1000 else f"{gramos/1000:g} kg"
            # Precio automático: primero precio por kg; si no existe, usa precio unidad para evitar $0.
            precio_unit_auto=float(pr.precio_kg or pr.precio_unidad or 0)
            costo_unit=float(pr.costo_kg or pr.costo_unidad or 0)
            precio_label="Precio por kg"
        elif modo=="Por litros":
            litros=st.number_input("Litros exactos", min_value=0.01, value=1.0, step=0.01, format="%.2f", key=f"litros_{pid}")
            cantidad_base=litros
            cantidad_txt=f"{litros:g} lts"
            # Precio automático para litros: primero precio unidad; si no existe, usa precio kg.
            precio_unit_auto=float(pr.precio_unidad or pr.precio_kg or 0)
            costo_unit=float(pr.costo_unidad or pr.costo_kg or 0)
            precio_label="Precio por litro"
        else:
            unidades=st.number_input("Unidades", min_value=1.0, value=1.0, step=1.0, key=f"unidades_{pid}")
            cantidad_base=unidades
            cantidad_txt=f"{unidades:g} u."
            precio_unit_auto=float(pr.precio_unidad or pr.precio_kg or 0)
            costo_unit=float(pr.costo_unidad or pr.costo_kg or 0)
            precio_label="Precio por unidad"
        st.markdown("#### ✍️ Edición manual antes de agregar")
        st.caption("El sistema calcula el total automáticamente con el producto, cliente y precio elegido. Si hace falta, podés corregir producto, precio o total manualmente antes de agregarlo al ticket.")
        nombre_ticket=st.text_input("Nombre del producto en ticket", value=str(pr.nombre), key=f"nombre_ticket_{pid}_{modo}")
        precio_key=f"precio_ticket_{pid}_{modo}_{float(cantidad_base):.3f}"
        precio_unit=st.number_input(precio_label, value=float(precio_unit_auto), step=100.0, key=precio_key)
        total_auto=precio_unit*cantidad_base
        total_key=f"total_ticket_{pid}_{modo}_{float(cantidad_base):.3f}_{float(precio_unit):.2f}"
        total_manual=st.number_input("Total del producto", value=float(round(total_auto,2)), step=100.0, key=total_key)
        costo=costo_unit*cantidad_base
        kpi("Subtotal automático / editable", money(total_manual), f"{nombre_ticket} · {cantidad_txt}")
        if st.button("➕ Agregar al ticket", use_container_width=True):
            st.session_state.cart.append({"producto_id":int(pid),"producto_nombre":nombre_ticket,"modo":modo,"cantidad_texto":cantidad_txt,"cantidad_base":float(cantidad_base),"precio_unitario":float(precio_unit),"costo_unitario":float(costo_unit),"total":round(float(total_manual),2),"costo_total":round(float(costo),2)})
            st.rerun()
    with c2:
        st.subheader("🧾 Ticket actual")
        if st.session_state.cart:
            cart=pd.DataFrame(st.session_state.cart); cart["Total $"]=cart.total.apply(money)
            st.dataframe(cart[["producto_nombre","cantidad_texto","precio_unitario","Total $"]],use_container_width=True,hide_index=True)
            total=sum(float(i["total"]) for i in st.session_state.cart)
            st.metric("Total",money(total))
            with st.expander("✏️ Editar / eliminar producto del ticket", expanded=False):
                idx=st.selectbox("Producto del ticket", list(range(len(st.session_state.cart))), format_func=lambda i: f"{i+1}. {st.session_state.cart[i]['producto_nombre']} - {money(st.session_state.cart[i]['total'])}")
                item=st.session_state.cart[idx]
                en=st.text_input("Editar nombre", value=item["producto_nombre"])
                ecant=st.text_input("Editar cantidad texto", value=item["cantidad_texto"])
                ebase=st.number_input("Cantidad base para descontar stock", value=float(item["cantidad_base"]), step=0.001, format="%.3f")
                epu=st.number_input("Editar precio unitario", value=float(item["precio_unitario"]), step=100.0)
                etot=st.number_input("Editar total", value=float(item["total"]), step=100.0)
                a1,a2=st.columns(2)
                if a1.button("Guardar edición del item", use_container_width=True):
                    st.session_state.cart[idx].update({"producto_nombre":en,"cantidad_texto":ecant,"cantidad_base":float(ebase),"precio_unitario":float(epu),"total":float(etot)})
                    st.success("Item actualizado."); st.rerun()
                if a2.button("Eliminar item", use_container_width=True):
                    st.session_state.cart.pop(idx); st.success("Item eliminado."); st.rerun()
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
                    if paid>0:
                        cur.execute("INSERT INTO cliente_pagos(fecha,cliente_id,monto,medio,observaciones) VALUES(?,?,?,?,?)", (now_str(),int(cid),paid,metodo,f"Pago ticket {ticket}"))
                        cur.execute("INSERT INTO caja(fecha,tipo,concepto,monto,medio,observaciones) VALUES(?,?,?,?,?,?)", (now_str(),"Ingreso",f"Cobro ticket {ticket}",paid,metodo,labels_c.get(str(cid),"")))
                    conn.commit()
                log_event("Ventas", "Ticket generado", f"{ticket} · Cliente {labels_c.get(str(cid),'')} · Total {money(total)} · Stock descontado")
                st.session_state.last_ticket=ticket; st.session_state.cart=[]; st.success("Venta guardada y stock descontado."); st.rerun()
        else:
            st.info("Agregá productos al ticket.")
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
        fact=float(vv.total.sum()) if not vv.empty else 0
        pag=float(pp.monto.sum()) if not pp.empty else 0
        fact_ini=float(getattr(c,"facturado_inicial",0) or 0)
        pag_ini=float(getattr(c,"pagado_inicial",0) or 0)
        debe_ini=float(getattr(c,"debe_inicial",0) or 0)
        fact += fact_ini
        pag += pag_ini
        debe=max((fact-pag)+debe_ini,0)
        debt_dates=[]
        if not vv.empty:
            for _,r in vv[vv.pagado < vv.total].iterrows():
                dt=parse_dt(r.fecha)
                if dt: debt_dates.append(dt)
        dias=max([(datetime.now()-d).days for d in debt_dates], default=0)
        dias=max(dias, int(getattr(c,"dias_deuda_inicial",0) or 0))
        rows.append({"id":c.id,"Cliente":c.nombre,"Facturado":fact,"Pagado":pag,"Debe":debe,"Facturado $":money(fact),"Pagado $":money(pag),"Debe $":money(debe),"Días deuda":dias,"Alerta":"⚠️ Más de 7 días sin pagar" if debe>0 and dias>=7 else "OK"})
    return pd.DataFrame(rows)

def clientes_page():
    banner(); header("Clientes", "Alta, edición, eliminación, pagos parciales, deuda y alertas.")
    with st.expander("➕ Agregar cliente", expanded=False):
        c1,c2,c3=st.columns(3)
        with c1:
            n=st.text_input("Nombre cliente")
            tel=st.text_input("Teléfono")
        with c2:
            tipo=st.text_input("Tipo", value="Cliente")
            zona=st.text_input("Zona")
        with c3:
            estado=st.selectbox("Estado",["Activo","Inactivo","Cuenta corriente"])
            obs=st.text_area("Observaciones")
        st.markdown("### 💳 Saldo inicial del cliente")
        s1,s2,s3,s4=st.columns(4)
        with s1: fact_ini=st.number_input("Facturado inicial",0.0,step=100.0,key="alta_fact_ini")
        with s2: pag_ini=st.number_input("Pagado inicial",0.0,step=100.0,key="alta_pag_ini")
        with s3: debe_ini=st.number_input("Debe inicial",0.0,step=100.0,key="alta_debe_ini")
        with s4: dias_ini=st.number_input("Días deuda inicial",0,step=1,key="alta_dias_ini")
        if st.button("Guardar cliente", use_container_width=True):
            if n.strip():
                exec_sql("INSERT OR IGNORE INTO clientes(nombre,tipo,zona,telefono,estado,observaciones,facturado_inicial,pagado_inicial,debe_inicial,dias_deuda_inicial) VALUES(?,?,?,?,?,?,?,?,?,?)", (n,tipo,zona,tel,estado,obs,fact_ini,pag_ini,debe_ini,int(dias_ini)))
                log_event("Clientes", "Alta cliente", f"{n} · Facturado inicial {money(fact_ini)} · Pagado inicial {money(pag_ini)} · Debe inicial {money(debe_ini)}")
                st.success("Cliente guardado."); st.rerun()
    cl=clientes_df(); res=clientes_resumen()
    if not res.empty: st.dataframe(res[["Cliente","Facturado $","Pagado $","Debe $","Días deuda","Alerta"]],use_container_width=True,hide_index=True)
    if cl.empty: return
    labels=id_label(cl)
    cid_str=st.selectbox("Editar / eliminar cliente", list(labels.keys()), format_func=lambda x: labels.get(str(x),str(x)))
    cid=selected_int(cid_str)
    row=cl[cl.id==cid].iloc[0]
    c1,c2,c3=st.columns(3)
    with c1: en=st.text_input("Nombre",value=row.nombre); etel=st.text_input("Teléfono",value=row.telefono)
    with c2: etipo=st.text_input("Tipo",value=row.tipo); ezona=st.text_input("Zona",value=row.zona)
    with c3: eest=st.selectbox("Estado cliente",["Activo","Inactivo","Cuenta corriente"], index=["Activo","Inactivo","Cuenta corriente"].index(row.estado) if row.estado in ["Activo","Inactivo","Cuenta corriente"] else 0); eobs=st.text_area("Observaciones cliente", value=row.observaciones)
    st.markdown("### ✏️ Editar saldo inicial / cuenta corriente")
    sc1,sc2,sc3,sc4=st.columns(4)
    with sc1: efact=st.number_input("Editar facturado",0.0,value=float(getattr(row,"facturado_inicial",0) or 0),step=100.0,key=f"efact_{cid}")
    with sc2: epag=st.number_input("Editar pagado",0.0,value=float(getattr(row,"pagado_inicial",0) or 0),step=100.0,key=f"epag_{cid}")
    with sc3: edebe=st.number_input("Editar debe",0.0,value=float(getattr(row,"debe_inicial",0) or 0),step=100.0,key=f"edebe_{cid}")
    with sc4: edias=st.number_input("Editar días deuda",0,value=int(getattr(row,"dias_deuda_inicial",0) or 0),step=1,key=f"edias_{cid}")
    a,b=st.columns(2)
    if a.button("💾 Guardar cambios cliente", use_container_width=True):
        exec_sql("UPDATE clientes SET nombre=?,tipo=?,zona=?,telefono=?,estado=?,observaciones=?,facturado_inicial=?,pagado_inicial=?,debe_inicial=?,dias_deuda_inicial=? WHERE id=?", (en,etipo,ezona,etel,eest,eobs,efact,epag,edebe,int(edias),int(cid)))
        log_event("Clientes", "Edición cliente", en)
        st.success("Cliente actualizado."); st.rerun()
    if b.button("🗑️ Eliminar cliente", use_container_width=True):
        exec_sql("DELETE FROM clientes WHERE id=?", (int(cid),))
        log_event("Clientes", "Eliminación cliente", str(row.nombre))
        st.success("Cliente eliminado."); st.rerun()
    st.subheader("💵 Registrar pago parcial")
    pago_cid_str=st.selectbox("Cliente que realiza el pago", list(labels.keys()), format_func=lambda x: labels.get(str(x),str(x)), key="pago_cliente_select")
    pago_cid=selected_int(pago_cid_str)
    c1,c2,c3=st.columns(3)
    with c1: monto=st.number_input("Monto pago",0.0,step=100.0)
    with c2: medio=st.selectbox("Medio pago",["Efectivo","Transferencia","Mercado Pago","Cheque","Otro"])
    with c3: obsp=st.text_input("Observación pago")
    if st.button("Guardar pago cliente", use_container_width=True):
        exec_sql("INSERT INTO cliente_pagos(fecha,cliente_id,monto,medio,observaciones) VALUES(?,?,?,?,?)", (now_str(),int(pago_cid),monto,medio,obsp)); exec_sql("INSERT INTO caja(fecha,tipo,concepto,monto,medio,observaciones) VALUES(?,?,?,?,?,?)", (now_str(),"Ingreso",f"Pago cliente {labels.get(str(pago_cid),'')}",monto,medio,obsp)); log_event("Clientes", "Pago parcial", f"{labels.get(str(pago_cid),'')} · {money(monto)}"); st.success("Pago guardado."); st.rerun()
    pagos=pagos_clientes_df()
    if not pagos.empty:
        st.subheader("Pagos cargados")
        st.dataframe(pagos[["id","fecha","cliente","monto","medio","observaciones"]],use_container_width=True,hide_index=True)
        did=st.number_input("ID de pago a eliminar",0,step=1)
        if st.button("Eliminar pago cliente", use_container_width=True) and did>0:
            exec_sql("DELETE FROM cliente_pagos WHERE id=?",(int(did),)); log_event("Clientes", "Eliminación pago cliente", f"Pago ID {did}"); st.rerun()

def proveedores_page():
    banner(); header("Proveedores", "Agregar, editar, eliminar proveedores, compras manuales, pagos y deudas.")
    with st.expander("➕ Agregar proveedor", expanded=False):
        n=st.text_input("Nombre proveedor"); tel=st.text_input("Teléfono proveedor"); zona=st.text_input("Zona proveedor"); obs=st.text_area("Observaciones proveedor")
        if st.button("Guardar proveedor") and n.strip():
            exec_sql("INSERT OR IGNORE INTO proveedores(nombre,telefono,zona,observaciones) VALUES(?,?,?,?)", (n,tel,zona,obs))
            log_event("Proveedores", "Alta proveedor", n)
            st.success("Proveedor guardado."); st.rerun()
    pr=proveedores_df(); labels=id_label(pr) if not pr.empty else {}
    if not pr.empty: st.dataframe(pr,use_container_width=True,hide_index=True)
    if labels:
        pid_str=st.selectbox("Proveedor para editar/eliminar", list(labels.keys()), format_func=lambda x: labels.get(str(x),str(x)))
        pid=selected_int(pid_str)
        row=pr[pr.id==pid]
        if not row.empty:
            row=row.iloc[0]
            c1,c2=st.columns(2)
            with c1: en=st.text_input("Editar nombre proveedor",value=str(row.nombre)); etel=st.text_input("Editar teléfono",value=str(row.telefono))
            with c2: ez=st.text_input("Editar zona",value=str(row.zona)); eo=st.text_area("Editar observaciones",value=str(row.observaciones))
            a,b=st.columns(2)
            if a.button("Guardar proveedor",use_container_width=True):
                exec_sql("UPDATE proveedores SET nombre=?,telefono=?,zona=?,observaciones=? WHERE id=?",(en,etel,ez,eo,int(pid)))
                log_event("Proveedores", "Edición proveedor", en)
                st.success("Proveedor actualizado."); st.rerun()
            if b.button("Eliminar proveedor",use_container_width=True):
                exec_sql("DELETE FROM proveedores WHERE id=?",(int(pid),))
                log_event("Proveedores", "Eliminación proveedor", str(row.nombre))
                st.success("Proveedor eliminado."); st.rerun()

        st.subheader("💸 Registrar pago al proveedor seleccionado")
        st.caption("Seleccionás el proveedor arriba y acá cargás un pago manual: por bolsa, litro, unidad, caja o el total directo. También queda reflejado como egreso de caja.")
        p1,p2,p3=st.columns(3)
        with p1:
            pago_prod=st.text_input("Producto / concepto del pago", placeholder="Ej: Harina 000 25kg", key=f"pago_prod_{pid}")
            pago_unidad=st.selectbox("Unidad del pago", ["bolsa","litro","unidad","caja","kg","bulto","otro"], key=f"pago_unidad_{pid}")
        with p2:
            pago_cant=st.number_input("Cantidad pagada", min_value=0.0, step=1.0, key=f"pago_cant_{pid}")
            pago_unit=st.number_input("Precio/costo por unidad", min_value=0.0, step=100.0, key=f"pago_unit_{pid}")
        with p3:
            pago_auto=float(pago_cant)*float(pago_unit)
            pago_monto=st.number_input("Total pagado", min_value=0.0, value=float(pago_auto), step=100.0, key=f"pago_monto_{pid}_{pago_auto}")
            pago_medio=st.selectbox("Medio de pago proveedor", ["Efectivo","Transferencia","Mercado Pago","Cheque","Otro"], key=f"pago_medio_{pid}")
        pago_obs=st.text_input("Observación del pago", placeholder="Ej: pago parcial / cancelación / anticipo", key=f"pago_obs_{pid}")
        if st.button("Guardar pago al proveedor seleccionado", use_container_width=True, key=f"guardar_pago_proveedor_{pid}"):
            if pago_monto>0:
                detalle_pago=f"{pago_prod} · {pago_cant:g} {pago_unidad} · {pago_obs}".strip()
                exec_sql("INSERT INTO proveedor_pagos(fecha,proveedor_id,monto,medio,observaciones) VALUES(?,?,?,?,?)", (now_str(),int(pid),float(pago_monto),pago_medio,detalle_pago))
                exec_sql("INSERT INTO caja(fecha,tipo,concepto,monto,medio,observaciones) VALUES(?,?,?,?,?,?)", (now_str(),"Egreso",f"Pago proveedor {labels.get(str(pid),'')}",float(pago_monto),pago_medio,detalle_pago))
                log_event("Proveedores", "Pago proveedor", f"{labels.get(str(pid),'')} · {money(pago_monto)} · {detalle_pago}")
                log_event("Caja", "Egreso proveedor", f"{labels.get(str(pid),'')} · {money(pago_monto)}")
                st.success("Pago guardado y egreso registrado en caja."); st.rerun()
            else:
                st.warning("Ingresá un total pagado mayor a cero.")
        st.subheader("🧾 Cargar compra manual")
        c1,c2,c3=st.columns(3)
        with c1: prod=st.text_input("Producto comprado manual", placeholder="Ej: Harina 000 25kg")
        with c2: cant=st.number_input("Cantidad comprada",0.0,step=1.0); unidad=st.selectbox("Unidad compra",["kg","litro","unidad","bolsa","caja","bulto"])
        with c3: costo=st.number_input("Costo total",0.0,step=100.0); pag=st.number_input("Pagado al proveedor",0.0,step=100.0)
        det=st.text_input("Detalle compra")
        if st.button("Guardar compra proveedor",use_container_width=True):
            exec_sql("INSERT INTO compras(fecha,proveedor_id,producto,cantidad,unidad,costo_total,pagado,detalle) VALUES(?,?,?,?,?,?,?,?)",(now_str(),int(pid),prod,cant,unidad,costo,pag,det))
            log_event("Proveedores", "Compra proveedor", f"{labels.get(str(pid),'')} · {prod} · {cant} {unidad} · {money(costo)}")
            if pag>0:
                exec_sql("INSERT INTO proveedor_pagos(fecha,proveedor_id,monto,medio,observaciones) VALUES(?,?,?,?,?)",(now_str(),int(pid),pag,"Transferencia",f"Pago compra {prod}"))
                exec_sql("INSERT INTO caja(fecha,tipo,concepto,monto,medio,observaciones) VALUES(?,?,?,?,?,?)",(now_str(),"Egreso",f"Pago proveedor {labels.get(str(pid),'')}",pag,"Transferencia",prod))
                log_event("Caja", "Egreso proveedor", f"{labels.get(str(pid),'')} · {money(pag)}")
            st.success("Compra guardada."); st.rerun()
    compras=compras_df(); pagos=pagos_proveedores_df()
    if not compras.empty:
        st.subheader("Compras cargadas")
        st.dataframe(compras,use_container_width=True,hide_index=True)
        opts=["0"]+[str(int(x)) for x in compras.id.tolist()]
        delc_str=st.selectbox("Compra para eliminar", opts, format_func=lambda x: "Elegir compra" if x=="0" else f"ID {x}")
        delc=selected_int(delc_str)
        if st.button("Eliminar compra",use_container_width=True) and delc>0:
            exec_sql("DELETE FROM compras WHERE id=?",(int(delc),)); log_event("Proveedores", "Eliminación compra", f"Compra ID {delc}"); st.success("Compra eliminada."); st.rerun()
    if not pagos.empty:
        st.subheader("Pagos a proveedores")
        st.dataframe(pagos,use_container_width=True,hide_index=True)
        opts=["0"]+[str(int(x)) for x in pagos.id.tolist()]
        dp_str=st.selectbox("Pago proveedor para eliminar", opts, format_func=lambda x: "Elegir pago" if x=="0" else f"ID {x}")
        dp=selected_int(dp_str)
        if st.button("Eliminar pago proveedor",use_container_width=True) and dp>0:
            exec_sql("DELETE FROM proveedor_pagos WHERE id=?",(int(dp),)); log_event("Proveedores", "Eliminación pago proveedor", f"Pago ID {dp}"); st.success("Pago eliminado."); st.rerun()

def caja_page():
    banner(); header("Caja", "Ingresos y egresos editables y eliminables.")
    c1,c2,c3=st.columns(3)
    with c1: tipo=st.selectbox("Tipo movimiento",["Ingreso","Egreso"]); monto=st.number_input("Monto",0.0,step=100.0)
    with c2: concepto=st.text_input("Concepto"); medio=st.selectbox("Medio",["Efectivo","Transferencia","Mercado Pago","Cheque","Otro"])
    with c3: obs=st.text_area("Observaciones caja")
    if st.button("Guardar movimiento caja",use_container_width=True):
        exec_sql("INSERT INTO caja(fecha,tipo,concepto,monto,medio,observaciones) VALUES(?,?,?,?,?,?)",(now_str(),tipo,concepto,monto,medio,obs))
        log_event("Caja", f"Movimiento {tipo}", f"{concepto} · {money(monto)} · {medio}")
        st.success("Movimiento guardado."); st.rerun()
    caja=caja_df()
    if not caja.empty:
        caja["Monto $"]=caja.monto.apply(money); st.dataframe(caja,use_container_width=True,hide_index=True)
        opts=["0"]+[str(int(x)) for x in caja.id.tolist()]
        cid_str=st.selectbox("Movimiento de caja para editar/eliminar", opts, format_func=lambda x: "Elegir movimiento" if x=="0" else f"ID {x}")
        cid=selected_int(cid_str)
        if cid>0:
            row=caja[caja.id==cid]
            if not row.empty:
                r=row.iloc[0]; c1,c2=st.columns(2)
                with c1: em=st.number_input("Nuevo monto",value=float(r.monto),step=100.0); ec=st.text_input("Nuevo concepto",value=str(r.concepto))
                with c2: et=st.selectbox("Nuevo tipo",["Ingreso","Egreso"],index=0 if r.tipo=="Ingreso" else 1); emedio=st.selectbox("Nuevo medio",["Efectivo","Transferencia","Mercado Pago","Cheque","Otro"], index=["Efectivo","Transferencia","Mercado Pago","Cheque","Otro"].index(r.medio) if r.medio in ["Efectivo","Transferencia","Mercado Pago","Cheque","Otro"] else 0); eo=st.text_input("Nueva observación",value=str(r.observaciones))
                a,b=st.columns(2)
                if a.button("Guardar edición caja",use_container_width=True):
                    exec_sql("UPDATE caja SET tipo=?,concepto=?,monto=?,medio=?,observaciones=? WHERE id=?",(et,ec,em,emedio,eo,int(cid)))
                    log_event("Caja", "Edición movimiento", f"ID {cid} · {et} · {money(em)}")
                    st.success("Movimiento actualizado."); st.rerun()
                if b.button("Eliminar movimiento caja",use_container_width=True):
                    exec_sql("DELETE FROM caja WHERE id=?",(int(cid),)); log_event("Caja", "Eliminación movimiento", f"ID {cid}"); st.success("Movimiento eliminado."); st.rerun()

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
    banner(); header("Logística", "Organizar recorridos, entregas, visitas diarias y agenda semanal con cliente guardado o manual.")
    cl=clientes_df(); labels=id_label(cl) if not cl.empty else {}
    with st.expander("➕ Agendar visita / entrega", expanded=True):
        c1,c2,c3=st.columns(3)
        with c1:
            fecha=st.date_input("Fecha programada", value=datetime.now().date())
            usar_cliente=st.radio("Cliente", ["Cliente guardado","Cliente manual"], horizontal=True)
        with c2:
            if usar_cliente=="Cliente guardado" and labels:
                cid_str=st.selectbox("Seleccionar cliente guardado", list(labels.keys()), format_func=lambda x: labels.get(str(x),str(x)), key="log_cliente_guardado")
                cid=selected_int(cid_str); manual=""
            else:
                cid=None; manual=st.text_input("Cliente manual")
            zona=st.text_input("Zona")
        with c3:
            direccion=st.text_input("Dirección")
            estado=st.selectbox("Estado",["Pendiente","Visitado","No visitado","Reprogramar","Entregado"])
        detalle=st.text_area("Detalle / pedido / recorrido")
        if estado=="Reprogramar":
            nueva_fecha=st.date_input("Nueva fecha reprogramada", value=datetime.now().date()+timedelta(days=1))
            detalle = (detalle or "") + f" | Reprogramado para {nueva_fecha.strftime('%d/%m/%Y')}"
        if st.button("Guardar visita",use_container_width=True):
            exec_sql("INSERT INTO logistica(fecha,cliente_id,cliente_manual,zona,direccion,detalle,estado) VALUES(?,?,?,?,?,?,?)",(fecha.strftime("%d/%m/%Y"),None if not cid else int(cid),manual,zona,direccion,detalle,estado)); log_event("Logística", "Alta visita/entrega", f"{fecha.strftime('%d/%m/%Y')} · {manual or (labels.get(str(cid),'') if cid else '')} · {estado}"); st.success("Visita guardada."); st.rerun()
    l=logistica_df()
    if not l.empty:
        l["Cliente final"]=l.apply(lambda r: r.cliente if pd.notna(r.cliente) and r.cliente else r.cliente_manual,axis=1)
        st.dataframe(l[["id","fecha","Cliente final","zona","direccion","detalle","estado"]],use_container_width=True,hide_index=True)
        ids=[str(int(x)) for x in l.id.tolist()]
        lid_str=st.selectbox("Visita para editar/eliminar", ["0"]+ids, format_func=lambda x: "Elegir visita" if str(x)=="0" else f"ID {x}")
        lid=selected_int(lid_str)
        if lid>0:
            row=l[l.id==lid]
            if not row.empty:
                r=row.iloc[0]
                c1,c2,c3=st.columns(3)
                with c1:
                    try: default_date=parse_dt(r.fecha).date() if parse_dt(r.fecha) else datetime.now().date()
                    except Exception: default_date=datetime.now().date()
                    ef=st.date_input("Fecha visita", value=default_date, key=f"fecha_edit_{lid}")
                    ecliente_manual=st.text_input("Cliente manual", value="" if pd.isna(r.cliente_manual) else str(r.cliente_manual))
                with c2:
                    ez=st.text_input("Zona visita",value="" if pd.isna(r.zona) else str(r.zona))
                    ed=st.text_input("Dirección visita",value="" if pd.isna(r.direccion) else str(r.direccion))
                with c3:
                    ee=st.selectbox("Estado visita",["Pendiente","Visitado","No visitado","Reprogramar","Entregado"], index=["Pendiente","Visitado","No visitado","Reprogramar","Entregado"].index(r.estado) if r.estado in ["Pendiente","Visitado","No visitado","Reprogramar","Entregado"] else 0)
                edet=st.text_area("Detalle visita",value="" if pd.isna(r.detalle) else str(r.detalle))
                a,b=st.columns(2)
                if a.button("Guardar edición logística",use_container_width=True):
                    exec_sql("UPDATE logistica SET fecha=?,cliente_manual=?,zona=?,direccion=?,detalle=?,estado=? WHERE id=?",(ef.strftime("%d/%m/%Y"),ecliente_manual,ez,ed,edet,ee,int(lid)))
                    log_event("Logística", "Edición visita/entrega", f"ID {lid} · {ef.strftime('%d/%m/%Y')} · {ee}")
                    st.success("Logística actualizada."); st.rerun()
                if b.button("Eliminar visita",use_container_width=True):
                    exec_sql("DELETE FROM logistica WHERE id=?",(int(lid),)); log_event("Logística", "Eliminación visita", f"ID {lid}"); st.success("Visita eliminada."); st.rerun()
    else: st.info("No hay visitas agendadas.")

def reportes_page():
    banner(); header("Reportes", "Estadísticas históricas y movimiento diario completo.")
    v=ventas_df(); c=caja_df(); its=items_df(); p=productos_df(False); act=actividad_df(); log=logistica_df(); pagos=pagos_clientes_df(); compras=compras_df()
    hoy=datetime.now().date()
    c1,c2,c3,c4=st.columns(4)
    venta_dia = v[v.fecha.apply(lambda x: (parse_dt(x) or datetime.min).date()==hoy)].total.sum() if not v.empty else 0
    ingreso_dia = c[(c.tipo=="Ingreso") & (c.fecha.apply(lambda x: (parse_dt(x) or datetime.min).date()==hoy))].monto.sum() if not c.empty else 0
    egreso_dia = c[(c.tipo=="Egreso") & (c.fecha.apply(lambda x: (parse_dt(x) or datetime.min).date()==hoy))].monto.sum() if not c.empty else 0
    with c1: st.metric("Venta día", money(venta_dia))
    with c2: st.metric("Ganancia día", money(profit_between(hoy,hoy)))
    with c3: st.metric("Ingresos caja día", money(ingreso_dia))
    with c4: st.metric("Egresos caja día", money(egreso_dia))
    st.subheader("📌 Movimiento completo del día")
    movimientos=[]
    if not act.empty:
        for _,r in act.iterrows():
            dt=parse_dt(r.fecha)
            if dt and dt.date()==hoy:
                movimientos.append({"Fecha":r.fecha,"Módulo":r.modulo,"Movimiento":r.accion,"Detalle":r.detalle})
    if not v.empty:
        for _,r in v.iterrows():
            dt=parse_dt(r.fecha)
            if dt and dt.date()==hoy:
                movimientos.append({"Fecha":r.fecha,"Módulo":"Ventas","Movimiento":"Ticket","Detalle":f"{r.ticket} · {r.cliente} · {money(r.total)}"})
    if not c.empty:
        for _,r in c.iterrows():
            dt=parse_dt(r.fecha)
            if dt and dt.date()==hoy:
                movimientos.append({"Fecha":r.fecha,"Módulo":"Caja","Movimiento":r.tipo,"Detalle":f"{r.concepto} · {money(r.monto)} · {r.medio}"})
    if not log.empty:
        log["Cliente final"]=log.apply(lambda r: r.cliente if pd.notna(r.cliente) and r.cliente else r.cliente_manual,axis=1)
        for _,r in log.iterrows():
            dt=parse_dt(r.fecha)
            if dt and dt.date()==hoy:
                movimientos.append({"Fecha":r.fecha,"Módulo":"Logística","Movimiento":r.estado,"Detalle":f"{r['Cliente final']} · {r.zona} · {r.direccion} · {r.detalle}"})
    mov_df=pd.DataFrame(movimientos).sort_values("Fecha", ascending=False) if movimientos else pd.DataFrame(columns=["Fecha","Módulo","Movimiento","Detalle"])
    st.dataframe(mov_df,use_container_width=True,hide_index=True)
    st.download_button("⬇️ Descargar reporte diario CSV", data=mov_df.to_csv(index=False).encode("utf-8-sig"), file_name=f"reporte_diario_{datetime.now().strftime('%Y%m%d')}.csv", mime="text/csv", use_container_width=True)
    if not v.empty:
        tmp=v.copy(); tmp["Día"]=tmp.fecha.apply(lambda x: (parse_dt(x) or datetime.now()).strftime("%d/%m/%Y")); d=tmp.groupby("Día",as_index=False).agg(Facturación=("total","sum"),Tickets=("ticket","count")); d["Facturación $"]=d.Facturación.apply(money)
        st.subheader("Facturación diaria histórica"); st.dataframe(d,use_container_width=True,hide_index=True); st.plotly_chart(fig_style(px.bar(d,x="Día",y="Facturación",title="Facturación diaria")),use_container_width=True)
    if not its.empty:
        st.subheader("Histórico de ventas por producto"); st.dataframe(its,use_container_width=True,hide_index=True)
    if not p.empty:
        st.subheader("Stock real actual"); st.dataframe(p[["nombre","categoria","stock","unidad_stock","Precio unidad $","Precio kg $","Costo unidad $","Costo kg $"]],use_container_width=True,hide_index=True)
    if not compras.empty:
        st.subheader("Compras históricas a proveedores"); st.dataframe(compras,use_container_width=True,hide_index=True)

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
