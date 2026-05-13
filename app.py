
import streamlit as st
import pandas as pd
import plotly.express as px
import random

st.set_page_config(page_title="DEL CAMPO DISTRIBUIDORA", page_icon="🏢", layout="wide", initial_sidebar_state="expanded")

APP_NAME = "DEL CAMPO DISTRIBUIDORA"
DEMO_USER = "demo"
DEMO_PASS = "demo123"
WHATSAPP_LINK = "https://wa.me/TUNUMERO?text=Hola,%20quiero%20solicitar%20acceso%20completo%20a%20la%20demo%20de%20DEL%20CAMPO%20DISTRIBUIDORA"

random.seed(7)
productos = pd.DataFrame({
    "Código": ["P-1001","P-1002","P-1003","P-1004","P-1005","P-1006","P-1007","P-1008"],
    "Producto": ["Harina 000 x 25kg","Aceite girasol 900ml","Azúcar 1kg","Yerba mate 1kg","Arroz largo fino 1kg","Fideos 500g","Leche 1L","Gaseosa 2.25L"],
    "Categoría": ["Harinas","Almacén","Almacén","Infusiones","Almacén","Pastas","Lácteos","Bebidas"],
    "Proveedor": ["Molino del Sur","Aceitera Norte","Dulce Campo","Yerbas del Litoral","Granos SRL","Pastas Centro","Lácteos Norte","Bebidas Premium"],
    "Stock": [380,155,90,52,430,220,71,39],
    "Stock mínimo": [120,80,100,60,150,100,80,60],
    "Precio venta": [18500,1950,1350,4600,1450,1050,1250,2850],
    "Estado": ["OK","OK","Bajo","Bajo","OK","OK","Bajo","Crítico"]
})
clientes = pd.DataFrame({
    "Cliente": ["Autoservicio Sol","Kiosco Central","Despensa Norte","Mayorista Río","Almacén Don Luis","Supermercado Avenida","Distribuidora Oeste"],
    "Zona": ["CABA","GBA Sur","GBA Norte","CABA","GBA Oeste","GBA Norte","GBA Oeste"],
    "Tipo": ["Comercio","Minorista","Comercio","Mayorista","Comercio","Mayorista","Revendedor"],
    "Estado": ["Activo","Activo","Inactivo","Activo","Activo","Activo","Activo"],
    "Deuda": [120000,45000,0,280000,90000,160000,75000],
    "Última compra": ["2026-05-10","2026-05-09","2026-03-18","2026-05-08","2026-05-07","2026-05-11","2026-05-06"]
})
ventas_mes = pd.DataFrame({"Día": list(range(1,31)), "Ventas": [random.randint(180000,620000) for _ in range(30)], "Ganancia": [random.randint(45000,210000) for _ in range(30)]})
ventas_categoria = pd.DataFrame({"Categoría": ["Almacén","Harinas","Bebidas","Lácteos","Pastas","Infusiones"], "Importe": [4200000,3600000,2100000,1450000,980000,760000]})
pedidos = pd.DataFrame({
    "Pedido": ["PED-2401","PED-2402","PED-2403","PED-2404","PED-2405","PED-2406"],
    "Cliente": ["Autoservicio Sol","Kiosco Central","Mayorista Río","Almacén Don Luis","Supermercado Avenida","Distribuidora Oeste"],
    "Estado": ["Pendiente","En preparación","En reparto","Entregado","Pendiente","En reparto"],
    "Importe": [250000,98000,430000,125000,345000,287000],
    "Zona": ["CABA","GBA Sur","CABA","GBA Oeste","GBA Norte","GBA Oeste"]
})
rutas = pd.DataFrame({"Ruta":["Ruta CABA","Ruta GBA Norte","Ruta GBA Sur","Ruta Oeste"],"Chofer":["Martín Gómez","Lucas Pérez","Diego Suárez","Sergio Díaz"],"Vehículo":["Furgón 01","Camión 02","Furgón 03","Camión 04"],"Pedidos":[14,9,11,8],"Estado":["En reparto","Pendiente","En preparación","En reparto"],"Entrega estimada":["17:30","18:10","16:45","19:00"]})
chat_demo = pd.DataFrame({"Hora":["09:12","09:34","10:05","10:28","11:15"],"Canal":["General","Logística","Privado jefe","Ventas","Depósito"],"Usuario":["Admin","Martín","Laura","Sofía","Nicolás"],"Mensaje":["Revisar pedidos pendientes antes del mediodía.","Ruta CABA salió con 14 entregas.","Cliente Mayorista Río pidió actualización de saldo.","Nueva venta cargada para Autoservicio Sol.","Stock bajo en gaseosas y yerba."]})

if "logged" not in st.session_state: st.session_state.logged = False
if "page" not in st.session_state: st.session_state.page = "Dashboard"

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');
html, body, [class*="css"] {font-family:'Inter',sans-serif;}
.stApp {background: radial-gradient(circle at top left, rgba(212,175,55,.16), transparent 32%), linear-gradient(135deg,#070707 0%,#111 55%,#050505 100%); color:#F8F1D7;}
section[data-testid="stSidebar"]{background:linear-gradient(180deg,#050505,#111);border-right:1px solid rgba(212,175,55,.28);}
section[data-testid="stSidebar"] *{color:#F8F1D7!important;}
[data-testid="stHeader"]{background:rgba(5,5,5,0);}
.demo-banner{background:linear-gradient(90deg,rgba(212,175,55,.25),rgba(245,213,106,.10));border:1px solid rgba(245,213,106,.45);color:#FFE89A;padding:14px 18px;border-radius:18px;font-weight:800;margin-bottom:18px;box-shadow:0 12px 32px rgba(212,175,55,.10);}
.card{background:linear-gradient(180deg,rgba(22,22,22,.95),rgba(10,10,10,.95));border:1px solid rgba(212,175,55,.22);border-radius:24px;padding:22px;box-shadow:0 18px 44px rgba(0,0,0,.45);transition:.25s ease;}
.card:hover{transform:translateY(-3px);border-color:rgba(245,213,106,.45);}
.kpi-label{color:#bba762;font-size:13px;font-weight:700;text-transform:uppercase;}
.kpi-value{color:#F5D56A;font-size:30px;font-weight:900;margin-top:4px;}
.kpi-note{color:#e7dca8;font-size:13px;margin-top:6px;}
.module-header{background:linear-gradient(135deg,#171717,#080808);border:1px solid rgba(245,213,106,.28);border-radius:26px;padding:24px;margin-bottom:18px;box-shadow:0 20px 55px rgba(0,0,0,.42);}
.module-header h2{color:#F5D56A;font-size:30px;font-weight:900;margin:0;}
.module-header p{color:#c8b46a;margin:7px 0 0 0;}
.locked-box{background:rgba(245,213,106,.08);border:1px dashed rgba(245,213,106,.45);color:#FFE89A;border-radius:18px;padding:16px;margin:12px 0;font-weight:700;}
.stButton>button{background:linear-gradient(135deg,#B98A1E,#F5D56A);color:#111!important;border:0;border-radius:14px;font-weight:900;padding:12px 18px;}
.stButton>button:disabled{background:#2b2b2b!important;color:#807244!important;border:1px solid rgba(245,213,106,.18);}
.stTextInput input,.stNumberInput input,.stSelectbox div[data-baseweb="select"],.stTextArea textarea{background:#111!important;border:1px solid rgba(245,213,106,.25)!important;color:#F8F1D7!important;}
[data-testid="stMetric"]{background:linear-gradient(180deg,rgba(22,22,22,.95),rgba(10,10,10,.95));border:1px solid rgba(212,175,55,.22);border-radius:20px;padding:18px;}
</style>
""", unsafe_allow_html=True)

def demo_banner():
    st.markdown('<div class="demo-banner">⚠️ DEMO COMERCIAL — ACCESO LIMITADO · Solo visualización. Carga, edición, guardado, chat real y reportes completos están bloqueados.</div>', unsafe_allow_html=True)
def locked_notice():
    st.markdown('<div class="locked-box">🔒 Función bloqueada en esta demo. Para utilizar esta herramienta con datos reales, solicite acceso completo.</div>', unsafe_allow_html=True)
def header(title, subtitle):
    st.markdown(f'<div class="module-header"><h2>{title}</h2><p>{subtitle}</p></div>', unsafe_allow_html=True)
def kpi(label, value, note):
    st.markdown(f'<div class="card"><div class="kpi-label">{label}</div><div class="kpi-value">{value}</div><div class="kpi-note">{note}</div></div>', unsafe_allow_html=True)
def styled_fig(fig, height=390):
    fig.update_layout(template="plotly_dark", height=height, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#F8F1D7"), title_font=dict(color="#F5D56A", size=20), margin=dict(l=20,r=20,t=55,b=25))
    return fig

def login():
    st.markdown("<br><br>", unsafe_allow_html=True)
    c1,c2,c3=st.columns([1,1.15,1])
    with c2:
        st.markdown('<div class="card" style="text-align:center;"><div style="font-size:54px;">🏢</div><div style="font-size:34px;font-weight:900;color:#F5D56A;">DEL CAMPO</div><div style="font-size:18px;font-weight:800;color:#F8F1D7;">DISTRIBUIDORA</div><div style="color:#c8b46a;margin-top:10px;">Demo comercial de gestión empresarial</div></div>', unsafe_allow_html=True)
        st.info("Acceso demo: usuario **demo** · contraseña **demo123**")
        user=st.text_input("Usuario", value="demo")
        pwd=st.text_input("Contraseña", type="password", value="demo123")
        if st.button("Ingresar a la demo", use_container_width=True):
            if user==DEMO_USER and pwd==DEMO_PASS:
                st.session_state.logged=True; st.rerun()
            else:
                st.error("Acceso no autorizado para esta demo.")
        st.link_button("📞 Solicitar acceso completo", WHATSAPP_LINK, use_container_width=True)

def sidebar():
    with st.sidebar:
        st.markdown("## 🏢 DEL CAMPO")
        st.markdown("**DISTRIBUIDORA**")
        st.caption("Demo comercial bloqueada")
        st.markdown("---")
        pages=["Dashboard","Productos","Ventas","Clientes","Caja","Logística","Chat interno","Reportes","Inteligencia","Configuración"]
        icons={"Dashboard":"📊","Productos":"📦","Ventas":"🧾","Clientes":"👥","Caja":"💰","Logística":"🚚","Chat interno":"💬","Reportes":"📑","Inteligencia":"🧠","Configuración":"⚙️"}
        for p in pages:
            if st.button(f"{icons[p]} {p}", use_container_width=True):
                st.session_state.page=p; st.rerun()
        st.markdown("---")
        st.success("Modo visualización")
        st.caption("Sin guardado · Sin edición · Sin uso operativo")
        if st.button("Cerrar demo", use_container_width=True):
            st.session_state.logged=False; st.rerun()

def dashboard():
    demo_banner(); header("Panel principal","Vista ejecutiva para analizar ventas, caja, stock, clientes y logística.")
    c1,c2,c3,c4,c5=st.columns(5)
    for col,args in zip([c1,c2,c3,c4,c5],[("Ventas del día","$ 685.000","+14% vs ayer"),("Ventas del mes","$ 12.850.000","+22% mensual"),("Caja del día","$ 1.240.000","Saldo positivo"),("Pedidos activos","42","18 en reparto"),("Stock crítico","5","Reposición sugerida")]):
        with col: kpi(*args)
    col1,col2=st.columns([2,1])
    with col1:
        st.plotly_chart(styled_fig(px.line(ventas_mes,x="Día",y=["Ventas","Ganancia"],markers=True,title="Evolución mensual de ventas y ganancias"),420), use_container_width=True)
    with col2:
        mix=pd.DataFrame({"Tipo":["Blanco","Negro","Cuenta corriente","Transferencia"],"Importe":[4200000,1800000,2600000,3100000]})
        st.plotly_chart(styled_fig(px.pie(mix,names="Tipo",values="Importe",hole=.55,title="Distribución comercial"),420), use_container_width=True)
    c1,c2=st.columns(2)
    with c1:
        st.subheader("🚨 Alertas operativas"); st.warning("5 productos por debajo del stock mínimo."); st.info("18 pedidos se encuentran en reparto."); st.error("3 clientes superaron el límite de crédito."); st.success("Caja diaria con saldo positivo.")
    with c2:
        st.subheader("📋 Pedidos recientes"); st.dataframe(pedidos,use_container_width=True,hide_index=True)

def productos_page():
    demo_banner(); header("Productos","Catálogo comercial, precios, stock y alertas de reposición."); locked_notice()
    f1,f2,f3=st.columns(3)
    with f1: st.text_input("Buscar producto", disabled=True, placeholder="Bloqueado en demo")
    with f2: st.selectbox("Categoría", ["Todas"]+sorted(productos["Categoría"].unique()), disabled=True)
    with f3: st.selectbox("Estado", ["Todos","OK","Bajo","Crítico"], disabled=True)
    st.dataframe(productos,use_container_width=True,hide_index=True)
    st.subheader("➕ Carga de producto")
    c1,c2,c3,c4=st.columns(4)
    with c1: st.text_input("Nombre", disabled=True)
    with c2: st.text_input("Categoría", disabled=True)
    with c3: st.number_input("Stock inicial", min_value=0, disabled=True)
    with c4: st.number_input("Precio", min_value=0, disabled=True)
    st.button("Guardar producto", disabled=True, use_container_width=True)

def ventas_page():
    demo_banner(); header("Ventas","Carga de operaciones, control de cobros y seguimiento comercial."); locked_notice()
    c1,c2=st.columns([1,1])
    with c1:
        st.subheader("🧾 Nueva venta")
        for label, opts in [("Cliente",clientes["Cliente"]),("Producto",productos["Producto"]),("Tipo de venta",["Blanco","Negro"]),("Método de pago",["Efectivo","Transferencia","Cuenta corriente","Cheque","Mercado Pago"])]:
            st.selectbox(label, opts, disabled=True)
        st.number_input("Cantidad", min_value=1, value=1, disabled=True)
        st.button("Registrar venta", disabled=True, use_container_width=True)
    with c2:
        st.plotly_chart(styled_fig(px.bar(ventas_categoria,x="Categoría",y="Importe",text_auto=True,title="Ventas por categoría")), use_container_width=True)
    st.subheader("📋 Operaciones recientes"); st.dataframe(pedidos,use_container_width=True,hide_index=True)

def clientes_page():
    demo_banner(); header("Clientes","Cartera comercial, deuda, actividad y ranking de compradores."); locked_notice()
    c1,c2,c3=st.columns(3)
    with c1: st.metric("Clientes activos","184","+12")
    with c2: st.metric("Clientes con deuda","37","-4")
    with c3: st.metric("Deuda total","$ 2.850.000","+6%")
    st.dataframe(clientes,use_container_width=True,hide_index=True)
    ranking=pd.DataFrame({"Cliente":clientes["Cliente"],"Facturación":[930000,420000,180000,1400000,510000,980000,760000]})
    st.plotly_chart(styled_fig(px.bar(ranking.sort_values("Facturación"),x="Facturación",y="Cliente",orientation="h",title="Ranking de clientes")),use_container_width=True)

def caja_page():
    demo_banner(); header("Caja","Ingresos, egresos, gastos, pagos y saldo diario."); locked_notice()
    for col,args in zip(st.columns(4),[("Ingresos","$ 1.240.000","+18%"),("Egresos","$ 420.000","-7%"),("Saldo final","$ 820.000","+12%"),("Transferencias","$ 510.000","+9%")]):
        with col: st.metric(*args)
    mov=pd.DataFrame({"Hora":["09:20","10:15","11:40","13:10","15:35","16:20"],"Tipo":["Ingreso","Ingreso","Egreso","Ingreso","Egreso","Ingreso"],"Detalle":["Venta efectivo","Transferencia cliente","Pago proveedor","Cobro cuenta corriente","Combustible","Venta mayorista"],"Importe":[220000,185000,-95000,310000,-45000,420000]})
    st.dataframe(mov,use_container_width=True,hide_index=True)

def logistica_page():
    demo_banner(); header("Logística","Repartos, rutas, choferes, vehículos y estados de entrega."); locked_notice()
    st.dataframe(rutas,use_container_width=True,hide_index=True)
    st.plotly_chart(styled_fig(px.bar(rutas,x="Ruta",y="Pedidos",color="Estado",title="Pedidos por ruta")),use_container_width=True)
    c1,c2,c3=st.columns(3)
    with c1: st.info("🚚 Ruta CABA: 14 entregas activas")
    with c2: st.warning("⏳ Ruta GBA Norte pendiente de salida")
    with c3: st.success("✅ Ruta Oeste en reparto")

def chat_page():
    demo_banner(); header("Chat interno","Comunicación interna para administración, ventas, depósito y reparto."); locked_notice()
    st.caption("Vista simulada: el envío de mensajes y adjuntos está bloqueado en esta demo.")
    st.dataframe(chat_demo,use_container_width=True,hide_index=True)
    st.selectbox("Canal",["General","Ventas","Logística","Depósito","Privado jefe"],disabled=True)
    st.file_uploader("Adjuntar archivo",disabled=True)
    st.text_area("Mensaje",disabled=True,placeholder="Bloqueado en demo")
    st.button("Enviar mensaje",disabled=True,use_container_width=True)

def reportes_page():
    demo_banner(); header("Reportes","Informes comerciales, stock, deuda, caja y logística."); locked_notice()
    st.selectbox("Tipo de reporte",["Ventas","Stock","Clientes","Caja","Logística"],disabled=True)
    st.plotly_chart(styled_fig(px.area(ventas_mes,x="Día",y="Ventas",title="Reporte visual de ventas")),use_container_width=True)
    st.download_button("Descargar reporte",data="Reporte bloqueado en demo",file_name="demo_bloqueada.txt",disabled=True)

def inteligencia_page():
    demo_banner(); header("Inteligencia de negocio","Indicadores estratégicos para tomar mejores decisiones.")
    c1,c2=st.columns(2)
    with c1:
        margen=pd.DataFrame({"Categoría":["Harinas","Almacén","Bebidas","Lácteos","Pastas"],"Margen":[31,24,38,21,28]})
        st.plotly_chart(styled_fig(px.bar(margen,x="Categoría",y="Margen",title="Margen promedio por categoría")),use_container_width=True)
    with c2:
        proy=pd.DataFrame({"Mes":["Ene","Feb","Mar","Abr","May","Jun"],"Proyección":[8.2,8.7,9.4,10.1,12.8,13.6]})
        st.plotly_chart(styled_fig(px.line(proy,x="Mes",y="Proyección",markers=True,title="Proyección de ventas estimadas")),use_container_width=True)
    st.success("Recomendación demo: reforzar stock en bebidas y harinas antes del cierre mensual.")

def configuracion_page():
    demo_banner(); header("Configuración","Parámetros generales de empresa, usuarios y permisos."); locked_notice()
    st.text_input("Nombre de la empresa",value="DEL CAMPO DISTRIBUIDORA",disabled=True)
    st.text_input("CUIT",value="Bloqueado en demo",disabled=True)
    st.selectbox("Tema visual",["Negro + Dorado Premium"],disabled=True)
    st.checkbox("Activar ventas",value=True,disabled=True)
    st.checkbox("Activar stock",value=True,disabled=True)
    st.checkbox("Activar chat interno",value=True,disabled=True)
    st.button("Guardar configuración",disabled=True,use_container_width=True)

if not st.session_state.logged:
    login()
else:
    sidebar()
    pages={"Dashboard":dashboard,"Productos":productos_page,"Ventas":ventas_page,"Clientes":clientes_page,"Caja":caja_page,"Logística":logistica_page,"Chat interno":chat_page,"Reportes":reportes_page,"Inteligencia":inteligencia_page,"Configuración":configuracion_page}
    pages[st.session_state.page]()
