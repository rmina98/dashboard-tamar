import os
import pandas as pd
import numpy as np
from datetime import datetime, date, timedelta
import streamlit as st
import plotly.graph_objects as go

# Configuración de la página
st.set_page_config(page_title="Dashboard TAMAR BE", layout="wide")
st.title("📊 Dashboard: TAMAR Break-Even")

# =========================================================
# 📂 1. CARGA Y PREPARACIÓN DE DATOS
# =========================================================
@st.cache_data
def cargar_excel():
    if not os.path.exists("Inputs.xlsx"):
        st.error("❌ Archivo 'Inputs.xlsx' no encontrado en la carpeta.")
        st.stop()
    
    xls = pd.ExcelFile("Inputs.xlsx")
    df_lec = pd.read_excel(xls, sheet_name='Lecaps')
    df_let = pd.read_excel(xls, sheet_name='Letam')
    df_precios = pd.read_excel(xls, sheet_name='Precios')
    df_tamar = pd.read_excel(xls, sheet_name='Tamar')
    df_feriados = pd.read_excel(xls, sheet_name='Feriados')
    
    feriados = set(pd.to_datetime(df_feriados["Feriados"]).dt.date)
    df_tamar['Fecha'] = pd.to_datetime(df_tamar['Fecha']).dt.date
    df_tamar = df_tamar[df_tamar['Tamar'] > 0].sort_values('Fecha').reset_index(drop=True)
    
    df_precios['Fecha_dt'] = pd.to_datetime(df_precios['Fecha']).dt.date
    
    return df_lec, df_let, df_precios, df_tamar, feriados

df_lec, df_let, df_precios, df_tamar, feriados = cargar_excel()

# =========================================================
# 📅 2. FUNCIONES DE CALENDARIO
# =========================================================
def es_habil(f, feriados):
    return f.weekday() < 5 and f not in feriados

def proximo_habil(f, feriados):
    actual = f
    while not es_habil(actual, feriados):
        actual += timedelta(days=1)
    return actual

def restar_dias_habiles(f, dias, feriados):
    actual = f
    restados = 0
    while restados < dias:
        actual -= timedelta(days=1)
        if es_habil(actual, feriados):
            restados += 1
    return actual

def contar_dias_habiles(f_inicio, f_fin, feriados):
    actual = f_inicio
    dias = 0
    while actual < f_fin:
        actual += timedelta(days=1)
        if es_habil(actual, feriados):
            dias += 1
    return dias

def calcular_fecha_liq(f_op, plazo_t, feriados):
    actual = f_op
    sumados = 0
    while sumados < plazo_t:
        actual += timedelta(days=1)
        if es_habil(actual, feriados):
            sumados += 1
    return proximo_habil(actual, feriados)

def dias360_excel(f_inicio, f_fin):
    d1, m1, y1 = f_inicio.day, f_inicio.month, f_inicio.year
    d2, m2, y2 = f_fin.day, f_fin.month, f_fin.year
    if d1 == 31: d1 = 30
    if d2 == 31 and d1 >= 30: d2 = 30
    return (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)

# =========================================================
# 🧮 3. MOTOR DE CÁLCULO
# =========================================================
def calcular_tamar_be(ticker_lec, ticker_let, p_lec, p_let, f_op, df_lec, df_let, df_tamar, feriados, plazo_t=1):
    row_lec = df_lec[df_lec['Ticker'] == ticker_lec].iloc[0]
    row_let = df_let[df_let['Ticker'] == ticker_let].iloc[0]
    f_liq = calcular_fecha_liq(f_op, plazo_t, feriados)
    
    # =================================================================
    # A. LECAP / BONCAP (TASA FIJA)
    # =================================================================
    f_vto_lec = pd.to_datetime(row_lec['Vencimiento']).date()
    f_em_lec = pd.to_datetime(row_lec['Emisión']).date()
    tem_lec_emision = row_lec['Tasa']
    d360_lec = dias360_excel(f_em_lec, f_vto_lec)
    
    # VPV de la Lecap
    if 'VPV' in row_lec and pd.notnull(row_lec['VPV']):
        vpv_lec = float(row_lec['VPV'])
    else:
        vpv_lec = 100.0 * ((1 + tem_lec_emision) ** (d360_lec / 30.0))
        if ticker_lec == 'S29E7': vpv_lec = 111.6660
            
    f_cobro_lec = proximo_habil(f_vto_lec, feriados)
    dias_cartera_lec = (f_cobro_lec - f_liq).days
    
    # Rendimiento y TIR de la Lecap
    rend_lec = (vpv_lec / p_lec) - 1
    tir_lec = ((1 + rend_lec) ** (365.0 / dias_cartera_lec)) - 1
    tna_lec = rend_lec * (365.0 / dias_cartera_lec)
    
    # =================================================================
    # B. LETAM (TASA VARIABLE) - VPV OBJETIVO
    # =================================================================
    f_vto_let = pd.to_datetime(row_let['Vencimiento']).date()
    f_em_let = pd.to_datetime(row_let['Emisión']).date()
    margen_let = row_let['Tasa']
    f_cobro_let = proximo_habil(f_vto_let, feriados)
    dias_cartera_let = (f_cobro_let - f_liq).days
    
    # VPV que necesita la Letam para empatarle a la TIR de la Lecap
    vpv_objetivo = p_let * ((1 + tir_lec) ** (dias_cartera_let / 365.0))
    
    # =================================================================
    # C. FIXING Y VALOR TÉCNICO (VTC) ACUMULADO
    # =================================================================
    # Lag de 10 días hábiles
    f_inicio_fixing = restar_dias_habiles(f_em_let, 10, feriados)
    f_fin_fixing = restar_dias_habiles(f_vto_let, 10, feriados)
    f_corte_dev = min(f_op, f_fin_fixing)
    
    # Promedio TAMAR ya devengada (conocida)
    df_dev = df_tamar[(df_tamar['Fecha'] >= f_inicio_fixing) & (df_tamar['Fecha'] <= f_corte_dev)]
    if not df_dev.empty:
        dev_tna = df_dev['Tamar'].mean() / 100.0
    else:
        dev_tna = 0.2413 # Fallback de tasa referencial si no hay datos
        
    # Calcular cuánto capitalizó la Letam desde la emisión hasta hoy (Valor Técnico)
    d360_pasados = max(0, dias360_excel(f_em_let, f_liq))
    tna_pasada = dev_tna + margen_let
    tem_pasada = tna_pasada * (30.0 / 365.0)
    vtc_hoy = 100.0 * ((1 + tem_pasada) ** (d360_pasados / 30.0))
    
    # =================================================================
    # D. DESPEJE DE LA TAMAR BREAK-EVEN FUTURA
    # =================================================================
    d360_restante = max(1, dias360_excel(f_liq, f_vto_let))
    
    if f_op < f_fin_fixing:
        # Qué TEM futura necesito desde el VTC hoy hasta el VPV objetivo
        tem_futura_req = ((vpv_objetivo / vtc_hoy) ** (30.0 / d360_restante)) - 1
        
        # Despejo la TAMAR (TNA) de esa TEM, restándole el margen de la Letam
        tamar_be = (tem_futura_req * (365.0 / 30.0)) - margen_let
    else:
        # Entró en los últimos 10 días hábiles, es tasa fija.
        tamar_be = 0.0 
        
    # TAMAR Mercado Hoy para comparar
    df_hist_mercado = df_tamar[df_tamar['Fecha'] <= f_op]
    tamar_mercado = df_hist_mercado['Tamar'].iloc[-1] / 100.0 if not df_hist_mercado.empty else dev_tna
    
    spread = (tamar_mercado - tamar_be) * 100 if (f_op < f_fin_fixing) else 0.0
        
    return {
        'fecha': f_op, 'tir_lec': tir_lec, 'tna_lec': tna_lec, 'vpv_objetivo': vpv_objetivo, 
        'tamar_be': tamar_be, 'tamar_mercado': tamar_mercado, 'spread_pkt': spread,
        'es_fija': f_op >= f_fin_fixing
    }

# =========================================================
# 🎛️ 4. INTERFAZ DE USUARIO (DASHBOARD)
# =========================================================
ultima_fila = df_precios.iloc[-1]
fecha_sim = ultima_fila['Fecha_dt']

st.sidebar.header("⚙️ Parámetros de Simulación")
st.sidebar.write(f"**Última fecha en datos:** {fecha_sim.strftime('%d/%m/%Y')}")

# Selección de Bonos y Precios
lecap_elegida = st.sidebar.selectbox("1️⃣ Tasa Fija (Lecap/Boncap):", df_lec['Ticker'].tolist())
def_precio_lec = float(ultima_fila[lecap_elegida] * 100 if ultima_fila[lecap_elegida] < 10 else ultima_fila[lecap_elegida])
precio_lecap_sim = st.sidebar.number_input("Precio Lecap:", value=def_precio_lec, step=0.10, format="%.2f")

st.sidebar.markdown("---")

letam_elegida = st.sidebar.selectbox("2️⃣ Tasa Variable (Letam):", df_let['Ticker'].tolist())
def_precio_let = float(ultima_fila[letam_elegida] * 100 if ultima_fila[letam_elegida] < 10 else ultima_fila[letam_elegida])
precio_letam_sim = st.sidebar.number_input("Precio Letam:", value=def_precio_let, step=0.10, format="%.2f")

st.sidebar.markdown("---")

# Selector de Rango de Fechas para la Histórica
st.sidebar.subheader("📅 Rango Histórico a Graficar")
fecha_min_general = df_precios['Fecha_dt'].min()
fecha_max_general = df_precios['Fecha_dt'].max()

rango_fechas = st.sidebar.date_input(
    "Seleccioná el período:",
    value=(fecha_min_general, fecha_max_general),
    min_value=fecha_min_general,
    max_value=fecha_max_general
)

if isinstance(rango_fechas, (tuple, list)) and len(rango_fechas) == 2:
    f_desde, f_hasta = rango_fechas
else:
    f_desde, f_hasta = fecha_min_general, fecha_max_general

# =========================================================
# 🚀 5. EJECUCIÓN Y RENDERIZADO
# =========================================================
# Simulación a Hoy / Precio ingresado
res_sim = calcular_tamar_be(lecap_elegida, letam_elegida, precio_lecap_sim, precio_letam_sim, fecha_sim, 
                            df_lec, df_let, df_tamar, feriados)

# Métricas Principales
col1, col2, col3 = st.columns(3)
col1.metric(label=f"TNA {lecap_elegida} (Fija)", value=f"{res_sim['tna_lec']*100:.2f}%")
col2.metric(label=f"VPV Req {letam_elegida} (Var)", value=f"${res_sim['vpv_objetivo']:.2f}")

if res_sim['es_fija']:
    col3.metric(label="Spread TAMAR BE vs Mkt", value="Ya es Fija")
    st.markdown("---")
    st.info("⚠️ **El período de fixing finalizó:** La Letam se encuentra dentro de los 10 días hábiles previos al vencimiento. Ya devengó la totalidad de la TAMAR y ahora opera como un instrumento a Tasa Fija, por lo que no corresponde calcular una Break-Even futura.")
else:
    col3.metric(label="Spread TAMAR BE vs Mkt", value=f"{res_sim['spread_pkt']:+.2f} pkt")
    st.markdown("---")
    col_a, col_b = st.columns(2)
    col_a.info(f"**TAMAR Break-Even Futura:** {res_sim['tamar_be']*100:.2f}%")
    col_b.info(f"**TAMAR Mercado Hoy:** {res_sim['tamar_mercado']*100:.2f}%")

# Filtrado de Precios para la Histórica
df_precios_filtrado = df_precios[(df_precios['Fecha_dt'] >= f_desde) & (df_precios['Fecha_dt'] <= f_hasta)]

historico = []
for _, row in df_precios_filtrado.iterrows():
    f_o = row['Fecha_dt']
    p_l = row[lecap_elegida] * 100 if row[lecap_elegida] < 10 else row[lecap_elegida]
    p_t = row[letam_elegida] * 100 if row[letam_elegida] < 10 else row[letam_elegida]
    
    r = calcular_tamar_be(lecap_elegida, letam_elegida, p_l, p_t, f_o, df_lec, df_let, df_tamar, feriados)
    historico.append(r)

df_hist = pd.DataFrame(historico)

# Gráfico Histórico Interactivo
st.subheader("📈 Evolución Histórica: BE vs Mercado")

if not df_hist.empty:
    # Filtramos del gráfico los días donde ya cerró el fixing, para que la línea no caiga a 0% arruinando el eje Y.
    df_hist_plot = df_hist[~df_hist['es_fija']].copy()
    
    if not df_hist_plot.empty:
        df_hist_plot['be_pct'] = df_hist_plot['tamar_be'] * 100
        df_hist_plot['mkt_pct'] = df_hist_plot['tamar_mercado'] * 100

        fig = go.Figure()

        fig.add_trace(go.Scatter(
            x=df_hist_plot['fecha'],
            y=df_hist_plot['be_pct'],
            mode='lines',
            name=f'TAMAR BE ({lecap_elegida} vs {letam_elegida})',
            line=dict(color='#d62728', width=2.5),
            hovertemplate='<b>Fecha:</b> %{x|%d/%m/%Y}<br><b>TAMAR BE:</b> %{y:.2f}%<extra></extra>'
        ))

        fig.add_trace(go.Scatter(
            x=df_hist_plot['fecha'],
            y=df_hist_plot['mkt_pct'],
            mode='lines',
            name='TAMAR Mkt Real',
            line=dict(color='#1f77b4', width=2.5, dash='dash'),
            hovertemplate='<b>Fecha:</b> %{x|%d/%m/%Y}<br><b>TAMAR Mkt Real:</b> %{y:.2f}%<extra></extra>'
        ))

        fig.update_layout(
            hovermode='x unified',
            yaxis_title='TNA (%)',
            template='plotly_white',
            height=450,
            margin=dict(l=10, r=10, t=30, b=10),
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1
            )
        )

        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("⚠️ Todos los datos en este rango de fechas corresponden al período donde la Letam ya operaba como tasa fija (cierre de fixing).")
else:
    st.warning("No hay datos en el rango de fechas seleccionado.")