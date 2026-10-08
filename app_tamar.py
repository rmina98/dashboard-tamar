import os
import pandas as pd
import numpy as np
from datetime import datetime, date, timedelta
import streamlit as st
import plotly.graph_objects as go


# =========================================================
# CONFIGURACIÓN
# =========================================================

st.set_page_config(
    page_title="Dashboard TAMAR Break-Even",
    layout="wide"
)

st.title("📊 Dashboard: TAMAR Break-Even")


# =========================================================
# 1. CARGA Y PREPARACIÓN DE DATOS
# =========================================================

@st.cache_data
def cargar_excel():

    if not os.path.exists("Inputs.xlsx"):
        st.error("❌ Archivo 'Inputs.xlsx' no encontrado en la carpeta.")
        st.stop()

    xls = pd.ExcelFile("Inputs.xlsx")

    df_lec = pd.read_excel(xls, sheet_name="Lecaps")
    df_let = pd.read_excel(xls, sheet_name="Letam")
    df_precios = pd.read_excel(xls, sheet_name="Precios")
    df_tamar = pd.read_excel(xls, sheet_name="Tamar")
    df_feriados = pd.read_excel(xls, sheet_name="Feriados")

    # -----------------------------------------------------
    # Feriados
    # -----------------------------------------------------

    feriados = set(
        pd.to_datetime(df_feriados["Feriados"]).dt.date
    )

    # -----------------------------------------------------
    # TAMAR
    # -----------------------------------------------------

    df_tamar["Fecha"] = pd.to_datetime(
        df_tamar["Fecha"]
    ).dt.date

    df_tamar = (
        df_tamar[df_tamar["Tamar"] > 0]
        .sort_values("Fecha")
        .reset_index(drop=True)
    )

    # -----------------------------------------------------
    # Precios
    # -----------------------------------------------------

    df_precios["Fecha_dt"] = pd.to_datetime(
        df_precios["Fecha"]
    ).dt.date

    df_precios = df_precios.sort_values(
        "Fecha_dt"
    ).reset_index(drop=True)

    return (
        df_lec,
        df_let,
        df_precios,
        df_tamar,
        feriados
    )


df_lec, df_let, df_precios, df_tamar, feriados = cargar_excel()


# =========================================================
# 2. FUNCIONES DE CALENDARIO
# =========================================================

def es_habil(f, feriados):
    """
    Determina si una fecha es día hábil.
    """
    return (
        f.weekday() < 5
        and f not in feriados
    )


def proximo_habil(f, feriados):
    """
    Si la fecha no es hábil, avanza hasta el próximo
    día hábil.
    """

    actual = f

    while not es_habil(actual, feriados):
        actual += timedelta(days=1)

    return actual


def restar_dias_habiles(f, dias, feriados):
    """
    Resta una cantidad de días hábiles.
    """

    actual = f
    restados = 0

    while restados < dias:

        actual -= timedelta(days=1)

        if es_habil(actual, feriados):
            restados += 1

    return actual


def contar_dias_habiles(f_inicio, f_fin, feriados):
    """
    Cuenta días hábiles entre dos fechas.
    """

    actual = f_inicio
    dias = 0

    while actual < f_fin:

        actual += timedelta(days=1)

        if es_habil(actual, feriados):
            dias += 1

    return dias


def generar_dias_habiles(f_inicio, f_fin, feriados):
    """
    Devuelve una lista con todos los días hábiles
    entre f_inicio y f_fin, inclusive.
    """

    fechas = []

    actual = f_inicio

    while actual <= f_fin:

        if es_habil(actual, feriados):
            fechas.append(actual)

        actual += timedelta(days=1)

    return fechas


def calcular_fecha_liq(f_op, plazo_t, feriados):
    """
    Calcula la fecha de liquidación según T+0, T+1, etc.
    """

    actual = f_op
    sumados = 0

    while sumados < plazo_t:

        actual += timedelta(days=1)

        if es_habil(actual, feriados):
            sumados += 1

    return proximo_habil(actual, feriados)


# =========================================================
# 3. CONVENCIÓN 30/360
# =========================================================

def dias360_excel(f_inicio, f_fin):
    """
    Convención 30/360 similar a Excel.
    """

    d1 = f_inicio.day
    m1 = f_inicio.month
    y1 = f_inicio.year

    d2 = f_fin.day
    m2 = f_fin.month
    y2 = f_fin.year

    if d1 == 31:
        d1 = 30

    if d2 == 31 and d1 >= 30:
        d2 = 30

    return (
        (y2 - y1) * 360
        + (m2 - m1) * 30
        + (d2 - d1)
    )


# =========================================================
# 4. FÓRMULA CONTRACTUAL DE LA LETAM
# =========================================================

def calcular_tamar_tem(tamar, margen):
    """
    Calcula la TAMAR TEM según la fórmula contractual:

        TAMAR TEM =
        [(1 + (TAMAR + MARGEN)/(365/32))
         ^ ((365/32)*(1/12))] - 1

    tamar y margen deben estar expresados como decimales.

    Ejemplo:
        TAMAR = 30%  -> 0.30
        Margen = 1%  -> 0.01
    """

    base = 1 + (tamar + margen) / (365.0 / 32.0)

    exponente = (365.0 / 32.0) * (1.0 / 12.0)

    tem = (
        base ** exponente
    ) - 1

    return tem


def factor_letam(tamar, margen, dias360):
    """
    Factor de capitalización contractual de la Letam:

        VPV = VNO *
              (1 + TAMAR TEM) ^ ((DÍAS/360)*12)

    usando la fórmula contractual de TAMAR TEM.

    Se simplifica algebraicamente a:

        Factor =
        [1 + (TAMAR + MARGEN)/(365/32)]
        ^ [(365/32)*(DÍAS/360)]
    """

    base = (
        1
        + (tamar + margen) / (365.0 / 32.0)
    )

    exponente = (
        (365.0 / 32.0)
        * (dias360 / 360.0)
    )

    return base ** exponente


def valor_letam(tamar, margen, dias360, vno=100.0):
    """
    Valor de la Letam para una determinada TAMAR.
    """

    return (
        vno
        * factor_letam(
            tamar,
            margen,
            dias360
        )
    )


# =========================================================
# 5. OBTENER TAMAR MÁS RECIENTE
# =========================================================

def obtener_ultima_tamar(f_op, df_tamar, fallback=0.2413):
    """
    Obtiene la última TAMAR publicada hasta la fecha
    de operación.

    Devuelve la tasa en formato decimal.
    """

    df_hist = df_tamar[
        df_tamar["Fecha"] <= f_op
    ]

    if df_hist.empty:
        return fallback

    return (
        float(df_hist.iloc[-1]["Tamar"])
        / 100.0
    )


# =========================================================
# 6. MOTOR PRINCIPAL
# =========================================================

def calcular_tamar_be(
    ticker_lec,
    ticker_let,
    p_lec,
    p_let,
    f_op,
    df_lec,
    df_let,
    df_tamar,
    feriados,
    plazo_t=1
):

    # -----------------------------------------------------
    # Buscar instrumentos
    # -----------------------------------------------------

    row_lec_df = df_lec[
        df_lec["Ticker"] == ticker_lec
    ]

    row_let_df = df_let[
        df_let["Ticker"] == ticker_let
    ]

    if row_lec_df.empty:
        raise ValueError(
            f"No se encontró la Lecap {ticker_lec}"
        )

    if row_let_df.empty:
        raise ValueError(
            f"No se encontró la Letam {ticker_let}"
        )

    row_lec = row_lec_df.iloc[0]
    row_let = row_let_df.iloc[0]

    # -----------------------------------------------------
    # Fecha de liquidación
    # -----------------------------------------------------

    f_liq = calcular_fecha_liq(
        f_op,
        plazo_t,
        feriados
    )

    # =====================================================
    # A. LECAP / BONCAP
    # =====================================================

    f_vto_lec = pd.to_datetime(
        row_lec["Vencimiento"]
    ).date()

    f_em_lec = pd.to_datetime(
        row_lec["Emisión"]
    ).date()

    tem_lec_emision = float(
        row_lec["Tasa"]
    )

    d360_lec = dias360_excel(
        f_em_lec,
        f_vto_lec
    )

    # -----------------------------------------------------
    # VPV de la Lecap
    # -----------------------------------------------------

    if (
        "VPV" in row_lec.index
        and pd.notnull(row_lec["VPV"])
    ):

        vpv_lec = float(
            row_lec["VPV"]
        )

    else:

        vpv_lec = (
            100.0
            * (
                1 + tem_lec_emision
            ) ** (
                d360_lec / 30.0
            )
        )

        # Caso particular de tu archivo original
        if ticker_lec == "S29E7":
            vpv_lec = 111.6660

    # -----------------------------------------------------
    # Fecha de cobro Lecap
    # -----------------------------------------------------

    f_cobro_lec = proximo_habil(
        f_vto_lec,
        feriados
    )

    dias_cartera_lec = (
        f_cobro_lec - f_liq
    ).days

    if dias_cartera_lec <= 0:
        raise ValueError(
            "La Lecap ya venció o no tiene plazo positivo "
            "desde la fecha de liquidación."
        )

    # -----------------------------------------------------
    # Retorno de la Lecap
    # -----------------------------------------------------

    rend_lec = (
        vpv_lec / p_lec
    ) - 1

    # TIR efectiva anual ACT/365
    tir_lec = (
        (1 + rend_lec)
        ** (365.0 / dias_cartera_lec)
    ) - 1

    # TNA equivalente
    tna_lec = (
        rend_lec
        * (365.0 / dias_cartera_lec)
    )

    # =====================================================
    # B. LETAM
    # =====================================================

    f_vto_let = pd.to_datetime(
        row_let["Vencimiento"]
    ).date()

    f_em_let = pd.to_datetime(
        row_let["Emisión"]
    ).date()

    margen_let = float(
        row_let["Tasa"]
    )

    # -----------------------------------------------------
    # Fecha de cobro Letam
    # -----------------------------------------------------

    f_cobro_let = proximo_habil(
        f_vto_let,
        feriados
    )

    dias_cartera_let = (
        f_cobro_let - f_liq
    ).days

    if dias_cartera_let <= 0:
        raise ValueError(
            "La Letam ya venció o no tiene plazo positivo."
        )

    # -----------------------------------------------------
    # Retorno que debería obtener la Letam
    # para empatar a la Lecap
    # -----------------------------------------------------

    vpv_objetivo = (
        p_let
        * (
            1 + tir_lec
        ) ** (
            dias_cartera_let / 365.0
        )
    )

    # =====================================================
    # C. PERÍODO DE FIXING TAMAR
    # =====================================================

    # El contrato determina la TAMAR:
    #
    # desde 10 días hábiles antes de la emisión
    # hasta 10 días hábiles antes del vencimiento.
    # -----------------------------------------------------

    f_inicio_fixing = restar_dias_habiles(
        f_em_let,
        10,
        feriados
    )

    f_fin_fixing = restar_dias_habiles(
        f_vto_let,
        10,
        feriados
    )

    # Todos los días hábiles del período de fixing
    dias_fixing = generar_dias_habiles(
        f_inicio_fixing,
        f_fin_fixing,
        feriados
    )

    n_total_fixings = len(
        dias_fixing
    )

    # -----------------------------------------------------
    # Fixings ya conocidos
    # -----------------------------------------------------

    dias_fixing_conocidos = [
        f
        for f in dias_fixing
        if f <= f_op
    ]

    # -----------------------------------------------------
    # Fixings futuros
    # -----------------------------------------------------

    dias_fixing_futuros = [
        f
        for f in dias_fixing
        if f > f_op
    ]

    n_conocidos = len(
        dias_fixing_conocidos
    )

    n_futuros = len(
        dias_fixing_futuros
    )

    # =====================================================
    # D. TAMAR HISTÓRICA CONOCIDA
    # =====================================================

    df_dev = df_tamar[
        (
            df_tamar["Fecha"]
            >= f_inicio_fixing
        )
        &
        (
            df_tamar["Fecha"]
            <= f_op
        )
        &
        (
            df_tamar["Fecha"]
            <= f_fin_fixing
        )
    ].copy()

    # Solo nos interesan fechas de fixing que
    # efectivamente estén dentro del período.
    df_dev = df_dev[
        df_dev["Fecha"].isin(
            dias_fixing
        )
    ]

    if not df_dev.empty:

        # TAMAR expresada como decimal
        tamar_conocidas = (
            df_dev["Tamar"]
            .astype(float)
            / 100.0
        )

        suma_tamar_conocida = (
            tamar_conocidas.sum()
        )

        tamar_promedio_conocida = (
            tamar_conocidas.mean()
        )

    else:

        suma_tamar_conocida = 0.0

        tamar_promedio_conocida = np.nan

    # =====================================================
    # E. TAMAR PROYECTADA
    # =====================================================

    tamar_mercado = obtener_ultima_tamar(
        f_op,
        df_tamar
    )

    # La última TAMAR disponible se usa como
    # proyección para los fixings futuros.
    tamar_proyectada = tamar_mercado

    # =====================================================
    # F. TAMAR PROMEDIO PROYECTADA
    # =====================================================

    if n_total_fixings > 0:

        tamar_promedio_proyectada = (
            suma_tamar_conocida
            + (
                tamar_proyectada
                * n_futuros
            )
        ) / n_total_fixings

    else:

        tamar_promedio_proyectada = tamar_proyectada

    # =====================================================
    # G. VALOR TEÓRICO DE LA LETAM CON TAMAR PROYECTADA
    # =====================================================

    d360_total_let = dias360_excel(
        f_em_let,
        f_vto_let
    )

    vpv_proyectado_let = valor_letam(
        tamar=tamar_promedio_proyectada,
        margen=margen_let,
        dias360=d360_total_let,
        vno=100.0
    )

    # =====================================================
    # H. TAMAR PROMEDIO BREAK-EVEN
    # =====================================================

    # Queremos encontrar la TAMAR promedio total X
    # que haga que:
    #
    # VPV_objetivo =
    # 100 * Factor_Letam(X)
    #
    # Despejamos X.

    if d360_total_let <= 0:

        tamar_be_promedio = np.nan

    else:

        exponent_inverse = (
            360.0
            / (
                (365.0 / 32.0)
                * d360_total_let
            )
        )

        base_be = (
            vpv_objetivo / 100.0
        ) ** exponent_inverse

        tamar_be_promedio = (
            (base_be - 1)
            * (365.0 / 32.0)
            - margen_let
        )

    # =====================================================
    # I. TAMAR FUTURA BREAK-EVEN
    # =====================================================

    if n_futuros > 0:

        # La TAMAR promedio final tiene que ser:
        #
        # TAMAR_BE_promedio =
        #
        # (suma TAMAR conocida
        #  + TAMAR futura * N futuros)
        # / N total

        tamar_be_futura = (
            (
                tamar_be_promedio
                * n_total_fixings
            )
            - suma_tamar_conocida
        ) / n_futuros

        es_fija = False

    else:

        # Ya conocemos todos los fixings
        # necesarios para determinar la TAMAR final.

        tamar_be_futura = np.nan

        es_fija = True

    # =====================================================
    # J. SPREAD VS TAMAR DE MERCADO
    # =====================================================

    if not es_fija:

        spread = (
            tamar_mercado
            - tamar_be_futura
        ) * 100.0

    else:

        spread = 0.0

    # =====================================================
    # K. TAMAR FINAL CONOCIDA / PROYECTADA
    # =====================================================

    if es_fija:

        tamar_final_conocida = (
            suma_tamar_conocida
            / n_total_fixings
            if n_total_fixings > 0
            else np.nan
        )

    else:

        tamar_final_conocida = (
            tamar_promedio_proyectada
        )

    # =====================================================
    # L. RESULTADO
    # =====================================================

    return {

        # Fechas
        "fecha": f_op,
        "fecha_liq": f_liq,
        "fecha_inicio_fixing": f_inicio_fixing,
        "fecha_fin_fixing": f_fin_fixing,

        # Lecap
        "tir_lec": tir_lec,
        "tna_lec": tna_lec,
        "vpv_lec": vpv_lec,

        # Letam
        "vpv_objetivo": vpv_objetivo,
        "vpv_proyectado_let": vpv_proyectado_let,
        "margen_let": margen_let,

        # Fixing
        "n_total_fixings": n_total_fixings,
        "n_conocidos": n_conocidos,
        "n_futuros": n_futuros,

        # TAMAR
        "tamar_promedio_conocida": tamar_promedio_conocida,
        "tamar_proyectada": tamar_proyectada,
        "tamar_promedio_proyectada": tamar_promedio_proyectada,
        "tamar_be_promedio": tamar_be_promedio,
        "tamar_be": tamar_be_futura,
        "tamar_mercado": tamar_mercado,
        "spread_pkt": spread,

        # Estado
        "es_fija": es_fija
    }


# =========================================================
# 7. INTERFAZ DE USUARIO
# =========================================================

ultima_fila = df_precios.iloc[-1]

fecha_sim = ultima_fila["Fecha_dt"]


st.sidebar.header(
    "⚙️ Parámetros de Simulación"
)

st.sidebar.write(
    f"**Última fecha en datos:** "
    f"{fecha_sim.strftime('%d/%m/%Y')}"
)


# =========================================================
# SELECCIÓN LECAP
# =========================================================

lecap_elegida = st.sidebar.selectbox(
    "1️⃣ Tasa Fija (Lecap/Boncap):",
    df_lec["Ticker"].tolist()
)


valor_lec_excel = ultima_fila[
    lecap_elegida
]

if pd.isna(valor_lec_excel):

    st.error(
        f"No hay precio de {lecap_elegida} "
        "para la última fecha disponible."
    )

    st.stop()


def_precio_lec = float(
    valor_lec_excel * 100
    if valor_lec_excel < 10
    else valor_lec_excel
)


precio_lecap_sim = st.sidebar.number_input(
    "Precio Lecap:",
    value=def_precio_lec,
    step=0.10,
    format="%.2f"
)


st.sidebar.markdown("---")


# =========================================================
# SELECCIÓN LETAM
# =========================================================

letam_elegida = st.sidebar.selectbox(
    "2️⃣ Tasa Variable (Letam):",
    df_let["Ticker"].tolist()
)


valor_let_excel = ultima_fila[
    letam_elegida
]

if pd.isna(valor_let_excel):

    st.error(
        f"No hay precio de {letam_elegida} "
        "para la última fecha disponible."
    )

    st.stop()


def_precio_let = float(
    valor_let_excel * 100
    if valor_let_excel < 10
    else valor_let_excel
)


precio_letam_sim = st.sidebar.number_input(
    "Precio Letam:",
    value=def_precio_let,
    step=0.10,
    format="%.2f"
)


st.sidebar.markdown("---")


# =========================================================
# RANGO HISTÓRICO
# =========================================================

st.sidebar.subheader(
    "📅 Rango Histórico a Graficar"
)


fecha_min_general = (
    df_precios["Fecha_dt"].min()
)

fecha_max_general = (
    df_precios["Fecha_dt"].max()
)


rango_fechas = st.sidebar.date_input(
    "Seleccioná el período:",
    value=(
        fecha_min_general,
        fecha_max_general
    ),
    min_value=fecha_min_general,
    max_value=fecha_max_general
)


if (
    isinstance(rango_fechas, (tuple, list))
    and len(rango_fechas) == 2
):

    f_desde, f_hasta = rango_fechas

else:

    f_desde = fecha_min_general
    f_hasta = fecha_max_general


# =========================================================
# 8. SIMULACIÓN ACTUAL
# =========================================================

try:

    res_sim = calcular_tamar_be(
        lecap_elegida,
        letam_elegida,
        precio_lecap_sim,
        precio_letam_sim,
        fecha_sim,
        df_lec,
        df_let,
        df_tamar,
        feriados
    )

except Exception as e:

    st.error(
        f"❌ Error en el cálculo: {e}"
    )

    st.stop()


# =========================================================
# 9. MÉTRICAS PRINCIPALES
# =========================================================

col1, col2, col3 = st.columns(3)


col1.metric(
    label=f"TNA {lecap_elegida} (Fija)",
    value=f"{res_sim['tna_lec'] * 100:.2f}%"
)


col2.metric(
    label=f"TAMAR BE Futura {letam_elegida}",
    value=(
        "Ya es Fija"
        if res_sim["es_fija"]
        else f"{res_sim['tamar_be'] * 100:.2f}%"
    )
)


if res_sim["es_fija"]:

    col3.metric(
        label="Estado",
        value="Fixing completo"
    )

else:

    col3.metric(
        label="Spread TAMAR BE vs Mkt",
        value=f"{res_sim['spread_pkt']:+.2f} pkt"
    )


st.markdown("---")


# =========================================================
# 10. INFORMACIÓN DETALLADA
# =========================================================

if res_sim["es_fija"]:

    st.info(
        "⚠️ **El período de fixing finalizó:** "
        "ya se conocen todos los fixings TAMAR necesarios "
        "para determinar la TAMAR contractual de la Letam."
    )

else:

    col_a, col_b = st.columns(2)

    col_a.info(
        f"**TAMAR Break-Even Futura:** "
        f"{res_sim['tamar_be'] * 100:.2f}%"
    )

    col_b.info(
        f"**TAMAR Mercado Hoy:** "
        f"{res_sim['tamar_mercado'] * 100:.2f}%"
    )


# =========================================================
# 11. DATOS DEL FIXING
# =========================================================

with st.expander("🔎 Detalle del cálculo TAMAR"):

    c1, c2, c3, c4 = st.columns(4)

    c1.metric(
        "Fixings totales",
        f"{res_sim['n_total_fixings']}"
    )

    c2.metric(
        "Fixings conocidos",
        f"{res_sim['n_conocidos']}"
    )

    c3.metric(
        "Fixings futuros",
        f"{res_sim['n_futuros']}"
    )

    c4.metric(
        "Margen Letam",
        f"{res_sim['margen_let'] * 100:.2f}%"
    )

    st.write(
        f"**Inicio fixing:** "
        f"{res_sim['fecha_inicio_fixing'].strftime('%d/%m/%Y')}"
    )

    st.write(
        f"**Fin fixing:** "
        f"{res_sim['fecha_fin_fixing'].strftime('%d/%m/%Y')}"
    )

    st.write(
        f"**TAMAR última disponible:** "
        f"{res_sim['tamar_mercado'] * 100:.2f}%"
    )

    st.write(
        f"**TAMAR promedio conocida:** "
        f"{res_sim['tamar_promedio_conocida'] * 100:.2f}%"
        if not pd.isna(
            res_sim["tamar_promedio_conocida"]
        )
        else "**TAMAR promedio conocida:** N/A"
    )

    st.write(
        f"**TAMAR promedio proyectada:** "
        f"{res_sim['tamar_promedio_proyectada'] * 100:.2f}%"
    )

    if not res_sim["es_fija"]:

        st.write(
            f"**TAMAR futura BE:** "
            f"{res_sim['tamar_be'] * 100:.2f}%"
        )


# =========================================================
# 12. HISTÓRICO
# =========================================================

df_precios_filtrado = df_precios[
    (
        df_precios["Fecha_dt"]
        >= f_desde
    )
    &
    (
        df_precios["Fecha_dt"]
        <= f_hasta
    )
]


historico = []


for _, row in df_precios_filtrado.iterrows():

    f_o = row["Fecha_dt"]

    precio_l = row[lecap_elegida]

    precio_t = row[letam_elegida]

    if pd.isna(precio_l) or pd.isna(precio_t):
        continue

    p_l = (
        precio_l * 100
        if precio_l < 10
        else precio_l
    )

    p_t = (
        precio_t * 100
        if precio_t < 10
        else precio_t
    )

    try:

        r = calcular_tamar_be(
            lecap_elegida,
            letam_elegida,
            p_l,
            p_t,
            f_o,
            df_lec,
            df_let,
            df_tamar,
            feriados
        )

        historico.append(r)

    except Exception:
        continue


df_hist = pd.DataFrame(
    historico
)


# =========================================================
# 13. GRÁFICO HISTÓRICO
# =========================================================

st.subheader(
    "📈 Evolución Histórica: BE vs Mercado"
)


if not df_hist.empty:

    # Solo mostramos días donde todavía
    # existía TAMAR futura por determinar.

    df_hist_plot = df_hist[
        ~df_hist["es_fija"]
    ].copy()


    if not df_hist_plot.empty:

        df_hist_plot["be_pct"] = (
            df_hist_plot["tamar_be"]
            * 100
        )

        df_hist_plot["mkt_pct"] = (
            df_hist_plot["tamar_mercado"]
            * 100
        )


        fig = go.Figure()


        # -------------------------------------------------
        # TAMAR BE
        # -------------------------------------------------

        fig.add_trace(
            go.Scatter(
                x=df_hist_plot["fecha"],
                y=df_hist_plot["be_pct"],
                mode="lines",
                name=(
                    f"TAMAR BE "
                    f"({lecap_elegida} vs "
                    f"{letam_elegida})"
                ),
                line=dict(
                    color="#d62728",
                    width=2.5
                ),
                hovertemplate=(
                    "<b>Fecha:</b> "
                    "%{x|%d/%m/%Y}"
                    "<br>"
                    "<b>TAMAR BE:</b> "
                    "%{y:.2f}%"
                    "<extra></extra>"
                )
            )
        )


        # -------------------------------------------------
        # TAMAR MERCADO
        # -------------------------------------------------

        fig.add_trace(
            go.Scatter(
                x=df_hist_plot["fecha"],
                y=df_hist_plot["mkt_pct"],
                mode="lines",
                name="TAMAR Mkt Real",
                line=dict(
                    color="#1f77b4",
                    width=2.5,
                    dash="dash"
                ),
                hovertemplate=(
                    "<b>Fecha:</b> "
                    "%{x|%d/%m/%Y}"
                    "<br>"
                    "<b>TAMAR Mkt:</b> "
                    "%{y:.2f}%"
                    "<extra></extra>"
                )
            )
        )


        # -------------------------------------------------
        # Layout
        # -------------------------------------------------

        fig.update_layout(
            hovermode="x unified",
            yaxis_title="TNA (%)",
            template="plotly_white",
            height=450,
            margin=dict(
                l=10,
                r=10,
                t=30,
                b=10
            ),
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1
            )
        )


        st.plotly_chart(
            fig,
            use_container_width=True
        )


    else:

        st.warning(
            "⚠️ Todos los datos en este rango "
            "corresponden al período donde la "
            "Letam ya tiene el fixing completo."
        )


else:

    st.warning(
        "No hay datos en el rango de fechas seleccionado."
    )