"""Dashboard de ventas para una cadena de restaurantes (Streamlit + Plotly).

Encabezados esperados (Excel o CSV), formato DataventasEnero:
tienda, fecha_proceso, Hora_Digitado, familia, Sub_Familia, id_articulo, productos, precio, cantidad
(TOTAL es opcional: si falta se calcula cantidad x precio; la hora sale de Hora_Digitado).
También se entiende el formato del notebook (IDTIENDA, FDIGITADO, DESCRIPCION, DESCRIP2, PRECIO...).

El archivo se busca junto a este app.py (o en la carpeta Origen) con un nombre que empiece
por "Dataventas", por ejemplo DataventasEnero.xlsx. También se puede subir desde la barra
lateral. Si no hay datos, se muestran datos de demostración.
"""
import io
import warnings
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
from scipy import stats

st.set_page_config(page_title="Dashboard de Ventas", page_icon="🍽️", layout="wide")

# ----------------------------------------------------------------------
# Configuración: nombres de columnas (los mismos del notebook)
# ----------------------------------------------------------------------
ID_FECHA = "FECHA"                  # fecha y hora en que se digitó la venta (Hora_Digitado)
ID_FECHA_PROCESO = "FECHA_PROCESO"  # día de negocio (opcional)
ID_HORA = "HORA"
ID_LOCAL = "LOCAL"
ID_DISTRITO = "DISTRITO"
ID_ZONA = "ZONA"
ID_ARTICULO = "ARTICULO"
ID_TIPO_ARTICULO = "TIPO_ARTICULO"  # familia
ID_SUBFAMILIA = "SUB_FAMILIA"
ID_CANTIDAD = "CANTIDAD"
ID_TOTAL = "TOTAL"
ID_PRECIO_UNITARIO = "PRECIO_UNITARIO"

# Encabezados de tus archivos (en mayúsculas) -> nombres internos del dashboard
ALIAS_COLUMNAS = {
    # DataventasEnero: tienda, fecha_proceso, Hora_Digitado, familia, Sub_Familia,
    #                  id_articulo, productos, precio, cantidad
    "TIENDA": ID_LOCAL,
    "HORA_DIGITADO": ID_FECHA,
    "FAMILIA": ID_TIPO_ARTICULO,
    "PRODUCTOS": ID_ARTICULO,
    # Dataventas del notebook: IDTIENDA, FECHAPROCESO, FDIGITADO, DESCRIPCION, DESCRIP2...
    "IDTIENDA": ID_LOCAL,
    "FDIGITADO": ID_FECHA,
    "FECHAPROCESO": ID_FECHA_PROCESO,
    "DESCRIP2": ID_ARTICULO,
    "DESCRIPCION": ID_TIPO_ARTICULO,
    "PRECIO": ID_PRECIO_UNITARIO,
}
CLAVES_COLUMNAS = set(ALIAS_COLUMNAS) | {
    ID_FECHA, ID_FECHA_PROCESO, ID_LOCAL, ID_ARTICULO, ID_TIPO_ARTICULO, ID_SUBFAMILIA,
    ID_CANTIDAD, ID_PRECIO_UNITARIO, ID_TOTAL,
}

# Se busca un archivo cuyo nombre empiece por alguno de estos prefijos
PREFIJOS_ARCHIVO = ("dataventas", "articles")
EXTENSIONES = (".csv", ".xlsx", ".xls")


def buscar_archivo():
    """Devuelve el archivo de datos más reciente junto a app.py o en la carpeta Origen."""
    candidatos = []
    for base in {Path(__file__).resolve().parent, Path.cwd()}:
        for carpeta in (base, base / "Origen"):
            if carpeta.is_dir():
                candidatos += [
                    p for p in carpeta.iterdir()
                    if p.is_file() and p.suffix.lower() in EXTENSIONES
                    and p.stem.lower().startswith(PREFIJOS_ARCHIVO)
                ]
    return max(candidatos, key=lambda p: p.stat().st_mtime) if candidatos else None

# Los 10 locales de muestra (coordenadas aproximadas del centro de cada distrito)
LOCALES = pd.DataFrame(
    [
        (3, "BREÑA", "CENTRO", -12.0570, -77.0530),
        (5, "RIMAC", "NORTE", -12.0290, -77.0430),
        (9, "SJL", "ESTE", -11.9700, -77.0050),
        (12, "SJM", "SUR", -12.1560, -76.9740),
        (14, "STA ANITA", "ESTE", -12.0450, -76.9720),
        (22, "LOS OLIVOS", "NORTE", -11.9700, -77.0720),
        (24, "SAN MIGUEL", "CENTRO", -12.0770, -77.0830),
        (25, "MIRAFLORES", "SUR", -12.1190, -77.0290),
        (26, "SAN BORJA", "SUR", -12.1000, -76.9990),
        (27, "BARRANCO", "SUR", -12.1400, -77.0210),
    ],
    columns=[ID_LOCAL, ID_DISTRITO, ID_ZONA, "LAT", "LON"],
)
ETIQUETA_LOCAL = {
    int(l): f"{l} · {d} ({z.title()})"
    for l, d, z in zip(LOCALES[ID_LOCAL], LOCALES[ID_DISTRITO], LOCALES[ID_ZONA])
}


def etiqueta(local):
    """Texto del local; si no es uno de los 10 de muestra aparece como DESCONOCIDO."""
    return ETIQUETA_LOCAL.get(int(local), f"{int(local)} · DESCONOCIDO")

ZONAS = ["CENTRO", "NORTE", "ESTE", "SUR"]
COLOR_ZONA = {"CENTRO": "#E4572E", "NORTE": "#29335C", "ESTE": "#F3A712", "SUR": "#2E9CCA"}
DIAS = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]
DIA_EN = dict(zip(DIAS, ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]))
TODOS_MESES = "Todos los meses"

st.markdown(
    """
    <style>
    div[data-testid="stMetric"] {
        background: rgba(128,128,128,0.08);
        border-radius: 12px;
        padding: 12px 16px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ----------------------------------------------------------------------
# Carga y preparación de datos
# ----------------------------------------------------------------------
def leer_tabla(origen, nombre):
    """Lee CSV o Excel; prueba utf-8 / latin-1 y separador , o ;"""
    if nombre.lower().endswith((".xlsx", ".xls")):
        # Si el libro tiene varias hojas (Hoja1, Hoja2...), se usa la que tiene los encabezados
        hojas = pd.read_excel(origen, sheet_name=None)

        def puntaje(t):
            cols = {str(c).strip().upper() for c in t.columns}
            return (len(cols & CLAVES_COLUMNAS), len(t))

        return max(hojas.values(), key=puntaje)
    # CSV: se detecta la codificación y el separador (; , tab o |) mirando el encabezado,
    # así una coma dentro de un nombre de producto no confunde la lectura.
    crudo = Path(origen).read_bytes() if isinstance(origen, (str, Path)) else origen.getvalue()
    try:
        texto = crudo.decode("utf-8-sig")
    except UnicodeDecodeError:
        texto = crudo.decode("latin-1")
    encabezado = texto.split("\n", 1)[0]
    sep = max((";", ",", "\t", "|"), key=encabezado.count)
    try:
        df = pd.read_csv(io.StringIO(texto), sep=sep)
        df.attrs["filas_omitidas"] = 0
    except pd.errors.ParserError:
        # Alguna fila tiene más o menos campos que el encabezado: se omite y se avisa cuántas fueron
        omitidas = []
        df = pd.read_csv(
            io.StringIO(texto), sep=sep, engine="python",
            on_bad_lines=lambda fila: omitidas.append(1),
        )
        df.attrs["filas_omitidas"] = len(omitidas)
    return df


def a_numero(serie):
    """Convierte a número; si viene como texto acepta coma decimal."""
    if not pd.api.types.is_numeric_dtype(serie):
        serie = serie.astype(str).str.strip().str.replace(",", ".", regex=False)
    return pd.to_numeric(serie, errors="coerce")


def parse_fecha(serie):
    """Convierte a fecha/hora. Acepta fechas de Excel, texto M/D/AAAA o D/M/AAAA."""
    if pd.api.types.is_datetime64_any_dtype(serie):
        return serie
    if pd.api.types.is_numeric_dtype(serie):  # número de serie de Excel
        return pd.to_datetime(serie, unit="D", origin="1899-12-30", errors="coerce")
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        mes_primero = pd.to_datetime(serie, errors="coerce")
        dia_primero = pd.to_datetime(serie, errors="coerce", dayfirst=True)
    # Se queda con la lectura que entiende más filas (en caso de empate, mes primero)
    return dia_primero if dia_primero.isna().sum() < mes_primero.isna().sum() else mes_primero


def preparar(df):
    omitidas = df.attrs.get("filas_omitidas", 0)
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip().str.upper()
    for origen, destino in ALIAS_COLUMNAS.items():
        if origen in df.columns and destino not in df.columns:
            df = df.rename(columns={origen: destino})

    for c in (ID_CANTIDAD, ID_PRECIO_UNITARIO, ID_TOTAL):
        if c in df.columns:
            df[c] = a_numero(df[c])

    # Igual que el notebook: si no hay TOTAL (o tiene vacíos) se calcula CANTIDAD x PRECIO
    if ID_TOTAL not in df.columns or df[ID_TOTAL].isna().any():
        if {ID_CANTIDAD, ID_PRECIO_UNITARIO} <= set(df.columns):
            df[ID_TOTAL] = df[ID_CANTIDAD] * df[ID_PRECIO_UNITARIO]

    requeridas = [ID_FECHA, ID_LOCAL, ID_ARTICULO, ID_TIPO_ARTICULO, ID_CANTIDAD, ID_TOTAL]
    faltan = [c for c in requeridas if c not in df.columns]
    if faltan:
        raise KeyError(
            f"Faltan columnas: {', '.join(faltan)}. Columnas encontradas: {', '.join(df.columns)}"
        )

    # Fecha y hora digitada (Hora_Digitado) y, si existe, el día de negocio (fecha_proceso)
    df[ID_FECHA] = parse_fecha(df[ID_FECHA])
    df = df.dropna(subset=[ID_FECHA])
    if ID_FECHA_PROCESO in df.columns:
        proceso = parse_fecha(df[ID_FECHA_PROCESO]).fillna(df[ID_FECHA])
    else:
        proceso = df[ID_FECHA]

    # Hora: igual que el notebook, sale de la hora de FDIGITADO.
    # Si la fecha no trae hora, se usa la columna HORA (acepta 13, "13:45:00" o 1345).
    if (df[ID_FECHA].dt.hour != 0).any() or (df[ID_FECHA].dt.minute != 0).any():
        hora = df[ID_FECHA].dt.hour
    elif ID_HORA in df.columns:
        hora_num = pd.to_numeric(df[ID_HORA], errors="coerce")
        if hora_num.notna().mean() > 0.9:
            hora = hora_num // 100 if hora_num.max() > 24 else hora_num
        else:
            hora = pd.to_numeric(df[ID_HORA].astype(str).str.extract(r"(\d{1,2})\s*:")[0], errors="coerce")
    else:
        hora = pd.Series(np.nan, index=df.index)
    df["HORA_N"] = hora.astype(float).round()

    # Local -> distrito, zona y coordenadas. Como en el notebook, un local que no está
    # en la tabla de los 10 queda como DESCONOCIDO (se mantiene en los totales).
    df[ID_LOCAL] = (
        pd.to_numeric(df[ID_LOCAL].astype(str).str.extract(r"(\d+)")[0], errors="coerce")
        .fillna(-1)
        .astype(int)
    )
    df = df.merge(LOCALES, on=ID_LOCAL, how="left", suffixes=("_ORIG", ""))
    df[ID_DISTRITO] = df[ID_DISTRITO].fillna("DESCONOCIDO")
    df[ID_ZONA] = df[ID_ZONA].fillna("DESCONOCIDO")

    for c in (ID_CANTIDAD, ID_TOTAL):
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    # Los nombres se dejan tal cual vienen (igual que el notebook, p. ej. "CERVEZAS " con espacio)
    df[ID_ARTICULO] = df[ID_ARTICULO].fillna("SIN DATO").astype(str)
    df[ID_TIPO_ARTICULO] = df[ID_TIPO_ARTICULO].fillna("SIN DATO").astype(str)
    if ID_SUBFAMILIA in df.columns:
        df[ID_SUBFAMILIA] = df[ID_SUBFAMILIA].fillna("SIN DATO").astype(str)

    # Columnas derivadas, calculadas con los dos criterios de día:
    #   DIG = día calendario de Hora_Digitado (como el notebook)
    #   NEG = día de negocio (fecha_proceso): una venta a las 00:30 cuenta en el día anterior
    for sufijo, momento in (("DIG", df[ID_FECHA]), ("NEG", proceso)):
        df[f"DIA_{sufijo}"] = momento.dt.normalize()
        df[f"MES_{sufijo}"] = momento.dt.to_period("M").astype(str)
        df[f"DIA_SEMANA_{sufijo}"] = pd.Categorical(
            momento.dt.dayofweek.map(dict(enumerate(DIAS))), categories=DIAS, ordered=True
        )
    df.attrs["filas_omitidas"] = omitidas
    return df


@st.cache_data(show_spinner="Cargando datos...")
def cargar_desde_ruta(ruta):
    return preparar(leer_tabla(ruta, ruta))


@st.cache_data(show_spinner="Procesando archivo...")
def cargar_desde_bytes(contenido, nombre):
    return preparar(leer_tabla(io.BytesIO(contenido), nombre))


@st.cache_data(show_spinner="Generando datos de demostración...")
def cargar_demo():
    """Datos inventados con los mismos encabezados que DataventasEnero."""
    rng = np.random.default_rng(7)
    catalogo = {
        "POLLOS BRASA": {
            "POLLOS": [("1/4 POLLO A LA BRASA", 24), ("1/2 POLLO A LA BRASA", 42), ("POLLO ENTERO", 78)],
        },
        "PARRILLAS Y CARNES": {
            "PARRILLAS": [("ANTICUCHOS", 28), ("PARRILLA PERSONAL", 55)],
            "CARNES": [("CHURRASCO", 38), ("LOMO SALTADO", 36)],
        },
        "PIZZAS": {
            "GIGANTES": [("SUPERGIGANTE AMERICANA+1 LT.", 48), ("SUPERGIGANTE HAWAIANA + 1LT GASEOSA", 60)],
            "PERSONALES": [("PIZZA PERSONAL PEPPERONI", 24)],
        },
        "PIQUEOS": {
            "PARA COMPARTIR": [("TABLA DE PIQUEOS", 45), ("CHICHARRON DE POLLO", 32)],
            "SNACKS": [("PAPAS FRITAS", 12), ("ALITAS BBQ", 28)],
        },
        "JUGOS Y REFRESCOS": {"JARRAS": [("CHICHA MORADA", 9), ("LIMONADA", 10)]},
        "CERVEZAS": {"CHOOP": [("CHOPP LITRO PILSEN", 24.9), ("CERVEZA ARTESANAL", 14)]},
        "POSTRES": {"DULCES": [("SUSPIRO LIMEÑO", 14), ("CHEESECAKE", 16)]},
    }
    items = [
        (fam, sub, art, precio)
        for fam, subs in catalogo.items()
        for sub, lista in subs.items()
        for art, precio in lista
    ]
    familias = np.array([i[0] for i in items])
    subfamilias = np.array([i[1] for i in items])
    productos = np.array([i[2] for i in items])
    precios = np.array([i[3] for i in items], dtype=float)
    peso_art = rng.uniform(0.5, 2.0, len(items))
    peso_art /= peso_art.sum()

    horas = np.arange(11, 26)  # 11:00 ... 01:59 (24 y 25 = después de medianoche)
    peso_hora = np.array([2, 6, 9, 8, 4, 2, 2, 4, 7, 9, 6, 2, 1.5, 1, 0.3])
    peso_hora /= peso_hora.sum()

    base_local = {int(l): rng.uniform(0.7, 1.4) for l in LOCALES[ID_LOCAL]}
    bloques = []
    for dia in pd.date_range("2026-09-01", "2026-09-30"):
        factor = 1.6 if dia.dayofweek >= 4 else 1.0
        for local, base in base_local.items():
            n = int(rng.poisson(110 * base * factor))
            idx = rng.choice(len(items), n, p=peso_art)
            h = rng.choice(horas, n, p=peso_hora)
            minutos = rng.integers(0, 60, n)
            digitado = dia + pd.to_timedelta(h, unit="h") + pd.to_timedelta(minutos, unit="m")
            bloques.append(
                pd.DataFrame(
                    {
                        "tienda": local,
                        "fecha_proceso": dia,
                        "Hora_Digitado": digitado,
                        "familia": familias[idx],
                        "Sub_Familia": subfamilias[idx],
                        "id_articulo": idx + 1000,
                        "productos": productos[idx],
                        "precio": precios[idx],
                        "cantidad": rng.integers(1, 5, n),
                    }
                )
            )
    return preparar(pd.concat(bloques, ignore_index=True))


def soles(x):
    return f"S/ {x:,.0f}"


def mostrar(fig, alto=380):
    fig.update_layout(height=alto, margin=dict(l=10, r=10, t=40, b=10))
    st.plotly_chart(fig, width="stretch")


# ----------------------------------------------------------------------
# Correlación: Pearson, Spearman y Kendall
# ----------------------------------------------------------------------
METODOS_CORR = {
    "Pearson": (stats.pearsonr, "Relación lineal entre dos variables. Es sensible a valores extremos."),
    "Spearman": (stats.spearmanr, "Relación monótona calculada con rangos. Resiste mejor los valores extremos."),
    "Kendall": (stats.kendalltau, "Relación monótona por pares concordantes. Buena con pocos datos o muchos empates."),
}
NIVELES_CORR = ["Local × día", "Producto", "Línea de venta"]
DEFECTO_CORR = {
    "Local × día": ("TRANSACCIONES", "TICKET_PROMEDIO"),
    "Producto": ("PRECIO_PROMEDIO", "UNIDADES"),
    "Línea de venta": ("PRECIO", "CANTIDAD"),
}


def tabla_correlacion(datos, nivel):
    """Tabla numérica sobre la que se calculan las correlaciones (una fila = un punto)."""
    hay_precio = ID_PRECIO_UNITARIO in datos.columns
    if nivel == "Local × día":
        agg = dict(
            VENTAS=(ID_TOTAL, "sum"), UNIDADES=(ID_CANTIDAD, "sum"),
            TRANSACCIONES=(ID_TOTAL, "count"), TICKET_PROMEDIO=(ID_TOTAL, "mean"),
        )
        if hay_precio:
            agg["PRECIO_PROMEDIO"] = (ID_PRECIO_UNITARIO, "mean")
        return datos.groupby([ID_LOCAL, "DIA"]).agg(**agg).reset_index(drop=True)
    if nivel == "Producto":
        agg = dict(
            VENTAS=(ID_TOTAL, "sum"), UNIDADES=(ID_CANTIDAD, "sum"),
            LOCALES_QUE_LO_VENDEN=(ID_LOCAL, "nunique"), DIAS_CON_VENTA=("DIA", "nunique"),
        )
        if hay_precio:
            agg["PRECIO_PROMEDIO"] = (ID_PRECIO_UNITARIO, "mean")
        return datos.groupby(ID_ARTICULO).agg(**agg).reset_index(drop=True)
    columnas = ([(ID_PRECIO_UNITARIO, "PRECIO")] if hay_precio else []) + [
        (ID_CANTIDAD, "CANTIDAD"), (ID_TOTAL, "TOTAL"), ("HORA_N", "HORA"),
    ]
    return datos[[c for c, _ in columnas]].rename(columns=dict(columnas))


def interpretar(coef, p):
    if np.isnan(coef):
        return "No se pudo calcular"
    fuerza = abs(coef)
    sig = "significativa" if p < 0.05 else "no significativa"
    if fuerza < 0.2:
        return f"Casi sin relación ({sig})"
    nivel = "débil" if fuerza < 0.4 else "moderada" if fuerza < 0.6 else "fuerte" if fuerza < 0.8 else "muy fuerte"
    return f"Correlación {nivel} {'positiva' if coef > 0 else 'negativa'} ({sig})"


def correlacionar(x, y):
    """Devuelve los pares válidos y una tabla con las tres pruebas (None si no se puede calcular)."""
    d = pd.DataFrame({"x": x, "y": y}).dropna()
    if len(d) < 3 or d["x"].nunique() < 2 or d["y"].nunique() < 2:
        return d, None
    filas = []
    for nombre, (prueba, _) in METODOS_CORR.items():
        res = prueba(d["x"], d["y"])
        coef, p = float(res[0]), float(res[1])
        filas.append({
            "Prueba": nombre, "Coeficiente": coef,
            "Valor p": "< 0.0001" if p < 1e-4 else f"{p:.4f}",
            "n": len(d), "Lectura": interpretar(coef, p),
        })
    return d, pd.DataFrame(filas)


# ----------------------------------------------------------------------
# Barra lateral: origen de datos y filtros
# ----------------------------------------------------------------------
st.sidebar.title("🍽️ Filtros")
archivo = st.sidebar.file_uploader("Subir datos (CSV o Excel)", type=["csv", "xlsx"])

es_demo = False
try:
    if archivo is not None:
        df = cargar_desde_bytes(archivo.getvalue(), archivo.name)
        fuente = f"Archivo subido: {archivo.name}"
    else:
        ruta = buscar_archivo()
        if ruta is not None:
            df = cargar_desde_ruta(str(ruta))
            fuente = f"Archivo: {ruta.name}"
        else:
            df = cargar_demo()
            fuente = "Datos de demostración"
            es_demo = True
except Exception as e:  # noqa: BLE001
    st.error(f"No pude cargar los datos: {e}")
    st.stop()

if df.empty:
    st.error("El archivo no tiene filas válidas. Revisa las columnas de fecha y de local.")
    st.stop()
filas_omitidas = df.attrs.get("filas_omitidas", 0)

# Criterio de día: solo se pregunta si hay ventas que cruzan la medianoche
usa_negocio = False
if (df["DIA_NEG"] != df["DIA_DIG"]).any():
    criterio = st.sidebar.radio(
        "Día de análisis",
        ["Fecha digitada", "Día de negocio (fecha_proceso)"],
        help=(
            "Una venta digitada a las 00:30 puede contarse en el día calendario (fecha digitada, "
            "como el notebook) o en el día en que abrió el turno (fecha_proceso)."
        ),
    )
    usa_negocio = criterio.startswith("Día de negocio")
sufijo = "NEG" if usa_negocio else "DIG"
df = df.rename(
    columns={f"DIA_{sufijo}": "DIA", f"MES_{sufijo}": "MES", f"DIA_SEMANA_{sufijo}": "DIA_SEMANA"}
)

meses = sorted(df["MES"].unique())
opciones_mes = [TODOS_MESES] + meses

# Mes por defecto: ignora los meses con muy pocas filas (ventas de madrugada que se pasan al
# mes siguiente). Si solo hay un mes real, se abre con "Todos los meses" para no perder esas filas.
filas_por_mes = df["MES"].value_counts()
meses_reales = sorted(filas_por_mes[filas_por_mes >= 0.05 * filas_por_mes.max()].index)
mes_defecto = 0 if len(meses_reales) <= 1 else opciones_mes.index(meses_reales[-1])
mes = st.sidebar.selectbox("Mes de análisis", opciones_mes, index=mes_defecto)
df_mes = df if mes == TODOS_MESES else df[df["MES"] == mes]

zonas_todas = ZONAS + (["DESCONOCIDO"] if (df[ID_ZONA] == "DESCONOCIDO").any() else [])
zonas = st.sidebar.multiselect("Zona", zonas_todas, default=zonas_todas)
loc_zona = df[[ID_LOCAL, ID_ZONA]].drop_duplicates()
locales_disp = sorted(int(x) for x in loc_zona.loc[loc_zona[ID_ZONA].isin(zonas), ID_LOCAL])
locales_sel = st.sidebar.multiselect(
    "Local", locales_disp, default=locales_disp, format_func=etiqueta
)

todas_familias = sorted(df_mes[ID_TIPO_ARTICULO].unique())
familias = st.sidebar.multiselect("Familia de producto", todas_familias, default=todas_familias)

dmin, dmax = df_mes["DIA"].min().date(), df_mes["DIA"].max().date()
rango = st.sidebar.date_input("Rango de fechas", (dmin, dmax), min_value=dmin, max_value=dmax)
if isinstance(rango, (tuple, list)) and len(rango) == 2:
    ini, fin = rango
elif isinstance(rango, (tuple, list)) and len(rango) == 1:
    ini = fin = rango[0]
else:
    ini = fin = rango

dias_sel = st.sidebar.multiselect("Días de la semana", DIAS, default=DIAS)

horas_validas = df_mes["HORA_N"].dropna()
if horas_validas.empty:
    h_ini, h_fin = 0, 23
else:
    hmin, hmax = int(horas_validas.min()), int(horas_validas.max())
    if hmin < hmax:
        h_ini, h_fin = st.sidebar.slider("Rango de horas", hmin, hmax, (hmin, hmax))
    else:
        h_ini, h_fin = hmin, hmax

st.sidebar.caption(f"📂 {fuente}")
if filas_omitidas:
    st.sidebar.warning(
        f"Se omitieron {filas_omitidas:,} filas del archivo con más o menos columnas que el encabezado."
    )
n_desc = int((df[ID_ZONA] == "DESCONOCIDO").sum())
if n_desc:
    st.sidebar.warning(f"{n_desc:,} filas son de locales fuera de los 10 de muestra (DESCONOCIDO).")

f = df_mes[
    df_mes[ID_ZONA].isin(zonas)
    & df_mes[ID_LOCAL].isin(locales_sel)
    & df_mes[ID_TIPO_ARTICULO].isin(familias)
    & df_mes["DIA"].between(pd.Timestamp(ini), pd.Timestamp(fin))
    & df_mes["DIA_SEMANA"].isin(dias_sel)
    & (df_mes["HORA_N"].between(h_ini, h_fin) | df_mes["HORA_N"].isna())
]

# ----------------------------------------------------------------------
# Encabezado y KPIs
# ----------------------------------------------------------------------
titulo_mes = meses_reales[0] if (mes == TODOS_MESES and len(meses_reales) == 1) else mes
st.title(f"🍽️ Dashboard de Ventas · {titulo_mes}")
if es_demo:
    st.info(
        "Estás viendo **datos de demostración**. Para usar los tuyos, guarda el archivo en "
        "`DataventasEnero.xlsx` (o .csv) junto a app.py, o súbelo en la barra lateral."
    )

if f.empty:
    st.warning("No hay ventas con esos filtros. Prueba ampliando zona, fechas u horas.")
    st.stop()

ventas = f[ID_TOTAL].sum()
unidades = f[ID_CANTIDAD].sum()
n_trans = len(f)
por_local = f.groupby(ID_LOCAL)[ID_TOTAL].sum()
local_top = int(por_local.idxmax())

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Ventas totales", soles(ventas))
c2.metric("Unidades vendidas", f"{unidades:,.0f}")
c3.metric("Ticket promedio", f"S/ {ventas / n_trans:,.2f}")
c4.metric("Transacciones", f"{n_trans:,}")
c5.metric("Local líder", f"Local {local_top}", etiqueta(local_top).split("· ")[1])

por_hora = f.dropna(subset=["HORA_N"]).groupby("HORA_N")[ID_TOTAL].sum()
hora_pico = f"{int(por_hora.idxmax()):02d}:00" if not por_hora.empty else "—"
por_dia_sem = f.groupby("DIA_SEMANA", observed=True)[ID_TOTAL].sum()
dia_pico = str(por_dia_sem.idxmax()) if not por_dia_sem.empty else "—"
st.caption(f"⏰ Hora pico: **{hora_pico}**  ·  📅 Día más fuerte: **{dia_pico}**")

tab_res, tab_tmp, tab_geo, tab_prod, tab_cor, tab_datos, tab_ver = st.tabs(
    ["📈 Resumen", "🕒 Temporal", "🗺️ Geográfico", "🍗 Productos", "📐 Correlación",
     "📋 Datos", "🧪 Verificación"]
)

# ----------------------------- Resumen --------------------------------
with tab_res:
    diario = f.groupby("DIA")[ID_TOTAL].sum().reset_index()
    fig = px.bar(
        diario, x="DIA", y=ID_TOTAL, title="Ventas por día",
        labels={"DIA": "Fecha", ID_TOTAL: "Ventas (S/)"}, color_discrete_sequence=["#2E9CCA"],
    )
    fig.add_scatter(
        x=diario["DIA"], y=diario[ID_TOTAL].rolling(7, min_periods=1).mean(),
        mode="lines", name="Promedio 7 días", line=dict(color="#E4572E", width=3),
    )
    mostrar(fig)

    izq, der = st.columns(2)
    with izq:
        fam = f.groupby(ID_TIPO_ARTICULO)[ID_TOTAL].sum().reset_index()
        fig = px.pie(fam, names=ID_TIPO_ARTICULO, values=ID_TOTAL, hole=0.5,
                     title="Participación por familia")
        mostrar(fig)
    with der:
        rank = por_local.reset_index()
        rank["LOCAL_TXT"] = rank[ID_LOCAL].map(etiqueta)
        rank = rank.sort_values(ID_TOTAL)
        fig = px.bar(rank, x=ID_TOTAL, y="LOCAL_TXT", orientation="h", title="Ranking de locales",
                     labels={ID_TOTAL: "Ventas (S/)", "LOCAL_TXT": ""}, color_discrete_sequence=["#29335C"])
        mostrar(fig)

# ----------------------------- Temporal -------------------------------
with tab_tmp:
    st.subheader("¿Cuándo se vende más? Día de la semana × hora")
    h = f.dropna(subset=["HORA_N"]).copy()
    if h.empty:
        st.info("No se pudo leer la hora en los datos.")
    else:
        h["HORA_N"] = h["HORA_N"].astype(int)
        piv = h.pivot_table(index="DIA_SEMANA", columns="HORA_N", values=ID_TOTAL,
                            aggfunc="sum", observed=False)
        piv.index = piv.index.astype(str)
        piv = piv.reindex(DIAS).fillna(0)
        # Con día de negocio, las horas de madrugada (0-5 h) van al final: son la noche de ese día
        orden = sorted(piv.columns, key=(lambda x: (x < 6, x)) if usa_negocio else None)
        piv = piv[orden]
        piv.columns = [f"{x}h" for x in orden]
        fig = px.imshow(piv, aspect="auto", color_continuous_scale="YlOrRd",
                        labels=dict(x="Hora del día", y="", color="Ventas (S/)"))
        mostrar(fig, 420)

    izq, der = st.columns(2)
    with izq:
        sem = f.groupby("DIA_SEMANA", observed=False)[ID_TOTAL].sum().reindex(DIAS).reset_index()
        fig = px.bar(sem, x="DIA_SEMANA", y=ID_TOTAL, title="Ventas por día de la semana",
                     labels={"DIA_SEMANA": "", ID_TOTAL: "Ventas (S/)"},
                     color_discrete_sequence=["#F3A712"])
        mostrar(fig)
    with der:
        hh = por_hora.reset_index()
        fig = px.line(hh, x="HORA_N", y=ID_TOTAL, markers=True, title="Ventas por hora",
                      labels={"HORA_N": "Hora", ID_TOTAL: "Ventas (S/)"},
                      color_discrete_sequence=["#E4572E"])
        fig.update_xaxes(dtick=1)
        mostrar(fig)

# ----------------------------- Geográfico -----------------------------
with tab_geo:
    ventas_loc = (
        f.groupby(ID_LOCAL)
        .agg(VENTAS=(ID_TOTAL, "sum"), UNIDADES=(ID_CANTIDAD, "sum"),
             TRANSACCIONES=(ID_TOTAL, "count"), TICKET=(ID_TOTAL, "mean"))
        .reset_index()
    )
    mapa = LOCALES[LOCALES[ID_LOCAL].isin(locales_sel)].merge(ventas_loc, on=ID_LOCAL, how="left")
    mapa = mapa.fillna({"VENTAS": 0, "UNIDADES": 0, "TRANSACCIONES": 0, "TICKET": 0})
    mapa["ETIQUETA"] = mapa[ID_LOCAL].map(etiqueta)
    mapa["TAMANO"] = mapa["VENTAS"].clip(lower=1)

    spread = max(mapa["LAT"].max() - mapa["LAT"].min(), mapa["LON"].max() - mapa["LON"].min())
    zoom = 10.0 if spread > 0.15 else 10.8 if spread > 0.08 else 11.8 if spread > 0.03 else 12.8
    opciones = dict(
        lat="LAT", lon="LON", size="TAMANO", color=ID_ZONA, color_discrete_map=COLOR_ZONA,
        hover_name="ETIQUETA", size_max=45, zoom=zoom,
        center=dict(lat=float(mapa["LAT"].mean()), lon=float(mapa["LON"].mean())),
        hover_data={"VENTAS": ":,.0f", "UNIDADES": ":,.0f", "TICKET": ":,.2f",
                    "LAT": False, "LON": False, "TAMANO": False, ID_ZONA: False},
    )
    st.subheader("Mapa de ventas por local")
    st.caption("El tamaño del círculo es proporcional a las ventas. Cambia la zona en la barra lateral y el mapa se reenfoca.")
    if mapa.empty:
        st.info("Los locales seleccionados no tienen coordenadas en el mapa.")
    else:
        try:
            fig = px.scatter_map(mapa, map_style="open-street-map", **opciones)
        except AttributeError:  # versiones antiguas de plotly
            fig = px.scatter_mapbox(mapa, mapbox_style="open-street-map", **opciones)
        mostrar(fig, 520)

    izq, der = st.columns(2)
    with izq:
        zona_df = f.groupby(ID_ZONA)[ID_TOTAL].sum().reset_index()
        fig = px.bar(zona_df, x=ID_ZONA, y=ID_TOTAL, color=ID_ZONA, color_discrete_map=COLOR_ZONA,
                     text_auto=",.0f", title="Ventas por zona", labels={ID_TOTAL: "Ventas (S/)"})
        fig.update_layout(showlegend=False)
        mostrar(fig)
    with der:
        tabla = mapa.sort_values("VENTAS", ascending=False)[
            [ID_LOCAL, ID_DISTRITO, ID_ZONA, "VENTAS", "UNIDADES", "TRANSACCIONES", "TICKET"]
        ]
        st.markdown("**Detalle por local**")
        st.dataframe(
            tabla, hide_index=True, width="stretch",
            column_config={
                "VENTAS": st.column_config.NumberColumn("Ventas (S/)", format="%.0f"),
                "UNIDADES": st.column_config.NumberColumn("Unidades", format="%.0f"),
                "TRANSACCIONES": st.column_config.NumberColumn("Transacciones", format="%d"),
                "TICKET": st.column_config.NumberColumn("Ticket prom.", format="%.2f"),
            },
        )

# ----------------------------- Productos ------------------------------
with tab_prod:
    medida = st.radio("Medir por", ["Ventas (S/)", "Unidades"], horizontal=True)
    col_m = ID_TOTAL if medida.startswith("Ventas") else ID_CANTIDAD

    prod = f.groupby([ID_TIPO_ARTICULO, ID_ARTICULO])[col_m].sum().reset_index()
    prod = prod[prod[col_m] > 0]
    # Familia > sub-familia > producto (la sub-familia solo si el archivo la trae)
    ruta_tree = [ID_TIPO_ARTICULO] + ([ID_SUBFAMILIA] if ID_SUBFAMILIA in f.columns else []) + [ID_ARTICULO]
    prod_tree = f.groupby(ruta_tree)[col_m].sum().reset_index()
    prod_tree = prod_tree[prod_tree[col_m] > 0]
    fig = px.treemap(
        prod_tree, path=ruta_tree, values=col_m,
        color=ID_TIPO_ARTICULO, color_discrete_sequence=px.colors.qualitative.Set2,
        title=f"Familias y productos ({medida}) · el tamaño indica cuánto aporta cada producto",
    )
    fig.update_traces(
        texttemplate="<b>%{label}</b><br>%{value:,.0f}<br>%{percentRoot:.1%}",
        hovertemplate=(
            "<b>%{label}</b><br>" + medida + ": %{value:,.0f}"
            "<br>%{percentParent:.1%} de su familia"
            "<br>%{percentRoot:.1%} del total<extra></extra>"
        ),
        textfont=dict(color="#1b1b1b"),
        marker=dict(line=dict(width=2)),
        root_color="rgba(0,0,0,0)",
    )
    fig.update_layout(uniformtext=dict(minsize=11, mode="hide"))
    mostrar(fig, 460)

    izq, der = st.columns(2)
    with izq:
        top = prod.nlargest(10, col_m).sort_values(col_m)
        fig = px.bar(top, x=col_m, y=ID_ARTICULO, orientation="h", color=ID_TIPO_ARTICULO,
                     title="Top 10 productos", labels={col_m: medida, ID_ARTICULO: ""})
        mostrar(fig, 420)
    with der:
        nivel = st.radio("Agrupar por", ["Zona", "Distrito"], horizontal=True)
        eje = ID_ZONA if nivel == "Zona" else ID_DISTRITO
        comb = f.groupby([eje, ID_TIPO_ARTICULO])[col_m].sum().reset_index()
        fig = px.bar(comb, x=eje, y=col_m, color=ID_TIPO_ARTICULO, barmode="stack",
                     title=f"{medida} por {nivel.lower()} y familia", labels={col_m: medida})
        mostrar(fig, 420)

# ----------------------------- Correlación ----------------------------
with tab_cor:
    st.subheader("Correlación: Pearson, Spearman y Kendall")
    st.caption(
        "Mide si dos variables se mueven juntas: +1 suben juntas, −1 una sube y la otra baja, "
        "0 sin relación. Usa los filtros de la barra lateral."
    )
    nivel_cor = st.radio(
        "¿Qué se compara?", NIVELES_CORR, horizontal=True,
        help=(
            "Local × día: cada punto es un local en un día. Producto: cada punto es un producto. "
            "Línea de venta: cada punto es una fila de la base."
        ),
    )
    tabla_cor = tabla_correlacion(f, nivel_cor)
    cols_cor = list(tabla_cor.columns)
    x_def, y_def = DEFECTO_CORR[nivel_cor]
    cx, cy = st.columns(2)
    x_var = cx.selectbox("Variable X", cols_cor, key=f"cor_x_{nivel_cor}",
                         index=cols_cor.index(x_def) if x_def in cols_cor else 0)
    y_var = cy.selectbox("Variable Y", cols_cor, key=f"cor_y_{nivel_cor}",
                         index=cols_cor.index(y_def) if y_def in cols_cor else min(1, len(cols_cor) - 1))

    if x_var == y_var:
        st.info("Elige dos variables distintas.")
    else:
        d_cor, res_cor = correlacionar(tabla_cor[x_var], tabla_cor[y_var])
        if res_cor is None:
            st.warning("No hay datos suficientes, o una de las variables no varía con los filtros actuales.")
        else:
            cols_m = st.columns(3)
            for col, fila in zip(cols_m, res_cor.to_dict("records")):
                col.metric(fila["Prueba"], f"{fila['Coeficiente']:+.3f}",
                           help=METODOS_CORR[fila["Prueba"]][1])
                col.caption(f"{fila['Lectura']} · p = {fila['Valor p']}")
            st.caption(f"{len(d_cor):,} puntos con datos en {x_var} y {y_var}.")

            coef = res_cor.set_index("Prueba")["Coeficiente"]
            if abs(coef["Pearson"] - coef["Spearman"]) > 0.2:
                st.info(
                    "Pearson y Spearman difieren bastante: la relación puede no ser lineal o haber "
                    "valores extremos que distorsionan a Pearson. En ese caso, fíate más de Spearman y Kendall."
                )

            muestra = d_cor if len(d_cor) <= 5000 else d_cor.sample(5000, random_state=0)
            sufijo_muestra = "" if len(d_cor) <= 5000 else " (se dibujan 5,000 puntos al azar)"
            fig = px.scatter(muestra, x="x", y="y", opacity=0.6,
                             labels={"x": x_var, "y": y_var},
                             title=f"{y_var} vs {x_var}{sufijo_muestra}",
                             color_discrete_sequence=["#2E9CCA"])
            pendiente, origen_y = np.polyfit(d_cor["x"], d_cor["y"], 1)
            xs = np.array([d_cor["x"].min(), d_cor["x"].max()])
            fig.add_scatter(x=xs, y=pendiente * xs + origen_y, mode="lines", name="Tendencia lineal",
                            line=dict(color="#E4572E", width=3))
            mostrar(fig, 420)

            st.dataframe(
                res_cor, hide_index=True, width="stretch",
                column_config={"Coeficiente": st.column_config.NumberColumn(format="%.3f")},
            )

    with st.expander("Matriz de correlación de todas las variables", expanded=False):
        if len(cols_cor) < 2:
            st.info("Se necesitan al menos dos variables.")
        else:
            for col, metodo in zip(st.columns(3), METODOS_CORR):
                mat = tabla_cor.corr(method=metodo.lower())
                fig = px.imshow(mat, zmin=-1, zmax=1, color_continuous_scale="RdBu_r",
                                text_auto=".2f", aspect="auto", title=metodo)
                fig.update_layout(coloraxis_showscale=False)
                with col:
                    mostrar(fig, 340)

    st.caption(
        "Ojo: que dos variables se muevan juntas no significa que una cause la otra. "
        "Un valor p menor a 0.05 indica que la relación difícilmente es casualidad, "
        "pero no que sea importante: mira también el tamaño del coeficiente."
    )

# ----------------------------- Datos ----------------------------------
with tab_datos:
    columnas = [ID_FECHA, "HORA_N", ID_LOCAL, ID_DISTRITO, ID_ZONA, ID_TIPO_ARTICULO]
    if ID_SUBFAMILIA in f.columns:
        columnas.append(ID_SUBFAMILIA)
    columnas += [ID_ARTICULO, ID_CANTIDAD, ID_TOTAL]
    st.caption(f"{len(f):,} filas con los filtros actuales (se muestran las primeras 500).")
    st.dataframe(f[columnas].head(500), hide_index=True, width="stretch")
    st.download_button(
        "⬇️ Descargar datos filtrados (CSV)",
        f[columnas].to_csv(index=False).encode("utf-8-sig"),
        file_name=f"ventas_{mes.replace(' ', '_')}.csv",
        mime="text/csv",
    )

# ----------------------------- Verificación ---------------------------
with tab_ver:
    st.subheader("Mismos cálculos que el notebook (Tarea 5)")
    if len(f) == len(df):
        st.success("Sin filtros: estas tablas deben coincidir con las que imprime el notebook.")
    else:
        st.warning(
            "Hay filtros activos. Para comparar con el notebook elige "
            f"«{TODOS_MESES}» y deja todas las zonas, locales, familias, fechas, días y horas."
        )

    v1, v2, v3 = st.columns(3)
    v1.metric("Filas", f"{len(f):,}")
    v2.metric("Ventas totales (S/)", f"{ventas:,.2f}")
    v3.metric("Unidades", f"{unidades:,.0f}")

    def tabla(t, con_indice=False):
        st.dataframe(t.round(2), hide_index=not con_indice, width="stretch")

    st.markdown("**Agregación temporal · por día de la semana**")
    t_dias = (
        f.groupby("DIA_SEMANA", observed=True)
        .agg(VENTAS_TOTALES=(ID_TOTAL, "sum"), TRANSACCIONES=(ID_TOTAL, "count"))
        .reset_index()
    )
    t_dias["DIA_SEMANA"] = t_dias["DIA_SEMANA"].astype(str).map(lambda d: f"{d} ({DIA_EN[d]})")
    tabla(t_dias)

    st.markdown("**Agregación temporal · por hora**")
    t_horas = (
        f.dropna(subset=["HORA_N"])
        .assign(HORA=lambda d: d["HORA_N"].astype(int))
        .groupby("HORA")
        .agg(VENTAS_TOTALES=(ID_TOTAL, "sum"), TRANSACCIONES=(ID_TOTAL, "count"))
        .reset_index()
    )
    tabla(t_horas)

    st.markdown("**Agregación espacial · por zona**")
    t_zona = (
        f.groupby(ID_ZONA)
        .agg(VENTAS_TOTALES=(ID_TOTAL, "sum"), UNIDADES_VENDIDAS=(ID_CANTIDAD, "sum"),
             VENTA_PROMEDIO=(ID_TOTAL, "mean"))
        .reset_index()
        .sort_values("VENTAS_TOTALES", ascending=False)
    )
    tabla(t_zona)

    st.markdown("**Agregación espacial · por distrito**")
    t_dist = (
        f.groupby(ID_DISTRITO)
        .agg(VENTAS_TOTALES=(ID_TOTAL, "sum"), UNIDADES_VENDIDAS=(ID_CANTIDAD, "sum"),
             VENTA_PROMEDIO=(ID_TOTAL, "mean"))
        .reset_index()
        .sort_values("VENTAS_TOTALES", ascending=False)
    )
    tabla(t_dist)

    st.markdown("**Agregación por atributo · familia de producto**")
    t_fam = (
        f.groupby(ID_TIPO_ARTICULO)
        .agg(VENTAS_TOTALES=(ID_TOTAL, "sum"), UNIDADES_TOTALES=(ID_CANTIDAD, "sum"))
        .reset_index()
        .sort_values("VENTAS_TOTALES", ascending=False)
        .rename(columns={ID_TIPO_ARTICULO: "FAMILIA"})
    )
    t_fam["PARTICIPACION_PCT"] = t_fam["VENTAS_TOTALES"] / t_fam["VENTAS_TOTALES"].sum() * 100
    tabla(t_fam)

    st.markdown("**Agregación combinada · zona × familia (ventas)**")
    t_comb = f.groupby([ID_ZONA, ID_TIPO_ARTICULO])[ID_TOTAL].sum().unstack().fillna(0)
    tabla(t_comb, con_indice=True)
