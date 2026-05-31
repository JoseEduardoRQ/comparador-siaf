import io
import re
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st


ASSETS_DIR = Path(__file__).parent / "assets"
LOGO_PRINCIPAL_PATH = ASSETS_DIR / "logo_principal_horizontal.png"
LOGO_CUADRADO_PATH = ASSETS_DIR / "logo_cuadrado.png"
LOGO_BLANCO_PATH = ASSETS_DIR / "logo_blanco.png"
LOGO_AZUL_CORPORATIVO_PATH = ASSETS_DIR / "logo_azul_corporativo.png"
COLUMNAS_COMPARACION = ["expediente siaf", "registro siaf", "siaf", "expediente"]
COLUMNAS_REPORTE = {
    "Expediente SIAF": ["expediente siaf", "expediente", "registro siaf", "siaf"],
    "Fase": ["fase"],
    "RB": ["rb"],
    "Cod. Doc.": ["cod doc", "codigo doc", "código doc", "cod documento", "codigo documento"],
    "Num. Doc.": ["num doc", "numero doc", "número doc", "num documento", "numero documento"],
    "Fecha Doc.": ["fecha doc", "fecha documento"],
    "Proveedor": ["proveedor", "ruc", "ruc proveedor"],
    "Nombre Proveedor": ["nombre proveedor", "nombre", "razon social", "razón social", "proveedor nombre"],
    "Clasificación": ["clasificacion", "clasificación"],
    "Mon.": ["mon", "moneda"],
    "Monto Origen": ["monto origen", "importe origen"],
    "Monto S/.": ["monto s", "monto s/", "monto soles", "importe", "monto", "importe s"],
    "Fecha Aprobación": ["fecha aprobacion", "fecha aprobación", "fecha aprobado"],
}


def cargar_excel(archivo):
    """Carga un archivo Excel sin modificar los datos originales."""
    try:
        dataframe = pd.read_excel(archivo, engine="openpyxl")
        dataframe.columns = dataframe.columns.astype(str).str.strip()
        return dataframe
    except Exception as error:
        st.error(f"No se pudo leer el archivo '{archivo.name}'. Verifica que sea un Excel valido.")
        st.caption(f"Detalle tecnico: {error}")
        return None


def normalizar_columna(nombre):
    """Normaliza encabezados para reconocerlos aunque vengan con saltos o simbolos."""
    texto = str(nombre)
    texto = texto.replace("\n", " ").replace("\r", " ").replace("\t", " ")
    texto = texto.replace("\u00a0", " ").replace("\u200b", "")
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(caracter for caracter in texto if not unicodedata.combining(caracter))
    texto = texto.lower().strip()
    texto = texto.replace(".", "")
    texto = re.sub(r"\s+", " ", texto)
    return texto


def limpiar_columna(serie):
    """Convierte el SIAF a texto y elimina espacios, vacios y decimales .0."""
    return (
        serie.fillna("")
        .astype(str)
        .str.strip()
        .str.replace("\u00a0", " ", regex=False)
        .str.replace("\u200b", "", regex=False)
        .str.replace(r"\s+", "", regex=True)
        .str.replace(r"\.0+$", "", regex=True)
    )


def convertir_monto(valor):
    """Convierte montos con coma o punto decimal a numero."""
    if pd.isna(valor):
        return 0.0

    if isinstance(valor, (int, float)):
        return float(valor)

    texto = str(valor).strip()
    if not texto:
        return 0.0

    texto = texto.replace("\u00a0", "").replace(" ", "")
    texto = re.sub(r"[^0-9,\.\-]", "", texto)

    if not texto or texto in {"-", ".", ","}:
        return 0.0

    if "," in texto and "." in texto:
        if texto.rfind(",") > texto.rfind("."):
            texto = texto.replace(".", "").replace(",", ".")
        else:
            texto = texto.replace(",", "")
    elif "," in texto:
        parte_decimal = texto.rsplit(",", 1)[-1]
        texto = texto.replace(",", ".") if len(parte_decimal) <= 2 else texto.replace(",", "")
    elif "." in texto:
        parte_decimal = texto.rsplit(".", 1)[-1]
        if len(parte_decimal) == 3 and texto.count(".") >= 1:
            texto = texto.replace(".", "")

    try:
        return float(texto)
    except ValueError:
        return 0.0


def buscar_columna(dataframe, opciones):
    """Busca una columna por coincidencia exacta o parecida usando nombres normalizados."""
    mapa_columnas = {
        normalizar_columna(columna): columna
        for columna in dataframe.columns
    }
    opciones_normalizadas = [normalizar_columna(opcion) for opcion in opciones]

    for opcion in opciones_normalizadas:
        if opcion in mapa_columnas:
            return mapa_columnas[opcion]

    alias_generales = {"monto", "importe", "siaf", "ruc", "rb", "mon", "fase", "nombre", "proveedor", "expediente"}

    for opcion in opciones_normalizadas:
        if opcion in alias_generales:
            continue
        for columna_normalizada, columna_original in mapa_columnas.items():
            if opcion and (opcion in columna_normalizada or columna_normalizada in opcion):
                return columna_original

    return None


def columna_siaf(dataframe):
    """Encuentra la columna que representa el Expediente SIAF."""
    return buscar_columna(dataframe, COLUMNAS_COMPARACION)


def columnas_detectadas(dataframe):
    """Devuelve columnas originales y normalizadas para mostrarlas en la app."""
    return pd.DataFrame(
        {
            "Columna leida": list(dataframe.columns),
            "Nombre normalizado": [
                normalizar_columna(columna) for columna in dataframe.columns
            ],
        }
    )


def validar_columna_siaf(dataframes):
    """Revisa que todos los archivos tengan una columna compatible con SIAF."""
    faltantes = [
        nombre
        for nombre, dataframe in dataframes.items()
        if columna_siaf(dataframe) is None
    ]

    if faltantes:
        st.error(
            "No se encontro una columna SIAF en: "
            + ", ".join(faltantes)
            + ". La app busca: expediente siaf, registro siaf, siaf o expediente."
        )
        return False

    return True


def comparar_archivos(df_base, df_comparacion):
    """Quita del archivo base los registros cuyo SIAF aparece en el otro archivo."""
    columna_base = columna_siaf(df_base)
    columna_comparacion = columna_siaf(df_comparacion)

    base_limpia = limpiar_columna(df_base[columna_base])
    comparacion_limpia = limpiar_columna(df_comparacion[columna_comparacion])

    valores_comparacion = set(comparacion_limpia[comparacion_limpia != ""])
    tiene_siaf = base_limpia != ""
    existe_en_comparacion = base_limpia.isin(valores_comparacion) & tiene_siaf

    pendientes = df_base.loc[tiene_siaf & ~existe_en_comparacion].copy()
    eliminados = int(existe_en_comparacion.sum())

    return pendientes, eliminados


def obtener_registros_eliminados(df_base, df_comparacion):
    """Obtiene solo las filas eliminadas para mostrarlas en reportes."""
    columna_base = columna_siaf(df_base)
    columna_comparacion = columna_siaf(df_comparacion)

    base_limpia = limpiar_columna(df_base[columna_base])
    comparacion_limpia = limpiar_columna(df_comparacion[columna_comparacion])
    valores_comparacion = set(comparacion_limpia[comparacion_limpia != ""])
    eliminados = base_limpia.isin(valores_comparacion) & (base_limpia != "")

    return df_base.loc[eliminados].copy()


def obtener_columna(dataframe, opciones, limpiar_siaf=False):
    """Obtiene una columna por varios nombres posibles o crea una columna vacia."""
    columna = buscar_columna(dataframe, opciones)
    if not columna:
        return pd.Series([""] * len(dataframe), index=dataframe.index)

    if limpiar_siaf:
        return limpiar_columna(dataframe[columna])

    return dataframe[columna]


def preparar_resultado(dataframe):
    """Ordena el resultado con las columnas finales solicitadas."""
    resultado = pd.DataFrame(index=dataframe.index)

    for columna_salida, alias in COLUMNAS_REPORTE.items():
        if columna_salida == "Expediente SIAF":
            resultado[columna_salida] = obtener_columna(dataframe, alias, limpiar_siaf=True)
        elif columna_salida in ["Monto Origen", "Monto S/."]:
            resultado[columna_salida] = obtener_columna(dataframe, alias).apply(convertir_monto)
        else:
            resultado[columna_salida] = obtener_columna(dataframe, alias)

    return resultado[list(COLUMNAS_REPORTE.keys())].reset_index(drop=True)


def calcular_total_monto(dataframe):
    """Suma Monto S/. y usa Monto Origen cuando Monto S/. esta vacio o en cero."""
    monto_soles = dataframe["Monto S/."].apply(convertir_monto)
    monto_origen = dataframe["Monto Origen"].apply(convertir_monto)
    monto_para_sumar = monto_soles.where(monto_soles != 0, monto_origen)
    return float(monto_para_sumar.sum())


def agregar_fila_total(dataframe):
    """Agrega una fila final TOTAL con la suma de Monto S/. o Monto Origen."""
    resultado = dataframe.copy()
    fila_total = {columna: "" for columna in resultado.columns}
    fila_total["Expediente SIAF"] = "TOTAL"
    fila_total["Monto S/."] = calcular_total_monto(resultado)

    return pd.concat([resultado, pd.DataFrame([fila_total])], ignore_index=True)


def crear_resumen_reporte(datos_resumen, fecha_generacion):
    """Crea los datos de la hoja Resumen."""
    return pd.DataFrame(
        [
            ["REPORTE DE COMPARACI\u00d3N SIAF", ""],
            ["Fecha de generacion", fecha_generacion.strftime("%Y-%m-%d")],
            ["Hora de generacion", fecha_generacion.strftime("%H:%M:%S")],
            ["Generado por", "Jose Rubina Quijano"],
            ["Especialidad", "Systems Engineering"],
            ["Sistema", "Sistema de Comparacion SIAF"],
            ["", ""],
            ["Devengados originales", datos_resumen["devengados_originales"]],
            ["Girados originales", datos_resumen["girados_originales"]],
            ["Pagados originales", datos_resumen["pagados_originales"]],
            ["Devengados pendientes", datos_resumen["devengados_pendientes"]],
            ["Girados pendientes", datos_resumen["girados_pendientes"]],
            ["Devengados eliminados", datos_resumen["devengados_eliminados"]],
            ["Girados eliminados", datos_resumen["girados_eliminados"]],
            ["Monto total Devengados pendientes", datos_resumen["monto_devengados"]],
            ["Monto total Girados pendientes", datos_resumen["monto_girados"]],
        ],
        columns=["Concepto", "Valor"],
    )


def crear_registros_eliminados(devengados_eliminados, girados_eliminados):
    """Crea una hoja con secciones de registros eliminados."""
    columnas = list(COLUMNAS_REPORTE.keys())
    filas = []

    filas.append(["A) Devengados eliminados porque ya estaban en Girados"] + [""] * (len(columnas) - 1))
    filas.append(columnas)
    filas.extend(devengados_eliminados.fillna("").values.tolist())
    filas.append([""] * len(columnas))
    filas.append(["B) Girados eliminados porque ya estaban en Pagados"] + [""] * (len(columnas) - 1))
    filas.append(columnas)
    filas.extend(girados_eliminados.fillna("").values.tolist())

    return pd.DataFrame(filas, columns=columnas)


def nombre_reporte_unico(fecha_generacion):
    """Genera un nombre unico para el reporte."""
    return f"Resultado_Comparacion_{fecha_generacion.strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"


def guardar_reporte_historico(archivo_excel, nombre_archivo):
    """Guarda automaticamente una copia del reporte en la carpeta reportes."""
    carpeta_reportes = Path(__file__).parent / "reportes"
    carpeta_reportes.mkdir(exist_ok=True)
    ruta_reporte = carpeta_reportes / nombre_archivo

    contador = 1
    while ruta_reporte.exists():
        base = ruta_reporte.stem
        ruta_reporte = carpeta_reportes / f"{base}_{contador}{ruta_reporte.suffix}"
        contador += 1

    ruta_reporte.write_bytes(archivo_excel)
    return ruta_reporte


def convertir_a_excel(hojas, datos_resumen=None, registros_eliminados=None, fecha_generacion=None):
    """Convierte las hojas a un archivo Excel con fila TOTAL al final."""
    salida = io.BytesIO()
    fecha_generacion = fecha_generacion or datetime.now()

    with pd.ExcelWriter(salida, engine="xlsxwriter") as writer:
        for nombre_hoja, dataframe in hojas.items():
            dataframe_excel = agregar_fila_total(dataframe)
            dataframe_excel.to_excel(writer, index=False, sheet_name=nombre_hoja)

            workbook = writer.book
            worksheet = writer.sheets[nombre_hoja]
            formato_encabezado = workbook.add_format(
                {
                    "bold": True,
                    "font_color": "white",
                    "bg_color": "#1F4E79",
                    "border": 1,
                }
            )
            formato_total = workbook.add_format(
                {
                    "bold": True,
                    "bg_color": "#EAF2F8",
                    "border": 1,
                }
            )
            formato_monto = workbook.add_format({"num_format": "#,##0.00"})

            for numero_columna, valor in enumerate(dataframe_excel.columns):
                worksheet.write(0, numero_columna, valor, formato_encabezado)
                ancho = max(12, min(35, len(str(valor)) + 6))
                worksheet.set_column(numero_columna, numero_columna, ancho)

            for columna_monto in ["Monto Origen", "Monto S/."]:
                indice = dataframe_excel.columns.get_loc(columna_monto)
                worksheet.set_column(indice, indice, 15, formato_monto)

            fila_total_excel = len(dataframe_excel)
            worksheet.set_row(fila_total_excel, None, formato_total)

        workbook = writer.book

        if datos_resumen:
            resumen = crear_resumen_reporte(datos_resumen, fecha_generacion)
            resumen.to_excel(writer, index=False, sheet_name="Resumen")
            worksheet = writer.sheets["Resumen"]
            formato_titulo = workbook.add_format(
                {
                    "bold": True,
                    "font_size": 16,
                    "font_color": "white",
                    "bg_color": "#0F4C81",
                    "align": "center",
                }
            )
            formato_encabezado = workbook.add_format(
                {"bold": True, "font_color": "white", "bg_color": "#1B6B93", "border": 1}
            )
            formato_monto = workbook.add_format({"num_format": "#,##0.00"})
            worksheet.merge_range("A1:B1", "REPORTE DE COMPARACI\u00d3N SIAF", formato_titulo)
            worksheet.write(4, 0, "Concepto", formato_encabezado)
            worksheet.write(4, 1, "Valor", formato_encabezado)
            worksheet.set_column(0, 0, 38)
            worksheet.set_column(1, 1, 22, formato_monto)

        if registros_eliminados is not None:
            registros_eliminados.to_excel(writer, index=False, sheet_name="Registros Eliminados")
            worksheet = writer.sheets["Registros Eliminados"]
            formato_seccion = workbook.add_format(
                {"bold": True, "font_color": "white", "bg_color": "#0F4C81", "border": 1}
            )
            formato_encabezado = workbook.add_format(
                {"bold": True, "font_color": "white", "bg_color": "#1B6B93", "border": 1}
            )
            for fila_idx, valor in enumerate(registros_eliminados.iloc[:, 0].astype(str), start=1):
                if valor.startswith("A)") or valor.startswith("B)"):
                    worksheet.set_row(fila_idx, None, formato_seccion)
                elif valor == "Expediente SIAF":
                    worksheet.set_row(fila_idx, None, formato_encabezado)
            for numero_columna, valor in enumerate(registros_eliminados.columns):
                ancho = max(12, min(32, len(str(valor)) + 6))
                worksheet.set_column(numero_columna, numero_columna, ancho)

    salida.seek(0)
    return salida.getvalue()


def filtrar_tabla(dataframe, texto_busqueda):
    """Busca texto dentro de cualquier columna de una tabla."""
    if not texto_busqueda:
        return dataframe

    texto = texto_busqueda.strip().lower()
    tabla_texto = dataframe.fillna("").astype(str)
    coincidencias = tabla_texto.apply(
        lambda columna: columna.str.lower().str.contains(texto, na=False, regex=False)
    ).any(axis=1)

    return dataframe.loc[coincidencias]


def mostrar_tarjeta(titulo, valor, es_monto=False):
    """Muestra una tarjeta simple con un indicador."""
    texto_valor = f"S/ {valor:,.2f}" if es_monto else f"{valor:,}"
    st.markdown(
        f"""
        <div class="metric-card">
            <p>{titulo}</p>
            <strong>{texto_valor}</strong>
        </div>
        """,
        unsafe_allow_html=True,
    )


def logo_data_uri(ruta_logo):
    """Devuelve el logo en formato embebido para la interfaz."""
    if not ruta_logo.exists():
        return ""
    import base64

    contenido = base64.b64encode(ruta_logo.read_bytes()).decode("utf-8")
    return f"data:image/png;base64,{contenido}"


def html_logo(clase, ruta_logo=LOGO_PRINCIPAL_PATH, alt="Jose Rubina Quijano Systems Engineering"):
    """Construye el HTML del logo si existe."""
    origen = logo_data_uri(ruta_logo)
    if not origen:
        return ""
    return f'<img class="{clase}" src="{origen}" alt="{alt}">'


def mostrar_encabezado_principal():
    """Muestra el encabezado principal con logo."""
    st.markdown(
        f"""
        <div class="app-header">
            <div class="brand-header">
                {html_logo("brand-logo-horizontal", LOGO_PRINCIPAL_PATH)}
                <div>
                    <h1>Sistema de Comparaci\u00f3n SIAF</h1>
                    <p>Devengados, Girados y Pagados</p>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def mostrar_footer():
    """Muestra el footer de identidad visual."""
    st.markdown(
        """
        <div class="footer-brand">
            <strong>Sistema de Comparaci\u00f3n SIAF v1.0</strong><br>
            Desarrollado por Jos\u00e9 Rubina Quijano<br>
            Systems Engineering
        </div>
        """,
        unsafe_allow_html=True,
    )


def obtener_credenciales():
    """Lee credenciales desde st.secrets."""
    try:
        usuario = st.secrets["APP_USER"]
        contrasena = st.secrets["APP_PASSWORD"]
    except Exception:
        return None, None

    return usuario, contrasena


def aplicar_estilos():
    """Aplica estilos visuales de la interfaz."""
    st.markdown(
        """
        <style>
        :root {
            color-scheme: light;
        }
        html,
        body,
        .stApp,
        [data-testid="stAppViewContainer"] {
            background: #F4F7FA !important;
            color: #1F2937 !important;
        }
        .main .block-container {
            padding-top: 1.25rem;
            padding-bottom: 2rem;
            max-width: 1260px;
        }
        .stApp {
            background: #F4F7FA !important;
        }
        [data-testid="stSidebar"] {
            background: #0B2545 !important;
        }
        [data-testid="stSidebar"] label,
        [data-testid="stSidebar"] p,
        [data-testid="stSidebar"] h1,
        [data-testid="stSidebar"] h2,
        [data-testid="stSidebar"] h3 {
            color: #F8FAFC !important;
        }
        [data-testid="stSidebar"] span,
        [data-testid="stSidebar"] small,
        [data-testid="stSidebar"] div {
            color: inherit;
        }
        .stButton > button,
        [data-testid="stFileUploader"] button {
            background: #0F4C81 !important;
            color: #FFFFFF !important;
            border: 1px solid #0F4C81 !important;
            border-radius: 8px !important;
            font-weight: 700 !important;
            box-shadow: 0 3px 10px rgba(15, 76, 129, 0.18);
        }
        .stButton > button:hover,
        [data-testid="stFileUploader"] button:hover {
            background: #1B6B93 !important;
            color: #FFFFFF !important;
            border-color: #1B6B93 !important;
        }
        .stDownloadButton > button {
            background: #16A34A !important;
            color: #FFFFFF !important;
            border: 1px solid #16A34A !important;
            border-radius: 8px !important;
            font-weight: 700 !important;
            box-shadow: 0 3px 10px rgba(22, 163, 74, 0.18);
        }
        .stDownloadButton > button:hover {
            background: #15803D !important;
            color: #FFFFFF !important;
            border-color: #15803D !important;
        }
        [data-testid="stSidebar"] .stButton button[kind="primary"],
        [data-testid="stSidebar"] [data-testid="stBaseButton-primary"] {
            background: #DC2626 !important;
            color: #FFFFFF !important;
            border: 1px solid #DC2626 !important;
        }
        [data-testid="stSidebar"] .stButton button[kind="primary"]:hover,
        [data-testid="stSidebar"] [data-testid="stBaseButton-primary"]:hover {
            background: #B91C1C !important;
            color: #FFFFFF !important;
            border-color: #B91C1C !important;
        }
        [data-testid="stSidebar"] .stButton button[kind="secondary"],
        [data-testid="stSidebar"] [data-testid="stBaseButton-secondary"] {
            background: #F59E0B !important;
            color: #FFFFFF !important;
            border: 1px solid #F59E0B !important;
        }
        [data-testid="stSidebar"] .stButton button[kind="secondary"]:hover,
        [data-testid="stSidebar"] [data-testid="stBaseButton-secondary"]:hover {
            background: #D97706 !important;
            color: #FFFFFF !important;
            border-color: #D97706 !important;
        }
        [data-testid="stFileUploader"] {
            background: #FFFFFF !important;
            border: 1px solid #E5E7EB !important;
            border-radius: 8px;
            padding: .55rem;
            box-shadow: 0 4px 12px rgba(15, 23, 42, 0.08);
        }
        [data-testid="stFileUploader"] *,
        [data-testid="stFileUploader"] section *,
        [data-testid="stFileUploaderDropzone"] *,
        [data-testid="stFileUploaderFile"] * {
            color: #132A43 !important;
        }
        [data-testid="stFileUploader"] button,
        [data-testid="stFileUploader"] button *,
        [data-testid="stSidebar"] [data-testid="stFileUploader"] button,
        [data-testid="stSidebar"] [data-testid="stFileUploader"] button *,
        [data-testid="stSidebar"] [data-testid="stFileUploader"] [data-testid="stBaseButton-secondary"],
        [data-testid="stSidebar"] [data-testid="stFileUploader"] [data-testid="stBaseButton-secondary"] * {
            background: #0F4C81 !important;
            color: #FFFFFF !important;
            border-color: #0F4C81 !important;
        }
        [data-testid="stFileUploader"] button:hover,
        [data-testid="stSidebar"] [data-testid="stFileUploader"] button:hover,
        [data-testid="stSidebar"] [data-testid="stFileUploader"] [data-testid="stBaseButton-secondary"]:hover {
            background: #1B6B93 !important;
            color: #FFFFFF !important;
            border-color: #1B6B93 !important;
        }
        [data-testid="stFileUploader"] small,
        [data-testid="stFileUploader"] [data-testid="stMarkdownContainer"] p {
            color: #334155 !important;
        }
        .stTextInput input {
            background: #FFFFFF !important;
            color: #1F2937 !important;
            border: 1px solid #E5E7EB !important;
            border-radius: 8px !important;
        }
        .stTextInput label {
            color: #1F2937 !important;
            font-weight: 700 !important;
        }
        [data-testid="stSidebar"] .stTextInput label,
        [data-testid="stSidebar"] .stTextInput p {
            color: #F8FAFC !important;
        }
        [data-testid="stAlert"] {
            border-radius: 8px !important;
            border-width: 1px !important;
            box-shadow: none !important;
        }
        [data-testid="stAlert"] * {
            font-weight: 600 !important;
        }
        [data-testid="stAlert"][kind="success"],
        [data-testid="stAlert"][data-baseweb="notification"][kind="success"] {
            background: #DCFCE7 !important;
            border-color: #86EFAC !important;
            color: #166534 !important;
        }
        [data-testid="stAlert"][kind="warning"],
        [data-testid="stAlert"][data-baseweb="notification"][kind="warning"] {
            background: #FEF3C7 !important;
            border-color: #FCD34D !important;
            color: #92400E !important;
        }
        [data-testid="stAlert"][kind="error"],
        [data-testid="stAlert"][data-baseweb="notification"][kind="error"] {
            background: #FEE2E2 !important;
            border-color: #FCA5A5 !important;
            color: #991B1B !important;
        }
        [data-testid="stAlert"][kind="info"],
        [data-testid="stAlert"][data-baseweb="notification"][kind="info"] {
            background: #DBEAFE !important;
            border-color: #93C5FD !important;
            color: #1E40AF !important;
        }
        [data-testid="stAlert"][kind="success"] *,
        [data-testid="stAlert"][kind="warning"] *,
        [data-testid="stAlert"][kind="error"] *,
        [data-testid="stAlert"][kind="info"] * {
            color: inherit !important;
        }
        .app-header {
            background: linear-gradient(135deg, #0B2545 0%, #1B6B93 100%);
            border-radius: 8px;
            color: white;
            padding: 1.85rem 2rem;
            margin-bottom: 1.25rem;
            box-shadow: 0 10px 24px rgba(11, 37, 69, 0.18);
        }
        .app-header h1 {
            color: white;
            font-size: 2.15rem;
            margin: 0 0 .3rem 0;
            letter-spacing: 0;
        }
        .app-header p {
            color: #F4F7FA;
            font-size: 1.05rem;
            margin: 0;
        }
        .brand-header {
            display: flex;
            align-items: center;
            gap: 1rem;
        }
        .brand-logo-horizontal {
            width: min(260px, 34vw);
            max-height: 86px;
            object-fit: contain;
            border-radius: 8px;
            box-shadow: 0 8px 18px rgba(0, 0, 0, 0.18);
            flex: 0 0 auto;
        }
        .brand-login-logo {
            width: min(360px, 88%);
            max-height: 170px;
            object-fit: contain;
            border-radius: 8px;
            display: block;
            margin: 0 auto 1rem auto;
            box-shadow: 0 8px 22px rgba(15, 23, 42, 0.14);
        }
        .sidebar-brand {
            background: #FFFFFF;
            border: 1px solid #E5E7EB;
            border-radius: 8px;
            padding: .85rem;
            margin-bottom: 1rem;
            text-align: center;
            box-shadow: 0 4px 12px rgba(15, 23, 42, 0.10);
        }
        .sidebar-brand img {
            width: 92px;
            height: 92px;
            object-fit: contain;
            border-radius: 8px;
            margin-bottom: .55rem;
        }
        .sidebar-brand p {
            color: #1F2937 !important;
            margin: .1rem 0 !important;
            font-weight: 700;
        }
        .sidebar-brand span {
            color: #475569 !important;
            display: block;
            font-size: .86rem;
        }
        .footer-brand {
            color: #475569;
            text-align: center;
            padding: 1.25rem 0 .5rem 0;
            font-size: .92rem;
        }
        .footer-brand strong {
            color: #1F2937;
        }
        @media (max-width: 720px) {
            .brand-header {
                align-items: flex-start;
            }
            .brand-header {
                flex-direction: column;
                gap: .7rem;
            }
            .brand-logo-horizontal {
                width: min(320px, 100%);
                max-height: 96px;
            }
            .app-header {
                padding: 1.25rem;
            }
        }
        .section-panel {
            background: #FFFFFF;
            border: 1px solid #E5E7EB;
            border-radius: 8px;
            padding: 1.1rem 1.2rem;
            margin-bottom: 1rem;
            box-shadow: 0 3px 12px rgba(15, 23, 42, 0.05);
        }
        .metric-card {
            background: #FFFFFF;
            border: 1px solid #E5E7EB;
            border-radius: 8px;
            border-top: 5px solid #0F4C81;
            box-shadow: 0 5px 16px rgba(15, 23, 42, 0.09);
            padding: 1rem;
            min-height: 122px;
        }
        .metric-card p {
            color: #475569;
            font-size: .86rem;
            margin: 0 0 .55rem 0;
            font-weight: 700;
        }
        .metric-card strong {
            color: #1F2937;
            font-size: 1.62rem;
            line-height: 1.2;
            word-break: break-word;
        }
        .status-box {
            background: #ECFDF5;
            border: 1px solid #A7F3D0;
            border-radius: 8px;
            color: #065F46;
            padding: .8rem 1rem;
            margin: .5rem 0;
        }
        .login-box {
            background: #FFFFFF;
            border: 1px solid #DFE7EF;
            border-radius: 8px;
            box-shadow: 0 10px 28px rgba(15, 23, 42, 0.12);
            padding: 1.5rem;
            margin-top: 2rem;
        }
        .login-icon {
            width: 54px;
            height: 54px;
            border-radius: 50%;
            background: #EAF2F8;
            color: #123B63;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.7rem;
            margin-bottom: .75rem;
        }
        .file-status {
            background: #FFFFFF;
            border: 1px solid #E5E7EB;
            border-radius: 8px;
            padding: .85rem 1rem;
            margin-bottom: 1rem;
            box-shadow: 0 3px 12px rgba(15, 23, 42, 0.05);
        }
        .file-status p {
            margin: .25rem 0;
            color: #1F2937;
            font-weight: 600;
        }
        .summary-box {
            background: #F8FAFC;
            border-left: 4px solid #1F6F8B;
            border-radius: 8px;
            padding: 1rem 1.25rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def login_correcto(usuario, contrasena):
    """Valida el acceso ignorando mayusculas en el usuario."""
    usuario_autorizado, contrasena_autorizada = obtener_credenciales()
    if not usuario_autorizado or not contrasena_autorizada:
        st.error("No se encontraron credenciales configuradas.")
        return False

    # En Streamlit Cloud estas credenciales se configuran desde Secrets.
    return (
        usuario.strip().lower() == usuario_autorizado.strip().lower()
        and contrasena == contrasena_autorizada
    )


def mostrar_login():
    """Muestra la pantalla de acceso."""
    mostrar_encabezado_principal()

    izquierda, centro, derecha = st.columns([1, 1.1, 1])
    with centro:
        st.markdown('<div class="login-box">', unsafe_allow_html=True)
        st.markdown(html_logo("brand-login-logo", LOGO_PRINCIPAL_PATH), unsafe_allow_html=True)
        st.subheader("Acceso al sistema")
        st.markdown(
            """
            <div style="text-align:center;margin-bottom:1rem;">
                <h2 style="color:#1F2937;margin:.2rem 0;font-size:1.35rem;">Sistema de Comparaci\u00f3n SIAF</h2>
                <p style="color:#475569;margin:.1rem 0;">Devengados, Girados y Pagados</p>
                <p style="color:#475569;margin:.4rem 0 0 0;font-size:.92rem;">Versi\u00f3n 1.0</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.caption("Bienvenida al Sistema de Comparaci\u00f3n SIAF")
        usuario = st.text_input("Usuario")
        contrasena = st.text_input("Contrase\u00f1a", type="password")

        if st.button("Ingresar", type="primary", use_container_width=True):
            if login_correcto(usuario, contrasena):
                st.session_state["autenticado"] = True
                st.session_state["usuario"] = "Erika"
                st.success("Bienvenida Erika")
                st.rerun()
            else:
                st.error("Usuario o contrase\u00f1a incorrectos")
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown(
        """
        <div class="footer-brand">
            \u00a9 2026 Jos\u00e9 Rubina Quijano<br>
            Systems Engineering
        </div>
        """,
        unsafe_allow_html=True,
    )


def cerrar_sesion():
    """Cierra la sesion actual."""
    st.session_state["autenticado"] = False
    st.session_state.pop("usuario", None)
    st.rerun()


def limpiar_formulario():
    """Limpia los archivos cargados para volver a empezar."""
    version = st.session_state.get("uploader_version", 0)
    for clave in [
        f"archivo_devengados_{version}",
        f"archivo_girados_{version}",
        f"archivo_pagados_{version}",
        "comparar_presionado",
    ]:
        st.session_state.pop(clave, None)
    st.session_state["uploader_version"] = version + 1
    st.rerun()


def porcentaje(parte, total):
    """Calcula un porcentaje para el dashboard."""
    if not total:
        return 0
    return (parte / total) * 100


def mostrar_tarjeta_dashboard(titulo, valor, detalle="", es_monto=False):
    """Muestra una tarjeta con valor y detalle."""
    texto_valor = f"S/ {valor:,.2f}" if es_monto else f"{valor:,}"
    st.markdown(
        f"""
        <div class="metric-card">
            <p>{titulo}</p>
            <strong>{texto_valor}</strong>
            <p>{detalle}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def mostrar_columnas_detectadas(dataframes):
    """Muestra las columnas leidas y su version normalizada."""
    with st.expander("Columnas detectadas en los archivos", expanded=True):
        tabs = st.tabs(list(dataframes.keys()))
        for tab, (nombre, dataframe) in zip(tabs, dataframes.items()):
            with tab:
                columna_detectada = columna_siaf(dataframe)
                if columna_detectada:
                    st.success(f"Columna SIAF detectada: {columna_detectada}")
                else:
                    st.warning("No se detecto columna SIAF en este archivo.")
                st.dataframe(columnas_detectadas(dataframe), use_container_width=True, hide_index=True)


def main():
    st.set_page_config(
        page_title="Sistema de Comparaci\u00f3n SIAF",
        page_icon=str(LOGO_CUADRADO_PATH) if LOGO_CUADRADO_PATH.exists() else ":bar_chart:",
        layout="wide",
    )
    aplicar_estilos()

    if not st.session_state.get("autenticado"):
        mostrar_login()
        return

    mostrar_encabezado_principal()

    with st.sidebar:
        st.markdown(
            f"""
            <div class="sidebar-brand">
                {html_logo("sidebar-logo", LOGO_CUADRADO_PATH)}
                <p>Jos\u00e9 Rubina Quijano</p>
                <span>Systems Engineering</span>
                <span>Versi\u00f3n 1.0</span>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.success("Bienvenida Erika")
        st.button("Cerrar sesi\u00f3n", on_click=cerrar_sesion, use_container_width=True, type="primary")
        st.divider()
        st.header("Carga de archivos")
        version = st.session_state.get("uploader_version", 0)
        archivo_devengados = st.file_uploader("Devengados.xlsx", type=["xlsx"], key=f"archivo_devengados_{version}")
        archivo_girados = st.file_uploader("Girados.xlsx", type=["xlsx"], key=f"archivo_girados_{version}")
        archivo_pagados = st.file_uploader("Pagados.xlsx", type=["xlsx"], key=f"archivo_pagados_{version}")
        st.divider()
        st.button("Nueva comparaci\u00f3n", on_click=limpiar_formulario, use_container_width=True, type="secondary")

    archivos_completos = archivo_devengados and archivo_girados and archivo_pagados
    st.markdown('<div class="file-status">', unsafe_allow_html=True)
    if archivos_completos:
        st.success("\u2714 Archivos cargados correctamente")
    else:
        st.warning("\u26a0 Falta cargar archivo")
    estado_devengados = "\u2714 Devengados cargado" if archivo_devengados else "\u26a0 Devengados pendiente"
    estado_girados = "\u2714 Girados cargado" if archivo_girados else "\u26a0 Girados pendiente"
    estado_pagados = "\u2714 Pagados cargado" if archivo_pagados else "\u26a0 Pagados pendiente"
    st.markdown(
        f"""
        <p>{estado_devengados}</p>
        <p>{estado_girados}</p>
        <p>{estado_pagados}</p>
        """,
        unsafe_allow_html=True,
    )
    st.markdown("</div>", unsafe_allow_html=True)

    if not archivo_devengados or not archivo_girados or not archivo_pagados:
        faltantes = []
        if not archivo_devengados:
            faltantes.append("Devengados.xlsx")
        if not archivo_girados:
            faltantes.append("Girados.xlsx")
        if not archivo_pagados:
            faltantes.append("Pagados.xlsx")

        st.warning("Faltan archivos por cargar: " + ", ".join(faltantes))
        st.info("Carga los tres archivos Excel para iniciar la comparaci\u00f3n.")
        with st.expander("\u00bfC\u00f3mo funciona?"):
            st.markdown(
                """
                **Devengados pendientes:** Son registros devengados que a\u00fan no aparecen en girados.

                **Girados pendientes:** Son registros girados que a\u00fan no aparecen en pagados.
                """
            )
        return

    if not st.button("Comparar", type="primary", use_container_width=True):
        st.info("Presiona Comparar para procesar los archivos cargados.")
        return

    with st.spinner("Procesando archivos..."):
        df_devengados = cargar_excel(archivo_devengados)
        df_girados = cargar_excel(archivo_girados)
        df_pagados = cargar_excel(archivo_pagados)

    if df_devengados is None or df_girados is None or df_pagados is None:
        st.error("No fue posible continuar. Revisa que los archivos sean Excel validos.")
        return

    st.success("\u2714 Archivos cargados correctamente")
    st.info("\u2139 Los archivos originales no fueron modificados")

    dataframes = {
        "Devengados": df_devengados,
        "Girados": df_girados,
        "Pagados": df_pagados,
    }

    mostrar_columnas_detectadas(dataframes)

    if not validar_columna_siaf(dataframes):
        return

    with st.spinner("Realizando comparaci\u00f3n..."):
        devengados_pendientes_base, eliminados_devengados = comparar_archivos(
            df_devengados, df_girados
        )
        girados_pendientes_base, eliminados_girados = comparar_archivos(
            df_girados, df_pagados
        )

        devengados_eliminados_base = obtener_registros_eliminados(df_devengados, df_girados)
        girados_eliminados_base = obtener_registros_eliminados(df_girados, df_pagados)
        devengados_pendientes = preparar_resultado(devengados_pendientes_base)
        girados_pendientes = preparar_resultado(girados_pendientes_base)
        devengados_eliminados_reporte = preparar_resultado(devengados_eliminados_base)
        girados_eliminados_reporte = preparar_resultado(girados_eliminados_base)

    st.success("\u2714 Comparaci\u00f3n realizada con \u00e9xito")
    total_devengados = calcular_total_monto(devengados_pendientes)
    total_girados = calcular_total_monto(girados_pendientes)
    fecha_generacion = datetime.now()
    nombre_archivo_reporte = nombre_reporte_unico(fecha_generacion)
    datos_resumen = {
        "devengados_originales": len(df_devengados),
        "girados_originales": len(df_girados),
        "pagados_originales": len(df_pagados),
        "devengados_pendientes": len(devengados_pendientes),
        "girados_pendientes": len(girados_pendientes),
        "devengados_eliminados": eliminados_devengados,
        "girados_eliminados": eliminados_girados,
        "monto_devengados": total_devengados,
        "monto_girados": total_girados,
    }
    registros_eliminados = crear_registros_eliminados(
        devengados_eliminados_reporte, girados_eliminados_reporte
    )

    st.markdown('<div class="section-panel">', unsafe_allow_html=True)
    st.subheader("Dashboard")
    dev_procesados = len(df_devengados) - eliminados_devengados
    gir_procesados = len(df_girados) - eliminados_girados
    tarjetas = [
        ("\U0001f4c4 Devengados originales", len(df_devengados), "Base cargada", False),
        ("\U0001f4c4 Girados originales", len(df_girados), "Base cargada", False),
        ("\U0001f4c4 Pagados originales", len(df_pagados), "Base cargada", False),
        (
            "\u2705 Devengados pendientes",
            len(devengados_pendientes),
            f"Devengados procesados: {dev_procesados:,} de {len(df_devengados):,} ({porcentaje(dev_procesados, len(df_devengados)):.1f}%)",
            False,
        ),
        (
            "\u2705 Girados pendientes",
            len(girados_pendientes),
            f"Girados procesados: {gir_procesados:,} de {len(df_girados):,} ({porcentaje(gir_procesados, len(df_girados)):.1f}%)",
            False,
        ),
        (
            "\U0001f5d1 Total eliminado de Devengados",
            eliminados_devengados,
            f"{porcentaje(eliminados_devengados, len(df_devengados)):.1f}% del archivo",
            False,
        ),
        (
            "\U0001f5d1 Total eliminado de Girados",
            eliminados_girados,
            f"{porcentaje(eliminados_girados, len(df_girados)):.1f}% del archivo",
            False,
        ),
        ("\U0001f4b0 Monto Devengados", total_devengados, "Suma del reporte", True),
        ("\U0001f4b0 Monto Girados", total_girados, "Suma del reporte", True),
    ]

    columnas_tarjetas = st.columns(3)
    for indice, (titulo, valor, detalle, es_monto) in enumerate(tarjetas):
        with columnas_tarjetas[indice % 3]:
            mostrar_tarjeta_dashboard(titulo, valor, detalle, es_monto)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="section-panel">', unsafe_allow_html=True)
    st.subheader("Resumen Ejecutivo")
    resumen_ejecutivo = [
        ("\U0001f4c4 Devengados originales", len(df_devengados), False),
        ("\U0001f4c4 Girados originales", len(df_girados), False),
        ("\U0001f4c4 Pagados originales", len(df_pagados), False),
        ("\u2705 Devengados pendientes", len(devengados_pendientes), False),
        ("\u2705 Girados pendientes", len(girados_pendientes), False),
        ("\U0001f5d1 Devengados eliminados", eliminados_devengados, False),
        ("\U0001f5d1 Girados eliminados", eliminados_girados, False),
        ("\U0001f4b0 Total Devengados", total_devengados, True),
        ("\U0001f4b0 Total Girados", total_girados, True),
    ]
    columnas_resumen = st.columns(3)
    for indice, (titulo, valor, es_monto) in enumerate(resumen_ejecutivo):
        with columnas_resumen[indice % 3]:
            mostrar_tarjeta(titulo, valor, es_monto)
    st.markdown("</div>", unsafe_allow_html=True)

    st.divider()

    with st.spinner("Generando reporte..."):
        archivo_excel = convertir_a_excel(
            {
                "Devengados Pendientes": devengados_pendientes,
                "Girados Pendientes": girados_pendientes,
            },
            datos_resumen=datos_resumen,
            registros_eliminados=registros_eliminados,
            fecha_generacion=fecha_generacion,
        )
        ruta_reporte = guardar_reporte_historico(archivo_excel, nombre_archivo_reporte)

    st.success("\u2714 Reporte generado correctamente")
    st.info(f"Ruta del archivo generado: {ruta_reporte}")
    st.info("\u2139 Los archivos originales no fueron modificados")

    st.markdown('<div class="section-panel">', unsafe_allow_html=True)
    st.subheader("Resultados")
    tab_devengados, tab_girados = st.tabs(["Devengados Pendientes", "Girados Pendientes"])

    with tab_devengados:
        busqueda_devengados = st.text_input("Buscar en Devengados Pendientes", key="buscar_devengados")
        tabla_devengados = filtrar_tabla(devengados_pendientes, busqueda_devengados)
        st.dataframe(tabla_devengados, use_container_width=True, hide_index=True)

    with tab_girados:
        busqueda_girados = st.text_input("Buscar en Girados Pendientes", key="buscar_girados")
        tabla_girados = filtrar_tabla(girados_pendientes, busqueda_girados)
        st.dataframe(tabla_girados, use_container_width=True, hide_index=True)

    if st.button("Ver registros eliminados", use_container_width=True):
        st.session_state["ver_registros_eliminados"] = not st.session_state.get(
            "ver_registros_eliminados", False
        )

    if st.session_state.get("ver_registros_eliminados", False):
        st.subheader("Registros Eliminados")
        tab_elim_dev, tab_elim_gir = st.tabs(
            [
                "Devengados eliminados porque ya estaban en Girados",
                "Girados eliminados porque ya estaban en Pagados",
            ]
        )
        with tab_elim_dev:
            st.dataframe(devengados_eliminados_reporte, use_container_width=True, hide_index=True)
        with tab_elim_gir:
            st.dataframe(girados_eliminados_reporte, use_container_width=True, hide_index=True)

    st.download_button(
        "Descargar " + nombre_archivo_reporte,
        data=archivo_excel,
        file_name=nombre_archivo_reporte,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
    st.markdown("</div>", unsafe_allow_html=True)

    st.divider()
    with st.expander("\u00bfC\u00f3mo funciona?"):
        st.markdown(
            """
            **Devengados pendientes:** Son registros devengados que a\u00fan no aparecen en girados.

            **Girados pendientes:** Son registros girados que a\u00fan no aparecen en pagados.
            """
        )
    mostrar_footer()


if __name__ == "__main__":
    main()
