# =============================================================================
# AsesorIA — Paso 1 del MVP
# App Streamlit: subir el Excel de la cartera y mostrar el resumen
# (valor total, nº de posiciones y gráfico de tarta por categoría).
# =============================================================================

import io                    # Para guardar los gráficos como imagen "en memoria" (sin archivos)
import unicodedata          # Para quitar acentos al comparar nombres de columna
from datetime import date   # Para la fecha de hoy en la portada del PDF

import pandas as pd          # Lee el Excel y lo convierte en una tabla manejable
import matplotlib.pyplot as plt  # Dibuja el gráfico de tarta
import streamlit as st       # Crea la interfaz web

# reportlab: librería para construir el PDF pieza a pieza (párrafos, imágenes, tablas)
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    Image,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

# -----------------------------------------------------------------------------
# 1. CONFIGURACIÓN DE LA PÁGINA
# -----------------------------------------------------------------------------
# Esto debe ser lo PRIMERO que se ejecuta de Streamlit: define el título de la
# pestaña del navegador y que el contenido ocupe un ancho centrado y cómodo.
st.set_page_config(page_title="AsesorIA — Análisis de carteras", page_icon="📊")

# Color corporativo (azul oscuro estilo banca privada, definido en CLAUDE.md)
AZUL_CORPORATIVO = "#1F3864"

# Disclaimer legal obligatorio (CLAUDE.md): en la app y al pie de cada página del PDF
DISCLAIMER = (
    "Documento de uso interno para el asesor. No constituye recomendación de inversión."
)

# Nombres de los meses para escribir la fecha en español sin depender
# de la configuración regional del ordenador
MESES_ES = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
]


# -----------------------------------------------------------------------------
# 2. FUNCIONES AUXILIARES
# -----------------------------------------------------------------------------

def normalizar(texto: str) -> str:
    """Convierte un texto a minúsculas y sin acentos.

    Ejemplo: 'Categoría' -> 'categoria'.
    Sirve para reconocer las columnas del Excel aunque el asesor las escriba
    un poco diferente (con/sin acentos, mayúsculas, símbolos como €...).
    """
    texto = str(texto).lower().strip()
    # NFKD separa cada letra de su acento; luego nos quedamos solo con la letra
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in texto if not unicodedata.combining(c))


# Diccionario de columnas que esperamos encontrar.
# Clave = nombre interno que usaremos en el código.
# Valor = palabras clave que deben aparecer en el nombre de la columna del Excel.
COLUMNAS_ESPERADAS = {
    "isin": ["isin"],
    "nombre": ["nombre"],
    "categoria": ["categoria"],
    "divisa": ["divisa"],
    "participaciones": ["participaciones"],
    "valor_liquidativo": ["liquidativo"],
    "valor_mercado": ["mercado"],
    "ter": ["ter"],
    "clase": ["clase"],
}


def identificar_columnas(df: pd.DataFrame) -> dict:
    """Busca en el Excel las columnas que necesitamos, con tolerancia.

    Devuelve un diccionario {nombre_interno: nombre_real_en_el_excel}.
    Si una columna no aparece, simplemente no estará en el diccionario
    (más adelante decidimos si es imprescindible o no).
    """
    encontradas = {}
    for col_real in df.columns:
        col_normalizada = normalizar(col_real)
        for nombre_interno, palabras_clave in COLUMNAS_ESPERADAS.items():
            if nombre_interno in encontradas:
                continue  # Ya la encontramos antes, no la pisamos
            if all(palabra in col_normalizada for palabra in palabras_clave):
                encontradas[nombre_interno] = col_real
                break
    return encontradas


def leer_cartera(archivo) -> pd.DataFrame:
    """Lee el Excel subido y devuelve una tabla limpia con las posiciones.

    Pasos:
    1. Leer el Excel con pandas.
    2. Identificar las columnas (aunque tengan nombres aproximados).
    3. Renombrarlas a nombres internos estándar.
    4. Eliminar filas que no son fondos reales (p. ej. la fila 'TOTAL').
    5. Convertir las columnas numéricas a números de verdad.
    6. Calcular el valor de mercado si falta (participaciones × valor liquidativo).
    """
    df = pd.read_excel(archivo)

    columnas = identificar_columnas(df)

    # Comprobación mínima: sin estas columnas no podemos hacer el resumen
    imprescindibles = ["isin", "categoria"]
    faltan = [c for c in imprescindibles if c not in columnas]
    if faltan:
        raise ValueError(
            f"No encuentro estas columnas en el Excel: {', '.join(faltan)}. "
            "Revisa que el archivo tenga el formato esperado."
        )

    # Nos quedamos solo con las columnas reconocidas y las renombramos
    # al nombre interno (así el resto del código no depende del Excel exacto)
    df = df[list(columnas.values())]
    df.columns = list(columnas.keys())

    # Quitar filas sin ISIN: así eliminamos la fila 'TOTAL' del final
    # y cualquier fila vacía que venga en el Excel
    df = df.dropna(subset=["isin"]).copy()

    # Convertir a número las columnas numéricas.
    # errors="coerce" significa: si un valor no es convertible (p. ej. texto),
    # lo deja como vacío (NaN) en lugar de romper la app.
    for col in ["participaciones", "valor_liquidativo", "valor_mercado", "ter"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Si falta el valor de mercado en alguna fila pero tenemos participaciones
    # y valor liquidativo, lo calculamos nosotros (multiplicación)
    if "valor_mercado" not in df.columns:
        df["valor_mercado"] = pd.NA
    puede_calcularse = (
        df["valor_mercado"].isna()
        & df.get("participaciones", pd.Series(dtype=float)).notna()
        & df.get("valor_liquidativo", pd.Series(dtype=float)).notna()
    )
    df.loc[puede_calcularse, "valor_mercado"] = (
        df.loc[puede_calcularse, "participaciones"]
        * df.loc[puede_calcularse, "valor_liquidativo"]
    )

    return df


def formato_euros(cantidad: float) -> str:
    """Formatea un número como euros al estilo español: 200.515,86 €"""
    # Python formatea como 200,515.86 (estilo inglés); intercambiamos los signos
    texto = f"{cantidad:,.2f}"
    texto = texto.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{texto} €"


def generar_pdf(
    nombre_cliente: str,
    cartera: pd.DataFrame,
    valor_total: float,
    png_tarta: io.BytesIO,
    datos_costes: dict | None,
    banderas: dict,
    total_banderas: int,
) -> bytes:
    """Construye el informe PDF y lo devuelve como bytes (listo para descargar).

    reportlab funciona como una "cadena de montaje": vamos añadiendo elementos
    (párrafos, imágenes, tablas, saltos de página) a una lista, y al final
    doc.build() los coloca en páginas A4 automáticamente.
    """
    buffer = io.BytesIO()  # El PDF se construye en memoria, sin tocar el disco
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        topMargin=2 * cm,
        bottomMargin=2.2 * cm,  # Hueco extra abajo para el disclaimer
        leftMargin=2 * cm,
        rightMargin=2 * cm,
        title="Informe de análisis de cartera",
    )

    # --- Estilos de texto (tipografía limpia, azul corporativo) ---
    azul = colors.HexColor(AZUL_CORPORATIVO)
    estilo_portada = ParagraphStyle(
        "portada", fontName="Helvetica-Bold", fontSize=24,
        textColor=azul, alignment=1, leading=30,  # alignment=1 → centrado
    )
    estilo_seccion = ParagraphStyle(
        "seccion", fontName="Helvetica-Bold", fontSize=14,
        textColor=azul, spaceBefore=16, spaceAfter=8,
    )
    estilo_subseccion = ParagraphStyle(
        "subseccion", fontName="Helvetica-Bold", fontSize=11,
        textColor=azul, spaceBefore=10, spaceAfter=4,
    )
    estilo_normal = ParagraphStyle(
        "normal", fontName="Helvetica", fontSize=10, leading=14,
    )
    estilo_centrado = ParagraphStyle(
        "centrado", parent=estilo_normal, fontSize=13, alignment=1,
    )
    estilo_destacado = ParagraphStyle(
        "destacado", fontName="Helvetica-Bold", fontSize=17,
        textColor=azul, alignment=1, spaceBefore=10, spaceAfter=12,
    )

    # --- Pie de página: disclaimer en TODAS las páginas ---
    def pie_de_pagina(canvas, _doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(colors.grey)
        canvas.drawCentredString(A4[0] / 2, 1.2 * cm, DISCLAIMER)
        canvas.drawRightString(A4[0] - 2 * cm, 1.2 * cm, f"Pág. {canvas.getPageNumber()}")
        canvas.restoreState()

    hoy = date.today()
    fecha_texto = f"{hoy.day} de {MESES_ES[hoy.month - 1]} de {hoy.year}"

    elementos = []

    # ========== PORTADA ==========
    elementos.append(Spacer(1, 7 * cm))  # Hueco para centrar verticalmente
    elementos.append(Paragraph("Informe de análisis de cartera", estilo_portada))
    elementos.append(Spacer(1, 1.5 * cm))
    elementos.append(Paragraph(nombre_cliente, estilo_centrado))
    elementos.append(Spacer(1, 0.4 * cm))
    elementos.append(Paragraph(fecha_texto, estilo_centrado))
    elementos.append(PageBreak())  # Salto: lo siguiente empieza en página nueva

    # ========== 1. RESUMEN DE LA CARTERA ==========
    elementos.append(Paragraph("1. Resumen de la cartera", estilo_seccion))
    elementos.append(Paragraph(
        f"Valor total: <b>{formato_euros(valor_total)}</b> &nbsp;&nbsp;·&nbsp;&nbsp; "
        f"Número de posiciones: <b>{len(cartera)}</b>",
        estilo_normal,
    ))
    elementos.append(Spacer(1, 0.4 * cm))
    png_tarta.seek(0)  # Rebobinar la imagen en memoria antes de leerla
    elementos.append(Image(png_tarta, width=12 * cm, height=8.6 * cm))

    # ========== 2. ANÁLISIS DE COSTES ==========
    if datos_costes is not None:
        elementos.append(Paragraph("2. Análisis de costes", estilo_seccion))
        ter_txt = f"{datos_costes['ter_ponderado']:.2f}".replace(".", ",")
        elementos.append(Paragraph(
            f"TER medio ponderado: <b>{ter_txt}%</b> &nbsp;&nbsp;·&nbsp;&nbsp; "
            f"Coste anual estimado: <b>{formato_euros(datos_costes['coste_anual'])}</b>",
            estilo_normal,
        ))
        elementos.append(Spacer(1, 0.3 * cm))
        elementos.append(Paragraph(
            f"Con una rentabilidad bruta del 5% anual, en 10 años la cartera actual "
            f"alcanzaría {formato_euros(datos_costes['valor_final_actual'])}, frente a "
            f"{formato_euros(datos_costes['valor_final_indexada'])} con un TER de "
            f"referencia del 0,40% (cartera indexada).",
            estilo_normal,
        ))
        elementos.append(Paragraph(
            f"Coste de oportunidad a 10 años: "
            f"{formato_euros(datos_costes['coste_oportunidad'])}",
            estilo_destacado,
        ))
        png_proyeccion = datos_costes["png_proyeccion"]
        png_proyeccion.seek(0)
        elementos.append(Image(png_proyeccion, width=14 * cm, height=8 * cm))

    # ========== 3. BANDERAS ROJAS ==========
    elementos.append(Paragraph("3. Banderas rojas", estilo_seccion))
    titulos_revisiones = {
        "caros": "Fondos con TER elevado",
        "concentracion": "Concentración por categoría",
        "solapamiento": "Solapamiento de fondos",
    }
    for clave, titulo in titulos_revisiones.items():
        elementos.append(Paragraph(titulo, estilo_subseccion))
        avisos = banderas.get(clave, [])
        if not avisos:
            elementos.append(Paragraph("Sin incidencias.", estilo_normal))
        else:
            for aviso in avisos:
                elementos.append(Paragraph(f"• {aviso}", estilo_normal))
    elementos.append(Paragraph(
        f"{total_banderas} banderas rojas detectadas", estilo_destacado,
    ))

    # ========== 4. TABLA DE POSICIONES ==========
    elementos.append(Paragraph("4. Posiciones de la cartera", estilo_seccion))

    # Estilo pequeño para que los nombres largos quepan haciendo varias líneas
    estilo_celda = ParagraphStyle("celda", fontName="Helvetica", fontSize=8, leading=10)

    filas = [["ISIN", "Fondo", "Categoría", "Valor mercado", "TER"]]
    for _, posicion in cartera.iterrows():
        valor_txt = (
            formato_euros(posicion["valor_mercado"])
            if pd.notna(posicion["valor_mercado"]) else "—"
        )
        ter_pos = posicion.get("ter")
        ter_pos_txt = (
            f"{ter_pos:.2f}%".replace(".", ",") if pd.notna(ter_pos) else "—"
        )
        filas.append([
            str(posicion.get("isin", "")),
            Paragraph(str(posicion.get("nombre", "")), estilo_celda),
            Paragraph(str(posicion.get("categoria", "")), estilo_celda),
            valor_txt,
            ter_pos_txt,
        ])

    # Anchos de columna: suman 17 cm = ancho útil de un A4 con márgenes de 2 cm
    tabla = Table(filas, colWidths=[3 * cm, 6 * cm, 3.5 * cm, 3 * cm, 1.5 * cm])
    tabla.setStyle(TableStyle([
        # Cabecera: fondo azul corporativo con letra blanca
        ("BACKGROUND", (0, 0), (-1, 0), azul),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        # Números alineados a la derecha (valor y TER)
        ("ALIGN", (3, 0), (4, -1), "RIGHT"),
        # Rejilla gris suave y un poco de aire en cada celda
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#B7C3DC")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    elementos.append(tabla)

    # Montar el documento: reportlab reparte los elementos en páginas
    # y llama a pie_de_pagina en cada una (primera y siguientes)
    doc.build(elementos, onFirstPage=pie_de_pagina, onLaterPages=pie_de_pagina)
    return buffer.getvalue()


# -----------------------------------------------------------------------------
# 3. INTERFAZ DE LA APP
# -----------------------------------------------------------------------------

st.title("📊 AsesorIA — Análisis de carteras")
st.markdown("Sube el Excel con las posiciones de la cartera de tu cliente.")

# Nombre del cliente para la portada del informe (opcional)
nombre_cliente = st.text_input("Nombre del cliente (para el informe)", value="Cliente")
if not nombre_cliente.strip():
    nombre_cliente = "Cliente"  # Si lo dejan vacío, usamos el valor por defecto

# Componente para subir el archivo (solo acepta .xlsx)
archivo_subido = st.file_uploader("Archivo Excel de la cartera", type=["xlsx"])

if archivo_subido is None:
    # Mientras no haya archivo, mostramos una pista y paramos aquí
    st.info("⬆️ Sube un archivo para empezar (por ejemplo, cartera_ejemplo_cliente.xlsx).")
    st.stop()

# --- A partir de aquí ya tenemos un archivo subido ---
try:
    cartera = leer_cartera(archivo_subido)
except Exception as error:
    st.error(f"No he podido leer el Excel: {error}")
    st.stop()

if cartera.empty:
    st.warning("El Excel no contiene ninguna posición válida.")
    st.stop()

# -----------------------------------------------------------------------------
# 4. RESUMEN DE LA CARTERA
# -----------------------------------------------------------------------------

st.header("Resumen de la cartera")

valor_total = cartera["valor_mercado"].sum()
num_posiciones = len(cartera)

# st.columns crea columnas lado a lado; st.metric muestra un dato destacado
col1, col2 = st.columns(2)
col1.metric("Valor total", formato_euros(valor_total))
col2.metric("Nº de posiciones", num_posiciones)

# --- Gráfico de tarta: distribución por categoría ---
st.subheader("Distribución por categoría")

# Agrupamos por categoría y sumamos el valor de mercado de cada una,
# ordenado de mayor a menor para que el gráfico se lea bien
por_categoria = (
    cartera.groupby("categoria")["valor_mercado"].sum().sort_values(ascending=False)
)

# Paleta de azules degradados a partir del color corporativo
colores = plt.cm.Blues_r(
    [0.15 + 0.6 * i / max(len(por_categoria) - 1, 1) for i in range(len(por_categoria))]
)

figura, eje = plt.subplots(figsize=(7, 5))
eje.pie(
    por_categoria,
    labels=por_categoria.index,
    autopct="%1.1f%%",          # Muestra el porcentaje dentro de cada porción
    startangle=90,               # Empieza la primera porción arriba
    colors=colores,
    textprops={"fontsize": 9},
)
eje.set_title("Distribución por categoría", color=AZUL_CORPORATIVO, fontsize=12)
st.pyplot(figura)

# Guardamos el gráfico como imagen PNG "en memoria" para reutilizarlo en el PDF
png_tarta = io.BytesIO()
figura.savefig(png_tarta, format="png", dpi=150, bbox_inches="tight")

# -----------------------------------------------------------------------------
# 5. ANÁLISIS DE COSTES
# -----------------------------------------------------------------------------

st.header("Análisis de costes")

# Parámetros del análisis (definidos en CLAUDE.md / petición del usuario)
TER_REFERENCIA = 0.40      # TER de una cartera indexada de referencia, en %
RENTABILIDAD_BRUTA = 5.0   # Rentabilidad bruta anual supuesta, en %
HORIZONTE_ANIOS = 10       # Años de la proyección

# Resultados del análisis de costes para el PDF.
# Empieza en None: si no hay datos de TER, el PDF omitirá esta sección.
datos_costes = None

# Para el TER ponderado solo podemos usar fondos que tengan TER y valor:
# descartamos filas donde falte alguno de los dos datos
con_ter = cartera.dropna(subset=["ter", "valor_mercado"])

if con_ter.empty:
    st.warning("Ningún fondo tiene dato de TER: no puedo calcular los costes.")
else:
    # Si algún fondo no tiene TER, avisamos de que queda fuera del cálculo
    sin_ter = len(cartera) - len(con_ter)
    if sin_ter > 0:
        st.caption(f"⚠️ {sin_ter} posición(es) sin TER, excluida(s) del cálculo de costes.")

    # --- 1. TER medio ponderado ---
    # Cada fondo pesa según su valor de mercado: es la media de los TER
    # "repartida" según cuánto dinero hay en cada fondo.
    # Fórmula: suma(TER_fondo × valor_fondo) / suma(valores)
    valor_con_ter = con_ter["valor_mercado"].sum()
    ter_ponderado = (con_ter["ter"] * con_ter["valor_mercado"]).sum() / valor_con_ter

    # --- 2. Coste anual estimado en euros ---
    # El TER está en % (ej. 1,45), por eso dividimos entre 100
    coste_anual = valor_total * ter_ponderado / 100

    col1, col2 = st.columns(2)
    col1.metric("TER medio ponderado", f"{ter_ponderado:.2f} %".replace(".", ","))
    col2.metric("Coste anual estimado", formato_euros(coste_anual))

    # --- 3. Proyección a 10 años ---
    # Simplificación habitual: rentabilidad neta anual = rentabilidad bruta − TER.
    # Cada año el valor se multiplica por (1 + rentabilidad_neta) → interés compuesto.
    def proyectar(ter_anual: float) -> list:
        """Devuelve el valor de la cartera año a año (del 0 al horizonte)."""
        rentabilidad_neta = (RENTABILIDAD_BRUTA - ter_anual) / 100
        return [
            valor_total * (1 + rentabilidad_neta) ** anio
            for anio in range(HORIZONTE_ANIOS + 1)
        ]

    evolucion_actual = proyectar(ter_ponderado)     # Cartera con sus costes actuales
    evolucion_indexada = proyectar(TER_REFERENCIA)  # Misma cartera con TER 0,40%

    # El "coste de oportunidad" es lo que el cliente deja de ganar en 10 años
    # por pagar el TER actual en vez del de referencia
    coste_oportunidad = evolucion_indexada[-1] - evolucion_actual[-1]

    # Preparamos los TER como texto con coma decimal (estilo español)
    ter_actual_txt = f"{ter_ponderado:.2f}".replace(".", ",")
    ter_ref_txt = f"{TER_REFERENCIA:.2f}".replace(".", ",")

    st.subheader(f"Proyección a {HORIZONTE_ANIOS} años")
    st.markdown(
        f"""Con una rentabilidad bruta del {RENTABILIDAD_BRUTA:.0f}% anual,
la cartera actual (TER {ter_actual_txt}%) alcanzaría
**{formato_euros(evolucion_actual[-1])}**, frente a
**{formato_euros(evolucion_indexada[-1])}** con un TER de referencia
del {ter_ref_txt}% (cartera indexada)."""
    )

    # Coste de oportunidad destacado en grande, en el azul corporativo
    st.markdown(
        f"""<div style="text-align:center; padding: 1em; border: 2px solid {AZUL_CORPORATIVO};
        border-radius: 8px; margin: 0.5em 0 1em 0;">
        <span style="font-size: 1.0em; color: {AZUL_CORPORATIVO};">Coste de oportunidad a {HORIZONTE_ANIOS} años</span><br>
        <span style="font-size: 2.2em; font-weight: bold; color: {AZUL_CORPORATIVO};">{formato_euros(coste_oportunidad)}</span>
        </div>""",
        unsafe_allow_html=True,
    )

    # --- 4. Gráfico de líneas: evolución de ambas carteras ---
    anios = list(range(HORIZONTE_ANIOS + 1))
    figura2, eje2 = plt.subplots(figsize=(7, 4))
    eje2.plot(
        anios, evolucion_indexada,
        color=AZUL_CORPORATIVO, linewidth=2.5,
        label=f"Cartera indexada (TER {TER_REFERENCIA:.2f}%)".replace(".", ","),
    )
    eje2.plot(
        anios, evolucion_actual,
        color="#8EA9DB", linewidth=2.5, linestyle="--",
        label=f"Cartera actual (TER {ter_ponderado:.2f}%)".replace(".", ","),
    )
    eje2.set_xlabel("Años")
    eje2.set_ylabel("Valor de la cartera")
    # Formatear el eje vertical como euros con separador de miles español
    eje2.yaxis.set_major_formatter(
        plt.FuncFormatter(lambda valor, _: f"{valor:,.0f} €".replace(",", "."))
    )
    eje2.legend(fontsize=9)
    eje2.grid(alpha=0.3)  # Rejilla suave para leer mejor los valores
    eje2.set_title(
        f"Evolución proyectada a {HORIZONTE_ANIOS} años",
        color=AZUL_CORPORATIVO, fontsize=12,
    )
    st.pyplot(figura2)

    # Guardamos el gráfico de proyección en memoria y empaquetamos todos
    # los resultados de costes para reutilizarlos en el informe PDF
    png_proyeccion = io.BytesIO()
    figura2.savefig(png_proyeccion, format="png", dpi=150, bbox_inches="tight")
    datos_costes = {
        "ter_ponderado": ter_ponderado,
        "coste_anual": coste_anual,
        "valor_final_actual": evolucion_actual[-1],
        "valor_final_indexada": evolucion_indexada[-1],
        "coste_oportunidad": coste_oportunidad,
        "png_proyeccion": png_proyeccion,
    }

# -----------------------------------------------------------------------------
# 6. BANDERAS ROJAS
# -----------------------------------------------------------------------------
# Tres revisiones automáticas que alertan al asesor de posibles problemas:
# fondos caros, exceso de concentración y fondos solapados (misma categoría).

st.header("🚩 Banderas rojas")

# Umbrales definidos en CLAUDE.md
UMBRAL_TER_CARO = 1.5        # TER en % a partir del cual un fondo se considera caro
UMBRAL_CONCENTRACION = 40.0  # % máximo razonable de una categoría sobre el total

total_banderas = 0  # Contador para el resumen final de la sección

# Aquí guardamos el texto de cada aviso para incluirlo después en el PDF
banderas_pdf = {"caros": [], "concentracion": [], "solapamiento": []}

# Para nombrar cada fondo usamos su nombre si la columna existe; si no, el ISIN
etiqueta = "nombre" if "nombre" in cartera.columns else "isin"

# --- 1. Fondos caros (TER > 1,5%) ---
st.subheader("Fondos con TER elevado")
# Nota: si un fondo no tiene TER (NaN), la comparación da False y queda fuera
fondos_caros = cartera[cartera["ter"] > UMBRAL_TER_CARO]
if fondos_caros.empty:
    st.success("✅ Sin incidencias: ningún fondo supera el 1,5% de TER.")
else:
    for _, fondo in fondos_caros.iterrows():
        ter_txt = f"{fondo['ter']:.2f}".replace(".", ",")
        aviso = (
            f"{fondo[etiqueta]} — TER {ter_txt}%: "
            "clase cara, revisar si existe clase limpia."
        )
        st.warning(f"⚠️ {aviso}")
        banderas_pdf["caros"].append(aviso)
    total_banderas += len(fondos_caros)

# --- 2. Concentración (categoría > 40% del total) ---
st.subheader("Concentración por categoría")
# Peso de cada categoría en % sobre el valor total de la cartera
pesos = cartera.groupby("categoria")["valor_mercado"].sum() / valor_total * 100
concentradas = pesos[pesos > UMBRAL_CONCENTRACION]
if concentradas.empty:
    st.success("✅ Sin incidencias: ninguna categoría supera el 40% de la cartera.")
else:
    for categoria, peso in concentradas.items():
        peso_txt = f"{peso:.1f}".replace(".", ",")
        aviso = (
            f"{categoria} concentra el {peso_txt}% de la cartera "
            f"(por encima del {UMBRAL_CONCENTRACION:.0f}% recomendado)."
        )
        st.warning(f"⚠️ {aviso}")
        banderas_pdf["concentracion"].append(aviso)
    total_banderas += len(concentradas)

# --- 3. Solapamiento (2 o más fondos en la misma categoría) ---
st.subheader("Solapamiento de fondos")
# Agrupamos por categoría y juntamos los nombres de los fondos en una lista
fondos_por_categoria = cartera.groupby("categoria")[etiqueta].apply(list)
solapadas = fondos_por_categoria[fondos_por_categoria.str.len() >= 2]
if solapadas.empty:
    st.success("✅ Sin incidencias: no hay categorías con varios fondos solapados.")
else:
    for categoria, fondos in solapadas.items():
        lista_fondos = ", ".join(str(f) for f in fondos)
        aviso = (
            f"{categoria} tiene {len(fondos)} fondos que pueden solaparse: "
            f"{lista_fondos}."
        )
        st.warning(f"⚠️ {aviso}")
        banderas_pdf["solapamiento"].append(aviso)
    total_banderas += len(solapadas)

# --- Contador resumen de la sección ---
if total_banderas == 0:
    st.success("✅ **0 banderas rojas detectadas.** Cartera sin incidencias.")
else:
    st.markdown(
        f"""<div style="text-align:center; padding: 0.8em; border: 2px solid {AZUL_CORPORATIVO};
        border-radius: 8px; margin: 0.5em 0 1em 0;">
        <span style="font-size: 1.6em; font-weight: bold; color: {AZUL_CORPORATIVO};">
        🚩 {total_banderas} banderas rojas detectadas</span>
        </div>""",
        unsafe_allow_html=True,
    )

# --- Tabla de posiciones (para que el asesor verifique los datos leídos) ---
st.subheader("Posiciones detectadas")
st.dataframe(cartera, hide_index=True)

# -----------------------------------------------------------------------------
# 7. DESCARGA DEL INFORME PDF
# -----------------------------------------------------------------------------

st.divider()  # Línea separadora antes del bloque de descarga

try:
    # Generamos el PDF en memoria con todo lo calculado arriba
    pdf_bytes = generar_pdf(
        nombre_cliente=nombre_cliente,
        cartera=cartera,
        valor_total=valor_total,
        png_tarta=png_tarta,
        datos_costes=datos_costes,
        banderas=banderas_pdf,
        total_banderas=total_banderas,
    )
    # Nombre del archivo descargado: informe_Nombre_Del_Cliente.pdf
    nombre_archivo = f"informe_{nombre_cliente.strip().replace(' ', '_')}.pdf"
    st.download_button(
        label="📄 Descargar informe PDF",
        data=pdf_bytes,
        file_name=nombre_archivo,
        mime="application/pdf",  # Le dice al navegador que el archivo es un PDF
        type="primary",          # Botón destacado en color
    )
except Exception as error:
    st.error(f"No he podido generar el PDF: {error}")

# -----------------------------------------------------------------------------
# 8. DISCLAIMER LEGAL (obligatorio en todo output, ver CLAUDE.md)
# -----------------------------------------------------------------------------
st.caption(DISCLAIMER)
