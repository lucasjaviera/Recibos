"""
============================================================================
 Extractor de Recibos de Sueldo -> Excel consolidado
============================================================================
App Streamlit que procesa múltiples recibos de sueldo en PDF (MINERA COMIRNA)
y genera un único Excel con las columnas:

    Legajo | Apellido y nombre | Código | Concepto | Unidad | Importe liquidado

Soporta DOS estructuras de recibo distintas:

  - FORMATO A (Enero a Mayo): recibo apaisado, con el detalle repetido
    dos veces (original + duplicado) en la misma página. Las columnas de
    importe son "Hab.C/Desc.", "Hab.S/Desc." y "Deducciones": en cada fila
    solo una de las tres tiene valor, y esa se unifica en "Importe liquidado".

  - FORMATO B (Junio en adelante): recibo vertical, una sola copia por
    página, con una tabla de aportes del EMPLEADOR arriba y la tabla del
    empleado abajo (cabecera "CÓDIGO ... MONTO"). El "MONTO" se mapea
    directamente a "Importe liquidado".

La extracción NO usa extract_tables() (que en estos recibos pierde la
alineación fila-columna), sino un enfoque por coordenadas de palabras
(extract_words) que reconstruye cada fila a partir del código de concepto.

Autor: Data Engineering - automatización contable
============================================================================
"""

import io
import re
from collections import defaultdict

import pandas as pd
import pdfplumber
import streamlit as st

# ---------------------------------------------------------------------------
# CONSTANTES / CONFIGURACIÓN DE EXTRACCIÓN
# ---------------------------------------------------------------------------

# Un código de concepto es siempre un número de exactamente 4 dígitos (0010, 0300...).
# Anclar las filas en este patrón descarta automáticamente TOTALES, NETO,
# SUB TOTAL, la tabla de aportes del empleador (sin código) y demás ruido.
COD_RE = re.compile(r"^\d{4}$")

# CUIL/CUIT para poder cortar el nombre en el Formato A (xx-xxxxxxxx-x).
CUIL_RE = re.compile(r"^\d{2}-\d{7,8}-\d$")

# Tolerancia vertical (en px) para agrupar palabras en una misma "fila".
# En el Formato A la columna "Deducciones" se renderiza ~3px más abajo que el
# resto de la fila, por eso agrupamos con una holgura pequeña.
Y_TOL = 5

# Nombres de columnas del DataFrame final (en el orden pedido).
COLUMNS = ["Legajo", "Apellido y nombre", "Código", "Concepto", "Unidad", "Importe liquidado"]


# ---------------------------------------------------------------------------
# HELPERS DE PARSEO
# ---------------------------------------------------------------------------

def _to_float(texto):
    """
    Convierte un importe en formato argentino ("1.397.224,50", "-430.706,70")
    a float. Devuelve None si el texto no es un número válido.
    """
    if texto is None:
        return None
    t = texto.strip().replace("$", "").replace(" ", "")
    if t in ("", "-"):
        return None
    # Miles con punto, decimales con coma -> formato estándar.
    t = t.replace(".", "").replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return None


def _es_importe(texto):
    """True si el token parece un importe numérico (permite signo y separadores)."""
    return bool(re.fullmatch(r"-?[\d.]+,\d{2}", (texto or "").strip()))


def _agrupar_en_filas(words):
    """
    Agrupa una lista de 'words' (dicts de pdfplumber con top/x0/text) en filas
    lógicas, usando la coordenada vertical 'top' con tolerancia Y_TOL.

    Devuelve una lista de filas; cada fila es una lista de (x0, texto) ordenada
    por x0 (izquierda a derecha).
    """
    if not words:
        return []

    # Ordenamos por posición vertical y vamos abriendo filas nuevas cuando el
    # salto respecto de la fila actual supera la tolerancia.
    words = sorted(words, key=lambda w: (w["top"], w["x0"]))
    filas = []
    fila_actual = []
    top_ref = None

    for w in words:
        if top_ref is None or abs(w["top"] - top_ref) <= Y_TOL:
            fila_actual.append(w)
            # El 'top_ref' es el de la primera palabra de la fila (la fila real,
            # no la columna desalineada que viene un poco más abajo).
            if top_ref is None:
                top_ref = w["top"]
        else:
            filas.append(fila_actual)
            fila_actual = [w]
            top_ref = w["top"]

    if fila_actual:
        filas.append(fila_actual)

    # Normalizamos cada fila a (x0, texto) ordenado por x.
    return [sorted([(round(w["x0"]), w["text"]) for w in f]) for f in filas]


# ---------------------------------------------------------------------------
# EXTRACCIÓN FORMATO A (Enero a Mayo) - recibo apaisado y duplicado
# ---------------------------------------------------------------------------

def _extraer_formato_a(page):
    """
    Extrae las filas de conceptos de una página en Formato A.

    El recibo aparece dos veces (izquierda y derecha); recortamos la mitad
    izquierda para no duplicar datos. Reconstruye cada fila por coordenadas y
    unifica Hab.C/Desc. / Hab.S/Desc. / Deducciones en un único importe.
    """
    # La mitad derecha es un duplicado exacto: recortamos un poco antes del
    # centro para no arrastrar fragmentos del segundo recibo.
    ancho_util = page.width * 0.49
    left = page.crop((0, 0, ancho_util, page.height))
    filas = _agrupar_en_filas(left.extract_words())

    # --- Datos de cabecera: Legajo + Apellido y nombre --------------------
    legajo, nombre = None, None
    for i, fila in enumerate(filas):
        textos = [t for _, t in fila]
        if "Legajo" in textos and any(t.startswith("Apellido") for t in textos):
            # La fila de datos es la siguiente al encabezado.
            if i + 1 < len(filas):
                datos = filas[i + 1]
                legajo = datos[0][1]  # token más a la izquierda
                # El nombre son los tokens entre el legajo y el CUIL.
                nombre_tokens = []
                for _, t in datos[1:]:
                    if CUIL_RE.match(t):
                        break
                    nombre_tokens.append(t)
                nombre = " ".join(nombre_tokens).strip()
            break

    # --- Localizar cabecera de la tabla de conceptos y sus columnas -------
    x_unidades = None
    header_idx = None
    for i, fila in enumerate(filas):
        textos = [t for _, t in fila]
        if "Cod" in textos and any(t.startswith("Unidades") for t in textos):
            header_idx = i
            for x, t in fila:
                if t.startswith("Unidades"):
                    x_unidades = x
            break

    registros = []
    if header_idx is None:
        return registros  # sin tabla reconocible

    # --- Recorrer filas de conceptos hasta TOTALES ------------------------
    for fila in filas[header_idx + 1:]:
        textos = [t for _, t in fila]
        if any(t.startswith("TOTALES") for t in textos):
            break

        code = fila[0][1]
        if not COD_RE.match(code):
            continue  # no es una fila de concepto válida

        concepto_tokens, unidad, importe = [], None, None
        # Frontera de columnas: todo lo que esté a la izquierda de "Unidades"
        # (con un margen) es texto del concepto; a la derecha, números.
        x_borde_concepto = (x_unidades or 175) - 8

        for x, t in fila[1:]:
            if x < x_borde_concepto and not _es_importe(t):
                concepto_tokens.append(t)
            elif _es_importe(t):
                val = _to_float(t)
                # La primera columna numérica (col. Unidades) es la unidad;
                # cualquier importe a la derecha es el monto liquidado.
                if x < (x_unidades or 175) + 30 and unidad is None and abs(val) < 100000:
                    unidad = val
                else:
                    importe = val  # Hab.C/Desc, Hab.S/Desc o Deducciones

        registros.append({
            "Legajo": legajo,
            "Apellido y nombre": nombre,
            "Código": code,
            "Concepto": " ".join(concepto_tokens).strip(),
            "Unidad": unidad,
            "Importe liquidado": importe,
        })

    return registros


# ---------------------------------------------------------------------------
# EXTRACCIÓN FORMATO B (Junio en adelante) - recibo vertical
# ---------------------------------------------------------------------------

def _extraer_formato_b(page):
    """
    Extrae las filas de conceptos de una página en Formato B.

    Hay una tabla de aportes del EMPLEADOR arriba (cabecera "CONCEPTO...MONTO"
    SIN código) y la tabla del empleado abajo (cabecera "CÓDIGO...MONTO").
    Anclamos en la cabecera con "CÓDIGO" y tomamos el último número de cada
    fila como "MONTO" -> Importe liquidado.
    """
    filas = _agrupar_en_filas(page.extract_words())

    # --- Datos de cabecera: Apellido y nombre + Legajo --------------------
    legajo, nombre = None, None
    for i, fila in enumerate(filas):
        textos = [t for _, t in fila]
        if "APELLIDO" in textos and "LEGAJO" in textos:
            if i + 1 < len(filas):
                datos = filas[i + 1]  # p.ej: Mensual 06 2026 ANDRADE, JOSE MIGUEL 00000013 $ ...
                nombre_tokens = []
                for _, t in datos:
                    if re.fullmatch(r"\d{6,}", t):  # legajo (6+ dígitos)
                        legajo = t
                        break
                    # Saltamos periodo/mes/año; el nombre suele venir en MAYÚSCULAS.
                    if t in ("Mensual", "Quincenal") or re.fullmatch(r"\d{1,4}", t):
                        continue
                    nombre_tokens.append(t)
                nombre = " ".join(nombre_tokens).strip().rstrip(",")
                # Reponemos la coma tradicional "APELLIDO, NOMBRE" si se perdió.
                nombre = re.sub(r"\s{2,}", " ", nombre)
            break

    # --- Localizar la cabecera de la tabla del EMPLEADO (con CÓDIGO) -------
    x_unidad = None
    header_idx = None
    for i, fila in enumerate(filas):
        textos = [t for _, t in fila]
        if "CÓDIGO" in textos and "MONTO" in textos:
            header_idx = i
            for x, t in fila:
                if t == "UNIDAD":
                    x_unidad = x
            break

    registros = []
    if header_idx is None:
        return registros

    # --- Recorrer filas de conceptos hasta el final de la tabla -----------
    for fila in filas[header_idx + 1:]:
        textos = [t for _, t in fila]
        # La tabla del empleado termina en la fila de COMPOSICIÓN SALARIAL o
        # "Son: ...". OJO: no cortar por "SUELDO", porque "SUELDO MENSUAL" es un
        # concepto válido de la propia tabla.
        if any(t.startswith(("COMPOSICIÓN", "Son:")) for t in textos):
            break

        code = fila[0][1]
        if not COD_RE.match(code):
            continue

        concepto_tokens, unidad, importe = [], None, None
        numeros = []  # (x, valor) de los tokens numéricos de la fila
        x_borde = (x_unidad or 226) - 10

        for x, t in fila[1:]:
            if _es_importe(t):
                numeros.append((x, _to_float(t)))
            elif x < x_borde:
                concepto_tokens.append(t)

        if numeros:
            # MONTO = último número de la fila (columna más a la derecha).
            importe = numeros[-1][1]
            # UNIDAD = primer número SOLO si cae en la columna UNIDAD (~226-300).
            x0_prim, val_prim = numeros[0]
            if x0_prim < 300 and len(numeros) >= 2:
                unidad = val_prim
            elif x0_prim < 300 and len(numeros) == 1:
                # Un único número que está en la columna UNIDAD (raro): es unidad,
                # no monto. En la práctica el único número es siempre MONTO.
                pass

        registros.append({
            "Legajo": legajo,
            "Apellido y nombre": nombre,
            "Código": code,
            "Concepto": " ".join(concepto_tokens).strip(),
            "Unidad": unidad,
            "Importe liquidado": importe,
        })

    return registros


# ---------------------------------------------------------------------------
# DETECCIÓN DE FORMATO
# ---------------------------------------------------------------------------

def _detectar_formato(texto_pagina):
    """
    Decide el formato de una página a partir de su texto.
      -> "A" si aparecen las columnas Hab.C/Desc. / Hab.S/Desc. / Deducciones.
      -> "B" si aparece la cabecera CÓDIGO ... MONTO.
      -> None si no se reconoce.
    """
    t = texto_pagina or ""
    tl = t.lower()
    if "hab.c/desc" in tl or "hab.s/desc" in tl or "deducciones" in tl:
        return "A"
    if "código" in tl and "monto" in tl:
        return "B"
    return None


# ---------------------------------------------------------------------------
# FUNCIÓN PRINCIPAL DE EXTRACCIÓN (por archivo)
# ---------------------------------------------------------------------------

def procesar_pdf(file_bytes, nombre_archivo):
    """
    Procesa un PDF completo (potencialmente muchas páginas/recibos) y devuelve
    una tupla (lista_de_registros, lista_de_avisos).

    Cada 'aviso' es un string describiendo una página que no se pudo leer o
    cuyo formato no se reconoció, sin cortar el resto del procesamiento.
    """
    registros, avisos = [], []

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for n_pag, page in enumerate(pdf.pages, start=1):
            # try/except por PÁGINA: un recibo roto no tumba el archivo entero.
            try:
                texto = page.extract_text() or ""
                formato = _detectar_formato(texto)

                if formato == "A":
                    registros_pag = _extraer_formato_a(page)
                elif formato == "B":
                    registros_pag = _extraer_formato_b(page)
                else:
                    avisos.append(
                        f"⚠️ {nombre_archivo} (pág. {n_pag}): formato no reconocido, se omite."
                    )
                    continue

                if not registros_pag:
                    avisos.append(
                        f"⚠️ {nombre_archivo} (pág. {n_pag}): no se extrajeron conceptos."
                    )
                registros.extend(registros_pag)

            except Exception as e:  # noqa: BLE001 - queremos continuar siempre
                avisos.append(f"❌ {nombre_archivo} (pág. {n_pag}): error al leer -> {e}")
                continue

    return registros, avisos


def construir_dataframe(registros):
    """
    Arma el DataFrame final a partir de los registros crudos, aplicando la
    limpieza y los tipos de datos pedidos:
      - Legajo / Código / Apellido y nombre: string (conservando ceros a izq.)
      - Unidad / Importe liquidado: float
    Descarta filas totalmente vacías o sin importe.
    """
    if not registros:
        return pd.DataFrame(columns=COLUMNS)

    df = pd.DataFrame(registros, columns=COLUMNS)

    # Tipado string y limpieza de espacios sobrantes.
    for col in ["Legajo", "Apellido y nombre", "Código", "Concepto"]:
        df[col] = (
            df[col].astype("string")
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
        )

    # Tipado numérico (float) para Unidad e Importe liquidado.
    df["Unidad"] = pd.to_numeric(df["Unidad"], errors="coerce")
    df["Importe liquidado"] = pd.to_numeric(df["Importe liquidado"], errors="coerce")

    # Descartar encabezados repetidos / filas basura sin código válido.
    df = df[df["Código"].str.fullmatch(r"\d{4}", na=False)]

    # Descartar filas sin ningún importe (nulos que pdfplumber pudo colar).
    df = df.dropna(subset=["Importe liquidado"])

    return df.reset_index(drop=True)


def dataframe_a_excel_bytes(df):
    """
    Serializa el DataFrame a un Excel en memoria (BytesIO), sin tocar disco.
    Fuerza Legajo y Código como texto para preservar los ceros a la izquierda.
    """
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Recibos")
        # Formato de texto en las columnas Legajo (A) y Código (C).
        ws = writer.sheets["Recibos"]
        for fila in ws.iter_rows(min_row=2, min_col=1, max_col=3):
            fila[0].number_format = "@"  # Legajo
            fila[2].number_format = "@"  # Código
    buffer.seek(0)
    return buffer.getvalue()


# ===========================================================================
# INTERFAZ STREAMLIT
# ===========================================================================

def render_ui():
    """Construye toda la interfaz de la aplicación."""
    st.set_page_config(page_title="Recibos PDF → Excel", page_icon="📄", layout="wide")

    st.title("📄 Recibos de sueldo (PDF) → Excel consolidado")
    st.caption(
        "Subí uno o varios recibos en PDF. La app detecta automáticamente el "
        "formato (Ene–May / Jun en adelante), unifica el importe liquidado y "
        "genera un único Excel."
    )

    # 1) CARGA DE ARCHIVOS ---------------------------------------------------
    archivos = st.file_uploader(
        "Arrastrá o seleccioná los PDF",
        type=["pdf"],
        accept_multiple_files=True,
        help="Podés subir varios recibos a la vez.",
    )

    if not archivos:
        st.info("Esperando archivos PDF…")
        return

    # 2) PROCESAMIENTO CON SPINNER / PROGRESO -------------------------------
    todos_registros, todos_avisos = [], []
    barra = st.progress(0.0, text="Procesando recibos…")

    with st.spinner("Extrayendo datos de los PDF…"):
        for idx, archivo in enumerate(archivos, start=1):
            try:
                regs, avisos = procesar_pdf(archivo.getvalue(), archivo.name)
                todos_registros.extend(regs)
                todos_avisos.extend(avisos)
            except Exception as e:  # noqa: BLE001
                # try/except por ARCHIVO: si uno falla entero, seguimos con el resto.
                st.error(f"No se pudo procesar «{archivo.name}»: {e}")
            barra.progress(idx / len(archivos), text=f"Procesados {idx}/{len(archivos)}")

    barra.empty()

    # 3) AVISOS (páginas problemáticas) -------------------------------------
    for aviso in todos_avisos:
        if aviso.startswith("❌"):
            st.error(aviso)
        else:
            st.warning(aviso)

    # 4) DATAFRAME + PREVIEW -------------------------------------------------
    df = construir_dataframe(todos_registros)

    if df.empty:
        st.error("No se pudo extraer ningún concepto de los archivos cargados.")
        return

    c1, c2, c3 = st.columns(3)
    c1.metric("Filas extraídas", len(df))
    c2.metric("Legajos", df["Legajo"].nunique())
    c3.metric("Archivos", len(archivos))

    st.subheader("Previsualización")
    st.dataframe(df, use_container_width=True, hide_index=True)

    # 5) DESCARGA DEL EXCEL --------------------------------------------------
    excel_bytes = dataframe_a_excel_bytes(df)
    st.download_button(
        label="⬇️ Descargar Excel consolidado",
        data=excel_bytes,
        file_name="recibos_consolidado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )


if __name__ == "__main__":
    render_ui()
