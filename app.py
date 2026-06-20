import streamlit as st
import pandas as pd
import io

st.set_page_config(
    page_title="Extractor de Guías",
    page_icon="📦",
    layout="centered"
)

st.title("📦 Extractor de Guías")
st.markdown("Sube tu archivo Excel con la hoja **DATOS** y descarga el TXT con las guías marcadas.")

# ---------------------------------------------------------------------------
# Lógica de búsqueda
# ---------------------------------------------------------------------------

import re

def val(df, fila, col):
    if fila < 0 or fila >= len(df):
        return ""
    v = df.iloc[fila, col]
    return str(v).strip() if pd.notna(v) else ""

def extraer_1314_de_texto(texto):
    """Extrae el primer código 1314XXXXXXXX que aparezca en un texto."""
    match = re.search(r'1314\d+', texto)
    return match.group(0) if match else None

def extraer_descripcion(texto):
    """Extrae solo el texto descriptivo desde la primera letra, sin números ni símbolos iniciales."""
    match = re.search(r'(?<![A-Z0-9])([A-ZÁÉÍÓÚÑA-záéíóúñ][^0-9/]{2,})', texto)
    return match.group(1).strip() if match else texto.strip()

def es_cambio_de_grupo(v):
    """Detecta si un valor en columna C representa una nueva Guia Master
    (tiene contenido y NO empieza por 1314)."""
    return v != "" and not v.startswith("1314")

def buscar_1314(df, fila_idx, col_c, col_x):
    """
    Jerarquía de búsqueda:
    1. Busca hacia ABAJO en col C hasta encontrar 1314 o celda vacía.
    2. Si no, busca hacia ARRIBA en col C hasta encontrar 1314 o cambio de grupo.
    3. Si no, extrae 1314 con regex de col X en las filas del mismo grupo.
    """

    # --- Determinar los límites del grupo ---
    # Límite superior: fila de la Master más cercana hacia arriba (inclusive)
    limite_arriba = 0
    for i in range(fila_idx - 1, -1, -1):
        v = val(df, i, col_c)
        if es_cambio_de_grupo(v):
            limite_arriba = i
            break

    # Límite inferior: próxima celda vacía o cambio de grupo hacia abajo
    limite_abajo = len(df) - 1
    for i in range(fila_idx + 1, len(df)):
        v = val(df, i, col_c)
        if v == "" or es_cambio_de_grupo(v):
            limite_abajo = i - 1
            break

    # PASO 1: buscar hacia abajo en col C dentro del grupo
    for i in range(fila_idx, limite_abajo + 1):
        v = val(df, i, col_c)
        if v.startswith("1314"):
            return v

    # PASO 2: buscar hacia arriba en col C dentro del grupo
    for i in range(fila_idx - 1, limite_arriba - 1, -1):
        v = val(df, i, col_c)
        if v.startswith("1314"):
            return v

    # PASO 3: buscar con regex en col X de todas las filas del grupo
    for i in range(limite_arriba, limite_abajo + 1):
        x = val(df, i, col_x)
        codigo = extraer_1314_de_texto(x)
        if codigo:
            return codigo

    return None

# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

archivo = st.file_uploader("Selecciona el archivo Excel", type=["xlsx", "xlsm", "xls"])

if archivo:
    st.success(f"Archivo cargado: **{archivo.name}**")

    if st.button("⚙️ Procesar y generar TXT", type="primary"):

        COL_C, COL_D, COL_H, COL_L, COL_X = 2, 3, 7, 11, 23

        with st.spinner("Procesando..."):
            try:
                df = pd.read_excel(archivo, sheet_name="DATOS", header=None, dtype=str)
            except Exception as e:
                st.error(f"No se pudo leer la hoja 'DATOS': {e}")
                st.stop()

            if df.shape[1] <= COL_X:
                st.error(f"El archivo no tiene suficientes columnas (se requiere hasta la columna X).")
                st.stop()

            resultados = []
            errores = []

            for i in range(len(df)):
                h = val(df, i, COL_H)
                if h.upper() == "X":
                    guia   = val(df, i, COL_D)
                    val_l  = val(df, i, COL_L)
                    val_x  = extraer_descripcion(val(df, i, COL_X))
                    codigo = buscar_1314(df, i, COL_C, COL_X)

                    if codigo:
                        resultados.append(f"{codigo}-{guia}-{val_l}-{val_x}")
                    else:
                        errores.append(f"Fila {i+1}: sin código 1314 para guía '{guia}'")

        # Resultados
        col1, col2 = st.columns(2)
        col1.metric("✅ Registros exportados", len(resultados))
        col2.metric("⚠️ Sin código 1314", len(errores))

        if resultados:
            contenido = "\n".join(resultados)
            nombre_txt = archivo.name.rsplit(".", 1)[0] + "_guias.txt"

            st.download_button(
                label="⬇️ Descargar TXT",
                data=contenido.encode("utf-8"),
                file_name=nombre_txt,
                mime="text/plain",
                type="primary"
            )

            with st.expander(f"Ver contenido ({len(resultados)} líneas)"):
                st.code(contenido, language=None)

        if errores:
            with st.expander(f"⚠️ Filas sin código 1314 ({len(errores)})"):
                for e in errores:
                    st.warning(e)

        if not resultados and not errores:
            st.info("No se encontraron guías marcadas con 'X' en la columna H.")
