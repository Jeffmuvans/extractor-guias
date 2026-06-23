import streamlit as st
import pandas as pd
import re

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

def buscar_1314(df, fila_idx, col_c, col_x):
    """
    Jerarquía de búsqueda:
    1. Col C de la fila tiene 1314 → retorna directo.
    2. Col C está vacía → busca hacia arriba hasta encontrar un valor;
       si ese valor es 1314 lo retorna, si no es None.
    3. Col C empieza con 1166 →
       a. Busca 1314 en col X de la misma fila.
       b. Mira col C de la fila de abajo; si es 1314 lo retorna.
       c. Busca 1314 en col X de la fila de abajo.
    4. Col C tiene otro valor → busca 1314 en col X de la misma fila.
    """
    c = val(df, fila_idx, col_c)

    # CASO 1: col C tiene el 1314 directo
    if c.startswith("1314"):
        return c

    # CASO 2: col C vacía → buscar hacia arriba hasta primer valor
    if c == "":
        for i in range(fila_idx - 1, -1, -1):
            v = val(df, i, col_c)
            if v != "":
                return v if v.startswith("1314") else None
        return None

    # CASO 3: col C empieza con 1166
    if c.startswith("1166"):
        # 3a. Col X de la misma fila
        x = extraer_1314_de_texto(val(df, fila_idx, col_x))
        if x:
            return x
        # 3b. Col C y X de la fila de abajo
        c_abajo = val(df, fila_idx + 1, col_c)
        if c_abajo.startswith("1314"):
            return c_abajo
        x_abajo = extraer_1314_de_texto(val(df, fila_idx + 1, col_x))
        if x_abajo:
            return x_abajo
        # 3c. Col C y X de la fila de arriba
        c_arriba = val(df, fila_idx - 1, col_c)
        if c_arriba.startswith("1314"):
            return c_arriba
        x_arriba = extraer_1314_de_texto(val(df, fila_idx - 1, col_x))
        if x_arriba:
            return x_arriba
        return None

    # CASO 4: otro valor → X misma fila, luego C y X de las 2 filas siguientes
    x = extraer_1314_de_texto(val(df, fila_idx, col_x))
    if x:
        return x
    for i in range(fila_idx + 1, fila_idx + 3):
        c_sig = val(df, i, col_c)
        if c_sig.startswith("1314"):
            return c_sig
        x_sig = extraer_1314_de_texto(val(df, i, col_x))
        if x_sig:
            return x_sig
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
                st.error("El archivo no tiene suficientes columnas (se requiere hasta la columna X).")
                st.stop()

            # Extraer encabezado desde fila 2 (índice 1)
            def celda(col):
                try:
                    v = df.iloc[1, col]
                    return str(v).strip() if pd.notna(v) else ""
                except:
                    return ""

            fecha        = celda(2)   # C2
            hora_inicio  = celda(3)   # D2
            hora_fin     = celda(4)   # E2
            doc_interm   = celda(7)   # H2
            nit          = celda(9)   # J2

            resultados = []
            errores = []

            for i in range(len(df)):
                h = val(df, i, COL_H)
                if h.upper() == "X":
                    guia  = val(df, i, COL_D)
                    val_l = val(df, i, COL_L)
                    val_x = extraer_descripcion(val(df, i, COL_X))
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
            encabezado = "FECHA=" + fecha + "\n" + \
                         "HORA_INICIO=" + hora_inicio + "\n" + \
                         "HORA_FIN=" + hora_fin + "\n" + \
                         "DOC_INTERMEDIARIO=" + doc_interm + "\n" + \
                         "NIT=" + nit + "\n" + \
                         "-----------------------------"
            contenido = encabezado + "\n" + "\n".join(resultados)
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
