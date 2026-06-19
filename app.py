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

def val(df, fila, col):
    if fila < 0 or fila >= len(df):
        return ""
    v = df.iloc[fila, col]
    return str(v).strip() if pd.notna(v) else ""

def buscar_1314(df, fila_idx, col_c, col_x):
    c_actual = val(df, fila_idx, col_c)

    # CASO 1: C vacía → buscar hacia arriba
    if c_actual == "":
        for i in range(fila_idx - 1, -1, -1):
            v = val(df, i, col_c)
            if v.startswith("1314"):
                return v
        return None

    # CASO 2: C ya tiene 1314
    if c_actual.startswith("1314"):
        return c_actual

    # CASO 3: C tiene otro valor → buscar hacia abajo alternando C y X
    fila_busqueda = fila_idx
    while fila_busqueda < len(df):
        x_val = val(df, fila_busqueda, col_x)
        if x_val.startswith("1314"):
            return x_val

        fila_busqueda += 1
        if fila_busqueda >= len(df):
            break

        c_sig = val(df, fila_busqueda, col_c)
        if c_sig == "":
            break
        if c_sig.startswith("1314"):
            return c_sig

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
                    val_x  = val(df, i, COL_X)
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
