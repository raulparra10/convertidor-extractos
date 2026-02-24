import streamlit as st
import pdfplumber
import pandas as pd
import io

def limpiar_monto(valor):
    """Limpia los valores monetarios quitando comas y convirtiéndolos a float."""
    if pd.isna(valor) or valor == '' or valor is None:
        return 0.0
    valor_str = str(valor).replace(',', '').replace(' ', '').strip()
    try:
        return float(valor_str)
    except ValueError:
        return valor

def procesar_pdf(archivo_pdf):
    todas_las_filas = []
    
    # 1. Extraer absolutamente todas las filas grandes
    with pdfplumber.open(archivo_pdf) as pdf:
        for pagina in pdf.pages:
            tablas = pagina.extract_tables()
            for tabla in tablas:
                if tabla:
                    for fila in tabla:
                        # Solo capturamos filas que tengan al menos 8 columnas 
                        # (ignora recuadros pequeños de metadatos)
                        if len(fila) >= 8:
                            # Limpiar texto (quitar saltos de línea molestos)
                            fila_limpia = [str(c).replace('\n', ' ').strip() if c else "" for c in fila]
                            todas_las_filas.append(fila_limpia)

    if not todas_las_filas:
        return None

    # 2. Identificar cuál es la fila de encabezados
    idx_encabezado = -1
    for i, fila in enumerate(todas_las_filas):
        texto_unido = " ".join(fila).lower()
        if "debito" in texto_unido or "débito" in texto_unido:
            idx_encabezado = i
            break

    if idx_encabezado == -1:
        return None

    encabezados = todas_las_filas[idx_encabezado]
    
    # 3. Filtrar los datos reales (todo lo que está debajo del encabezado)
    datos_finales = []
    for fila in todas_las_filas[idx_encabezado + 1:]:
        texto_unido = " ".join(fila).lower()
        
        # Ignoramos si el sistema repite el encabezado en otra página
        if "debito" in texto_unido or "débito" in texto_unido:
            continue
            
        # Ignoramos filas que estén 100% vacías
        if all(c == "" for c in fila):
            continue
            
        datos_finales.append(fila)

    # 4. Asegurar que no haya descuadres (que todas las filas midan igual)
    num_cols = len(encabezados)
    datos_cuadrados = []
    for fila in datos_finales:
        if len(fila) == num_cols:
            datos_cuadrados.append(fila)
        elif len(fila) < num_cols:
            datos_cuadrados.append(fila + [""] * (num_cols - len(fila)))
        else:
            datos_cuadrados.append(fila[:num_cols])

    # 5. Armar la tabla de Pandas y dar formato de número
    df = pd.DataFrame(datos_cuadrados, columns=encabezados)
    
    for col in df.columns:
        nombre_col = col.lower()
        if "debito" in nombre_col or "débito" in nombre_col or "credito" in nombre_col or "crédito" in nombre_col or "saldo" in nombre_col:
            df[col] = df[col].apply(limpiar_monto)

    # 6. Crear el archivo Excel
    buffer_excel = io.BytesIO()
    with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Extracto Proveedor')
        
    buffer_excel.seek(0)
    return buffer_excel

# --- INTERFAZ WEB ---
st.set_page_config(page_title="Conversor de Extractos", page_icon="📊", layout="centered")

st.title("📊 Conversor de Extractos a Excel")
st.markdown("Sube el extracto en formato PDF generado por el sistema y obtén una tabla de Excel limpia y estructurada.")

archivo_subido = st.file_uploader("Selecciona el archivo PDF", type=["pdf"])

if archivo_subido is not None:
    st.info("Procesando el documento, por favor espera...")
    
    try:
        excel_generado = procesar_pdf(archivo_subido)
        
        if excel_generado:
            st.success("¡Conversión exitosa!")
            
            st.download_button(
                label="📥 Descargar archivo Excel",
                data=excel_generado,
                file_name="Extracto_Convertido.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No se encontraron tablas estructuradas en el documento PDF.")
            
    except Exception as e:
        st.error(f"Ocurrió un error interno al procesar el archivo: {e}")
