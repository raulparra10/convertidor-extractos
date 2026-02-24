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
    datos_totales = []
    encabezados = None
    
    with pdfplumber.open(archivo_pdf) as pdf:
        for pagina in pdf.pages:
            tablas = pagina.extract_tables()
            
            for tabla in tablas:
                if not tabla:
                    continue
                    
                for fila in tabla:
                    # Limpiamos saltos de línea y valores nulos
                    fila_limpia = [str(c).replace('\n', ' ').strip() if c else "" for c in fila]
                    texto_fila = " ".join(fila_limpia).lower()
                    
                    # 1. BUSCADOR INTELIGENTE: Identificamos la cabecera por sus palabras clave
                    if "débito" in texto_fila and "saldo" in texto_fila:
                        if not encabezados:
                            encabezados = fila_limpia
                        continue  # Saltamos la cabecera para no meterla como un dato más
                        
                    # 2. EXTRACCIÓN: Si ya encontramos los encabezados, empezamos a guardar los datos
                    if encabezados:
                        # Solo guardamos si la fila no está completamente vacía
                        if any(celda != "" for celda in fila_limpia):
                            # Evitamos guardar cabeceras repetidas si aparecen en la página 2, 3, etc.
                            if "débito" not in texto_fila: 
                                datos_totales.append(fila_limpia)

    # Si no encontró nada, retorna None
    if not datos_totales or not encabezados:
        return None
        
    # 3. CUADRAR TABLA: Nos aseguramos de que todas las filas tengan la misma cantidad de columnas
    datos_cuadrados = []
    num_cols = len(encabezados)
    for fila in datos_totales:
        if len(fila) == num_cols:
            datos_cuadrados.append(fila)
        elif len(fila) < num_cols:
            datos_cuadrados.append(fila + [""] * (num_cols - len(fila)))
        else:
            datos_cuadrados.append(fila[:num_cols])

    # 4. CREAR EXCEL: Armamos el DataFrame
    df = pd.DataFrame(datos_cuadrados, columns=encabezados)
    
    # Limpiamos las columnas numéricas para que Excel las reconozca como moneda
    for col in df.columns:
        if "débito" in col.lower() or "crédito" in col.lower() or "saldo" in col.lower():
            df[col] = df[col].apply(limpiar_monto)
            
    # Generamos el archivo en memoria
    buffer_excel = io.BytesIO()
    with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Extracto Proveedor')
        
    buffer_excel.seek(0)
    return buffer_excel

# --- INTERFAZ WEB CON STREAMLIT ---
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
        st.error(f"Ocurrió un error al procesar el archivo: {e}")
