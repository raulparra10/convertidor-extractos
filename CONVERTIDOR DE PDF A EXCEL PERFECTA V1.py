import streamlit as st
import pdfplumber
import pandas as pd
import io

def limpiar_monto(valor):
    """Limpia los valores monetarios quitando comas y convirtiéndolos a float."""
    if pd.isna(valor) or valor == '' or valor is None:
        return 0.0
    # Convertir a texto, quitar comas de miles y espacios en blanco
    valor_str = str(valor).replace(',', '').replace(' ', '').strip()
    try:
        return float(valor_str)
    except ValueError:
        return valor

def procesar_pdf(archivo_pdf):
    """Extrae las tablas del PDF, las limpia y devuelve un archivo Excel en memoria."""
    datos_totales = []
    
    # Abrir el PDF
    with pdfplumber.open(archivo_pdf) as pdf:
        for i, pagina in enumerate(pdf.pages):
            # Extraer la tabla de la página actual
            tabla = pagina.extract_table()
            
            if tabla:
                if not datos_totales:
                    # Si es la primera tabla, guardamos todo incluyendo los encabezados
                    datos_totales.extend(tabla)
                else:
                    # Para las siguientes páginas, omitimos la fila de encabezados
                    datos_totales.extend(tabla[1:])
    
    # Si no se encontró ninguna tabla, retornamos None
    if not datos_totales:
        return None
        
    # Extraer y limpiar encabezados (eliminar saltos de línea)
    encabezados = [str(col).replace('\n', ' ').strip() if col else f"Col_Vacia_{i}" 
                   for i, col in enumerate(datos_totales[0])]
    
    # Crear el DataFrame de pandas con el resto de los datos
    df = pd.DataFrame(datos_totales[1:], columns=encabezados)
    
    # Limpiar saltos de línea en todas las celdas para que Excel no cree filas dobles
    df = df.replace('\n', ' ', regex=True)
    
    # Convertir las columnas financieras a formato numérico
    columnas_moneda = ['Débito', 'Crédito', 'Saldo']
    for col in columnas_moneda:
        if col in df.columns:
            df[col] = df[col].apply(limpiar_monto)
            
    # Generar el archivo Excel en memoria (BytesIO) para poder descargarlo en la web
    buffer_excel = io.BytesIO()
    with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Extracto Proveedor')
        
    buffer_excel.seek(0)
    return buffer_excel

# --- INTERFAZ WEB CON STREAMLIT ---

# Configuración básica de la página
st.set_page_config(page_title="Conversor de Extractos", page_icon="📊", layout="centered")

st.title("📊 Conversor de Extractos a Excel")
st.markdown("Sube el extracto en formato PDF generado por el sistema y obtén una tabla de Excel limpia y estructurada.")

# Widget para que el usuario suba el archivo PDF
archivo_subido = st.file_uploader("Selecciona el archivo PDF", type=["pdf"])

if archivo_subido is not None:
    st.info("Procesando el documento, por favor espera...")
    
    try:
        # Llamar a la función principal de procesamiento
        excel_generado = procesar_pdf(archivo_subido)
        
        if excel_generado:
            st.success("¡Conversión exitosa!")
            
            # Botón de descarga
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