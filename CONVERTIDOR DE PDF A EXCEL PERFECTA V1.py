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
    """Extrae las tablas del PDF, las limpia y devuelve un archivo Excel en memoria."""
    datos_totales = []
    encabezados = None
    
    with pdfplumber.open(archivo_pdf) as pdf:
        for pagina in pdf.pages:
            # Extraer TODAS las tablas de la página
            tablas = pagina.extract_tables()
            
            for tabla in tablas:
                # Ignorar tablas vacías
                if not tabla or len(tabla) == 0:
                    continue
                
                # EL TRUCO: Solo tomar en cuenta las tablas grandes (más de 6 columnas)
                # Esto evita capturar los recuadros pequeños de "Hora", "Usuario", etc.
                if len(tabla[0]) > 6:
                    if not encabezados:
                        # Guardar la cabecera la primera vez
                        encabezados = tabla[0]
                        datos_totales.extend(tabla[1:])
                    else:
                        # Si el encabezado se repite en la siguiente página, lo saltamos
                        if tabla[0][0] == encabezados[0]:
                            datos_totales.extend(tabla[1:])
                        else:
                            datos_totales.extend(tabla)
    
    if not datos_totales or not encabezados:
        return None
        
    # Limpiar encabezados (eliminar saltos de línea)
    encabezados = [str(col).replace('\n', ' ').strip() if col else f"Col_Vacia_{i}" 
                   for i, col in enumerate(encabezados)]
    
    # Igualar el tamaño de todas las filas para que Pandas no falle
    datos_cuadrados = []
    for fila in datos_totales:
        if len(fila) == len(encabezados):
            datos_cuadrados.append(fila)
        elif len(fila) < len(encabezados):
            # Rellenar con espacios vacíos si la fila es más corta
            datos_cuadrados.append(fila + [None] * (len(encabezados) - len(fila)))
        else:
            # Recortar si la fila es más larga
            datos_cuadrados.append(fila[:len(encabezados)])

    # Crear el DataFrame
    df = pd.DataFrame(datos_cuadrados, columns=encabezados)
    df = df.replace('\n', ' ', regex=True)
    
    # Aplicar formato financiero
    columnas_moneda = ['Débito', 'Crédito', 'Saldo']
    for col in columnas_moneda:
        if col in df.columns:
            df[col] = df[col].apply(limpiar_monto)
            
    # Generar el Excel
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
