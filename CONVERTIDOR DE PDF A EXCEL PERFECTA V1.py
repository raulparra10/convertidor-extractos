import streamlit as st
import pdfplumber
import pandas as pd
import io
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

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
    
    # 1. Extracción pura: Buscar todas las tablas del documento
    with pdfplumber.open(archivo_pdf) as pdf:
        for pagina in pdf.pages:
            tablas = pagina.extract_tables()
            for tabla in tablas:
                if tabla:
                    for fila in tabla:
                        # Filtrar: Solo nos interesan filas anchas (evita los recuadros de fecha/hora)
                        if len(fila) >= 10:
                            fila_limpia = [str(c).replace('\n', ' ').strip() if c else "" for c in fila]
                            todas_las_filas.append(fila_limpia)

    if not todas_las_filas:
        return None

    # 2. Identificar la fila de los títulos (Headers)
    idx_encabezado = -1
    for i, fila in enumerate(todas_las_filas):
        texto_unido = " ".join(fila).lower()
        # Buscamos palabras clave de la tabla principal
        if "comprobante" in texto_unido and "débito" in texto_unido:
            idx_encabezado = i
            break

    if idx_encabezado == -1:
        return None

    encabezados_originales = todas_las_filas[idx_encabezado]
    # Aseguramos nombres de columnas únicos y limpios
    encabezados = [col if col != "" else f"Columna_{i}" for i, col in enumerate(encabezados_originales)]
    
    # 3. Filtrar los datos reales ignorando repeticiones de títulos en otras páginas
    datos_finales = []
    for fila in todas_las_filas[idx_encabezado + 1:]:
        texto_unido = " ".join(fila).lower()
        if "comprobante" in texto_unido and "débito" in texto_unido:
            continue
        if all(c == "" for c in fila):
            continue
        datos_finales.append(fila)

    # 4. Cuadrar la tabla para evitar errores de Pandas
    num_cols = len(encabezados)
    datos_cuadrados = []
    for fila in datos_finales:
        if len(fila) == num_cols:
            datos_cuadrados.append(fila)
        elif len(fila) < num_cols:
            datos_cuadrados.append(fila + [""] * (num_cols - len(fila)))
        else:
            datos_cuadrados.append(fila[:num_cols])

    # 5. Crear el DataFrame y limpiar montos
    df = pd.DataFrame(datos_cuadrados, columns=encabezados)
    for col in df.columns:
        nombre_col = col.lower()
        if "débito" in nombre_col or "crédito" in nombre_col or "saldo" in nombre_col or "debito" in nombre_col or "credito" in nombre_col:
            df[col] = df[col].apply(limpiar_monto)

    # 6. DISEÑO CORPORATIVO Y EXPORTACIÓN A EXCEL
    buffer_excel = io.BytesIO()
    with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Extracto Proveedor')
        
        workbook = writer.book
        worksheet = writer.sheets['Extracto Proveedor']
        
        # Estilos Corporativos
        color_fondo_header = PatternFill(start_color="002060", end_color="002060", fill_type="solid") # Azul oscuro corporativo
        fuente_header = Font(color="FFFFFF", bold=True, name="Calibri", size=11)
        alineacion_centro = Alignment(horizontal="center", vertical="center", wrap_text=True)
        alineacion_izquierda = Alignment(horizontal="left", vertical="center")
        borde_fino = Border(left=Side(style='thin', color="BFBFBF"), 
                            right=Side(style='thin', color="BFBFBF"), 
                            top=Side(style='thin', color="BFBFBF"), 
                            bottom=Side(style='thin', color="BFBFBF"))

        # Aplicar estilo a los Títulos (Header)
        for col_num, value in enumerate(df.columns.values):
            celda = worksheet.cell(row=1, column=col_num+1)
            celda.fill = color_fondo_header
            celda.font = fuente_header
            celda.alignment = alineacion_centro
            celda.border = borde_fino

        # Activar Filtros Automáticos en la primera fila
        worksheet.auto_filter.ref = worksheet.dimensions

        # Aplicar estilo a las filas de datos y formato de moneda
        for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
            for celda in row:
                celda.border = borde_fino
                celda.alignment = alineacion_izquierda
                
                # Formato de moneda para columnas financieras (buscando las últimas 3 columnas generalmente)
                nombre_columna = df.columns[celda.column - 1].lower()
                if "débito" in nombre_columna or "crédito" in nombre_columna or "saldo" in nombre_columna or "debito" in nombre_columna or "credito" in nombre_columna:
                    celda.number_format = '#,##0.00'
                    celda.alignment = Alignment(horizontal="right", vertical="center")

        # Ajustar el ancho de las columnas automáticamente
        for col in worksheet.columns:
            max_length = 0
            columna_letra = col[0].column_letter
            for celda in col:
                try:
                    if len(str(celda.value)) > max_length:
                        max_length = len(str(celda.value))
                except:
                    pass
            # Dar un poco de margen, pero poner un límite máximo para que no sea inmensa
            ancho_ajustado = min((max_length + 2), 45)
            worksheet.column_dimensions[columna_letra].width = ancho_ajustado
            
        # Congelar la primera fila para que al hacer scroll los títulos sigan visibles
        worksheet.freeze_panes = "A2"

    buffer_excel.seek(0)
    return buffer_excel

# --- INTERFAZ WEB ---
st.set_page_config(page_title="Conversor de Extractos", page_icon="📊", layout="centered")

# CSS para darle un toque más limpio a la web
st.markdown("""
    <style>
    .main {background-color: #F8F9FA;}
    h1 {color: #002060;}
    </style>
    """, unsafe_allow_html=True)

st.title("📊 Conversor de Extractos a Excel")
st.markdown("Sube el extracto en formato PDF generado por el sistema administrativo. Obtendrás una tabla estructurada, con diseño corporativo y lista para conciliar.")

archivo_subido = st.file_uploader("Selecciona el archivo PDF", type=["pdf"])

if archivo_subido is not None:
    st.info("Procesando la estructura del documento...")
    
    try:
        excel_generado = procesar_pdf(archivo_subido)
        
        if excel_generado:
            st.success("¡Conversión y formateo exitosos!")
            
            st.download_button(
                label="📥 Descargar Reporte en Excel",
                data=excel_generado,
                file_name="Extracto_Proveedor_Estructurado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No se logró estructurar la tabla. Verifica que el PDF sea el reporte original.")
            
    except Exception as e:
        st.error(f"Ocurrió un error interno: {e}")
