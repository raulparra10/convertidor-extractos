import streamlit as st
import pdfplumber
import pandas as pd
import io
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

def limpiar_monto(valor):
    """Convierte texto de moneda a número float manejando comas y paréntesis."""
    if pd.isna(valor) or valor == '' or valor is None:
        return 0.0
    # Eliminar comas de miles y espacios
    s = str(valor).replace(',', '').replace(' ', '').strip()
    # Manejar saldos negativos representados con signo menos
    try:
        return float(s)
    except ValueError:
        return 0.0

def procesar_pdf(archivo_pdf):
    todas_las_filas = []
    encabezados = None
    
    with pdfplumber.open(archivo_pdf) as pdf:
        for pagina in pdf.pages:
            tablas = pagina.extract_tables()
            for tabla in tablas:
                if not tabla: continue
                
                for fila in tabla:
                    # Limpiar ruidos y saltos de línea en cada celda
                    fila_limpia = [str(c).replace('\n', ' ').strip() if c else "" for c in fila]
                    texto_fila = " ".join(fila_limpia).lower()
                    
                    # 1. IDENTIFICAR CABECERA: Buscamos las columnas clave del extracto
                    if "débito" in texto_fila and "saldo" in texto_fila:
                        if not encabezados:
                            encabezados = fila_limpia
                        continue 
                    
                    # 2. CAPTURAR DATOS: Si ya tenemos cabecera, guardamos la fila
                    if encabezados:
                        # Omitimos filas vacías, el "Saldo inicial" o repeticiones de cabecera
                        if any(fila_limpia) and "saldo inicial" not in texto_fila and "débito" not in texto_fila:
                            # Aseguramos que la fila tenga el mismo largo que el encabezado
                            if len(fila_limpia) == len(encabezados):
                                todas_las_filas.append(fila_limpia)
                            elif len(fila_limpia) > len(encabezados):
                                todas_las_filas.append(fila_limpia[:len(encabezados)])
                            else:
                                todas_las_filas.append(fila_limpia + [""] * (len(encabezados) - len(fila_limpia)))

    if not todas_las_filas: return None

    # Crear DataFrame
    df = pd.DataFrame(todas_las_filas, columns=encabezados)
    
    # Limpiar columnas financieras (Débito, Crédito, Saldo)
    cols_financieras = [c for c in df.columns if any(k in c.lower() for k in ["débito", "crédito", "saldo"])]
    for col in cols_financieras:
        df[col] = df[col].apply(limpiar_monto)

    # GENERACIÓN DE EXCEL CON ESTILO CORPORATIVO
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Extracto Detallado')
        
        wb = writer.book
        ws = writer.sheets['Extracto Detallado']
        
        # Estilos: Azul Oscuro y Blanco para encabezados
        header_fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        center_alig = Alignment(horizontal="center", vertical="center")
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_alig
            cell.border = border

        # Filtros, Inmovilizar panel y Formato Numérico
        ws.auto_filter.ref = ws.dimensions
        ws.freeze_panes = "A2"
        
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                cell.border = border
                # Aplicar formato contable a las columnas de dinero
                if ws.cell(row=1, column=cell.column).value in cols_financieras:
                    cell.number_format = '#,##0'
                    cell.alignment = Alignment(horizontal="right")

        # Ajuste automático de columnas
        for col in ws.columns:
            max_length = max(len(str(cell.value)) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = min(max_length + 2, 50)

    return output.getvalue()

# INTERFAZ STREAMLIT
st.set_page_config(page_title="Convertidor Perfecta", layout="centered")
st.title("📊 Conversor de Extractos a Excel")
st.markdown("Herramienta optimizada para reportes de **Perfecta Automotores S.A.**")

archivo = st.file_uploader("Sube el PDF del Extracto", type=["pdf"])

if archivo:
    with st.spinner("Estructurando datos..."):
        resultado = procesar_pdf(archivo)
        if resultado:
            st.success("¡Tabla generada con éxito!")
            st.download_button(
                label="📥 Descargar Excel Profesional",
                data=resultado,
                file_name="Extracto_Krona_Procesado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("No se detectó la tabla. Asegúrate de que el PDF contiene las columnas Débito/Crédito.")
