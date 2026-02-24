import os
import pandas as pd
from pdf2image import convert_from_path
import pytesseract
import re
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- CONFIGURACIÓN ---
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
ruta_poppler = r'C:\Users\DELL Latitude 7200\Downloads\poppler-25.12.0\Library\bin'
ruta_pdf = r"C:\Users\DELL Latitude 7200\Desktop\Extracto Proveedor Detallado Krona.pdf"
destino = r"C:\Users\DELL Latitude 7200\Desktop\PERFECTA\RESULTADO SCRIPT PERFECTA"
ruta_excel = os.path.join(destino, "Extracto_Krona_Conciliado_Final.xlsx")

if not os.path.exists(destino): os.makedirs(destino)

def limpiar_monto(valor):
    if not valor: return 0.0
    limpio = str(valor).replace('.', '').replace(',', '').strip()
    try: return float(limpio)
    except: return 0.0

def procesar_extracto_perfecta():
    print("🔍 Analizando extracto con lógica de anclaje contable...")
    try:
        paginas = convert_from_path(ruta_pdf, dpi=300, poppler_path=ruta_poppler)
    except Exception as e:
        print(f"❌ Error Poppler: {e}"); return

    datos_finales = []

    for i, img in enumerate(paginas):
        texto = pytesseract.image_to_string(img, lang='spa', config='--psm 6')
        print(f"Procesando página {i+1}...")
        
        for linea in texto.split('\n'):
            linea = linea.strip()
            if not linea or any(x in linea for x in ["Comprobante", "Totales", "PERFECTA"]): continue
            
            # 1. Extraer los últimos 3 bloques numéricos (Débito, Crédito, Saldo)
            partes = linea.split()
            if len(partes) < 4: continue
            
            saldo = partes[-1]
            credito = partes[-2]
            debito = partes[-3]
            
            # 2. Extraer Comprobante (primer elemento)
            comprobante = partes[0]
            
            # 3. Identificar Fechas (DD-MM-YY)
            fechas = re.findall(r'\d{2}-\d{2}-\d{2}', linea)
            f_transac = fechas[0] if len(fechas) > 0 else ""
            f_pago = fechas[1] if len(fechas) > 1 else ""
            
            # 4. Identificar Asiento/Periodo (Ej: 843/202208)
            asiento = re.search(r'\d+/\d{6}', linea)
            asiento_val = asiento.group(0) if asiento else ""
            
            # 5. Capturar Nro. Planilla (el cero después de la fecha transac)
            # Buscamos el patrón: Fecha + 0
            nro_planilla = "0" if f_transac + " 0" in linea else ""
            
            # 6. Capturar Orden de Pago (números de 3-4 dígitos que no sean años)
            op_match = re.search(r'\s(\d{3,4})\s', linea)
            orden_pago = op_match.group(1) if op_match else ""

            # 7. El resto del texto entre el Asiento y el Débito es el Concepto
            concepto_match = re.search(f"{asiento_val or f_pago or nro_planilla}(.*?){debito}", linea)
            concepto = concepto_match.group(1).strip() if concepto_match else ""

            datos_finales.append({
                "Comprobante": comprobante,
                "Fecha Transac.": f_transac,
                "Nro. Planilla": nro_planilla,
                "Tipo Planilla": "",
                "Orden de Pago": orden_pago,
                "Fecha de Pago": f_pago,
                "Asiento/Periodo": asiento_val,
                "Descripción Concepto": concepto,
                "Débito": limpiar_monto(debito),
                "Crédito": limpiar_monto(credito),
                "Saldo": limpiar_monto(saldo)
            })

    df = pd.DataFrame(datos_finales)
    
    # --- DISEÑO CORPORATIVO ---
    with pd.ExcelWriter(ruta_excel, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Extracto Krona')
        ws = writer.sheets['Extracto Krona']
        
        # Estilos BMW/Perfecta
        header_fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")
            cell.border = border

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                cell.border = border
                if cell.column_letter in ['I', 'J', 'K']:
                    cell.number_format = '#,##0'
                    cell.alignment = Alignment(horizontal="right")

        anchos = {'A': 20, 'B': 15, 'C': 12, 'D': 12, 'E': 15, 'F': 15, 'G': 18, 'H': 50, 'I': 15, 'J': 15, 'K': 15}
        for col, ancho in anchos.items(): ws.column_dimensions[col].width = ancho

    print(f"\n✅ Proceso exitoso. Archivo generado: {ruta_excel}")

if __name__ == "__main__":
    procesar_extracto_perfecta()
