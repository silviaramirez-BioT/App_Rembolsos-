import os
import sys
import re
import io
import glob
import logging
import zipfile
import datetime
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape as xml_escape

# Silenciar advertencias de pdfminer y librerías de PDF
logging.getLogger('pdfminer').setLevel(logging.ERROR)
logging.getLogger('pypdfium2').setLevel(logging.ERROR)
logging.getLogger('pdfplumber').setLevel(logging.ERROR)

# Intentar importar librerías auxiliares para PDFs e imágenes (OCR)
try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    from rapidocr_onnxruntime import RapidOCR
    _ocr_engine = RapidOCR()
except Exception:
    _ocr_engine = None


def a_fecha_serial(fecha_iso):
    """
    Convierte una fecha 'YYYY-MM-DD' a número serial de Excel (base 1900).
    """
    if not fecha_iso:
        return None
    try:
        d = datetime.datetime.strptime(fecha_iso, '%Y-%m-%d').date()
        return (d - datetime.date(1899, 12, 30)).days
    except Exception:
        return None


def normalizar_fecha(texto_fecha):
    """
    Parsea diferentes formatos de texto de fecha a 'YYYY-MM-DD'.
    Soporta: YYYY-MM-DD, DD/MM/YYYY, DD/Mes/YYYY, Mes DD YY, etc.
    """
    if not texto_fecha:
        return ''
    
    texto_fecha = texto_fecha.strip()
    meses = {
        'ene': '01', 'feb': '02', 'mar': '03', 'abr': '04', 'may': '05', 'jun': '06',
        'jul': '07', 'ago': '08', 'sep': '09', 'oct': '10', 'nov': '11', 'dic': '12',
        'jan': '01', 'apr': '04', 'aug': '08', 'dec': '12'
    }

    # Formato ISO: YYYY-MM-DD
    m = re.search(r'(202\d)[-/](\d{1,2})[-/](\d{1,2})', texto_fecha)
    if m:
        y, mes, d = m.group(1), int(m.group(2)), int(m.group(3))
        return f"{y}-{mes:02d}-{d:02d}"

    # Formato DD/MM/YYYY
    m = re.search(r'(\d{1,2})[-/](\d{1,2})[-/](202\d)', texto_fecha)
    if m:
        d, mes, y = int(m.group(1)), int(m.group(2)), m.group(3)
        return f"{y}-{mes:02d}-{d:02d}"

    # Formato DD/Mes/YYYY o DD/Mes/YY (ej. 11/Ago/2026)
    m = re.search(r'(\d{1,2})[/-]([a-zA-Z]{3,4})[/-](202\d|\d{2})', texto_fecha)
    if m:
        d = int(m.group(1))
        m_txt = m.group(2).lower()[:3]
        mes_num = meses.get(m_txt, '08')
        y_raw = m.group(3)
        y = y_raw if len(y_raw) == 4 else f"20{y_raw}"
        return f"{y}-{mes_num}-{d:02d}"

    # Formato Mes DD YY (ej. AGO 10 26, AG0 12 26)
    m = re.search(r'(AG0|AGO|ENE|FEB|MAR|ABR|MAY|JUN|JUL|SEP|OCT|NOV|DIC)\s+(\d{1,2})\s+(\d{2})', texto_fecha, re.IGNORECASE)
    if m:
        m_txt = m.group(1).upper().replace('0', 'O')[:3].lower()
        mes_num = meses.get(m_txt, '08')
        d = int(m.group(2))
        y = f"20{m.group(3)}"
        return f"{y}-{mes_num}-{d:02d}"

    return ''


def extraer_datos_xml(ruta_xml):
    """
    Parsea el archivo XML CFDI (soporta CFDI 3.3 y 4.0) y extrae los campos requeridos
    para la plantilla de reembolsos.
    """
    try:
        tree = ET.parse(ruta_xml)
        root = tree.getroot()
        
        def strip_ns(tag):
            return tag.split('}')[-1] if '}' in tag else tag

        fecha_str = root.attrib.get('Fecha', '')
        total = float(root.attrib.get('Total', '0'))
        subtotal = float(root.attrib.get('SubTotal', '0'))
        serie = root.attrib.get('Serie', '').strip()
        folio = root.attrib.get('Folio', '').strip()
        
        if serie and folio:
            factura = f"{serie} {folio}"
        elif folio:
            factura = folio
        elif serie:
            factura = serie
        else:
            factura = ''

        emisor_nombre = ''
        emisor_rfc = ''
        conceptos_desc = []
        uuid = ''

        for elem in root.iter():
            tag = strip_ns(elem.tag)
            if tag == 'Emisor':
                emisor_nombre = elem.attrib.get('Nombre', '').strip()
                emisor_rfc = elem.attrib.get('Rfc', '').strip()
            elif tag == 'Concepto':
                desc = elem.attrib.get('Descripcion', '').strip()
                if desc:
                    conceptos_desc.append(desc)
            elif tag == 'TimbreFiscalDigital':
                uuid = elem.attrib.get('UUID', '').strip()

        concepto = ' / '.join(conceptos_desc) if conceptos_desc else 'Gasto'
        comentario = emisor_nombre if emisor_nombre else emisor_rfc
        fecha_iso = normalizar_fecha(fecha_str.split('T')[0] if 'T' in fecha_str else fecha_str)
        fecha_serial = a_fecha_serial(fecha_iso)

        return {
            'origen': 'xml',
            'archivo': os.path.basename(ruta_xml),
            'ruta_completa': ruta_xml,
            'concepto': concepto,
            'fecha_iso': fecha_iso,
            'fecha_serial': fecha_serial,
            'monto': total,
            'subtotal': subtotal,
            'comentario': comentario,
            'factura': factura,
            'uuid': uuid,
            'rfc_emisor': emisor_rfc,
            'nombre_emisor': emisor_nombre
        }

    except Exception as e:
        print(f"  [ERROR] Error procesando XML '{ruta_xml}': {e}")
        return None


def extraer_datos_pdf_cfdi(ruta_pdf):
    """
    Parsea una factura o resumen de facturas en PDF (CFDI) extrayendo los campos
    relevantes cuando no se cuenta con el XML individual.
    Soporta documentos de una y múltiples páginas (ej. resúmenes de viajes).
    """
    if not pdfplumber:
        return []

    records = []
    try:
        with pdfplumber.open(ruta_pdf) as pdf:
            for page_idx, page in enumerate(pdf.pages):
                p_text = page.extract_text() or ''
                if not p_text.strip():
                    continue

                # 1. UUID / Folio Fiscal
                # Prioridad 1: Cadena Original ||1.1|UUID|
                uuid = ''
                m_cadena_uuid = re.search(r'\|\|1\.1\|([0-9a-fA-F\-]{36})\|', p_text)
                if m_cadena_uuid:
                    uuid = m_cadena_uuid.group(1).strip()
                else:
                    m_ff = re.search(r'Folio Fiscal(?:\s*\(UUID\))?[:\s]+([0-9a-fA-F\-]{36})', p_text, re.IGNORECASE)
                    if m_ff:
                        uuid = m_ff.group(1).strip()
                    else:
                        m_u = re.search(r'([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})', p_text)
                        if m_u:
                            uuid = m_u.group(1).strip()

                # Si no tiene UUID y no parece factura, omitir
                if not uuid and 'CFDI' not in p_text and 'Factura' not in p_text and 'Comprobante' not in p_text:
                    continue

                # 2. Total (evitar coincidir con 'SubTotal')
                total = 0.0
                m_tot = re.findall(r'(?<!Sub)(?<!SUB)Total(?:\s+Comprobante|\s+M\.?N\.?)?\s*[:\$]?\s*\$?\s*([\d,]+\.\d{2})', p_text, re.IGNORECASE)
                if m_tot:
                    vals = [float(x.replace(',', '')) for x in m_tot]
                    total = max(vals)
                else:
                    # Buscar líneas con Total
                    m_money = re.findall(r'(?<!Sub)Total\s*\n\s*([\d,]+\.\d{2})', p_text, re.IGNORECASE)
                    if m_money:
                        total = float(m_money[-1].replace(',', ''))
                    else:
                        m_all_tot = re.findall(r'\$\s*([\d,]+\.\d{2})', p_text)
                        if m_all_tot:
                            total = float(m_all_tot[-1].replace(',', ''))

                # 3. Fecha
                fecha_iso = ''
                m_cadena_fec = re.search(r'\|\|1\.1\|[0-9a-fA-F\-]+\|(\d{4}-\d{2}-\d{2})', p_text)
                if m_cadena_fec:
                    fecha_iso = m_cadena_fec.group(1)
                else:
                    m_fec = re.search(r'Fecha(?:\s+de\s+Expedici[oó]n)?[:\s]+(\d{4}-\d{2}-\d{2}|\d{2}/\d{2}/\d{4})', p_text, re.IGNORECASE)
                    if m_fec:
                        fecha_iso = normalizar_fecha(m_fec.group(1))

                # 4. Emisor
                emisor = ''
                if 'SERVICIOS GASOLINEROS DE MEXICO' in p_text.upper():
                    emisor = 'SERVICIOS GASOLINEROS DE MEXICO'
                elif 'BM DEL NORTE' in p_text.upper():
                    emisor = 'BM DEL NORTE'
                elif 'ORSAN DEL NORTE' in p_text.upper():
                    emisor = 'ORSAN DEL NORTE'
                elif 'CONCESIONARIA AUTOPISTA MONTERREY-SALTILLO' in p_text.upper():
                    emisor = 'CONCESIONARIA AUTOPISTA MONTERREY-SALTILLO'
                elif 'FACTURIFY' in p_text.upper():
                    emisor = 'FACTURIFY'
                else:
                    m_em = re.search(r'EMISOR\s*\n\s*([^\n]+)', p_text, re.IGNORECASE)
                    if m_em:
                        emisor = m_em.group(1).strip()
                    else:
                        m_em2 = re.search(r'Emisor:\s*\n\s*([^\n]+)', p_text, re.IGNORECASE)
                        if m_em2:
                            emisor = m_em2.group(1).strip()

                # 5. Factura (Serie / Folio)
                factura = ''
                m_fac = re.search(r'Factura:\s*([^\n]+)', p_text, re.IGNORECASE)
                if m_fac:
                    factura = m_fac.group(1).strip()
                elif 'CLI' in p_text:
                    m_cli = re.search(r'(CLI\s*\d+)', p_text)
                    if m_cli:
                        factura = m_cli.group(1).strip()
                elif 'Serie-Folio:' in p_text:
                    m_sf = re.search(r'Serie-Folio:\s*([^\n]+)', p_text)
                    if m_sf:
                        factura = m_sf.group(1).strip()

                # 6. Concepto
                concepto = 'Gasto'
                p_upper = p_text.upper()
                if 'MAGNA' in p_upper:
                    concepto = 'MAGNA (LT)'
                elif 'PREMIUM' in p_upper:
                    concepto = 'PREMIUM (LT)'
                elif 'EXTRA' in p_upper:
                    concepto = 'EXTRA'
                elif 'PEAJE' in p_upper or 'AUTOPISTA' in p_upper:
                    concepto = 'Peaje'
                elif 'ACUMULADOR' in p_upper:
                    m_ac = re.search(r'(ACUMULADOR[^\n]+)', p_text, re.IGNORECASE)
                    concepto = m_ac.group(1).strip() if m_ac else 'ACUMULADOR HI-TEC 12V 730 CCA'
                elif 'ALQUILER DE VEHICULOS' in p_upper or 'TRANSPORTE DE PASAJEROS' in p_upper or 'TARIFA' in p_upper:
                    concepto = 'ALQUILER DE VEHICULOS (RENTA DE VEHICULOS, SERVICIOS PRIVADOS DE TRANSPORTE DE PASAJEROS)'

                if total > 0 or uuid:
                    fecha_serial = a_fecha_serial(fecha_iso)
                    records.append({
                        'origen': 'pdf_cfdi',
                        'archivo': os.path.basename(ruta_pdf),
                        'ruta_completa': ruta_pdf,
                        'page': page_idx + 1,
                        'concepto': concepto,
                        'fecha_iso': fecha_iso,
                        'fecha_serial': fecha_serial,
                        'monto': total,
                        'subtotal': total,
                        'comentario': emisor,
                        'factura': factura,
                        'uuid': uuid,
                        'rfc_emisor': '',
                        'nombre_emisor': emisor
                    })

    except Exception as e:
        print(f"  [ERROR] Error procesando PDF '{ruta_pdf}': {e}")

    return records


def extraer_datos_ticket(ruta_imagen_o_pdf):
    """
    Extrae la información de un ticket o recibo (foto o PDF escaneado)
    mediante OCR (RapidOCR) o texto del documento.
    """
    full_text = ''
    text_lines = []

    # Si es un PDF, intentar extraer texto directo o renderizar imagen para OCR
    if ruta_imagen_o_pdf.lower().endswith('.pdf'):
        if pdfplumber:
            try:
                with pdfplumber.open(ruta_imagen_o_pdf) as pdf:
                    for page in pdf.pages:
                        t = page.extract_text() or ''
                        if t.strip():
                            full_text += t + '\n'
            except Exception:
                pass

    # Si no hay suficiente texto o es imagen (.jpg, .jpeg, .png, etc.), aplicar OCR
    if (not full_text.strip() or not ruta_imagen_o_pdf.lower().endswith('.pdf')) and _ocr_engine:
        try:
            if ruta_imagen_o_pdf.lower().endswith('.pdf') and pdfplumber:
                with pdfplumber.open(ruta_imagen_o_pdf) as pdf:
                    for page in pdf.pages:
                        img = page.to_image(resolution=200).original
                        import numpy as np
                        result, _ = _ocr_engine(np.array(img))
                        if result:
                            lines = [r[1].strip() for r in result if r[1].strip()]
                            text_lines.extend(lines)
                            full_text += ' \n '.join(lines) + '\n'
            else:
                result, _ = _ocr_engine(ruta_imagen_o_pdf)
                if result:
                    lines = [r[1].strip() for r in result if r[1].strip()]
                    text_lines.extend(lines)
                    full_text = ' \n '.join(lines)
        except Exception as e:
            print(f"  [AVISO] Error OCR en '{ruta_imagen_o_pdf}': {e}")

    if not full_text.strip():
        return None

    # 1. Total / Monto (evitar coincidir con Subtotal)
    total = 0.0
    m_tot = re.findall(r'(?<!Sub)(?<!SUB)Total(?:\s+M\.?N\.?|\s+Comprobante)?\s*[:\$]?\s*\$?\s*([\d,]+\.\d{2})', full_text, re.IGNORECASE)
    if m_tot:
        for t in m_tot:
            val = float(t.replace(',', ''))
            if val > total:
                total = val

    if total == 0.0:
        m_money = re.findall(r'\$\s*([\d,]+\.\d{2})', full_text)
        if m_money:
            vals = [float(x.replace(',', '')) for x in m_money]
            total = max(vals)

    # 2. Fecha
    fecha_iso = normalizar_fecha(full_text)
    fecha_serial = a_fecha_serial(fecha_iso)

    # 3. Emisor / Proveedor
    f_upper = full_text.upper()
    emisor = ''
    if 'OXXOGAS' in f_upper or 'SERVICIOS GASOLINEROS' in f_upper:
        emisor = 'SERVICIOS GASOLINEROS DE MEXICO'
    elif 'BATTERY MASTER' in f_upper or 'BM DEL NORTE' in f_upper:
        emisor = 'BM DEL NORTE'
    elif 'ORSAN' in f_upper or 'MOBIL' in f_upper:
        emisor = 'ORSAN DEL NORTE'
    elif 'AUTOPISTA' in f_upper or 'CAMS' in f_upper:
        emisor = 'CONCESIONARIA AUTOPISTA MONTERREY-SALTILLO'
    elif 'UBER' in f_upper or 'DIDI' in f_upper or 'FACTURIFY' in f_upper:
        emisor = 'FACTURIFY'
    else:
        emisor = text_lines[0] if text_lines else 'Gasto'

    # 4. Concepto
    concepto = 'Gasto'
    if 'MAGNA' in f_upper:
        concepto = 'MAGNA (LT)'
    elif 'PREMIUM' in f_upper:
        concepto = 'PREMIUM (LT)'
    elif 'EXTRA' in f_upper:
        concepto = 'EXTRA'
    elif 'ACUMULADOR' in f_upper:
        m_ac = re.search(r'ACUMULADOR[^\n]+', full_text, re.IGNORECASE)
        concepto = m_ac.group(0).strip() if m_ac else 'ACUMULADOR HI-TEC 12V 730 CCA'
    elif 'AUTOPISTA' in f_upper or 'TRAYECTO' in f_upper or 'PEAJE' in f_upper:
        concepto = 'Peaje'
    elif 'ALQUILER' in f_upper or 'TRANSPORTE' in f_upper:
        concepto = 'ALQUILER DE VEHICULOS'

    # 5. Folio ticket / Factura
    folio = ''
    m_cli = re.search(r'(CLI\s*[-–]?\s*\d+)', full_text, re.IGNORECASE)
    if m_cli:
        folio = m_cli.group(1).replace('–', '-').replace('-', ' ').strip()
    else:
        m_fol = re.search(r'(?:Folio|Ticket|N\.sec\.?)\s*[:#]?\s*([A-Za-z0-9\-_/]+)', full_text, re.IGNORECASE)
        if m_fol:
            folio = f"Ticket {m_fol.group(1).strip()}"
        else:
            folio = "Ticket"

    # Folio Fiscal (si el ticket impreso trae el timbre fiscal)
    uuid = ''
    m_u = re.search(r'([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})', full_text)
    if m_u:
        uuid = m_u.group(1).strip()

    if total == 0.0:
        return None

    return {
        'origen': 'ticket',
        'archivo': os.path.basename(ruta_imagen_o_pdf),
        'ruta_completa': ruta_imagen_o_pdf,
        'concepto': concepto,
        'fecha_iso': fecha_iso,
        'fecha_serial': fecha_serial,
        'monto': total,
        'subtotal': total,
        'comentario': emisor,
        'factura': folio,
        'uuid': uuid,
        'rfc_emisor': '',
        'nombre_emisor': emisor
    }


def desduplicar_y_unificar_gastos(registros_xml, registros_pdf, registros_tickets):
    """
    Combina los registros provenientes de XMLs, PDFs y Tickets, evitando duplicados:
    - Si existe el XML oficial, este tiene la máxima prioridad.
    - Si una factura PDF corresponde al mismo UUID que un XML, se omite el duplicado.
    - Si un ticket corresponde a una factura (por UUID, folio/número o monto idéntico + proveedor/fecha),
      se evita duplicar el gasto.
    - Si un ticket no tiene factura correspondiente, se incluye como gasto por comprobar/reembolsar.
    """
    final_list = []
    uuids_registrados = set()

    def extraer_digitos_folio(folio_str):
        if not folio_str:
            return ''
        d = re.sub(r'\D', '', folio_str)
        return d if len(d) >= 4 else ''

    # 1. Agregar XMLs (Fuente primaria más exacta)
    for r in registros_xml:
        final_list.append(r)
        if r.get('uuid'):
            uuids_registrados.add(r['uuid'].lower())

    # 2. Agregar Facturas PDF que no estén en XMLs
    for r in registros_pdf:
        u = (r.get('uuid') or '').lower()
        if u and u in uuids_registrados:
            continue
        
        # Verificar coincidencia por Folio o por Monto exacto + Fecha
        ya_existe = False
        r_digitos = extraer_digitos_folio(r.get('factura', ''))
        for ex in final_list:
            ex_digitos = extraer_digitos_folio(ex.get('factura', ''))
            if (r_digitos and ex_digitos and r_digitos == ex_digitos) or (
                abs(ex['monto'] - r['monto']) < 0.01 and (
                    (ex['fecha_iso'] and ex['fecha_iso'] == r['fecha_iso']) or
                    (ex['comentario'] and r['comentario'] and ex['comentario'].lower() == r['comentario'].lower())
                )
            ):
                ya_existe = True
                break
        
        if not ya_existe:
            final_list.append(r)
            if u:
                uuids_registrados.add(u)

    # 3. Agregar Tickets que no correspondan a facturas ya añadidas
    for t in registros_tickets:
        u = (t.get('uuid') or '').lower()
        if u and u in uuids_registrados:
            continue
        
        ya_cubierto = False
        t_digitos = extraer_digitos_folio(t.get('factura', ''))
        for ex in final_list:
            ex_digitos = extraer_digitos_folio(ex.get('factura', ''))
            # Coincidencia por Folio numérico (ej. CLI 63308140)
            if t_digitos and ex_digitos and t_digitos == ex_digitos:
                ya_cubierto = True
                break

            # Coincidencia por Monto exacto (o subtotal) + Proveedor o Fecha
            if abs(ex['monto'] - t['monto']) < 0.01 or abs(ex.get('subtotal', 0.0) - t['monto']) < 0.01:
                if (ex['comentario'] and t['comentario'] and (
                    ex['comentario'].lower() in t['comentario'].lower() or
                    t['comentario'].lower() in ex['comentario'].lower()
                )) or (ex['fecha_iso'] and t['fecha_iso'] and ex['fecha_iso'] == t['fecha_iso']):
                    ya_cubierto = True
                    # Si el XML no tenía fecha de recibo pero el ticket sí, completar la fecha
                    if not ex.get('fecha_serial') and t.get('fecha_serial'):
                        ex['fecha_serial'] = t['fecha_serial']
                        ex['fecha_iso'] = t['fecha_iso']
                    break
        
        if not ya_cubierto:
            final_list.append(t)
            if u:
                uuids_registrados.add(u)

    # Ordenar cronológicamente por fecha
    final_list.sort(key=lambda x: (x.get('fecha_iso', '') or '9999-99-99', x.get('factura', '')))
    return final_list


def encontrar_plantilla_base(carpeta_raiz):
    """
    Busca un archivo de plantilla Excel limpio en la estructura de carpetas.
    Prioriza carpetas como 'zExtra' o archivos de plantilla 'Reembolsos XX'.
    """
    for root_dir, dirs, files in os.walk(carpeta_raiz):
        if 'zextra' in root_dir.lower():
            for f in files:
                if f.lower().endswith('.xlsx') and not f.startswith('~$'):
                    return os.path.join(root_dir, f)

    for root_dir, dirs, files in os.walk(carpeta_raiz):
        for f in files:
            if f.lower().endswith('.xlsx') and not f.startswith('~$'):
                if ' xx ' in f.lower() or 'plantilla' in f.lower():
                    return os.path.join(root_dir, f)

    for root_dir, dirs, files in os.walk(carpeta_raiz):
        for f in files:
            if f.lower().endswith('.xlsx') and not f.startswith('~$'):
                return os.path.join(root_dir, f)
    return None


def generar_excel_reembolso(plantilla_bytes, registros):
    """
    Genera un archivo Excel (.xlsx) válido inyectando los registros en la tabla Table1.
    Mantiene intactos todos los estilos, fuentes, bordes, formatos de celda y fórmulas.
    """
    N = len(registros)
    total_sum = sum(r['monto'] for r in registros) if registros else 0.0

    in_zip = zipfile.ZipFile(io.BytesIO(plantilla_bytes), 'r')
    out_buf = io.BytesIO()
    out_zip = zipfile.ZipFile(out_buf, 'w', compression=zipfile.ZIP_DEFLATED)

    rows_xml = []
    # Fila 1: Encabezados de tabla
    rows_xml.append('<row r="1" spans="1:8" x14ac:dyDescent="0.25"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>2</v></c><c r="C1" t="s"><v>3</v></c><c r="D1" t="s"><v>8</v></c><c r="E1" s="1" t="s"><v>1</v></c><c r="F1" t="s"><v>5</v></c><c r="G1" t="s"><v>6</v></c><c r="H1" t="s"><v>7</v></c></row>')

    if N > 0:
        for i, rec in enumerate(registros):
            r = i + 2
            concepto_safe = xml_escape(str(rec.get('concepto', '')))
            comentario_safe = xml_escape(str(rec.get('comentario', '')))
            factura_safe = xml_escape(str(rec.get('factura', '')))
            uuid_safe = xml_escape(str(rec.get('uuid', '')))
            
            c_a = f'<c r="A{r}" t="inlineStr"><is><t>{concepto_safe}</t></is></c>'
            
            # Fechas (serial numérico para aplicar formato de fecha Excel s="3", o inlineStr como fallback)
            if rec.get('fecha_serial'):
                c_b = f'<c r="B{r}" s="3"><v>{rec["fecha_serial"]}</v></c>'
                c_c = f'<c r="C{r}" s="3"><v>{rec["fecha_serial"]}</v></c>'
            elif rec.get('fecha_iso'):
                c_b = f'<c r="B{r}" s="3" t="inlineStr"><is><t>{rec["fecha_iso"]}</t></is></c>'
                c_c = f'<c r="C{r}" s="3" t="inlineStr"><is><t>{rec["fecha_iso"]}</t></is></c>'
            else:
                c_b = f'<c r="B{r}" s="3"/>'
                c_c = f'<c r="C{r}" s="3"/>'
                
            c_d = f'<c r="D{r}" s="3"/>'
            c_e = f'<c r="E{r}" s="1"><v>{rec["monto"]:.2f}</v></c>'
            c_f = f'<c r="F{r}" t="inlineStr"><is><t>{comentario_safe}</t></is></c>'
            c_g = f'<c r="G{r}" s="4" t="inlineStr"><is><t>{factura_safe}</t></is></c>'
            c_h = f'<c r="H{r}" s="4" t="inlineStr"><is><t>{uuid_safe}</t></is></c>'
            
            rows_xml.append(f'<row r="{r}" spans="1:8" x14ac:dyDescent="0.25">{c_a}{c_b}{c_c}{c_d}{c_e}{c_f}{c_g}{c_h}</row>')

        r_tot = N + 2
        r_dim = N + 3
        rows_xml.append(f'<row r="{r_tot}" spans="1:8" x14ac:dyDescent="0.25"><c r="A{r_tot}" t="s"><v>4</v></c><c r="E{r_tot}" s="2"><f>SUBTOTAL(109,Table1[Monto])</f><v>{total_sum:.2f}</v></c><c r="F{r_tot}" s="5"/></row>')
        rows_xml.append(f'<row r="{r_dim}" spans="1:8" x14ac:dyDescent="0.25"><c r="F{r_dim}" s="5"/></row>')
    else:
        # Plantilla vacía
        for r in range(2, 8):
            rows_xml.append(f'<row r="{r}" spans="1:8" x14ac:dyDescent="0.25"><c r="B{r}" s="3"/><c r="C{r}" s="3"/><c r="D{r}" s="3"/><c r="G{r}" s="4"/><c r="H{r}" s="4"/></row>')
        r_tot = 8
        r_dim = 9
        rows_xml.append(f'<row r="8" spans="1:8" x14ac:dyDescent="0.25"><c r="A8" t="s"><v>4</v></c><c r="E8" s="2"><f>SUBTOTAL(109,Table1[Monto])</f><v>0</v></c><c r="F8" s="5"/></row>')
        rows_xml.append(f'<row r="9" spans="1:8" x14ac:dyDescent="0.25"><c r="F9" s="5"/></row>')

    sheetData_content = ''.join(rows_xml)

    # Actualizar sheet1.xml
    sheet_xml_str = in_zip.read('xl/worksheets/sheet1.xml').decode('utf-8')
    sheet_xml_str = re.sub(r'<dimension ref="[^"]+"/>', f'<dimension ref="A1:H{r_dim}"/>', sheet_xml_str)
    sheet_xml_str = re.sub(r'<sheetData>.*?</sheetData>', f'<sheetData>{sheetData_content}</sheetData>', sheet_xml_str, flags=re.DOTALL)

    # Actualizar table1.xml
    table_xml_str = in_zip.read('xl/tables/table1.xml').decode('utf-8')
    table_xml_str = re.sub(r'ref="[^"]+"', f'ref="A1:H{r_tot}"', table_xml_str, count=1)
    table_xml_str = re.sub(r'<autoFilter ref="[^"]+"', f'<autoFilter ref="A1:H{r_tot-1}"', table_xml_str, count=1)

    # Actualizar calcChain.xml
    calc_xml_str = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="E{r_tot}" i="1" l="1"/></calcChain>'

    for item in in_zip.infolist():
        if item.filename == 'xl/worksheets/sheet1.xml':
            out_zip.writestr(item, sheet_xml_str.encode('utf-8'))
        elif item.filename == 'xl/tables/table1.xml':
            out_zip.writestr(item, table_xml_str.encode('utf-8'))
        elif item.filename == 'xl/calcChain.xml':
            out_zip.writestr(item, calc_xml_str.encode('utf-8'))
        else:
            out_zip.writestr(item, in_zip.read(item.filename))

    out_zip.close()
    in_zip.close()

    return out_buf.getvalue()


def procesar_reembolsos(carpeta_base):
    """
    Recorre las carpetas individuales de cada trabajador, extrae los comprobantes
    (XMLs, PDFs de facturas, imágenes y PDFs de tickets), y actualiza el archivo Excel
    en la carpeta de cada trabajador sin unificarlos.
    """
    carpeta_base = os.path.abspath(carpeta_base)
    print("=" * 75)
    print(f"PROCESANDO REEMBOLSOS EN: {carpeta_base}")
    print("=" * 75)

    if not os.path.isdir(carpeta_base):
        print(f"Error: La carpeta '{carpeta_base}' no existe.")
        return

    # Obtener plantilla base
    ruta_plantilla = encontrar_plantilla_base(carpeta_base)
    if not ruta_plantilla:
        print("Error: No se encontró ningún archivo de plantilla Excel (.xlsx).")
        return

    print(f"Plantilla base utilizada: {os.path.relpath(ruta_plantilla, carpeta_base)}\n")
    with open(ruta_plantilla, 'rb') as f:
        plantilla_bytes = f.read()

    # Identificar carpetas de trabajadores (subdirectorios directos de carpeta_base)
    subdirectorios = [d for d in os.listdir(carpeta_base) if os.path.isdir(os.path.join(carpeta_base, d))]
    
    if not subdirectorios:
        carpetas_a_procesar = [carpeta_base]
    else:
        carpetas_a_procesar = [os.path.join(carpeta_base, d) for d in sorted(subdirectorios)]

    total_trabajadores = 0
    total_registros_global = 0
    total_monto_global = 0.0

    for dir_path in carpetas_a_procesar:
        nombre_carpeta = os.path.basename(dir_path)
        
        # Ignorar carpeta de plantilla auxiliar 'zExtra'
        if nombre_carpeta.lower() == 'zextra':
            continue

        archivos = os.listdir(dir_path)
        archivos_xml = [f for f in archivos if f.lower().endswith('.xml')]
        archivos_pdf = [f for f in archivos if f.lower().endswith('.pdf')]
        archivos_img = [f for f in archivos if f.lower().endswith(('.jpg', '.jpeg', '.png', '.webp'))]
        archivos_excel = [f for f in archivos if f.lower().endswith('.xlsx') and not f.startswith('~$')]

        # 1. Extraer registros XML
        registros_xml = []
        for xml_file in archivos_xml:
            ruta_xml = os.path.join(dir_path, xml_file)
            datos = extraer_datos_xml(ruta_xml)
            if datos:
                registros_xml.append(datos)

        # 2. Extraer registros PDF (Facturas CFDI)
        registros_pdf = []
        for pdf_file in archivos_pdf:
            ruta_pdf = os.path.join(dir_path, pdf_file)
            datos_pdf_lista = extraer_datos_pdf_cfdi(ruta_pdf)
            registros_pdf.extend(datos_pdf_lista)

        # 3. Extraer registros de Tickets (Imágenes y PDFs de reportes/tickets)
        registros_tickets = []
        for img_file in archivos_img:
            ruta_img = os.path.join(dir_path, img_file)
            datos_ticket = extraer_datos_ticket(ruta_img)
            if datos_ticket:
                registros_tickets.append(datos_ticket)

        for pdf_file in archivos_pdf:
            # Si el PDF no fue reconocido como CFDI electrónico, intentar como ticket escaneado
            ruta_pdf = os.path.join(dir_path, pdf_file)
            if not any(r['archivo'] == pdf_file for r in registros_pdf):
                datos_ticket = extraer_datos_ticket(ruta_pdf)
                if datos_ticket:
                    registros_tickets.append(datos_ticket)

        # 4. Desduplicar y consolidar para este trabajador
        registros_finales = desduplicar_y_unificar_gastos(registros_xml, registros_pdf, registros_tickets)

        # 5. Determinar el archivo Excel destino en la carpeta de este trabajador
        if archivos_excel:
            # Mantener el nombre del archivo Excel existente en su carpeta
            nombre_excel_destino = archivos_excel[0]
        else:
            # Si no existía, asignarle el nombre estándar con su nombre
            nombre_excel_destino = f"Reembolsos {nombre_carpeta} WXX 2026XXXX.xlsx"

        ruta_excel_destino = os.path.join(dir_path, nombre_excel_destino)

        # Generar contenido del nuevo Excel
        nuevo_excel_bytes = generar_excel_reembolso(plantilla_bytes, registros_finales)

        # Escribir el archivo Excel en la carpeta del trabajador
        with open(ruta_excel_destino, 'wb') as f:
            f.write(nuevo_excel_bytes)

        monto_trabajador = sum(r['monto'] for r in registros_finales)
        print(f"Trabajador: {nombre_carpeta}")
        print(f"  -> Archivo Excel: {nombre_excel_destino}")
        print(f"  -> Registros procesados: {len(registros_finales)} (XMLs: {len(registros_xml)}, PDFs: {len(registros_pdf)}, Tickets: {len(registros_tickets)})")
        print(f"  -> Total reembolsos: ${monto_trabajador:,.2f}")
        for r in registros_finales:
            origen_badge = f"[{r['origen'].upper()}]"
            fec = r['fecha_iso'] or 'Sin fecha'
            fac = r['factura'] or 'S/F'
            em = r['comentario'] or r['concepto']
            print(f"     • {origen_badge:<10} {fec} | ${r['monto']:>8.2f} | {em[:30]:<30} | {fac}")
        print()

        total_trabajadores += 1
        total_registros_global += len(registros_finales)
        total_monto_global += monto_trabajador

    print("=" * 75)
    print("PROCESO COMPLETADO EXITOSAMENTE")
    print(f"Total trabajadores procesados: {total_trabajadores}")
    print(f"Total comprobantes registrados: {total_registros_global}")
    print(f"Monto total reembolsos: ${total_monto_global:,.2f}")
    print("=" * 75)


if __name__ == '__main__':
    if len(sys.argv) > 1:
        ruta_carpeta = sys.argv[1]
    else:
        carpeta_script = os.path.dirname(os.path.abspath(__file__))
        candidato = os.path.join(carpeta_script, 'W33 Agosto 16')
        if os.path.exists(candidato):
            ruta_carpeta = candidato
        else:
            ruta_carpeta = carpeta_script

    procesar_reembolsos(ruta_carpeta)
