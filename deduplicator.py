# deduplicator.py

import openpyxl
from openpyxl.styles import Font, Alignment, NamedStyle
from difflib import SequenceMatcher
from collections import defaultdict
import re
import datetime
from copy import deepcopy

# --- TODAS TUS FUNCIONES AUXILIARES INTACTAS ---
def norm_key(text): return re.sub(r'\W+', '', str(text).lower().strip()) if text else ""
def convert_html_entities(text):
    if not isinstance(text, str): return text
    html_entities = {'á':'á','é':'é','í':'í','ó':'ó','ú':'ú','ñ':'ñ','Á':'Á','É':'É','Í':'Í','Ó':'Ó','Ú':'Ú','Ñ':'Ñ','"':'"','"':'"','"':'"',''':"'",''':"'",'Â':'','â':'','€':'','™':''}
    for entity, char in html_entities.items(): text = text.replace(entity, char)
    return text
def normalize_title(title):
    if not isinstance(title, str): return ""
    title = convert_html_entities(title); title = re.sub(r'\s*\|\s*[\w\s]+$', '', title)
    return re.sub(r'\W+', ' ', title).lower().strip()
def corregir_texto(text):
    if not isinstance(text, str): return text
    text = convert_html_entities(text); text = re.sub(r'(<br>|\[\.\.\.\]|\s+)', ' ', text).strip()
    if match := re.search(r'[A-Z]', text): text = text[match.start():]
    if text and not text.endswith('...'): text = text.rstrip('.') + '...'
    return text
def extract_link(cell):
    if cell.hyperlink: return {"value": "Link", "url": cell.hyperlink.target}
    if cell.value and isinstance(cell.value, str):
        if match := re.search(r'=HYPERLINK\("([^"]+)"', cell.value): return {"value": "Link", "url": match.group(1)}
    return {"value": cell.value, "url": None}
def parse_date(fecha):
    if isinstance(fecha, datetime.datetime): return fecha.date()
    try: return datetime.datetime.strptime(str(fecha).split(" ")[0], "%Y-%m-%d").date()
    except (ValueError, TypeError): return None
def format_date_str(fecha_obj):
    if isinstance(fecha_obj, datetime.date): return fecha_obj.isoformat()
    return str(fecha_obj)[:10]
def es_internet(row): return norm_key(row.get(norm_key('Tipo de Medio'))) == 'internet'
def es_radio_o_tv(row): return norm_key(row.get(norm_key('Tipo de Medio'))) in {'radio', 'televisión'}
def mark_as_duplicate_to_delete(row):
    row['Mantener'] = "Eliminar"; row[norm_key('Tono')] = "Duplicada"; row[norm_key('Tema')] = "-"; row[norm_key('Temas Generales - Tema')] = "-"
def is_title_problematic(title):
    if not isinstance(title, str): return False
    if re.search(r'\s*\|\s*[\w\s]+$', title): return True
    if re.search(r'[Ââ€™"""'']', title): return True
    return False

# --- Función Principal CORREGIDA ---
def run_deduplication_process(wb, empresa_dict, internet_dict, region_dict):
    sheet = wb.active
    custom_link_style = NamedStyle(name="CustomLink", font=Font(color="000000", underline="none"), alignment=Alignment(horizontal="left"), number_format='@')
    if "CustomLink" not in wb.named_styles: wb.add_named_style(custom_link_style)
        
    headers = [cell.value for cell in sheet[1]]
    headers_norm = [norm_key(h) for h in headers]
    processed_rows = []

    for row_idx, row_cells in enumerate(sheet.iter_rows(min_row=2)):
        if all(c.value is None for c in row_cells): continue
        base_data = {headers_norm[i]: extract_link(cell) if headers_norm[i] in [norm_key('Link Nota'), norm_key('Link (Streaming - Imagen)')] else cell.value for i, cell in enumerate(row_cells)}
        
        base_data[norm_key('Título')] = convert_html_entities(str(base_data.get(norm_key('Título'), '')))
        base_data[norm_key('Resumen - Aclaracion')] = corregir_texto(base_data.get(norm_key('Resumen - Aclaracion')))
        tipo_medio_key = norm_key('Tipo de Medio'); tm_norm = norm_key(base_data.get(tipo_medio_key))
        if tm_norm in {'aire', 'cable'}: base_data[tipo_medio_key] = 'Televisión'
        elif tm_norm in {'am', 'fm'}: base_data[tipo_medio_key] = 'Radio'
        elif tm_norm == 'diario': base_data[tipo_medio_key] = 'Prensa'
        elif tm_norm == 'online': base_data[tipo_medio_key] = 'Internet'
        elif tm_norm == 'revista': base_data[tipo_medio_key] = 'Revista'
        
        link_nota_key, link_streaming_key = norm_key("Link Nota"), norm_key("Link (Streaming - Imagen)")
        tipo_medio_val = base_data.get(tipo_medio_key)
        if tipo_medio_val == "Internet": base_data[link_nota_key], base_data[link_streaming_key] = base_data.get(link_streaming_key), base_data.get(link_nota_key)
        elif tipo_medio_val in {"Prensa", "Revista"}:
            if (not base_data.get(link_nota_key) or not base_data.get(link_nota_key, {}).get('url')) and base_data.get(link_streaming_key, {}).get('url'): base_data[link_nota_key] = base_data.get(link_streaming_key)
            base_data[link_streaming_key] = None
        elif tipo_medio_val in {"Radio", "Televisión"}: base_data[link_streaming_key] = None
        
        menciones_key = norm_key('Menciones - Empresa'); menciones_str = str(base_data.get(menciones_key) or '')
        menciones = [m.strip() for m in menciones_str.split(';') if m.strip()]
        if not menciones:
            processed_rows.append(base_data)
        else:
            for mencion in menciones:
                new_row = deepcopy(base_data)
                mencion_limpia = mencion.lower().strip()
                new_row[menciones_key] = empresa_dict.get(mencion_limpia, mencion)
                processed_rows.append(new_row)

    # APLICAR MAPEOS DE INTERNET Y REGIÓN DESPUÉS DE LA EXPANSIÓN PERO ANTES DE DEDUPLICACIÓN
    medio_key = norm_key('Medio')
    tipo_medio_key = norm_key('Tipo de Medio')
    region_key = norm_key('Región')
    
    for row in processed_rows:
        # Mapeo de Internet
        if str(row.get(tipo_medio_key, '')).lower().strip() == 'internet':
            medio_val = str(row.get(medio_key, '')).lower().strip()
            if medio_val in internet_dict:
                row[medio_key] = internet_dict[medio_val]
        
        # Mapeo de Región (aplicar después del mapeo de Internet para usar el nombre correcto)
        medio_actual_val = str(row.get(medio_key, '')).lower().strip()
        row[region_key] = region_dict.get(medio_actual_val, "Online")

    # Inicializar campos de duplicación
    for row in processed_rows: row.update({'Duplicada': "FALSE", 'Posible Duplicada': "FALSE", 'Mantener': "Conservar"})

    # --- FASE 1: DUPLICADOS EXACTOS ---
    grupos_exactos = defaultdict(list)
    for idx, row in enumerate(processed_rows):
        key_tuple = (normalize_title(row.get(norm_key('Título'))), norm_key(row.get(norm_key('Medio'))), format_date_str(parse_date(row.get(norm_key('Fecha')))), norm_key(row.get(norm_key('Menciones - Empresa'))))
        if es_radio_o_tv(row): key_tuple += (str(row.get(norm_key('Hora'))),)
        grupos_exactos[key_tuple].append(idx)
    for indices in grupos_exactos.values():
        if len(indices) > 1:
            indices.sort(key=lambda i: (not is_title_problematic(processed_rows[i].get(norm_key('Título'))), '"' in str(processed_rows[i].get(norm_key('Título'), '')), processed_rows[i].get(norm_key('Hora')) or ''), reverse=True)
            for pos, idx in enumerate(indices):
                processed_rows[idx]['Duplicada'] = "Sí"
                if pos > 0: mark_as_duplicate_to_delete(processed_rows[idx])

    # --- FASE 2: POSIBLES DUPLICADOS POR TÍTULO ---
    titulo_groups = defaultdict(list)
    for idx, row in enumerate(processed_rows):
        if row['Duplicada'] == "FALSE":
            titulo_norm = normalize_title(row.get(norm_key('Título')))
            if titulo_norm: titulo_groups[titulo_norm].append(idx)
    
    for indices in titulo_groups.values():
        if len(indices) > 1:
            for i in range(len(indices)):
                for j in range(i+1, len(indices)):
                    idx1, idx2 = indices[i], indices[j]
                    row1, row2 = processed_rows[idx1], processed_rows[idx2]
                    
                    if (format_date_str(parse_date(row1.get(norm_key('Fecha')))) == format_date_str(parse_date(row2.get(norm_key('Fecha')))) and
                        norm_key(row1.get(norm_key('Menciones - Empresa'))) == norm_key(row2.get(norm_key('Menciones - Empresa')))):
                        
                        if es_radio_o_tv(row1) and es_radio_o_tv(row2):
                            if (norm_key(row1.get(norm_key('Medio'))) == norm_key(row2.get(norm_key('Medio'))) and
                                str(row1.get(norm_key('Hora'))) == str(row2.get(norm_key('Hora')))):
                                row1['Posible Duplicada'] = "Sí"; row2['Posible Duplicada'] = "Sí"
                                mark_as_duplicate_to_delete(row2)
                        else:
                            row1['Posible Duplicada'] = "Sí"; row2['Posible Duplicada'] = "Sí"
                            mark_as_duplicate_to_delete(row2)

    # --- FASE 3: DUPLICADOS POR SIMILITUD DE TÍTULO ---
    for idx, row in enumerate(processed_rows):
        if row['Duplicada'] == "FALSE" and row['Posible Duplicada'] == "FALSE":
            titulo_actual = normalize_title(row.get(norm_key('Título')))
            if not titulo_actual: continue
            
            for idx2, row2 in enumerate(processed_rows):
                if (idx2 != idx and row2['Duplicada'] == "FALSE" and row2['Posible Duplicada'] == "FALSE" and
                    format_date_str(parse_date(row.get(norm_key('Fecha')))) == format_date_str(parse_date(row2.get(norm_key('Fecha')))) and
                    norm_key(row.get(norm_key('Menciones - Empresa'))) == norm_key(row2.get(norm_key('Menciones - Empresa')))):
                    
                    titulo2 = normalize_title(row2.get(norm_key('Título')))
                    if titulo2 and SequenceMatcher(None, titulo_actual, titulo2).ratio() > 0.8:
                        row['Posible Duplicada'] = "Sí"; row2['Posible Duplicada'] = "Sí"
                        mark_as_duplicate_to_delete(row2)

    # --- GENERACIÓN DEL REPORTE FINAL ---
    final_order = ["ID Noticia", "Fecha", "Hora", "Medio", "Tipo de Medio", "Sección - Programa", "Región","Título", "Autor - Conductor", "Nro. Pagina", "Dimensión", "Duración - Nro. Caracteres", "CPE", "Tier", "Audiencia", "Tono", "Tema", "Temas Generales - Tema", "Resumen - Aclaracion", "Link Nota", "Link (Streaming - Imagen)", "Menciones - Empresa", "Duplicada", "Posible Duplicada", "Mantener"]
    new_wb = openpyxl.Workbook()
    new_sheet = new_wb.active
    new_sheet.title = "Hoja1"
    new_sheet.append(final_order)
    
    if "CustomLink" not in new_wb.named_styles: new_wb.add_named_style(custom_link_style)

    for row_data in processed_rows:
        if row_data['Mantener'] == 'Conservar':
            titulo_key = norm_key('Título')
            row_data[titulo_key] = re.sub(r'\s*\|\s*[\w\s]+$', '', str(row_data.get(titulo_key, ''))).strip()
        new_row_to_append = [row_data.get(norm_key(header))['value'] if isinstance(row_data.get(norm_key(header)), dict) and 'value' in row_data.get(norm_key(header)) else row_data.get(norm_key(header)) for header in final_order]
        new_sheet.append(new_row_to_append)
    
    link_nota_idx = final_order.index("Link Nota")
    link_streaming_idx = final_order.index("Link (Streaming - Imagen)")
    for i, row_cells in enumerate(new_sheet.iter_rows(min_row=2)):
        if i < len(processed_rows):
            processed = processed_rows[i]
            if link_nota := processed.get(norm_key("Link Nota")):
                if isinstance(link_nota, dict) and link_nota.get("url"):
                    cell = row_cells[link_nota_idx]; cell.hyperlink = link_nota["url"]; cell.value = "Link"; cell.style = "CustomLink"
            if link_stream := processed.get(norm_key("Link (Streaming - Imagen)")):
                 if isinstance(link_stream, dict) and link_stream.get("url"):
                    cell = row_cells[link_streaming_idx]; cell.hyperlink = link_stream["url"]; cell.value = "Link"; cell.style = "CustomLink"
    
    wb.remove(wb.active)
    
    summary = {"total_rows": len(processed_rows), "to_eliminate": sum(1 for r in processed_rows if r['Mantener'] == 'Eliminar'), "to_conserve": len(processed_rows) - sum(1 for r in processed_rows if r['Mantener'] == 'Eliminar'), "exact_duplicates": sum(1 for r in processed_rows if r['Duplicada'] == 'Sí'), "possible_duplicates": sum(1 for r in processed_rows if r['Posible Duplicada'] == 'Sí')}
    return new_wb, summary
