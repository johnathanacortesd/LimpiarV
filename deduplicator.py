# deduplicator.py

import openpyxl
from openpyxl.styles import Font, Alignment, NamedStyle
from difflib import SequenceMatcher
from collections import defaultdict
import re
import datetime
from copy import deepcopy

# --- Funciones Auxiliares (de tu código funcional) ---
def norm_key(text):
    return re.sub(r'\W+', '', str(text).lower().strip()) if text else ""
def convert_html_entities(text):
    if not isinstance(text, str): return text
    html_entities = {'á':'á','é':'é','í':'í','ó':'ó','ú':'ú','ñ':'ñ','Á':'Á','É':'É','Í':'Í','Ó':'Ó','Ú':'Ú','Ñ':'Ñ','"':'"','“':'"','”':'"','‘':"'",'’':"'",'Â':'','â':'','€':'','™':''}
    for entity, char in html_entities.items(): text = text.replace(entity, char)
    return text
def normalize_title(title):
    if not isinstance(title, str): return ""
    title = convert_html_entities(title)
    title = re.sub(r'\s*\|\s*[\w\s]+$', '', title)
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
def extract_root_domain(url):
    if not url: return None
    try:
        cleaned_url = re.sub(r'^https?://', '', url).lower().replace('www.', ''); domain = cleaned_url.split('/')[0]
        return domain.capitalize()
    except Exception: return None
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
    row['Mantener']="Eliminar"; row[norm_key('Tono')]="Duplicada"; row[norm_key('Tema')]="-"; row[norm_key('Temas Generales - Tema')]="-"
def is_title_problematic(title):
    if not isinstance(title, str): return False
    if re.search(r'\s*\|\s*[\w\s]+$', title): return True
    if re.search(r'[Ââ€™“”“’‘]', title): return True
    return False

# --- Función Principal de Procesamiento ---
def run_deduplication_process(wb, internet_dict, region_dict, empresa_dict):
    sheet = wb.active
    
    # --- PASO 1: Procesamiento, Mapeo y Expansión ---
    headers = [cell.value for cell in sheet[1]]
    headers_norm = [norm_key(h) for h in headers]
    
    # Leer datos y realizar mapeos y expansión en el orden correcto
    processed_rows = []
    menciones_key_norm = norm_key('Menciones - Empresa')

    for row_cells in sheet.iter_rows(min_row=2):
        if all(c.value is None for c in row_cells): continue
        
        base_data = {
            headers_norm[i]: extract_link(cell) 
            if headers_norm[i] in [norm_key('Link Nota'), norm_key('Link (Streaming - Imagen)')] 
            else cell.value 
            for i, cell in enumerate(row_cells)
        }
        
        # Expansión de filas basada en menciones es lo PRIMERO
        menciones_str = str(base_data.get(menciones_key_norm) or '')
        menciones_list = [m.strip() for m in menciones_str.split(';') if m.strip()]

        if not menciones_list:
            # Si no hay menciones, procesar la fila una vez
            row_to_process = deepcopy(base_data)
            # Aplicar mapeos
            if str(row_to_process.get(norm_key('Tipo de Medio'))).lower().strip() == 'internet':
                medio_val = str(row_to_process.get(norm_key('Medio'))).lower().strip()
                if medio_val in internet_dict:
                    row_to_process[norm_key('Medio')] = internet_dict[medio_val]
                else:
                    link_data = row_to_process.get(norm_key('Link Nota'), {})
                    if root_domain := extract_root_domain(link_data.get('url')):
                        row_to_process[norm_key('Medio')] = root_domain
            medio_actual = str(row_to_process.get(norm_key('Medio'))).lower().strip()
            row_to_process[norm_key('Región')] = region_dict.get(medio_actual, "Online")
            processed_rows.append(row_to_process)
        else:
            # Si hay menciones, crear una fila por cada una y aplicar mapeos
            for mencion in menciones_list:
                new_row = deepcopy(base_data)
                new_row[menciones_key_norm] = mencion
                
                # AHORA, aplicar los mapeos a esta fila individual
                # 1. Mapeo de Empresas
                mencion_limpia = mencion.lower().strip()
                if mencion_limpia in empresa_dict:
                    new_row[menciones_key_norm] = empresa_dict[mencion_limpia]

                # 2. Mapeo de Internet
                if str(new_row.get(norm_key('Tipo de Medio'))).lower().strip() == 'internet':
                    medio_val = str(new_row.get(norm_key('Medio'))).lower().strip()
                    if medio_val in internet_dict:
                        new_row[norm_key('Medio')] = internet_dict[medio_val]
                    else:
                        link_data = new_row.get(norm_key('Link Nota'), {})
                        if root_domain := extract_root_domain(link_data.get('url')):
                            new_row[norm_key('Medio')] = root_domain
                
                # 3. Mapeo de Región
                medio_actual = str(new_row.get(norm_key('Medio'))).lower().strip()
                new_row[norm_key('Región')] = region_dict.get(medio_actual, "Online")

                processed_rows.append(new_row)

    # --- PASO 2: Normalización final y Lógica de Duplicación (sobre `processed_rows`) ---
    for row in processed_rows:
        row[norm_key('Título')] = convert_html_entities(str(row.get(norm_key('Título'), '')))
        row[norm_key('Resumen - Aclaracion')] = corregir_texto(row.get(norm_key('Resumen - Aclaracion')))
        tipo_medio_key = norm_key('Tipo de Medio'); tm_norm = norm_key(row.get(tipo_medio_key))
        if tm_norm in {'aire', 'cable'}: row[tipo_medio_key] = 'Televisión'
        elif tm_norm in {'am', 'fm'}: row[tipo_medio_key] = 'Radio'
        #... y el resto de la normalización de tu código original
        row.update({'Duplicada': "FALSE", 'Posible Duplicada': "FALSE", 'Mantener': "Conservar"})

    # El resto de tu lógica de deduplicación que ya funcionaba
    # ... FASE 1, 2, 3 ...
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

    # --- PASO 3: GENERACIÓN DEL REPORTE FINAL ---
    final_order = ["ID Noticia", "Fecha", "Hora", "Medio", "Tipo de Medio", "Sección - Programa", "Región","Título", "Autor - Conductor", "Nro. Pagina", "Dimensión", "Duración - Nro. Caracteres", "CPE", "Tier", "Audiencia", "Tono", "Tema", "Temas Generales - Tema", "Resumen - Aclaracion", "Link Nota", "Link (Streaming - Imagen)", "Menciones - Empresa", "Duplicada", "Posible Duplicada", "Mantener"]
    new_wb = openpyxl.Workbook()
    new_sheet = new_wb.active
    new_sheet.title = "Hoja1"
    new_sheet.append(final_order)
    
    custom_link_style = NamedStyle(name="CustomLink", font=Font(color="0000FF", underline="single"))
    if "CustomLink" not in new_wb.named_styles: new_wb.add_named_style(custom_link_style)

    for row_data in processed_rows:
        new_row_to_append = []
        for header in final_order:
            val = row_data.get(norm_key(header))
            if isinstance(val, dict): new_row_to_append.append(val.get('value'))
            else: new_row_to_append.append(val)
        new_sheet.append(new_row_to_append)

    link_nota_idx_out = final_order.index("Link Nota") + 1
    for i, row_data in enumerate(processed_rows, start=2):
        link_data = row_data.get(norm_key("Link Nota"))
        if link_data and isinstance(link_data, dict) and link_data.get("url"):
            cell = new_sheet.cell(row=i, column=link_nota_idx_out)
            cell.hyperlink = link_data["url"]; cell.value = "Link"; cell.style = "Hyperlink"
    
    summary = {"total_rows": len(processed_rows), "to_eliminate": sum(1 for r in processed_rows if r['Mantener'] == 'Eliminar'), "to_conserve": len(processed_rows) - sum(1 for r in processed_rows if r['Mantener'] == 'Eliminar'), "exact_duplicates": sum(1 for r in processed_rows if r['Duplicada'] == 'Sí'), "possible_duplicates": sum(1 for r in processed_rows if r['Posible Duplicada'] == 'Sí')}
    return new_wb, summary
