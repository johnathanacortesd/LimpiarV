# deduplicator.py

import openpyxl
from openpyxl.styles import Font, Alignment, NamedStyle
from difflib import SequenceMatcher
from collections import defaultdict
import re
import datetime
from copy import deepcopy

# --- Funciones Auxiliares ---
def norm_key(text): return re.sub(r'\W+', '', str(text).lower().strip()) if text else ""

def normalize_title(title):
    if not isinstance(title, str): return ""
    title = re.sub(r'\s*\|\s*[\w\s]+$', '', title)
    return re.sub(r'\W+', ' ', title).lower().strip()

def extract_root_domain(url):
    if not url: return None
    try:
        cleaned_url = re.sub(r'^https?://', '', url).lower().replace('www.', '')
        domain = cleaned_url.split('/')[0]
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
    
    # PASO 1: LEER DATOS Y EXPANDIR FILAS
    headers = [cell.value for cell in sheet[1]]
    headers_norm = [norm_key(h) for h in headers]
    
    initial_rows = []
    for row_cells in sheet.iter_rows(min_row=2):
        if all(c.value is None for c in row_cells): continue
        row_data = {
            headers_norm[i]: {"value": cell.value, "url": cell.hyperlink.target if cell.hyperlink else None} 
            if headers_norm[i] in [norm_key('Link Nota'), norm_key('Link (Streaming - Imagen)')] 
            else cell.value 
            for i, cell in enumerate(row_cells)
        }
        initial_rows.append(row_data)

    expanded_rows = []
    menciones_key_norm = norm_key('Menciones - Empresa')
    for base_row in initial_rows:
        menciones_str = str(base_row.get(menciones_key_norm) or '')
        menciones = [m.strip() for m in menciones_str.split(';') if m.strip()]
        if not menciones:
            expanded_rows.append(base_row)
        else:
            for mencion in menciones:
                new_row = deepcopy(base_row)
                new_row[menciones_key_norm] = mencion
                expanded_rows.append(new_row)

    # PASO 2: APLICAR TODOS LOS MAPEOS A LAS FILAS YA EXPANDIDAS
    processed_rows = []
    for row in expanded_rows:
        # 1. Mapeo de Empresas
        mencion = str(row.get(menciones_key_norm, '')).lower().strip()
        if mencion in empresa_dict:
            row[menciones_key_norm] = empresa_dict[mencion]

        # 2. Mapeo de Internet
        if str(row.get(norm_key('Tipo de Medio'))).lower().strip() == 'internet':
            medio_val = str(row.get(norm_key('Medio'))).lower().strip()
            if medio_val in internet_dict:
                row[norm_key('Medio')] = internet_dict[medio_val]
            else:
                link_data = row.get(norm_key('Link Nota'), {})
                url = link_data.get('url')
                if root_domain := extract_root_domain(url):
                    row[norm_key('Medio')] = root_domain

        # 3. Mapeo de Región
        medio_actual = str(row.get(norm_key('Medio'))).lower().strip()
        row[norm_key('Región')] = region_dict.get(medio_actual, "Online")
        
        # 4. Normalización final de Título
        row[norm_key('Título')] = normalize_title(row.get(norm_key('Título')))
        
        processed_rows.append(row)

    # PASO 3: LÓGICA DE DUPLICACIÓN
    for row in processed_rows:
        row.update({'Duplicada': "FALSE", 'Posible Duplicada': "FALSE", 'Mantener': "Conservar"})

    # FASE 1: Duplicados Exactos
    grupos_exactos = defaultdict(list)
    for idx, row in enumerate(processed_rows):
        key_tuple = (row.get(norm_key('Título')), norm_key(row.get(norm_key('Medio'))), format_date_str(parse_date(row.get(norm_key('Fecha')))), norm_key(row.get(norm_key('Menciones - Empresa'))))
        if es_radio_o_tv(row): key_tuple += (str(row.get(norm_key('Hora'))),)
        grupos_exactos[key_tuple].append(idx)
    for indices in grupos_exactos.values():
        if len(indices) > 1:
            indices.sort(key=lambda i: (not is_title_problematic(processed_rows[i].get(norm_key('Título'))), '"' in str(processed_rows[i].get(norm_key('Título'), '')), processed_rows[i].get(norm_key('Hora')) or ''), reverse=True)
            for pos, idx in enumerate(indices):
                processed_rows[idx]['Duplicada'] = "Sí"
                if pos > 0: mark_as_duplicate_to_delete(processed_rows[idx])
    
    # (Las Fases 2 y 3 de deduplicación irían aquí)
    # ...

    # PASO 4: GENERACIÓN DEL REPORTE FINAL
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
