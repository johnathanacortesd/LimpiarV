# deduplicator.py (versión mejorada con mapeo de menciones)

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, NamedStyle
from difflib import SequenceMatcher
from collections import defaultdict
import re
import datetime
from copy import deepcopy

# --- Funciones Auxiliares ---
def norm_key(text):
    return re.sub(r'\W+', '', str(text).lower().strip()) if text else ""

def convert_html_entities(text):
    if not isinstance(text, str): return text
    html_entities = {
        'á': 'á', 'é': 'é', 'í': 'í', 'ó': 'ó', 'ú': 'ú', 'ñ': 'ñ',
        'Á': 'Á', 'É': 'É', 'Í': 'Í', 'Ó': 'Ó', 'Ú': 'Ú', 'Ñ': 'Ñ',
        '"': '"', '"': '"', '"': '"', ''': "'", ''': "'",
        'Â': '', 'â': '', '€': '', '™': ''
    }
    for entity, char in html_entities.items():
        text = text.replace(entity, char)
    return text

def normalize_title(title):
    if not isinstance(title, str): return ""
    title = convert_html_entities(title)
    title = re.sub(r'\s*\|\s*[\w\s]+$', '', title)
    return re.sub(r'\W+', ' ', title).lower().strip()

def corregir_texto(text):
    if not isinstance(text, str): return text
    text = convert_html_entities(text)
    text = re.sub(r'(<br>|\[\.\.\.\]|\s+)', ' ', text).strip()
    match = re.search(r'[A-Z]', text)
    if match: text = text[match.start():]
    if text and not text.endswith('...'): text = text.rstrip('.') + '...'
    return text

def extract_link(cell):
    if cell.hyperlink: return {"value": "Link", "url": cell.hyperlink.target}
    if cell.value and isinstance(cell.value, str):
        match = re.search(r'=HYPERLINK\("([^"]+)"', cell.value)
        if match: return {"value": "Link", "url": match.group(1)}
    return {"value": cell.value, "url": None}

def parse_date(fecha):
    if isinstance(fecha, datetime.datetime): return fecha.date()
    try: return datetime.datetime.strptime(str(fecha).split(" ")[0], "%Y-%m-%d").date()
    except (ValueError, TypeError): return None

def format_date_str(fecha_obj):
    if isinstance(fecha_obj, datetime.date): return fecha_obj.isoformat()
    return str(fecha_obj)[:10]

def es_internet(row):
    return norm_key(row.get(norm_key('Tipo de Medio'))) == 'internet'

def es_radio_o_tv(row):
    tm = norm_key(row.get(norm_key('Tipo de Medio')))
    return tm in {'radio', 'televisión'}

def mark_as_duplicate_to_delete(row):
    row['Mantener'] = "Eliminar"
    row[norm_key('Tono')] = "Duplicada"
    row[norm_key('Tema')] = "-"
    row[norm_key('Temas Generales - Tema')] = "-"

def is_title_problematic(title):
    if not isinstance(title, str): return False
    if re.search(r'\s*\|\s*[\w\s]+$', title): return True
    if re.search(r'[Ââ€™"""'']', title): return True
    return False

def apply_mentions_mapping(menciones_str, mentions_dict):
    """
    Aplica el mapeo de menciones a una cadena de menciones separadas por punto y coma.
    
    Args:
        menciones_str (str): Cadena con menciones separadas por ';'
        mentions_dict (dict): Diccionario de mapeo {original: normalizado}
    
    Returns:
        str: Cadena con menciones mapeadas
    """
    if not menciones_str or not mentions_dict:
        return menciones_str
    
    menciones_list = [m.strip() for m in str(menciones_str).split(';') if m.strip()]
    menciones_mapped = []
    
    for mencion in menciones_list:
        mencion_lower = mencion.lower().strip()
        mapped = False
        
        # Buscar coincidencia exacta primero
        if mencion_lower in mentions_dict:
            menciones_mapped.append(mentions_dict[mencion_lower])
            mapped = True
        else:
            # Buscar coincidencia parcial
            for original_key, mapped_value in mentions_dict.items():
                if original_key in mencion_lower or mencion_lower in original_key:
                    menciones_mapped.append(mapped_value)
                    mapped = True
                    break
        
        if not mapped:
            menciones_mapped.append(mencion)  # Mantener original si no hay mapeo
    
    return '; '.join(menciones_mapped)

# --- Función Principal de Deduplicación ---
def run_deduplication_process(wb, mentions_dict=None):
    """
    Ejecuta el proceso completo de deduplicación.
    
    Args:
        wb: Workbook de openpyxl
        mentions_dict (dict, optional): Diccionario para mapear menciones
    
    Returns:
        tuple: (workbook final, resumen estadístico)
    """
    sheet = wb.active
    
    # Configurar estilo para enlaces
    custom_link_style = NamedStyle(name="CustomLink")
    custom_link_style.font = Font(color="000000", underline="none")
    custom_link_style.alignment = Alignment(horizontal="left")
    custom_link_style.number_format = '@'
    if "CustomLink" not in wb.named_styles:
        wb.add_named_style(custom_link_style)
        
    # --- PASO 1: PROCESAMIENTO Y NORMALIZACIÓN ---
    headers = [cell.value for cell in sheet[1]]
    headers_norm = [norm_key(h) for h in headers]
    processed_rows = []
    mentions_mapping_stats = {'total_processed': 0, 'mapped_count': 0}

    for row_idx, row_cells in enumerate(sheet.iter_rows(min_row=2)):
        if all(c.value is None for c in row_cells): continue
        
        base_data = {}
        for i, cell in enumerate(row_cells):
            col_name = headers_norm[i]
            if col_name in [norm_key('Link Nota'), norm_key('Link (Streaming - Imagen)')]:
                base_data[col_name] = extract_link(cell)
            else:
                base_data[col_name] = cell.value

        # Normalizar título
        base_data[norm_key('Título')] = convert_html_entities(str(base_data.get(norm_key('Título'), '')))
        
        # Corregir texto de resumen
        base_data[norm_key('Resumen - Aclaracion')] = corregir_texto(base_data.get(norm_key('Resumen - Aclaracion')))
        
        # Normalizar tipo de medio
        tipo_medio_key = norm_key('Tipo de Medio')
        tm_norm = norm_key(base_data.get(tipo_medio_key))
        if tm_norm in {'aire', 'cable'}: base_data[tipo_medio_key] = 'Televisión'
        elif tm_norm in {'am', 'fm'}: base_data[tipo_medio_key] = 'Radio'
        elif tm_norm == 'diario': base_data[tipo_medio_key] = 'Prensa'
        elif tm_norm == 'online': base_data[tipo_medio_key] = 'Internet'
        elif tm_norm == 'revista': base_data[tipo_medio_key] = 'Revista'
        
        # Ajustar enlaces según tipo de medio
        link_nota_key = norm_key("Link Nota")
        link_streaming_key = norm_key("Link (Streaming - Imagen)")
        tipo_medio_val = base_data.get(tipo_medio_key)
        
        if tipo_medio_val == "Internet":
            base_data[link_nota_key], base_data[link_streaming_key] = base_data.get(link_streaming_key), base_data.get(link_nota_key)
        elif tipo_medio_val in {"Prensa", "Revista"}:
            if (not base_data.get(link_nota_key) or not base_data.get(link_nota_key, {}).get('url')) and base_data.get(link_streaming_key, {}).get('url'):
                base_data[link_nota_key] = base_data.get(link_streaming_key)
            base_data[link_streaming_key] = None
        elif tipo_medio_val in {"Radio", "Televisión"}:
            base_data[link_streaming_key] = None
        
        # Procesar menciones con mapeo
        menciones_key = norm_key('Menciones - Empresa')
        menciones_value = base_data.get(menciones_key)
        
        if menciones_value:
            menciones_str = str(menciones_value)
            mentions_mapping_stats['total_processed'] += 1
            
            # Aplicar mapeo de menciones si está disponible
            if mentions_dict:
                original_menciones = menciones_str
                mapped_menciones = apply_mentions_mapping(menciones_str, mentions_dict)
                if original_menciones != mapped_menciones:
                    mentions_mapping_stats['mapped_count'] += 1
                menciones_str = mapped_menciones
            
            # Dividir menciones y crear filas separadas
            menciones = [m.strip() for m in menciones_str.split(';') if m.strip()]
            if not menciones:
                processed_rows.append(base_data)
            else:
                for mencion in menciones:
                    new_row = deepcopy(base_data)
                    new_row[menciones_key] = mencion
                    processed_rows.append(new_row)
        else:
            processed_rows.append(base_data)
    
    # Inicializar campos de control
    for row in processed_rows:
        row.update({
            'Duplicada': "FALSE",
            'Posible Duplicada': "FALSE",
            'Mantener': "Conservar"
        })

    # --- PASO 2: DETECCIÓN DE DUPLICADOS ---
    
    # FASE 1: Duplicados Exactos
    grupos_exactos = defaultdict(list)
    for idx, row in enumerate(processed_rows):
        key_tuple = (
            normalize_title(row.get(norm_key('Título'))),
            norm_key(row.get(norm_key('Medio'))),
            format_date_str(parse_date(row.get(norm_key('Fecha')))),
            norm_key(row.get(norm_key('Menciones - Empresa')))
        )
        if es_radio_o_tv(row):
            key_tuple += (str(row.get(norm_key('Hora'))),)
        grupos_exactos[key_tuple].append(idx)
    
    exact_duplicates_count = 0
    for indices in grupos_exactos.values():
        if len(indices) > 1:
            exact_duplicates_count += len(indices)
            indices.sort(key=lambda i: (
                not is_title_problematic(processed_rows[i].get(norm_key('Título'))),
                '"' in str(processed_rows[i].get(norm_key('Título'), '')),
                processed_rows[i].get(norm_key('Hora')) or ''
            ), reverse=True)
            
            for pos, idx in enumerate(indices):
                processed_rows[idx]['Duplicada'] = "Sí"
                if pos > 0:
                    mark_as_duplicate_to_delete(processed_rows[idx])

    # FASE 2: Posibles Duplicados por Similitud (mismo día)
    SIMILARIDAD_MINIMA = 0.8
    grupos_posibles = defaultdict(list)
    for idx, row in enumerate(processed_rows):
        if row['Duplicada'] == "FALSE":
            key_tuple = (
                norm_key(row.get(norm_key('Menciones - Empresa'))),
                norm_key(row.get(norm_key('Medio'))),
                format_date_str(parse_date(row.get(norm_key('Fecha'))))
            )
            if es_radio_o_tv(row):
                key_tuple += (str(row.get(norm_key('Hora'))),)
            grupos_posibles[key_tuple].append(idx)
    
    possible_duplicates_count = 0
    for group in grupos_posibles.values():
        if len(group) < 2: continue
        
        # Algoritmo Union-Find para agrupar duplicados
        parent = {i: i for i in group}
        
        def find(x):
            while parent[x] != x:
                parent[x] = parent[parent[x]]
                x = parent[x]
            return x
        
        def union(x, y):
            rx, ry = find(x), find(y)
            parent[ry] = rx
        
        for i in range(len(group)):
            for j in range(i + 1, len(group)):
                idx_i, idx_j = group[i], group[j]
                if (processed_rows[idx_i]['Mantener'] == 'Eliminar' or 
                    processed_rows[idx_j]['Mantener'] == 'Eliminar'): 
                    continue
                
                title_i = normalize_title(processed_rows[idx_i].get(norm_key('Título')))
                title_j = normalize_title(processed_rows[idx_j].get(norm_key('Título')))
                
                if title_i and title_j and SequenceMatcher(None, title_i, title_j).ratio() >= SIMILARIDAD_MINIMA:
                    union(idx_i, idx_j)
        
        # Agrupar por clusters
        clusters = defaultdict(list)
        for i in group:
            clusters[find(i)].append(i)
        
        for cluster in clusters.values():
            if len(cluster) > 1:
                possible_duplicates_count += len(cluster)
                cluster.sort(key=lambda i: (
                    not is_title_problematic(processed_rows[i].get(norm_key('Título'))),
                    '"' in str(processed_rows[i].get(norm_key('Título'), '')),
                    processed_rows[i].get(norm_key('Hora')) or ''
                ), reverse=True)
                
                for pos, idx in enumerate(cluster):
                    processed_rows[idx]['Posible Duplicada'] = "Sí"
                    if pos > 0:
                        mark_as_duplicate_to_delete(processed_rows[idx])

    # FASE 3: Duplicados en Fechas Diferentes
    grupos_fecha_diff = defaultdict(list)
    for idx, row in enumerate(processed_rows):
        if row['Mantener'] == 'Conservar' and not es_radio_o_tv(row):
            titulo_norm = normalize_title(row.get(norm_key('Título')))
            if titulo_norm:
                key_tuple = (
                    titulo_norm,
                    norm_key(row.get(norm_key('Menciones - Empresa'))),
                    norm_key(row.get(norm_key('Medio')))
                )
                grupos_fecha_diff[key_tuple].append(idx)
    
    for key, group in grupos_fecha_diff.items():
        if len(group) < 2: continue
        
        is_internet_group = es_internet(processed_rows[group[0]])
        
        # Union-Find para fechas diferentes
        parent = {i: i for i in group}
        
        def find(x):
            while parent[x] != x:
                parent[x] = parent[parent[x]]
                x = parent[x]
            return x
        
        def union(x, y):
            rx, ry = find(x), find(y)
            parent[ry] = rx
        
        for i in range(len(group)):
            for j in range(i + 1, len(group)):
                idx_i, idx_j = group[i], group[j]
                date_i = parse_date(processed_rows[idx_i].get(norm_key('Fecha')))
                date_j = parse_date(processed_rows[idx_j].get(norm_key('Fecha')))
                
                if not date_i or not date_j: continue
                
                should_union = False
                if is_internet_group:
                    if abs((date_i - date_j).days) == 1:
                        should_union = True
                else:
                    if date_i != date_j:
                        should_union = True
                
                if should_union:
                    union(idx_i, idx_j)
        
        # Agrupar por clusters
        clusters = defaultdict(list)
        for i in group:
            clusters[find(i)].append(i)
        
        for cluster in clusters.values():
            if len(cluster) > 1:
                cluster.sort(key=lambda i: (
                    not is_title_problematic(processed_rows[i].get(norm_key('Título'))),
                    '"' in str(processed_rows[i].get(norm_key('Título'), '')),
                    format_date_str(parse_date(processed_rows[i].get(norm_key('Fecha'))))
                ), reverse=True)
                
                for pos, idx in enumerate(cluster):
                    processed_rows[idx]['Posible Duplicada'] = "Sí"
                    if pos > 0:
                        mark_as_duplicate_to_delete(processed_rows[idx])
    
    # --- PASO 3: GENERACIÓN DEL REPORTE FINAL ---
    final_order = [
        "ID Noticia", "Fecha", "Hora", "Medio", "Tipo de Medio", 
        "Sección - Programa", "Región", "Título", "Autor - Conductor", 
        "Nro. Pagina", "Dimensión", "Duración - Nro. Caracteres", "CPE", 
        "Tier", "Audiencia", "Tono", "Tema", "Temas Generales - Tema", 
        "Resumen - Aclaracion", "Link Nota", "Link (Streaming - Imagen)", 
        "Menciones - Empresa", "Duplicada", "Posible Duplicada", "Mantener"
    ]
    
    # Crear nueva hoja
    new_sheet = wb.create_sheet("Resultado")
    new_sheet.append(final_order)
    
    # Procesar filas para el resultado final
    for row_data in processed_rows:
        # Limpiar títulos de filas que se van a conservar
        if row_data['Mantener'] == 'Conservar':
            titulo_key = norm_key('Título')
            titulo_actual = str(row_data.get(titulo_key, ''))
            titulo_limpio = re.sub(r'\s*\|\s*[\w\s]+, '', titulo_actual).strip()
            row_data[titulo_key] = titulo_limpio
        
        # Preparar fila para insertar
        new_row_to_append = []
        for header in final_order:
            header_norm = norm_key(header)
            cell_value = row_data.get(header_norm)
            
            # Manejar enlaces
            if isinstance(cell_value, dict) and 'value' in cell_value:
                new_row_to_append.append(cell_value['value'])
            else:
                new_row_to_append.append(cell_value)
        
        new_sheet.append(new_row_to_append)
    
    # Aplicar hipervínculos
    link_nota_idx = final_order.index("Link Nota")
    link_streaming_idx = final_order.index("Link (Streaming - Imagen)")
    
    for i, row_cells in enumerate(new_sheet.iter_rows(min_row=2)):
        if i < len(processed_rows):
            processed = processed_rows[i]
            
            # Link Nota
            link_nota = processed.get(norm_key("Link Nota"))
            if link_nota and isinstance(link_nota, dict) and link_nota.get("url"):
                cell = row_cells[link_nota_idx]
                cell.hyperlink = link_nota["url"]
                cell.value = "Link"
                cell.style = "CustomLink"
            
            # Link Streaming
            link_stream = processed.get(norm_key("Link (Streaming - Imagen)"))
            if link_stream and isinstance(link_stream, dict) and link_stream.get("url"):
                cell = row_cells[link_streaming_idx]
                cell.hyperlink = link_stream["url"]
                cell.value = "Link"
                cell.style = "CustomLink"
    
    # Remover hoja original y renombrar
    wb.remove(sheet)
    new_sheet.title = "Hoja1"
    
    # --- PREPARAR RESUMEN ESTADÍSTICO ---
    total_rows = len(processed_rows)
    to_eliminate = sum(1 for row in processed_rows if row['Mantener'] == 'Eliminar')
    to_conserve = total_rows - to_eliminate
    
    summary = {
        "total_rows": total_rows,
        "to_eliminate": to_eliminate,
        "to_conserve": to_conserve,
        "exact_duplicates": exact_duplicates_count,
        "possible_duplicates": possible_duplicates_count,
        "mentions_mapping_stats": mentions_mapping_stats
    }
    
    return wb, summary
