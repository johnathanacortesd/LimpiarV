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
