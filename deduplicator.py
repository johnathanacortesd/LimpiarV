# deduplicator.py

import openpyxl
from openpyxl.styles import Font, Alignment, NamedStyle
from difflib import SequenceMatcher
from collections import defaultdict
import re
import datetime
from copy import deepcopy
import html

# --- CONSTANTES ---
# Usar constantes mejora la legibilidad y previene errores de tipeo.
CONSERVAR = "Conservar"
ELIMINAR = "Eliminar"
SI = "Sí"
NO = "No"
DUPLICADA = "Duplicada"
TONO_DUPLICADA = "Duplicada"
TEMA_VACIO = "-"

# --- FUNCIONES AUXILIARES ---
def norm_key(text): 
    """Normaliza un texto para usarlo como clave: minúsculas, sin espacios ni caracteres especiales."""
    return re.sub(r'\W+', '', str(text).lower().strip()) if text else ""

def convert_html_entities(text):
    """Convierte todas las entidades HTML a sus caracteres correspondientes."""
    if not isinstance(text, str):
        return text
    return html.unescape(text)

def normalize_title(title):
    """Normaliza un título para la comparación, limpiándolo completamente."""
    if not isinstance(title, str):
        return ""
    title = convert_html_entities(title)
    # Remover sufijos como "| Medio Name" que a veces quedan
    title = re.sub(r'\s*\|\s*[\w\s]+$', '', title)
    # Remover todos los caracteres no alfanuméricos y convertir a minúsculas
    title = re.sub(r'\W+', '', title.lower().strip())
    return title

def corregir_texto(text):
    """Limpia y corrige el texto del resumen."""
    if not isinstance(text, str): 
        return text
    text = convert_html_entities(text)
    text = text.replace('<br>', ' ').replace('[...]', ' ')
    text = re.sub(r'\s+', ' ', text).strip()
    
    # Buscar la primera letra mayúscula para empezar desde ahí
    if match := re.search(r'[A-Z]', text): 
        text = text[match.start():]
    
    # Asegurar que termine con puntos suspensivos si no es un final de oración
    if text and not text.endswith('...'):
        text = re.sub(r'[\.,;:]$', '', text.strip()).strip()
        text += '...'
    return text

def extract_link(cell):
    """Extrae un hipervínculo y su texto de una celda de Excel."""
    if cell.hyperlink and cell.hyperlink.target:
        return {"value": cell.value or "Link", "url": cell.hyperlink.target}
    if cell.value and isinstance(cell.value, str):
        if match := re.search(r'=HYPERLINK\("([^"]+)"', cell.value): 
            return {"value": "Link", "url": match.group(1)}
    return {"value": cell.value, "url": None}

def format_date(date_val):
    """Formatea una fecha a un string 'YYYY-MM-DD'."""
    if isinstance(date_val, datetime.datetime):
        return date_val.strftime('%Y-%m-%d')
    if isinstance(date_val, str):
        try:
            return datetime.datetime.strptime(date_val.split(' ')[0], '%Y-%m-%d').strftime('%Y-%m-%d')
        except ValueError:
            return str(date_val)
    return str(date_val)

def parse_date_obj(date_val):
    """Convierte un valor de fecha a un objeto datetime.date para comparaciones."""
    if isinstance(date_val, datetime.datetime):
        return date_val.date()
    if isinstance(date_val, datetime.date):
        return date_val
    if isinstance(date_val, str):
        try:
            return datetime.datetime.strptime(date_val.split(' ')[0], '%Y-%m-%d').date()
        except (ValueError, AttributeError):
            return datetime.date.min
    return datetime.date.min

def es_internet(row): 
    """Verifica si el 'Tipo de Medio' es 'Internet'."""
    return norm_key(row.get(norm_key('Tipo de Medio'))) == 'internet'

def is_title_problematic(title):
    """Detecta si un título es problemático (genérico, mala codificación)."""
    if not isinstance(title, str) or not title.strip():
        return True
    # Títulos genéricos
    if re.search(r'^(sin t[íi]tulo|untitled|no title)$', title.lower()):
        return True
    # Detectar mala codificación de caracteres
    if re.search(r'[Ââ€™"""'']', title):
        return True
    return False

def mark_as_duplicate_to_delete(row):
    """Marca una fila para ser eliminada y limpia sus campos."""
    row['Mantener'] = ELIMINAR
    row[norm_key('Tono')] = TONO_DUPLICADA
    row[norm_key('Tema')] = TEMA_VACIO
    row[norm_key('Temas Generales - Tema')] = TEMA_VACIO

def get_row_priority_score(row):
    """
    Calcula una tupla de puntuación para una fila.
    Una puntuación más baja es mejor (se ordenará de menor a mayor).
    El objetivo es que la fila a conservar tenga la puntuación más baja.
    """
    original_title = str(row.get('original_titulo', ''))
    cleaned_title = str(row.get(norm_key('Título'), ''))
    
    # Puntuación 0 (mejor) si el título original ya estaba limpio, 1 si no.
    cleanliness_score = 0 if original_title == cleaned_title else 1
    
    # Puntuación 0 si el título contiene comillas (prioridad alta), 1 si no.
    quotes_score = 0 if '"' in original_title else 1
    
    # Puntuación de fecha/hora (más reciente es mejor, por eso se niega)
    fecha_obj = parse_date_obj(row.get(norm_key('Fecha')))
    hora_val = row.get(norm_key('Hora')) or datetime.time.min
    
    # Índice de fila original como desempate final
    original_index = row.get('original_row_index', float('inf'))

    return (cleanliness_score, quotes_score, -fecha_obj.toordinal(), -hora_val.hour, -hora_val.minute, original_index)


# --- FUNCIÓN PRINCIPAL ---
def run_deduplication_process(wb, empresa_dict, internet_dict, region_dict):
    """
    Procesa el workbook aplicando expansión por menciones, mapeos y deduplicación completa.
    """
    sheet = wb.active

    # --- REQUISITO: ESTILO DE LINK MODIFICADO ---
    # Texto negro, sin subrayado.
    link_style_no_underline = NamedStyle(name="LinkNegroSinSubrayado", 
                                         font=Font(color="000000", underline=None), 
                                         alignment=Alignment(horizontal="left"))
    if link_style_no_underline.name not in wb.named_styles:
        wb.add_named_style(link_style_no_underline)
        
    headers = [cell.value for cell in sheet[1]]
    headers_norm = [norm_key(h) for h in headers]
    processed_rows = []

    # --- FASE 1: EXPANSIÓN POR MENCIONES Y LIMPIEZA INICIAL ---
    for row_idx, row_cells in enumerate(sheet.iter_rows(min_row=2)):
        if all(c.value is None for c in row_cells): 
            continue
            
        base_data = {'original_row_index': row_idx + 2}
        for i, cell in enumerate(row_cells):
            col_name = headers_norm[i]
            if col_name in [norm_key('Link Nota'), norm_key('Link (Streaming - Imagen)')]:
                base_data[col_name] = extract_link(cell)
            else:
                base_data[col_name] = cell.value

        titulo_key = norm_key('Título')
        original_title = str(base_data.get(titulo_key, ''))
        base_data['original_titulo'] = original_title
        base_data[titulo_key] = convert_html_entities(original_title)
        
        base_data[norm_key('Resumen - Aclaracion')] = corregir_texto(base_data.get(norm_key('Resumen - Aclaracion')))
        
        tipo_medio_key = norm_key('Tipo de Medio')
        tm_norm = norm_key(base_data.get(tipo_medio_key))
        if tm_norm in {'aire', 'cable'}: base_data[tipo_medio_key] = 'Televisión'
        elif tm_norm in {'am', 'fm'}: base_data[tipo_medio_key] = 'Radio'
        elif tm_norm == 'diario': base_data[tipo_medio_key] = 'Prensa'
        elif tm_norm == 'online': base_data[tipo_medio_key] = 'Internet'
        elif tm_norm == 'revista': base_data[tipo_medio_key] = 'Revista'
        
        link_nota_key, link_streaming_key = norm_key("Link Nota"), norm_key("Link (Streaming - Imagen)")
        if es_internet({'Tipo de Medio': base_data.get(tipo_medio_key)}):
            base_data[link_nota_key], base_data[link_streaming_key] = base_data.get(link_streaming_key), base_data.get(link_nota_key)
        else:
            base_data[link_streaming_key] = None

        menciones_key = norm_key('Menciones - Empresa')
        menciones_str = str(base_data.get(menciones_key) or '')
        menciones = [m.strip() for m in menciones_str.split(';') if m.strip()]
        
        if not menciones:
            processed_rows.append(base_data)
        else:
            for mencion in menciones:
                new_row = deepcopy(base_data)
                mencion_limpia = mencion.lower().strip()
                new_row[menciones_key] = empresa_dict.get(mencion_limpia, mencion)
                processed_rows.append(new_row)

    # --- FASE 2: MAPEOS Y PREPARACIÓN ---
    for row in processed_rows:
        if es_internet(row):
            medio_val = str(row.get(norm_key('Medio'), '')).lower().strip()
            if medio_val in internet_dict:
                row[norm_key('Medio')] = internet_dict[medio_val]
        
        medio_actual_val = str(row.get(norm_key('Medio'), '')).lower().strip()
        row[norm_key('Región')] = region_dict.get(medio_actual_val, "Online")

        row.update({'Duplicada': NO, 'Posible Duplicada': NO, 'Mantener': CONSERVAR})
        
        if is_title_problematic(row.get(norm_key('Título'))):
            row[DUPLICADA] = SI
            mark_as_duplicate_to_delete(row)

    # --- FASE 3: DETECTAR DUPLICADOS EXACTOS ---
    grupos_exactos = defaultdict(list)
    for idx, row in enumerate(processed_rows):
        if row['Mantener'] == ELIMINAR: continue
            
        key_parts = [
            normalize_title(row.get(norm_key('Título'))),
            norm_key(row.get(norm_key('Medio'))),
            norm_key(row.get(norm_key('Menciones - Empresa'))),
            format_date(row.get(norm_key('Fecha')))
        ]
        if not es_internet(row): key_parts.append(str(row.get(norm_key('Hora'))))
        
        grupos_exactos[tuple(key_parts)].append(idx)

    for indices in grupos_exactos.values():
        if len(indices) > 1:
            indices.sort(key=lambda i: get_row_priority_score(processed_rows[i]))
            for pos, idx in enumerate(indices):
                processed_rows[idx][DUPLICADA] = SI
                if pos > 0: mark_as_duplicate_to_delete(processed_rows[idx])

    # --- FASE 4: DETECTAR DUPLICADOS POR SIMILITUD ---
    SIMILARIDAD_MINIMA = 0.85
    grupos_para_similitud = defaultdict(list)
    
    for idx, row in enumerate(processed_rows):
        if row[DUPLICADA] == NO and row['Mantener'] == CONSERVAR:
            key_parts = [
                norm_key(row.get(norm_key('Medio'))),
                norm_key(row.get(norm_key('Menciones - Empresa'))),
                format_date(row.get(norm_key('Fecha')))
            ]
            if not es_internet(row): key_parts.append(str(row.get(norm_key('Hora'))))
            grupos_para_similitud[tuple(key_parts)].append(idx)

    for group in grupos_para_similitud.values():
        if len(group) < 2: continue
        
        # Agrupar por similitud de títulos con Union-Find
        parent = {i: i for i in group}
        def find(i):
            if parent[i] == i: return i
            parent[i] = find(parent[i])
            return parent[i]
        def union(i, j):
            root_i, root_j = find(i), find(j)
            if root_i != root_j: parent[root_j] = root_i
        
        for i in range(len(group)):
            for j in range(i + 1, len(group)):
                idx_i, idx_j = group[i], group[j]
                title_i = normalize_title(processed_rows[idx_i].get(norm_key('Título')))
                title_j = normalize_title(processed_rows[idx_j].get(norm_key('Título')))
                if title_i and title_j and SequenceMatcher(None, title_i, title_j).ratio() >= SIMILARIDAD_MINIMA:
                    union(idx_i, idx_j)
        
        clusters = defaultdict(list)
        for i in group: clusters[find(i)].append(i)
        
        for cluster in clusters.values():
            if len(cluster) > 1:
                cluster.sort(key=lambda i: get_row_priority_score(processed_rows[i]))
                for pos, idx in enumerate(cluster):
                    processed_rows[idx]['Posible Duplicada'] = SI
                    if pos > 0: mark_as_duplicate_to_delete(processed_rows[idx])
    
    # --- FASE 5: LIMPIEZA FINAL Y ORDENAMIENTO ---

    # Limpiar títulos de las filas que se conservan
    for row in processed_rows:
        if row['Mantener'] == CONSERVAR:
            titulo_key = norm_key('Título')
            title = str(row.get(titulo_key, ''))
            row[titulo_key] = re.sub(r'\s*\|\s*[\w\s]+$', '', title).strip()

    # --- REQUISITO: ORDENAMIENTO FINAL DEL REPORTE ---
    # Ordenar primero por 'Título' (A-Z) y luego por 'Medio' (A-Z)
    processed_rows.sort(key=lambda r: (
        str(r.get(norm_key('Título'), '')).lower(),
        str(r.get(norm_key('Medio'), '')).lower()
    ))
    
    # --- FASE 6: GENERACIÓN DEL REPORTE FINAL ---
    final_order = [
        "ID Noticia", "Fecha", "Hora", "Medio", "Tipo de Medio", "Sección - Programa", "Región",
        "Título", "Autor - Conductor", "Nro. Pagina", "Dimensión", "Duración - Nro. Caracteres", 
        "CPE", "Tier", "Audiencia", "Tono", "Tema", "Temas Generales - Tema", 
        "Resumen - Aclaracion", "Link Nota", "Link (Streaming - Imagen)", "Menciones - Empresa", 
        "Duplicada", "Posible Duplicada", "Mantener"
    ]
    
    new_wb = openpyxl.Workbook()
    new_sheet = new_wb.active
    new_sheet.title = "Resultado Depurado"
    new_sheet.append(final_order)
    
    if link_style_no_underline.name not in new_wb.named_styles:
        new_wb.add_named_style(link_style_no_underline)

    link_nota_idx = final_order.index("Link Nota")
    link_streaming_idx = final_order.index("Link (Streaming - Imagen)")

    for row_data in processed_rows:
        new_row_to_append = [row_data.get(norm_key(h), row_data.get(h)) for h in final_order]
        new_sheet.append(new_row_to_append)
        
        current_row_idx = new_sheet.max_row
        
        # Aplicar hipervínculos con el nuevo estilo
        link_nota_data = row_data.get(norm_key("Link Nota"))
        if isinstance(link_nota_data, dict) and link_nota_data.get("url"):
            cell = new_sheet.cell(row=current_row_idx, column=link_nota_idx + 1)
            cell.hyperlink = link_nota_data["url"]
            cell.value = "Link"
            cell.style = link_style_no_underline.name
        
        link_stream_data = row_data.get(norm_key("Link (Streaming - Imagen)"))
        if isinstance(link_stream_data, dict) and link_stream_data.get("url"):
            cell = new_sheet.cell(row=current_row_idx, column=link_streaming_idx + 1)
            cell.hyperlink = link_stream_data["url"]
            cell.value = "Link"
            cell.style = link_style_no_underline.name
    
    # Resumen para la app
    to_eliminate_count = sum(1 for r in processed_rows if r['Mantener'] == ELIMINAR)
    summary = {
        "total_rows": len(processed_rows),
        "to_eliminate": to_eliminate_count,
        "to_conserve": len(processed_rows) - to_eliminate_count,
        "exact_duplicates": sum(1 for r in processed_rows if r[DUPLICADA] == SI),
        "possible_duplicates": sum(1 for r in processed_rows if r['Posible Duplicada'] == SI and r[DUPLICADA] == NO)
    }
    
    return new_wb, summary
