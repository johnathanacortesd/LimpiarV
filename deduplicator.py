# deduplicator.py

import openpyxl
from openpyxl.styles import Font, Alignment, NamedStyle
from difflib import SequenceMatcher
from collections import defaultdict
import re
import datetime
from copy import deepcopy
import html

# --- FUNCIONES AUXILIARES ---
def norm_key(text): 
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
    # Remover sufijos como "| Medio Name"
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
    
    # Asegurar que termine con puntos suspensivos
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
            return datetime.date(1, 1, 1)
    return datetime.date(1, 1, 1)

def es_internet(row): 
    """Verifica si el 'Tipo de Medio' es 'Internet'."""
    return norm_key(row.get(norm_key('Tipo de Medio'))) == 'internet'

def es_radio_o_tv(row): 
    """Verifica si el 'Tipo de Medio' es 'Radio' o 'Televisión'."""
    return norm_key(row.get(norm_key('Tipo de Medio'))) in {'radio', 'televisión'}

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
    # Detectar sufijos problemáticos que no se limpiaron bien
    if re.search(r'\s*\|\s*[\w\s]+$', title):
        return True
    return False

def mark_as_duplicate_to_delete(row):
    """Marca una fila para ser eliminada."""
    row['Mantener'] = "Eliminar"
    row[norm_key('Tono')] = "Duplicada"
    row[norm_key('Tema')] = "-"
    row[norm_key('Temas Generales - Tema')] = "-"

def get_title_priority(row):
    """Asigna una puntuación de prioridad basada en el formato del título y el medio."""
    medio_key = norm_key('Medio')
    titulo_key = norm_key('Título')
    medio_norm = norm_key(row.get(medio_key))
    titulo_str = str(row.get(titulo_key, ''))
    
    # Priorizar ciertos medios con formatos específicos
    if medio_norm == norm_key('El Colombiano (Online)'):
        return 1 if '| El Colombiano' in titulo_str else 0
    if medio_norm == norm_key('El Nuevo Siglo (Online)'):
        return 1 if titulo_str.strip().endswith('El Nuevo Siglo') else 0
    return 0

def get_title_cleanliness_score(row):
    """
    Asigna una puntuación de "limpieza". Un título es "sucio" si su versión original
    es diferente a la versión limpia. Puntuación más baja es mejor.
    """
    original_title = str(row.get('original_titulo', ''))
    cleaned_title = str(row.get(norm_key('Título'), ''))
    return 0 if original_title == cleaned_title else 1

# --- FUNCIÓN PRINCIPAL ---
def run_deduplication_process(wb, empresa_dict, internet_dict, region_dict):
    """
    Procesa el workbook aplicando expansión por menciones, mapeos y deduplicación completa.
    """
    sheet = wb.active
    custom_link_style = NamedStyle(name="CustomLink", 
                                 font=Font(color="0000FF", underline="single"), 
                                 alignment=Alignment(horizontal="left"))
    if "CustomLink" not in wb.named_styles: 
        wb.add_named_style(custom_link_style)
        
    headers = [cell.value for cell in sheet[1]]
    headers_norm = [norm_key(h) for h in headers]
    processed_rows = []

    # --- FASE 1: EXPANSIÓN POR MENCIONES Y LIMPIEZA INICIAL ---
    for row_idx, row_cells in enumerate(sheet.iter_rows(min_row=2)):
        if all(c.value is None for c in row_cells): 
            continue
            
        # Crear diccionario base con datos de la fila
        base_data = {'original_row_index': row_idx + 2}
        for i, cell in enumerate(row_cells):
            col_name = headers_norm[i]
            if col_name in [norm_key('Link Nota'), norm_key('Link (Streaming - Imagen)')]:
                base_data[col_name] = extract_link(cell)
            else:
                base_data[col_name] = cell.value

        # Guardar título original antes de limpiarlo
        titulo_key = norm_key('Título')
        original_title = str(base_data.get(titulo_key, ''))
        base_data['original_titulo'] = original_title
        base_data[titulo_key] = convert_html_entities(original_title)
        
        # Limpiar el resumen
        base_data[norm_key('Resumen - Aclaracion')] = corregir_texto(
            base_data.get(norm_key('Resumen - Aclaracion'))
        )
        
        # Normalizar tipos de medio
        tipo_medio_key = norm_key('Tipo de Medio')
        tm_norm = norm_key(base_data.get(tipo_medio_key))
        if tm_norm in {'aire', 'cable'}: 
            base_data[tipo_medio_key] = 'Televisión'
        elif tm_norm in {'am', 'fm'}: 
            base_data[tipo_medio_key] = 'Radio'
        elif tm_norm == 'diario': 
            base_data[tipo_medio_key] = 'Prensa'
        elif tm_norm == 'online': 
            base_data[tipo_medio_key] = 'Internet'
        elif tm_norm == 'revista': 
            base_data[tipo_medio_key] = 'Revista'
        
        # Reorganizar links según tipo de medio
        link_nota_key, link_streaming_key = norm_key("Link Nota"), norm_key("Link (Streaming - Imagen)")
        tipo_medio_val = base_data.get(tipo_medio_key)
        
        if tipo_medio_val == "Internet":
            # Para internet: intercambiar los links
            base_data[link_nota_key], base_data[link_streaming_key] = (
                base_data.get(link_streaming_key), base_data.get(link_nota_key)
            )
        elif tipo_medio_val in {"Prensa", "Revista"}:
            # Si no hay link_nota pero sí streaming, mover streaming a nota
            is_link_nota_empty = (not base_data.get(link_nota_key) 
                                or not base_data.get(link_nota_key, {}).get('url'))
            has_streaming_link = base_data.get(link_streaming_key, {}).get('url')
            if is_link_nota_empty and has_streaming_link:
                base_data[link_nota_key] = base_data.get(link_streaming_key)
            base_data[link_streaming_key] = None
        elif tipo_medio_val in {"Radio", "Televisión"}: 
            # Para radio/TV: limpiar streaming
            base_data[link_streaming_key] = None
        
        # Expandir por menciones de empresa
        menciones_key = norm_key('Menciones - Empresa')
        menciones_str = str(base_data.get(menciones_key) or '')
        menciones = [m.strip() for m in menciones_str.split(';') if m.strip()]
        
        if not menciones:
            processed_rows.append(base_data)
        else:
            for mencion in menciones:
                new_row = deepcopy(base_data)
                # Aplicar mapeo de empresas
                mencion_limpia = mencion.lower().strip()
                new_row[menciones_key] = empresa_dict.get(mencion_limpia, mencion)
                processed_rows.append(new_row)

    # --- FASE 2: APLICAR MAPEOS DE INTERNET Y REGIÓN ---
    medio_key = norm_key('Medio')
    tipo_medio_key = norm_key('Tipo de Medio')
    region_key = norm_key('Región')
    
    for row in processed_rows:
        # Mapeo de Internet
        if str(row.get(tipo_medio_key, '')).lower().strip() == 'internet':
            medio_val = str(row.get(medio_key, '')).lower().strip()
            if medio_val in internet_dict:
                row[medio_key] = internet_dict[medio_val]
        
        # Mapeo de Región (aplicar después del mapeo de Internet)
        medio_actual_val = str(row.get(medio_key, '')).lower().strip()
        row[region_key] = region_dict.get(medio_actual_val, "Online")

    # --- FASE 3: INICIALIZAR CAMPOS DE DEDUPLICACIÓN ---
    for row in processed_rows:
        row.update({
            'Duplicada': "FALSE",
            'Posible Duplicada': "FALSE",
            'Mantener': "Conservar"
        })

    # --- FASE 4: MARCAR TÍTULOS PROBLEMÁTICOS ---
    for row in processed_rows:
        if is_title_problematic(row.get(norm_key('Título'))):
            row['Duplicada'] = "Sí"
            mark_as_duplicate_to_delete(row)

    # --- FASE 5: DETECTAR DUPLICADOS EXACTOS ---
    grupos_exactos = defaultdict(list)
    for idx, row in enumerate(processed_rows):
        if row['Mantener'] == 'Eliminar': 
            continue
            
        key_tuple = (
            normalize_title(row.get(norm_key('Título'))),
            norm_key(row.get(norm_key('Medio'))),
            norm_key(row.get(norm_key('Menciones - Empresa'))),
            format_date(row.get(norm_key('Fecha')))
        )
        
        # Para medios que no son internet, incluir la hora en la clave
        if not es_internet(row):
            key_tuple += (str(row.get(norm_key('Hora'))),)
            
        grupos_exactos[key_tuple].append(idx)

    # Procesar grupos de duplicados exactos
    for indices in grupos_exactos.values():
        if len(indices) > 1:
            # Ordenar por prioridad (el mejor queda primero)
            indices.sort(key=lambda i: processed_rows[i].get('original_row_index'))
            indices.sort(key=lambda i: '"' in str(processed_rows[i].get(norm_key('Título'), '')), reverse=True)
            indices.sort(key=lambda i: get_title_priority(processed_rows[i]), reverse=True)
            indices.sort(key=lambda i: get_title_cleanliness_score(processed_rows[i]))
            
            # Marcar todos como duplicados, eliminar todos excepto el primero
            for pos, idx in enumerate(indices):
                processed_rows[idx]['Duplicada'] = "Sí"
                if pos > 0:
                    mark_as_duplicate_to_delete(processed_rows[idx])

    # --- FASE 6: DETECTAR DUPLICADOS POR SIMILITUD ---
    SIMILARIDAD_MINIMA = 0.85
    grupos_para_similitud = defaultdict(list)
    
    # Agrupar noticias no duplicadas por criterios similares
    for idx, row in enumerate(processed_rows):
        if row['Duplicada'] == 'FALSE' and row['Mantener'] == 'Conservar':
            key_tuple = (
                norm_key(row.get(norm_key('Medio'))),
                norm_key(row.get(norm_key('Menciones - Empresa'))),
                format_date(row.get(norm_key('Fecha')))
            )
            
            # Para medios que no son internet, incluir la hora
            if not es_internet(row):
                key_tuple += (str(row.get(norm_key('Hora'))),)
                
            grupos_para_similitud[key_tuple].append(idx)

    # Procesar similitud dentro de cada grupo
    for group in grupos_para_similitud.values():
        if len(group) < 2: 
            continue
            
        # Usar Union-Find para encontrar clusters de títulos similares
        parent = {i: i for i in group}
        
        def find(x):
            if parent[x] == x: 
                return x
            parent[x] = find(parent[x])
            return parent[x]
        
        def union(x, y):
            rx, ry = find(x), find(y)
            if rx != ry: 
                parent[ry] = rx
        
        # Comparar todos los pares y unir los similares
        for i in range(len(group)):
            for j in range(i + 1, len(group)):
                idx_i, idx_j = group[i], group[j]
                row_i, row_j = processed_rows[idx_i], processed_rows[idx_j]
                
                title_i = normalize_title(row_i.get(norm_key('Título')))
                title_j = normalize_title(row_j.get(norm_key('Título')))
                
                if (title_i and title_j and 
                    SequenceMatcher(None, title_i, title_j).ratio() >= SIMILARIDAD_MINIMA):
                    union(idx_i, idx_j)
        
        # Crear clusters y procesar duplicados
        clusters = defaultdict(list)
        for i in group: 
            clusters[find(i)].append(i)
        
        for cluster in clusters.values():
            if len(cluster) > 1:
                # Ordenar por prioridad dentro del cluster
                cluster.sort(key=lambda i: (
                    parse_date_obj(processed_rows[i].get(norm_key('Fecha'))), 
                    processed_rows[i].get(norm_key('Hora')) or datetime.time(0, 0)
                ), reverse=True)
                cluster.sort(key=lambda i: '"' in str(processed_rows[i].get(norm_key('Título'), '')), reverse=True)
                cluster.sort(key=lambda i: get_title_priority(processed_rows[i]), reverse=True)
                cluster.sort(key=lambda i: get_title_cleanliness_score(processed_rows[i]))
                
                # Marcar como posibles duplicados
                for pos, idx in enumerate(cluster):
                    processed_rows[idx]['Posible Duplicada'] = "Sí"
                    if pos > 0 and processed_rows[idx]['Mantener'] != "Eliminar":
                        mark_as_duplicate_to_delete(processed_rows[idx])

    # --- FASE 7: GENERACIÓN DEL REPORTE FINAL ---
    final_order = [
        "ID Noticia", "Fecha", "Hora", "Medio", "Tipo de Medio", "Sección - Programa", "Región",
        "Título", "Autor - Conductor", "Nro. Pagina", "Dimensión", "Duración - Nro. Caracteres", 
        "CPE", "Tier", "Audiencia", "Tono", "Tema", "Temas Generales - Tema", 
        "Resumen - Aclaracion", "Link Nota", "Link (Streaming - Imagen)", "Menciones - Empresa", 
        "Duplicada", "Posible Duplicada", "Mantener"
    ]
    
    new_wb = openpyxl.Workbook()
    new_sheet = new_wb.active
    new_sheet.title = "Resultado"
    new_sheet.append(final_order)
    
    if "CustomLink" not in new_wb.named_styles: 
        new_wb.add_named_style(custom_link_style)

    # Ordenar filas por índice original para mantener el orden
    processed_rows.sort(key=lambda r: r.get('original_row_index', 0))

    # Aplicar limpieza final de títulos solo para filas que se conservan
    for row_data in processed_rows:
        if row_data['Mantener'] == 'Conservar':
            titulo_key = norm_key('Título')
            # Limpiar sufijos problemáticos del título final
            title = str(row_data.get(titulo_key, ''))
            title = re.sub(r'\s*\|\s*[\w\s]+$', '', title).strip()
            row_data[titulo_key] = title
        
        # Agregar fila al nuevo sheet
        new_row_to_append = []
        for header in final_order:
            key = norm_key(header)
            val = row_data.get(key, row_data.get(header, None))
            if isinstance(val, dict):
                new_row_to_append.append(val.get('value'))
            else:
                new_row_to_append.append(val)
        new_sheet.append(new_row_to_append)
    
    # Agregar hipervínculos
    link_nota_idx = final_order.index("Link Nota")
    link_streaming_idx = final_order.index("Link (Streaming - Imagen)")
    
    for i, row_cells in enumerate(new_sheet.iter_rows(min_row=2)):
        if i < len(processed_rows):
            processed = processed_rows[i]
            
            # Link Nota
            link_data = processed.get(norm_key("Link Nota"))
            if link_data and isinstance(link_data, dict) and link_data.get("url"):
                cell = row_cells[link_nota_idx]
                cell.hyperlink = link_data["url"]
                cell.value = "Link"
                cell.style = "CustomLink"
            
            # Link Streaming
            link_data_stream = processed.get(norm_key("Link (Streaming - Imagen)"))
            if link_data_stream and isinstance(link_data_stream, dict) and link_data_stream.get("url"):
                cell = row_cells[link_streaming_idx]
                cell.hyperlink = link_data_stream["url"]
                cell.value = "Link"
                cell.style = "CustomLink"
    
    # Eliminar la hoja original
    if wb.active in wb.worksheets:
        wb.remove(wb.active)
    
    # Calcular resumen
    summary = {
        "total_rows": len(processed_rows),
        "to_eliminate": sum(1 for r in processed_rows if r['Mantener'] == 'Eliminar'),
        "to_conserve": len(processed_rows) - sum(1 for r in processed_rows if r['Mantener'] == 'Eliminar'),
        "exact_duplicates": sum(1 for r in processed_rows if r['Duplicada'] == 'Sí'),
        "possible_duplicates": sum(1 for r in processed_rows if r['Posible Duplicada'] == 'Sí' and r['Duplicada'] == 'FALSE')
    }
    
    return new_wb, summary
