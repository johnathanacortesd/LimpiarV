# app.py (versión 3.2.0 - con Análisis IA Integrado)

# --- INICIO DE LAS IMPORTACIONES ---
import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment, NamedStyle
from difflib import SequenceMatcher
from collections import defaultdict, Counter
import re
import datetime
from copy import deepcopy
import html
import io
from itertools import combinations
import pandas as pd
import openai
import numpy as np
from sklearn.metrics.pairwise import cosine_similarity
from unidecode import unidecode
import time

# --- SECCIÓN 1: LÓGICA DE DEDUPLICACIÓN ---

# --- CONSTANTES Y FUNCIONES AUXILIARES (DEDUPLICACIÓN) ---
CONSERVAR = "Conservar"
ELIMINAR = "Eliminar"
SI = "Sí"
NO = "FALSE"
TONO_DUPLICADA = "Duplicada"
TEMA_VACIO = "-"

def norm_key(text):
    return re.sub(r'\W+', '', str(text).lower().strip()) if text else ""

def convert_html_entities(text):
    if not isinstance(text, str): return text
    return html.unescape(text)

def normalize_title(title):
    if not isinstance(title, str): return ""
    processed_title = convert_html_entities(title)
    processed_title = processed_title.replace('"', '"').replace('"', '"')
    processed_title = processed_title.replace("'", "'").replace("'", "'")
    processed_title = processed_title.lower().strip()
    processed_title = re.sub(r'\s*\|\s*[\w\s]+$', '', processed_title)
    processed_title = re.sub(r'\W+', '', processed_title)
    return processed_title

def corregir_texto(text):
    if not isinstance(text, str): return text
    text = convert_html_entities(text)
    text = text.replace('<br>', ' ').replace('[...]', ' ')
    text = re.sub(r'\s+', ' ', text).strip()
    if match := re.search(r'[A-Z]', text): text = text[match.start():]
    if text and not text.endswith('...'):
        text = re.sub(r'[\.,;:]$', '', text.strip()).strip() + '...'
    return text

def extract_link(cell):
    if cell.hyperlink and cell.hyperlink.target:
        return {"value": cell.value or "Link", "url": cell.hyperlink.target}
    if cell.value and isinstance(cell.value, str):
        if match := re.search(r'=HYPERLINK\("([^"]+)"', cell.value):
            return {"value": "Link", "url": match.group(1)}
    return {"value": cell.value, "url": None}

def format_date(date_val):
    if isinstance(date_val, datetime.datetime): return date_val.strftime('%Y-%m-%d')
    if isinstance(date_val, str):
        try: return datetime.datetime.strptime(date_val.split(' ')[0], '%Y-%m-%d').strftime('%Y-%m-%d')
        except ValueError: return str(date_val)
    return str(date_val)

def parse_date_obj(date_val):
    if isinstance(date_val, (datetime.datetime, datetime.date)): return date_val
    if isinstance(date_val, str):
        try:
            return datetime.datetime.strptime(date_val.split(' ')[0], '%d/%m/%Y' if '/' in date_val else '%Y-%m-%d')
        except (ValueError, AttributeError): return datetime.datetime.min
    return datetime.datetime.min

def parse_time_obj(time_val):
    if isinstance(time_val, datetime.time): return time_val
    if isinstance(time_val, datetime.datetime): return time_val.time()
    if isinstance(time_val, str):
        for fmt in ("%H:%M:%S", "%H:%M"):
            try: return datetime.datetime.strptime(time_val, fmt).time()
            except ValueError: continue
    return datetime.time.min

def es_internet(row):
    return norm_key(row.get(norm_key('Tipo de Medio'))) == 'internet'

def is_title_problematic(title):
    if not isinstance(title, str) or not title.strip(): return True
    if re.search(r'^(sin t[íi]tulo|untitled|no title)$', title.lower()): return True
    if re.search(r'[Ââ€™"""'']', title): return True
    if re.search(r'\s*\|\s*[\w\s]+$', title): return True
    return False

def get_quote_priority_score(row):
    title = str(row.get('original_titulo', ''))
    return 0 if '"' in title or "'" in title else 1

def mark_as_duplicate_to_delete(row):
    row['Mantener'] = ELIMINAR
    row[norm_key('Tono')] = TONO_DUPLICADA
    row[norm_key('Tema')] = TEMA_VACIO
    row[norm_key('Temas Generales - Tema')] = TEMA_VACIO

def get_title_priority(row):
    medio_norm, titulo_str = norm_key(row.get(norm_key('Medio'))), str(row.get(norm_key('Título'), ''))
    if medio_norm == norm_key('El Colombiano (Online)'): return 1 if '| El Colombiano' in titulo_str else 0
    if medio_norm == norm_key('El Nuevo Siglo (Online)'): return 1 if titulo_str.strip().endswith('El Nuevo Siglo') else 0
    return 0

def get_title_cleanliness_score(row):
    return 0 if str(row.get('original_titulo', '')) == str(row.get(norm_key('Título'), '')) else 1

def are_rows_similar(row_i, row_j):
    SIMILARIDAD_MINIMA = 0.85
    if norm_key(row_i.get(norm_key('Medio'))) != norm_key(row_j.get(norm_key('Medio'))): return False
    if norm_key(row_i.get(norm_key('Menciones - Empresa'))) != norm_key(row_j.get(norm_key('Menciones - Empresa'))): return False
    date_i, date_j = parse_date_obj(row_i.get(norm_key('Fecha'))), parse_date_obj(row_j.get(norm_key('Fecha')))
    if not (date_i and date_j and abs((date_i - date_j).days) <= 1): return False
    title_i, title_j = normalize_title(row_i.get(norm_key('Título'))), normalize_title(row_j.get(norm_key('Título')))
    if not (title_i and title_j and SequenceMatcher(None, title_i, title_j).ratio() >= SIMILARIDAD_MINIMA): return False
    es_internet_i, es_internet_j = es_internet(row_i), es_internet(row_j)
    if es_internet_i and es_internet_j: return True
    if es_internet_i != es_internet_j: return False
    time_i, time_j = parse_time_obj(row_i.get(norm_key('Hora'))), parse_time_obj(row_j.get(norm_key('Hora')))
    if time_i and time_j and abs((time_i.hour * 60 + time_i.minute) - (time_j.hour * 60 + time_j.minute)) > 5: return False
    return True

class UnionFind:
    def __init__(self, n): self.parent = list(range(n))
    def find(self, i):
        if self.parent[i] == i: return i
        self.parent[i] = self.find(self.parent[i]); return self.parent[i]
    def union(self, i, j):
        root_i, root_j = self.find(i), self.find(j)
        if root_i != root_j: self.parent[root_i] = root_j

def run_deduplication_process(wb, internet_dict, region_dict):
    sheet = wb.active
    custom_link_style = NamedStyle(name="CustomLink", font=Font(color="0000FF", underline="single"), alignment=Alignment(horizontal="left"))
    if "CustomLink" not in wb.named_styles: wb.add_named_style(custom_link_style)
    headers, headers_norm = [cell.value for cell in sheet[1]], [norm_key(h.value) for h in sheet[1]]
    processed_rows = []
    
    for row_idx, row_cells in enumerate(sheet.iter_rows(min_row=2), start=2):
        if all(c.value is None for c in row_cells): continue
        base_data = {'original_row_index': row_idx}
        for i, cell in enumerate(row_cells):
            col_name = headers_norm[i]
            if col_name in [norm_key('Link Nota'), norm_key('Link (Streaming - Imagen)')]: base_data[col_name] = extract_link(cell)
            else: base_data[col_name] = cell.value
        titulo_key, original_title = norm_key('Título'), str(base_data.get(norm_key('Título'), ''))
        base_data['original_titulo'], base_data[titulo_key] = original_title, convert_html_entities(original_title)
        base_data[norm_key('Resumen - Aclaracion')] = corregir_texto(base_data.get(norm_key('Resumen - Aclaracion')))
        tipo_medio_map = {'aire': 'Televisión', 'cable': 'Televisión', 'am': 'Radio', 'fm': 'Radio', 'diario': 'Prensa', 'online': 'Internet', 'revista': 'Revista'}
        tipo_medio_key, tm_norm = norm_key('Tipo de Medio'), norm_key(base_data.get(norm_key('Tipo de Medio')))
        base_data[tipo_medio_key] = tipo_medio_map.get(tm_norm, base_data.get(tipo_medio_key))
        link_nota_key, link_streaming_key, tipo_medio_val = norm_key("Link Nota"), norm_key("Link (Streaming - Imagen)"), base_data.get(tipo_medio_key)
        if tipo_medio_val == "Internet": base_data[link_nota_key], base_data[link_streaming_key] = (base_data.get(link_streaming_key), base_data.get(link_nota_key))
        elif tipo_medio_val in {"Prensa", "Revista"}:
            if (not base_data.get(link_nota_key, {}).get('url')) and base_data.get(link_streaming_key, {}).get('url'): base_data[link_nota_key] = base_data.get(link_streaming_key)
            base_data[link_streaming_key] = None
        elif tipo_medio_val in {"Radio", "Televisión"}: base_data[link_streaming_key] = None
        menciones_key, menciones_str = norm_key('Menciones - Empresa'), str(base_data.get(norm_key('Menciones - Empresa')) or '')
        menciones = [m.strip() for m in menciones_str.split(';') if m.strip()]
        if not menciones: processed_rows.append(base_data)
        else:
            for mencion in menciones:
                new_row = deepcopy(base_data); new_row[menciones_key] = mencion; processed_rows.append(new_row)
    
    medio_key, tipo_medio_key, region_key = norm_key('Medio'), norm_key('Tipo de Medio'), norm_key('Región')
    for row in processed_rows:
        if str(row.get(tipo_medio_key, '')).lower().strip() == 'internet':
            row[medio_key] = internet_dict.get(str(row.get(medio_key, '')).lower().strip(), row.get(medio_key))
        row[region_key] = region_dict.get(str(row.get(medio_key, '')).lower().strip(), "Online")
        
    problematic_title_indices = {idx for idx, row in enumerate(processed_rows) if is_title_problematic(row.get(norm_key('Título')))}
    
    for row in processed_rows: row.update({'Duplicada': NO, 'Posible Duplicada': NO, 'Mantener': CONSERVAR, 'ID Fila Conservada': ''})

    grupos_exactos = defaultdict(list)
    for idx, row in enumerate(processed_rows):
        hora_key = str(row.get(norm_key('Hora'))) if not es_internet(row) else None
        key_tuple = (normalize_title(row.get(norm_key('Título'))), norm_key(row.get(norm_key('Medio'))), norm_key(row.get(norm_key('Menciones - Empresa'))), format_date(row.get(norm_key('Fecha'))), hora_key)
        grupos_exactos[key_tuple].append(idx)
    
    for indices in grupos_exactos.values():
        if len(indices) > 1:
            indices.sort(key=lambda i: (is_title_problematic(processed_rows[i].get(norm_key('Título'))), get_quote_priority_score(processed_rows[i]), get_title_cleanliness_score(processed_rows[i]), -get_title_priority(processed_rows[i]), processed_rows[i]['original_row_index']))
            winner_idx = indices[0]
            winner_id = str(processed_rows[winner_idx].get(norm_key('ID Noticia'), '')) or f"Fila Original {processed_rows[winner_idx].get('original_row_index', 'N/A')}"
            processed_rows[winner_idx]['Duplicada'] = SI
            for loser_idx in indices[1:]:
                mark_as_duplicate_to_delete(processed_rows[loser_idx])
                processed_rows[loser_idx]['Duplicada'], processed_rows[loser_idx]['ID Fila Conservada'] = SI, winner_id
    
    candidates_indices = [idx for idx, row in enumerate(processed_rows) if row['Mantener'] == CONSERVAR]
    uf = UnionFind(len(candidates_indices))
    for i_idx, j_idx in combinations(range(len(candidates_indices)), 2):
        original_i, original_j = candidates_indices[i_idx], candidates_indices[j_idx]
        if are_rows_similar(processed_rows[original_i], processed_rows[original_j]):
            uf.union(i_idx, j_idx)
            
    clusters_by_root = defaultdict(list)
    for i, original_idx in enumerate(candidates_indices): clusters_by_root[uf.find(i)].append(original_idx)

    for cluster_list in clusters_by_root.values():
        if len(cluster_list) > 1:
            cluster_list.sort(key=lambda i: (is_title_problematic(processed_rows[i].get(norm_key('Título'))), get_quote_priority_score(processed_rows[i]), get_title_cleanliness_score(processed_rows[i]), -get_title_priority(processed_rows[i]), -datetime.datetime.combine(parse_date_obj(processed_rows[i].get(norm_key('Fecha'))), parse_time_obj(processed_rows[i].get(norm_key('Hora')))).timestamp(), processed_rows[i]['original_row_index']))
            winner_idx = cluster_list[0]
            winner_id = str(processed_rows[winner_idx].get(norm_key('ID Noticia'), '')) or f"Fila Original {processed_rows[winner_idx].get('original_row_index', 'N/A')}"
            for idx in cluster_list: processed_rows[idx]['Posible Duplicada'] = SI
            for loser_idx in cluster_list[1:]:
                mark_as_duplicate_to_delete(processed_rows[loser_idx])
                processed_rows[loser_idx]['ID Fila Conservada'] = winner_id

    for row in processed_rows:
        if row['Mantener'] == ELIMINAR and not row.get('ID Fila Conservada'):
            row['ID Fila Conservada'] = f"Eliminada sin par - Fila {row.get('original_row_index', 'N/A')}"
            if row.get(norm_key('Tono')) != TONO_DUPLICADA: row[norm_key('Tono')] = TONO_DUPLICADA
    
    final_order = ["ID Noticia", "Fecha", "Hora", "Medio", "Tipo de Medio", "Sección - Programa", "Región", "Título", "Autor - Conductor", "Nro. Pagina", "Dimensión", "Duración - Nro. Caracteres", "CPE", "Tier", "Audiencia", "Tono", "Tema", "Temas Generales - Tema", "Resumen - Aclaracion", "Link Nota", "Link (Streaming - Imagen)", "Menciones - Empresa", "Duplicada", "Posible Duplicada", "Mantener", "ID Fila Conservada"]
    main_wb, summary_wb = openpyxl.Workbook(), openpyxl.Workbook()
    main_sheet, summary_sheet = main_wb.active, summary_wb.active
    main_sheet.title, summary_sheet.title = "Resultado", "Resumen"
    main_sheet.append(final_order), summary_sheet.append(["Resumen"])
    if "CustomLink" not in main_wb.named_styles: main_wb.add_named_style(custom_link_style)
    
    processed_rows.sort(key=lambda r: r.get('original_row_index', 0))
    for row_data in processed_rows:
        row_data[norm_key('Título')] = re.sub(r'\s*\|\s*[\w\s]+$', '', str(row_data.get(norm_key('Título'), ''))).strip()
        new_row_to_append = [row_data.get(norm_key(h), row_data.get(h, '')) for h in final_order]
        for i, val in enumerate(new_row_to_append):
            if isinstance(val, dict): new_row_to_append[i] = val.get('value')
        main_sheet.append(new_row_to_append)
        if row_data['Mantener'] == CONSERVAR:
            title_val, resumen_val = row_data.get(norm_key('Título'), ""), row_data.get(norm_key('Resumen - Aclaracion'), "")
            summary_sheet.append([f"{str(title_val)} {str(resumen_val)}".strip()])
    
    link_nota_idx, link_streaming_idx = final_order.index("Link Nota") + 1, final_order.index("Link (Streaming - Imagen)") + 1
    for row_idx, row in enumerate(processed_rows, start=2):
        for col_idx, key in [(link_nota_idx, norm_key("Link Nota")), (link_streaming_idx, norm_key("Link (Streaming - Imagen)"))]:
            link_data = row.get(key)
            if isinstance(link_data, dict) and link_data.get("url"):
                cell = main_sheet.cell(row=row_idx, column=col_idx)
                cell.hyperlink, cell.value, cell.style = link_data["url"], "Link", "CustomLink"
    
    summary = {"total_rows": len(processed_rows), "to_eliminate": sum(1 for r in processed_rows if r['Mantener'] == ELIMINAR), "to_conserve": sum(1 for r in processed_rows if r['Mantener'] == CONSERVAR), "exact_duplicates_marked": sum(1 for r in processed_rows if r['Duplicada'] == SI), "possible_duplicates_total": sum(1 for r in processed_rows if r['Posible Duplicada'] == SI), "problematic_titles": len(problematic_title_indices)}
    return main_wb, summary_wb, summary


# --- SECCIÓN 2: LÓGICA DE ANÁLISIS CON IA ---

class ClasificadorTonoNoticias:
    def __init__(self, marca, client_openai):
        self.marca, self.client = marca, client_openai
        self.embeddings_cache, self.tonos_asignados_cache, self.grupos_similares_info = {}, {}, {}
        self.high_similarity_threshold = 0.90
        self.OPENAI_MODEL_EMBEDDING = 'text-embedding-3-small'
        self.OPENAI_MODEL_CLASIFICACION = "gpt-4.1-mini-2025-04-14"

    def limpiar_texto(self, texto, lower=False):
        if pd.isna(texto): return ""
        texto_procesado = str(texto).strip()
        if lower: texto_procesado = texto_procesado.lower()
        return re.sub(r'\s+', ' ', texto_procesado)

    def obtener_embedding(self, texto_original):
        texto_limpio_cache = self.limpiar_texto(texto_original, lower=True)
        if not texto_limpio_cache: return None
        if texto_limpio_cache in self.embeddings_cache: return self.embeddings_cache[texto_limpio_cache]
        try:
            time.sleep(0.02)
            response = self.client.embeddings.create(input=[texto_limpio_cache], model=self.OPENAI_MODEL_EMBEDDING)
            embedding = response.data[0].embedding
            self.embeddings_cache[texto_limpio_cache] = embedding
            return embedding
        except Exception as e:
            st.warning(f"Advertencia (Tono Embedding): {e}")
            return None

    def calcular_similitud(self, emb1, emb2):
        if emb1 is None or emb2 is None: return 0
        return cosine_similarity(np.array(emb1).reshape(1, -1), np.array(emb2).reshape(1, -1))[0][0]

    def clasificar_tono_noticia_gpt(self, resumen_noticia):
        texto_analisis = self.limpiar_texto(resumen_noticia)
        if not texto_analisis: return "Neutro"
        if texto_analisis in self.tonos_asignados_cache: return self.tonos_asignados_cache[texto_analisis]
        prompt = f"""Eres un analista de medios experto. Tu tarea es clasificar el TONO PREDOMINANTE de la siguiente NOTICIA sobre "{self.marca}". Consideraciones: - POSITIVO: logros, resultados sólidos. - NEGATIVO: problemas, controversias. - NEUTRO: mención factual, incidental, en listado. NOTICIA: --- {texto_analisis} --- Responde ÚNICAMENTE con: POSITIVO, NEGATIVO o NEUTRO."""
        try:
            time.sleep(0.1)
            response = self.client.chat.completions.create(model=self.OPENAI_MODEL_CLASIFICACION, messages=[{"role": "system", "content": "Tu respuesta debe ser una única palabra: POSITIVO, NEGATIVO o NEUTRO."}, {"role": "user", "content": prompt}], max_tokens=5, temperature=0.0)
            tono = response.choices[0].message.content.strip().capitalize()
            if tono not in ["Positivo", "Negativo", "Neutro"]: tono = "Neutro"
            self.tonos_asignados_cache[texto_analisis] = tono
            return tono
        except Exception as e:
            st.warning(f"Advertencia (Clasificación Tono): {e}")
            return "Neutro (Error)"

    def detectar_grupos_similares(self, df, embeddings_list):
        n = len(df)
        adj = [[] for _ in range(n)]
        for i in range(n):
            if embeddings_list[i] is None: continue
            for j in range(i + 1, n):
                if embeddings_list[j] is None: continue
                if self.calcular_similitud(embeddings_list[i], embeddings_list[j]) >= self.high_similarity_threshold:
                    adj[i].append(j)
                    adj[j].append(i)
        
        visitados, grupos = [False] * n, []
        for i in range(n):
            if not visitados[i] and embeddings_list[i] is not None:
                grupo_actual, q = [], [i]
                visitados[i] = True
                while q:
                    u = q.pop(0)
                    grupo_actual.append(u)
                    for v_idx in adj[u]:
                        if not visitados[v_idx]:
                            visitados[v_idx] = True
                            q.append(v_idx)
                if grupo_actual: grupos.append(sorted(grupo_actual))
        
        item_to_rep_map = {}
        for grupo in grupos:
            rep_idx = grupo[0]
            self.grupos_similares_info[rep_idx] = {'tono': None, 'miembros': grupo}
            for miembro_idx in grupo:
                item_to_rep_map[miembro_idx] = rep_idx
        return item_to_rep_map

    def procesar_dataframe(self, df, columna_texto):
        if columna_texto not in df.columns:
            st.error(f"Error interno: La columna '{columna_texto}' no se encuentra para el análisis de Tono.")
            return df
        df['Tono_IA'] = "No Procesado"
        total_rows = len(df)
        status_text = st.empty()

        embeddings_list = [self.obtener_embedding(row[columna_texto]) for _, row in df.iterrows()]
        item_to_rep_map = self.detectar_grupos_similares(df, embeddings_list)
        
        for idx in range(total_rows):
            status_text.text(f"Analizando Tono: Fila {idx+1}/{total_rows}")
            if df.at[idx, 'Tono_IA'] != "No Procesado": continue
            if idx in item_to_rep_map:
                rep_idx = item_to_rep_map[idx]
                if self.grupos_similares_info[rep_idx]['tono'] is None:
                    self.grupos_similares_info[rep_idx]['tono'] = self.clasificar_tono_noticia_gpt(df.iloc[rep_idx][columna_texto])
                df.at[idx, 'Tono_IA'] = self.grupos_similares_info[rep_idx]['tono']
            else:
                df.at[idx, 'Tono_IA'] = self.clasificar_tono_noticia_gpt(df.iloc[idx][columna_texto])
        status_text.empty()
        return df

class ClasificadorTemasAvanzado:
    def __init__(self, client_openai):
        self.client = client_openai
        self.embeddings_cache, self.temas_cache = {}, {}
        self.OPENAI_MODEL_EMBEDDING = 'text-embedding-3-small'
        self.OPENAI_MODEL_CLASIFICACION = "gpt-4.1-mini-2025-04-14"

    def limpiar_texto(self, texto):
        return re.sub(r'\s+', ' ', str(texto).strip()) if pd.notna(texto) else ""

    def extraer_keywords_clave(self, texto):
        if not texto: return []
        stop_words = {'el','la','de','que','y','en','un','es','se','no','te','lo','le','da','su','por','son','con','para','al','del','los','las','una','sobre','todo','también','tras','otro','algún','muy','fue','han','más','hasta','desde','está','entre','cuando','todo','esta','ser','son','tiene','le','me','mi','tu','ti','nos','os','les','si','ya','solo','ni','puede','han'}
        palabras = re.findall(r'\b[a-záéíóúñ]{3,}\b', texto.lower())
        return [p for p, freq in Counter([p for p in palabras if p not in stop_words and not p.isdigit()]).most_common(10)]

    def obtener_embedding(self, texto):
        texto_limpio_cache = self.limpiar_texto(texto).lower()
        if not texto_limpio_cache: return None
        if texto_limpio_cache in self.embeddings_cache: return self.embeddings_cache.get(texto_limpio_cache)
        try:
            time.sleep(0.02)
            response = self.client.embeddings.create(input=[texto_limpio_cache], model=self.OPENAI_MODEL_EMBEDDING)
            embedding = response.data[0].embedding
            self.embeddings_cache[texto_limpio_cache] = embedding
            return embedding
        except Exception as e:
            st.warning(f"Advertencia (Tema Embedding): {e}")
            return None

    def calcular_similitud_semantica(self, emb1, emb2):
        return 0 if emb1 is None or emb2 is None else cosine_similarity(np.array(emb1).reshape(1, -1), np.array(emb2).reshape(1, -1))[0][0]
    
    def clustering_jerarquico_avanzado(self, df, embeddings_list):
        valid_indices = [i for i, emb in enumerate(embeddings_list) if emb is not None]
        if len(valid_indices) < 2: return {f"individual_{i}": [i] for i in valid_indices}
        n_valid = len(valid_indices)
        similarity_matrix = np.zeros((n_valid, n_valid))
        for i in range(n_valid):
            similarity_matrix[i,i] = 1.0
            for j in range(i + 1, n_valid):
                sim = self.calcular_similitud_semantica(embeddings_list[valid_indices[i]], embeddings_list[valid_indices[j]])
                similarity_matrix[i, j] = similarity_matrix[j, i] = sim
        clusters_finales, processed_indices = {}, set()
        for threshold in [0.85, 0.75]:
            visited = set()
            for i, idx_i in enumerate(valid_indices):
                if idx_i in processed_indices or idx_i in visited: continue
                cluster, queue = [], [i]; visited.add(idx_i)
                while queue:
                    current_i = queue.pop(0)
                    cluster.append(valid_indices[current_i])
                    for j, idx_j in enumerate(valid_indices):
                        if idx_j not in visited and idx_j not in processed_indices and similarity_matrix[current_i, j] >= threshold:
                            visited.add(idx_j); queue.append(j)
                if len(cluster) >= 2:
                    clusters_finales[f"cluster_{threshold}_{len(clusters_finales)}"] = cluster
                    processed_indices.update(cluster)
        for idx in valid_indices:
            if idx not in processed_indices: clusters_finales[f"individual_{idx}"] = [idx]
        return clusters_finales

    def generar_tema_inteligente(self, textos_cluster, keywords_cluster):
        contexto_resumido = " | ".join([str(texto)[:200] for texto in textos_cluster[:5]])
        keywords_str = ", ".join(keywords_cluster[:15])
        contexto_hash = str(hash(contexto_resumido + keywords_str))
        if contexto_hash in self.temas_cache: return self.temas_cache[contexto_hash]
        prompt = f"""Eres un experto analista de contenido. Analiza el siguiente grupo de textos similares y genera un TEMA PRINCIPAL muy específico y descriptivo. INSTRUCCIONES: - El tema debe ser conciso (máximo 6 palabras) - Debe capturar la esencia común de todos los textos - Debe ser específico, no genérico - Enfócate en el aspecto más relevante y distintivo. PALABRAS CLAVE IDENTIFICADAS: {keywords_str}. TEXTOS DEL GRUPO: --- {contexto_resumido} ---. Genera únicamente el tema, sin explicaciones adicionales."""
        try:
            time.sleep(0.1)
            response = self.client.chat.completions.create(model=self.OPENAI_MODEL_CLASIFICACION, messages=[{"role": "system", "content": "Eres un experto en síntesis temática que genera temas precisos y específicos."}, {"role": "user", "content": prompt}], max_tokens=40, temperature=0.2)
            tema = re.sub(r'\"|Tema:|TEMA:', '', response.choices[0].message.content).strip()
            if not tema or len(tema) < 3: tema = " ".join(keywords_cluster[:3]).title() or "Tema General"
            self.temas_cache[contexto_hash] = tema
            return tema
        except Exception as e:
            st.warning(f"Advertencia (Generación Tema): {e}")
            return " ".join(keywords_cluster[:3]).title() or "Tema General"

    def generar_tema_individual(self, texto):
        if not texto or pd.isna(texto): return "Sin Contenido"
        texto_hash, keywords = str(hash(str(texto))), self.extraer_keywords_clave(texto)
        if texto_hash in self.temas_cache: return self.temas_cache[texto_hash]
        prompt = f"""Analiza el siguiente texto y genera un TEMA específico y descriptivo en máximo 5 palabras. El tema debe ser: - Específico y distintivo - Relacionado con el contenido principal - Conciso pero informativo. TEXTO: --- {str(texto)[:400]} ---. Genera únicamente el tema."""
        try:
            time.sleep(0.1)
            response = self.client.chat.completions.create(model=self.OPENAI_MODEL_CLASIFICACION, messages=[{"role": "system", "content": "Genera temas específicos y concisos."}, {"role": "user", "content": prompt}], max_tokens=30, temperature=0.3)
            tema = re.sub(r'\"|Tema:|TEMA:', '', response.choices[0].message.content).strip()
            if not tema or len(tema) < 3: tema = " ".join(keywords[:3]).title() or "Tema General"
            self.temas_cache[texto_hash] = tema
            return tema
        except Exception as e:
            st.warning(f"Advertencia (Tema Individual): {e}")
            return " ".join(keywords[:3]).title() or "Tema General"

    def procesar_dataframe(self, df, columna_texto):
        if columna_texto not in df.columns:
            st.error(f"Error interno: La columna '{columna_texto}' no se encuentra para el análisis de Tema.")
            return df
        df['Tema_IA'] = "No Procesado"
        total_rows = len(df)
        status_text = st.empty()
        
        status_text.text("Generando representaciones semánticas para Temas...")
        embeddings_list = [self.obtener_embedding(row[columna_texto]) if pd.notna(row[columna_texto]) else None for _, row in df.iterrows()]
        
        status_text.text("Identificando grupos temáticos...")
        clusters = self.clustering_jerarquico_avanzado(df, embeddings_list)
        
        processed_count = 0
        for cluster_id, indices in clusters.items():
            processed_count += len(indices)
            status_text.text(f"Generando temas inteligentes... {processed_count}/{total_rows}")
            if len(indices) > 1:
                textos_cluster = [df.iloc[idx][columna_texto] for idx in indices if pd.notna(df.iloc[idx][columna_texto])]
                top_keywords = [kw for kw, freq in Counter([kw for texto in textos_cluster for kw in self.extraer_keywords_clave(texto)]).most_common(10)]
                tema_cluster = self.generar_tema_inteligente(textos_cluster, top_keywords)
                for idx in indices:
                    df.at[idx, 'Tema_IA'] = tema_cluster
            elif indices:
                df.at[indices[0], 'Tema_IA'] = self.generar_tema_individual(df.iloc[indices[0]][columna_texto])
        
        df.loc[df['Tema_IA'] == "No Procesado", 'Tema_IA'] = df.loc[df['Tema_IA'] == "No Procesado", columna_texto].apply(self.generar_tema_individual)
        status_text.empty()
        return df

def generar_resumen_estrategico(client, df, marca_analizada):
    datos_contexto = ""
    df_conservadas = df[df['Mantener'] == 'Conservar'].copy()
    volumen_total = len(df_conservadas)
    if volumen_total == 0: return "No hay datos conservados para generar un resumen."
    
    dist_tono = df_conservadas['Tono_IA'].value_counts().to_string().replace('\n', ', ')
    datos_contexto += f"- Diagnóstico General: {volumen_total} menciones analizadas. Distribución de tono: {dist_tono}.\n"
    
    top_temas = df_conservadas['Tema_IA'].value_counts().nlargest(5).to_string().replace('\n', ', ')
    datos_contexto += f"- Temas Principales: {top_temas}.\n"
    
    col_audiencia, col_cpe = None, None
    for col in df.columns:
        if 'audiencia' in col.lower(): col_audiencia = col
        if 'cpe' in col.lower(): col_cpe = col
        
    df_negativas = df_conservadas[df_conservadas['Tono_IA'] == 'Negativo']
    if col_audiencia and not df_negativas.empty:
        riesgos = df_negativas.nlargest(3, col_audiencia)[['Tema_IA', col_audiencia]].copy()
        if not riesgos.empty:
            riesgos[col_audiencia] = riesgos[col_audiencia].apply(lambda x: f"{x:,.0f}")
            datos_contexto += f"- Riesgos Potenciales (Temas Negativos con Mayor Audiencia):\n{riesgos.to_string(index=False)}\n"
    if col_cpe and not df_negativas.empty:
        riesgos_cpe = df_negativas.nlargest(3, col_cpe)[['Tema_IA', col_cpe]].copy()
        if not riesgos_cpe.empty:
                 riesgos_cpe[col_cpe] = riesgos_cpe[col_cpe].apply(lambda x: f"${x:,.0f}")
                 datos_contexto += f"- Riesgos Potenciales (Temas Negativos con Mayor CPE):\n{riesgos_cpe.to_string(index=False)}\n"

    prompt_final = f"""A partir de los siguientes datos resumidos sobre la presencia mediática de "{marca_analizada}":\n---\n{datos_contexto}\n---\nAnaliza los datos y presenta un informe ejecutivo conciso en español.\n**Instrucciones:**\n1. **Diagnóstico General:** Una o dos oraciones que resuman la situación mediática general, considerando el volumen y el tono predominante.\n2. **Puntos Clave:** Enumera 2-3 hallazgos cruciales, conectando los temas más importantes con su tono y alcance (si hay datos de CPE/Audiencia).\n3. **Riesgos Potenciales:** Identifica explícitamente 1-2 riesgos evidentes en los datos, por ejemplo, temas negativos con alta audiencia o costo.\n**NO INCLUYAS RECOMENDACIONES.** Limítate a diagnosticar con un lenguaje claro y directo, en formato de texto plano."""
    try:
        response = client.chat.completions.create(model="gpt-4-turbo", messages=[{"role": "system", "content": "Eres un analista de medios experto en resumir datos para informes ejecutivos concisos y directos."}, {"role": "user", "content": prompt_final}], max_tokens=500, temperature=0.2)
        resumen = response.choices[0].message.content
        return re.sub(r'(\*\*|__|\*|_|#+\s*|^\s*[\*\-]\s*|^\s*\d+\.\s*)', '', resumen, flags=re.MULTILINE)
    except Exception as e:
        st.warning(f"No se pudo generar el resumen estratégico: {e}")
        return "El resumen ejecutivo no pudo ser generado debido a un error en la API."


# --- SECCIÓN 3: FLUJO PRINCIPAL DE LA APLICACIÓN STREAMLIT ---
st.set_page_config(page_title="Intelli-Clean | Depurador y Analizador IA", page_icon="🤖", layout="wide", initial_sidebar_state="expanded")

def check_password():
    def password_entered():
        try:
            st.session_state["password_correct"] = st.session_state["password"] == st.secrets.password.password
            del st.session_state["password"]
        except (AttributeError, KeyError):
            st.session_state["password_correct"] = False
    try:
        _ = st.secrets.password.password
    except (AttributeError, KeyError):
        st.error("🚨 ¡Error de configuración! Contraseña no definida en 'Secrets'.")
        return False
    if not st.session_state.get("password_correct", False):
        c1, c2, c3 = st.columns([1, 1, 1]);
        with c2:
            st.markdown("<h1 style='text-align: center;'>🤖</h1>", unsafe_allow_html=True)
            st.markdown("<h3 style='text-align: center;'>Intelli-Clean Access</h3>", unsafe_allow_html=True)
            st.text_input("Contraseña", type="password", on_change=password_entered, key="password", placeholder="Introduce la contraseña", label_visibility="collapsed")
            if 'password' in st.session_state and st.session_state.password != "" and not st.session_state.password_correct:
                   st.error("😕 Contraseña incorrecta.")
        return False
    return True

if check_password():
    st.title("✨ Intelli-Clean: Depurador y Analizador de Noticias")
    st.caption("Una herramienta para mapear, limpiar, deduplicar y analizar tus informes con precisión.")
    st.divider()

    with st.sidebar:
        st.header("📂 Paso 1: Carga y Depuración")
        uploaded_main_file = st.file_uploader("1. Informe Principal de Noticias", type="xlsx", key="main_file")
        uploaded_internet_map = st.file_uploader("2. Mapeo de Medios de Internet", type="xlsx", key="internet_map")
        uploaded_region_map = st.file_uploader("3. Mapeo de Regiones", type="xlsx", key="region_map")
        st.divider()
        process_button = st.button("🚀 Depurar Archivos", type="primary", use_container_width=True)
        st.divider()
        st.info("Versión del Script: 3.2.0-CON-IA")

    # Inicializar el estado de la sesión
    for key in ['dedup_complete', 'ai_complete', 'summary', 'main_stream', 'df_depurado', 'summary_wb', 'resumen_ejecutivo', 'ai_analyzed_stream']:
        if key not in st.session_state:
            st.session_state[key] = False if 'complete' in key else ({} if key == 'summary' else None)

    # Lógica del botón de depuración
    if process_button:
        st.session_state.dedup_complete, st.session_state.ai_complete = False, False
        if all([uploaded_main_file, uploaded_internet_map, uploaded_region_map]):
            with st.status("Iniciando proceso de deduplicación... ⏳", expanded=True) as status:
                try:
                    status.write("Cargando archivos y creando diccionarios de mapeo...")
                    wb_main = openpyxl.load_workbook(uploaded_main_file)
                    internet_dict = {str(r[0].value).lower().strip(): str(r[1].value) for r in openpyxl.load_workbook(uploaded_internet_map, data_only=True).active.iter_rows(min_row=2) if r[0].value and r[1].value}
                    region_dict = {str(r[0].value).lower().strip(): str(r[1].value) for r in openpyxl.load_workbook(uploaded_region_map, data_only=True).active.iter_rows(min_row=2) if r[0].value and r[1].value}
                    
                    status.write("🧠 Ejecutando limpieza y deduplicación...")
                    final_wb, summary_wb, summary = run_deduplication_process(wb_main, internet_dict, region_dict)
                    
                    status.update(label="✅ ¡Deduplicación completada!", state="complete", expanded=False)
                    st.session_state.summary = summary
                    main_stream = io.BytesIO()
                    final_wb.save(main_stream)
                    st.session_state.main_stream = main_stream
                    main_stream.seek(0)
                    st.session_state.df_depurado = pd.read_excel(main_stream)
                    st.session_state.summary_wb = summary_wb
                    st.session_state.dedup_complete = True
                except Exception as e:
                    status.update(label="❌ Error en el proceso de deduplicación", state="error", expanded=True)
                    st.error(f"Ha ocurrido un error inesperado: {e}")
                    st.exception(e)
        else:
            st.warning("⚠️ Por favor, asegúrate de cargar los tres archivos requeridos.")

    # Mostrar resultados de la depuración si está completa
    if st.session_state.dedup_complete:
        st.header("Resultados de la Deduplicación")
        col1, col2, col3 = st.columns(3)
        col1.metric("Filas Totales Procesadas", st.session_state.summary.get('total_rows', 0))
        col2.metric("👍 Filas para Conservar", st.session_state.summary.get('to_conserve', 0))
        col3.metric("🗑️ Filas para Eliminar", st.session_state.summary.get('to_eliminate', 0))
        st.download_button(
            "1. Descargar Informe Principal Depurado",
            st.session_state.main_stream.getvalue(),
            f"Informe_Depurado_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.divider()
        
        # Mostrar sección de análisis IA
        st.header("🤖 Paso 2: Análisis con IA (Opcional)")
        with st.form("ai_analysis_form"):
            try:
                api_key = st.secrets["OPENAI_API_KEY"]
                st.success("API Key cargada desde Secrets.", icon="✅")
            except (KeyError, AttributeError):
                api_key = st.text_input("Ingresa tu API Key de OpenAI", type="password", help="No se encontraron Secrets. Ingresa la API Key manualmente.")
            
            marca = st.text_input("➡️ Ingresa la Marca o Cliente para el análisis", placeholder="Ej: Coca-Cola")
            ai_submit_button = st.form_submit_button("🧠 Analizar y Diagnosticar con IA", type="primary", use_container_width=True)

        # Lógica del botón de análisis IA
        if ai_submit_button:
            if not api_key or not marca:
                st.error("Es necesario proporcionar la API Key de OpenAI y el nombre de la Marca.")
            else:
                try:
                    # Preparar datos para IA
                    stream = io.BytesIO()
                    st.session_state.summary_wb.save(stream)
                    stream.seek(0)
                    df_for_ai = pd.read_excel(stream)
                    
                    if df_for_ai.empty:
                        st.warning("No hay datos para analizar con IA (el archivo de resúmenes está vacío).")
                    else:
                        client = openai.OpenAI(api_key=api_key)
                        with st.status("Realizando análisis con IA...", expanded=True) as status:
                            # Clasificar Tono
                            status.write("Clasificando Tono...")
                            clasificador_tono = ClasificadorTonoNoticias(marca, client)
                            df_con_tono = clasificador_tono.procesar_dataframe(df_for_ai.copy(), 'Resumen')
                            
                            # Clasificar Tema
                            status.write("Identificando Temas...")
                            clasificador_tema = ClasificadorTemasAvanzado(client)
                            df_analizado_ia = clasificador_tema.procesar_dataframe(df_con_tono.copy(), 'Resumen')
                            
                            # Fusionar resultados
                            df_final_completo = st.session_state.df_depurado.copy()
                            mask = df_final_completo['Mantener'] == 'Conservar'
                            if len(df_final_completo[mask]) == len(df_analizado_ia):
                                df_final_completo.loc[mask, 'Tono_IA'] = df_analizado_ia['Tono_IA'].values
                                df_final_completo.loc[mask, 'Tema_IA'] = df_analizado_ia['Tema_IA'].values
                            else:
                                raise Exception(f"Error de coincidencia de filas: {len(df_final_completo[mask])} conservadas vs {len(df_analizado_ia)} analizadas.")

                            # Generar Resumen
                            status.write("Generando Resumen Ejecutivo...")
                            st.session_state.resumen_ejecutivo = generar_resumen_estrategico(client, df_final_completo, marca)
                            
                            status.update(label="¡Análisis IA completado!", state="complete", expanded=False)
                        
                        # Preparar archivo de descarga final
                        output_stream = io.BytesIO()
                        with pd.ExcelWriter(output_stream, engine='openpyxl') as writer:
                            df_final_completo.to_excel(writer, index=False, sheet_name='Resultados_IA')
                            if st.session_state.resumen_ejecutivo:
                                summary_df = pd.DataFrame({'Diagnostico General': [st.session_state.resumen_ejecutivo]})
                                summary_df.to_excel(writer, index=False, sheet_name='Diagnostico_General')
                                worksheet = writer.sheets['Diagnostico_General']
                                worksheet.column_dimensions['A'].width = 120
                                cell = worksheet['A2']
                                cell.alignment = Alignment(wrap_text=True, vertical='top')
                        st.session_state.ai_analyzed_stream = output_stream
                        st.session_state.ai_complete = True

                except openai.AuthenticationError:
                    st.error("Error de autenticación con OpenAI. Verifica que tu API Key sea correcta.")
                except Exception as e:
                    st.error(f"Ocurrió un error inesperado durante el análisis IA: {e}")
                    st.exception(e)

    # Mostrar resultados del análisis IA si está completo
    if st.session_state.ai_complete:
        st.header("✅ Resultados del Análisis IA")
        if st.session_state.resumen_ejecutivo:
            st.subheader("📝 Resumen Ejecutivo Estratégico")
            st.text_area("Diagnóstico General", value=st.session_state.resumen_ejecutivo, height=250, disabled=True, help="Este resumen fue generado por IA a partir de los datos analizados.")
        if st.session_state.ai_analyzed_stream:
            st.download_button(
                "2. Descargar Reporte Completo con Diagnóstico",
                st.session_state.ai_analyzed_stream.getvalue(),
                f"Reporte_Analizado_IA_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    elif not st.session_state.dedup_complete and not process_button:
        st.info("Carga los archivos en el menú de la izquierda y haz clic en 'Depurar Archivos' para comenzar.")
