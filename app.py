# app.py

import streamlit as st
import openpyxl
import io
import datetime
import pandas as pd
import openai
import numpy as np
from sklearn.metrics.pairwise import cosine_similarity
from unidecode import unidecode
import re
import time
from collections import Counter
from deduplicator import run_deduplication_process

# --- Configuración de la página ---
st.set_page_config(page_title="Intelli-Clean | Depurador y Analizador IA", page_icon="🤖", layout="wide", initial_sidebar_state="expanded")

# --- Autenticación (sin cambios) ---
def check_password():
    def password_entered():
        try:
            if st.session_state["password"] == st.secrets.password.password:
                st.session_state["password_correct"] = True; del st.session_state["password"]
            else: st.session_state["password_correct"] = False
        except (AttributeError, KeyError): st.session_state["password_correct"] = False
    try: _ = st.secrets.password.password
    except (AttributeError, KeyError):
        st.error("🚨 ¡Error de configuración! Contraseña no definida en 'Secrets'."); return False
    if "password_correct" not in st.session_state: st.session_state["password_correct"] = False
    if not st.session_state["password_correct"]:
        c1, c2, c3 = st.columns([1, 1, 1]);
        with c2:
            st.markdown("<h1 style='text-align: center;'>🤖</h1>", unsafe_allow_html=True)
            st.markdown("<h3 style='text-align: center;'>Intelli-Clean Access</h3>", unsafe_allow_html=True)
            st.text_input("Contraseña", type="password", on_change=password_entered, key="password", placeholder="Introduce la contraseña", label_visibility="collapsed")
            if 'password' in st.session_state and st.session_state.password != "" and not st.session_state.password_correct:
                 st.error("😕 Contraseña incorrecta.")
        return False
    return True

# --- CLASES DE ANÁLISIS DE IA (INTEGRADAS, SIN CAMBIOS) ---
class ClasificadorTonoNoticias:
    def __init__(self, marca, client_openai):
        self.marca=marca; self.client=client_openai; self.embeddings_cache={}; self.tonos_asignados_cache={}; self.grupos_similares_info={}
        self.high_similarity_threshold=0.90; self.OPENAI_MODEL_EMBEDDING='text-embedding-3-small'; self.OPENAI_MODEL_CLASIFICACION="gpt-4.1-mini-2025-04-14"
    def limpiar_texto(self, texto, lower=False):
        if pd.isna(texto): return ""
        texto_procesado=str(texto).strip()
        if lower: texto_procesado=texto_procesado.lower()
        return re.sub(r'\s+', ' ', texto_procesado)
    def obtener_embedding(self, texto_original):
        texto_limpio_cache=self.limpiar_texto(texto_original, lower=True)
        if not texto_limpio_cache: return None
        if texto_limpio_cache in self.embeddings_cache: return self.embeddings_cache[texto_limpio_cache]
        try:
            time.sleep(0.02); response=self.client.embeddings.create(input=[texto_limpio_cache], model=self.OPENAI_MODEL_EMBEDDING)
            embedding=response.data[0].embedding; self.embeddings_cache[texto_limpio_cache]=embedding; return embedding
        except Exception as e: st.warning(f"Advertencia (Tono Embedding): {e}"); return None
    def calcular_similitud(self, emb1, emb2):
        if emb1 is None or emb2 is None: return 0
        return cosine_similarity(np.array(emb1).reshape(1, -1), np.array(emb2).reshape(1, -1))[0][0]
    def clasificar_tono_noticia_gpt(self, resumen_noticia):
        texto_analisis=self.limpiar_texto(resumen_noticia)
        if not texto_analisis: return "Neutro"
        if texto_analisis in self.tonos_asignados_cache: return self.tonos_asignados_cache[texto_analisis]
        prompt=f"""Eres un analista de medios experto... Tu tarea es clasificar el TONO PREDOMINANTE de la siguiente NOTICIA sobre "{self.marca}". Consideraciones: - POSITIVO: logros, resultados sólidos. - NEGATIVO: problemas, controversias. - NEUTRO: mención factual, incidental, en listado. NOTICIA: --- {texto_analisis} --- Responde ÚNICAMENTE con: POSITIVO, NEGATIVO o NEUTRO."""
        try:
            time.sleep(0.1); response=self.client.chat.completions.create(model=self.OPENAI_MODEL_CLASIFICACION, messages=[{"role": "system", "content": "Tu respuesta debe ser una única palabra: POSITIVO, NEGATIVO o NEUTRO."}, {"role": "user", "content": prompt}], max_tokens=5, temperature=0.0)
            tono=response.choices[0].message.content.strip().capitalize()
            if tono not in ["Positivo", "Negativo", "Neutro"]: tono="Neutro"
            self.tonos_asignados_cache[texto_analisis]=tono; return tono
        except Exception as e: st.warning(f"Advertencia (Clasificación Tono): {e}"); return "Neutro (Error)"
    def detectar_grupos_similares(self, df, embeddings_list):
        n=len(df); adj=[[] for _ in range(n)]
        for i in range(n):
            if embeddings_list[i] is None: continue
            for j in range(i + 1, n):
                if embeddings_list[j] is None: continue
                if self.calcular_similitud(embeddings_list[i], embeddings_list[j]) >= self.high_similarity_threshold: adj[i].append(j); adj[j].append(i)
        visitados=[False] * n; grupos=[]
        for i in range(n):
            if not visitados[i] and embeddings_list[i] is not None:
                grupo_actual=[]; q=[i]; visitados[i]=True
                while q:
                    u=q.pop(0); grupo_actual.append(u)
                    for v_idx in adj[u]:
                        if not visitados[v_idx]: visitados[v_idx]=True; q.append(v_idx)
                if grupo_actual: grupos.append(sorted(grupo_actual))
        item_to_rep_map={}
        for grupo in grupos:
            rep_idx=grupo[0]; self.grupos_similares_info[rep_idx]={'tono': None, 'miembros': grupo}
            for miembro_idx in grupo: item_to_rep_map[miembro_idx]=rep_idx
        return item_to_rep_map
    def procesar_dataframe(self, df, columna_texto):
        df.columns = [unidecode(col.strip().lower()).replace(' ', '') for col in df.columns]
        if columna_texto not in df.columns: st.error(f"Error interno: La columna '{columna_texto}' no se encuentra en el dataframe procesado."); return df
        df['Tono']="No Procesado"
        total_rows=len(df)
        status_text = st.empty()
        with st.spinner("Analizando Tono..."):
            embeddings_list=[self.obtener_embedding(row[columna_texto]) for idx, row in df.iterrows()]
            item_to_rep_map=self.detectar_grupos_similares(df, embeddings_list)
            for idx in range(total_rows):
                status_text.text(f"Analizando Tono: Fila {idx+1}/{total_rows}")
                if df.at[idx, 'Tono'] != "No Procesado": continue
                if idx in item_to_rep_map:
                    rep_idx=item_to_rep_map[idx]
                    if self.grupos_similares_info[rep_idx]['tono'] is None:
                        tono_rep=self.clasificar_tono_noticia_gpt(df.iloc[rep_idx][columna_texto])
                        self.grupos_similares_info[rep_idx]['tono']=tono_rep
                    df.at[idx, 'Tono']=self.grupos_similares_info[rep_idx]['tono']
                else: df.at[idx, 'Tono']=self.clasificar_tono_noticia_gpt(df.iloc[idx][columna_texto])
        status_text.empty()
        return df

class ClasificadorTemasAvanzado:
    def __init__(self, client_openai):
        self.client = client_openai; self.embeddings_cache = {}; self.temas_cache = {}
        self.OPENAI_MODEL_EMBEDDING = 'text-embedding-3-small'; self.OPENAI_MODEL_CLASIFICACION = "gpt-4.1-mini-2025-04-14" 
    def limpiar_texto(self, texto):
        if pd.isna(texto): return ""
        return re.sub(r'\s+', ' ', str(texto).strip())
    def extraer_keywords_clave(self, texto):
        if not texto: return []
        stop_words = {'el','la','de','que','y','en','un','es','se','no','te','lo','le','da','su','por','son','con','para','al','del','los','las','una','sobre','todo','también','tras','otro','algún','muy','fue','han','más','hasta','desde','está','entre','cuando','todo','esta','ser','son','tiene','le','me','mi','tu','ti','nos','os','les','si','ya','solo','ni','puede','han'}
        palabras = re.findall(r'\b[a-záéíóúñ]{3,}\b', texto.lower())
        keywords = [p for p in palabras if p not in stop_words and not p.isdigit()]
        counter = Counter(keywords)
        return [palabra for palabra, freq in counter.most_common(10)]
    def obtener_embedding(self, texto):
        texto_limpio_cache = self.limpiar_texto(texto).lower()
        if not texto_limpio_cache: return None
        if texto_limpio_cache in self.embeddings_cache: return self.embeddings_cache[texto_limpio_cache]
        try:
            time.sleep(0.02); response = self.client.embeddings.create(input=[texto_limpio_cache], model=self.OPENAI_MODEL_EMBEDDING)
            embedding = response.data[0].embedding; self.embeddings_cache[texto_limpio_cache] = embedding; return embedding
        except Exception as e: st.warning(f"Advertencia (Tema Embedding): {e}"); return None
    def calcular_similitud_semantica(self, emb1, emb2):
        if emb1 is None or emb2 is None: return 0
        return cosine_similarity(np.array(emb1).reshape(1, -1), np.array(emb2).reshape(1, -1))[0][0]
    def clustering_jerarquico_avanzado(self, df, embeddings_list):
        valid_indices = [i for i, emb in enumerate(embeddings_list) if emb is not None]
        if len(valid_indices) < 2: return {f"individual_{i}": [i] for i in valid_indices}
        n_valid = len(valid_indices); similarity_matrix = np.zeros((n_valid, n_valid))
        for i in range(n_valid):
            for j in range(i+1, n_valid):
                sim = self.calcular_similitud_semantica(embeddings_list[valid_indices[i]], embeddings_list[valid_indices[j]])
                similarity_matrix[i][j] = sim; similarity_matrix[j][i] = sim
            similarity_matrix[i][i] = 1.0
        clusters_finales = {}; processed_indices = set()
        for threshold in [0.85, 0.75]:
            current_clusters = self._encontrar_clusters_por_threshold(similarity_matrix, valid_indices, threshold, processed_indices)
            for cluster_id, indices in current_clusters.items():
                if len(indices) >= 2: clusters_finales[f"cluster_{threshold}_{cluster_id}"] = indices; processed_indices.update(indices)
        for idx in valid_indices:
            if idx not in processed_indices: clusters_finales[f"individual_{idx}"] = [idx]
        return clusters_finales
    def _encontrar_clusters_por_threshold(self, similarity_matrix, valid_indices, threshold, processed_indices):
        clusters = {}; cluster_id = 0; visited = set()
        for i, idx_i in enumerate(valid_indices):
            if idx_i in processed_indices or idx_i in visited: continue
            cluster = []; queue = [i]; visited.add(idx_i)
            while queue:
                current_i = queue.pop(0); current_idx = valid_indices[current_i]; cluster.append(current_idx)
                for j, idx_j in enumerate(valid_indices):
                    if (idx_j not in visited and idx_j not in processed_indices and similarity_matrix[current_i][j] >= threshold):
                        visited.add(idx_j); queue.append(j)
            if len(cluster) >= 2: clusters[cluster_id] = cluster; cluster_id += 1
            else:
                for idx in cluster: visited.discard(idx)
        return clusters
    def generar_tema_inteligente(self, textos_cluster, keywords_cluster):
        contexto_resumido = " | ".join([str(texto)[:200] for texto in textos_cluster[:5]])
        keywords_str = ", ".join(keywords_cluster[:15])
        contexto_hash = str(hash(contexto_resumido + keywords_str))
        if contexto_hash in self.temas_cache: return self.temas_cache[contexto_hash]
        prompt = f"""Eres un experto analista de contenido. Analiza el siguiente grupo de textos similares y genera un TEMA PRINCIPAL muy específico y descriptivo. INSTRUCCIONES: - El tema debe ser conciso (máximo 6 palabras) - Debe capturar la esencia común de todos los textos - Debe ser específico, no genérico - Enfócate en el aspecto más relevante y distintivo. PALABRAS CLAVE IDENTIFICADAS: {keywords_str}. TEXTOS DEL GRUPO: --- {contexto_resumido} ---. Genera únicamente el tema, sin explicaciones adicionales."""
        try:
            time.sleep(0.1)
            response = self.client.chat.completions.create(model=self.OPENAI_MODEL_CLASIFICACION, messages=[{"role": "system", "content": "Eres un experto en síntesis temática que genera temas precisos y específicos."}, {"role": "user", "content": prompt}], max_tokens=40, temperature=0.2)
            tema = response.choices[0].message.content.strip().replace('"', '').replace('Tema:', '').replace('TEMA:', '').strip()
            if not tema or len(tema) < 3: tema = self._generar_tema_fallback(keywords_cluster)
            self.temas_cache[contexto_hash] = tema; return tema
        except Exception as e: st.warning(f"Advertencia (Generación Tema): {e}"); return self._generar_tema_fallback(keywords_cluster)
    def _generar_tema_fallback(self, keywords):
        if not keywords: return "Tema General"
        return " ".join(keywords[:3]).title()
    def generar_tema_individual(self, texto):
        if not texto or pd.isna(texto): return "Sin Contenido"
        texto_hash = str(hash(str(texto)))
        if texto_hash in self.temas_cache: return self.temas_cache[texto_hash]
        keywords = self.extraer_keywords_clave(texto)
        prompt = f"""Analiza el siguiente texto y genera un TEMA específico y descriptivo en máximo 5 palabras. El tema debe ser: - Específico y distintivo - Relacionado con el contenido principal - Conciso pero informativo. TEXTO: --- {str(texto)[:400]} ---. Genera únicamente el tema."""
        try:
            time.sleep(0.1)
            response = self.client.chat.completions.create(model=self.OPENAI_MODEL_CLASIFICACION, messages=[{"role": "system", "content": "Genera temas específicos y concisos."}, {"role": "user", "content": prompt}], max_tokens=30, temperature=0.3)
            tema = response.choices[0].message.content.strip().replace('"', '').replace('Tema:', '').replace('TEMA:', '').strip()
            if not tema or len(tema) < 3: tema = self._generar_tema_fallback(keywords)
            self.temas_cache[texto_hash] = tema; return tema
        except Exception as e: st.warning(f"Advertencia (Tema Individual): {e}"); return self._generar_tema_fallback(keywords)
    def procesar_dataframe(self, df, columna_texto):
        df.columns = [unidecode(col.strip().lower()).replace(' ', '') for col in df.columns]
        if columna_texto not in df.columns: st.error(f"Error interno: La columna '{columna_texto}' no se encuentra en el dataframe procesado."); return df
        df['Tema'] = "No Procesado"
        total_rows = len(df)
        status_text = st.empty()
        with st.spinner("Analizando Temas con IA Avanzada..."):
            status_text.text("Generando representaciones semánticas...")
            embeddings_list = [self.obtener_embedding(row[columna_texto]) if pd.notna(row[columna_texto]) else None for idx, row in df.iterrows()]
            status_text.text("Identificando grupos temáticos...")
            clusters = self.clustering_jerarquico_avanzado(df, embeddings_list)
            status_text.text("Generando temas inteligentes...")
            for cluster_id, indices in clusters.items():
                if len(indices) > 1:
                    textos_cluster = [df.iloc[idx][columna_texto] for idx in indices if pd.notna(df.iloc[idx][columna_texto])]
                    all_keywords = [kw for texto in textos_cluster for kw in self.extraer_keywords_clave(texto)]
                    top_keywords = [kw for kw, freq in Counter(all_keywords).most_common(10)]
                    tema_cluster = self.generar_tema_inteligente(textos_cluster, top_keywords)
                    for idx in indices: df.at[idx, 'Tema'] = tema_cluster
                elif indices: # Es un cluster individual
                    idx = indices[0]
                    texto = df.iloc[idx][columna_texto]
                    df.at[idx, 'Tema'] = self.generar_tema_individual(texto)
            elementos_sin_tema = df[df['Tema'] == "No Procesado"]
            if not elementos_sin_tema.empty:
                for idx in elementos_sin_tema.index:
                    texto = df.iloc[idx][columna_texto]
                    df.at[idx, 'Tema'] = self.generar_tema_individual(texto)
        status_text.empty()
        return df

# --- NUEVA FUNCIÓN PARA GENERAR RESUMEN EJECUTIVO ---
def generar_resumen_estrategico(client, df, marca_analizada):
    datos_contexto = ""
    
    volumen_total = len(df)
    dist_tono = df['Tono_IA'].value_counts().to_string().replace('\n', ', ')
    datos_contexto += f"- Diagnóstico General: {volumen_total} menciones analizadas. Distribución de tono: {dist_tono}.\n"

    top_temas = df['Tema_IA'].value_counts().nlargest(5).to_string().replace('\n', ', ')
    datos_contexto += f"- Temas Principales: {top_temas}.\n"

    # Buscar columnas de KPI de forma robusta
    col_audiencia, col_cpe = None, None
    for col in df.columns:
        if 'audiencia' in col.lower(): col_audiencia = col
        if 'cpe' in col.lower(): col_cpe = col

    if col_audiencia:
        riesgos = df[df['Tono_IA'] == 'Negativo'].nlargest(3, col_audiencia)[['Tema_IA', col_audiencia]].copy()
        if not riesgos.empty:
            riesgos[col_audiencia] = riesgos[col_audiencia].apply(lambda x: f"{x:,.0f}")
            datos_contexto += f"- Riesgos Potenciales (Temas Negativos con Mayor Audiencia):\n{riesgos.to_string(index=False)}\n"
    
    if col_cpe:
        riesgos_cpe = df[df['Tono_IA'] == 'Negativo'].nlargest(3, col_cpe)[['Tema_IA', col_cpe]].copy()
        if not riesgos_cpe.empty:
             riesgos_cpe[col_cpe] = riesgos_cpe[col_cpe].apply(lambda x: f"${x:,.0f}")
             datos_contexto += f"- Riesgos Potenciales (Temas Negativos con Mayor CPE):\n{riesgos_cpe.to_string(index=False)}\n"

    prompt_final = f"""
    A partir de los siguientes datos resumidos sobre la presencia mediática de "{marca_analizada}":
    ---
    {datos_contexto}
    ---
    Analiza los datos y presenta un informe ejecutivo conciso en español.
    **Instrucciones:**
    1. **Diagnóstico General:** Una o dos oraciones que resuman la situación mediática general, considerando el volumen y el tono predominante.
    2. **Puntos Clave:** Enumera 2-3 hallazgos cruciales, conectando los temas más importantes con su tono y alcance (si hay datos de CPE/Audiencia).
    3. **Riesgos Potenciales:** Identifica explícitamente 1-2 riesgos evidentes en los datos, por ejemplo, temas negativos con alta audiencia o costo.
    **NO INCLUYAS RECOMENDACIONES.** Limítate a diagnosticar con un lenguaje claro y directo, en formato de texto plano.
    """
    
    try:
        response = client.chat.completions.create(
            model="gpt-4.1-mini-2025-04-14",
            messages=[
                {"role": "system", "content": "Eres un analista de medios experto en resumir datos para informes ejecutivos concisos y directos."},
                {"role": "user", "content": prompt_final}
            ],
            max_tokens=500, temperature=0.2
        )
        return response.choices[0].message.content
    except Exception as e:
        st.warning(f"No se pudo generar el resumen estratégico: {e}")
        return "El resumen ejecutivo no pudo ser generado debido a un error en la API."

# --- FLUJO PRINCIPAL DE LA APLICACIÓN ---
if check_password():
    st.title("✨ Intelli-Clean: Depurador y Analizador de Noticias")
    st.caption("Una herramienta inteligente para mapear, limpiar, deduplicar y analizar tus informes con precisión.")
    st.divider()

    # --- Sidebar para carga de archivos ---
    with st.sidebar:
        st.header("📂 Paso 1: Carga y Depuración")
        uploaded_main_file = st.file_uploader("1. Informe Principal de Noticias", type="xlsx", key="main_file")
        uploaded_internet_map = st.file_uploader("2. Mapeo de Medios de Internet", type="xlsx", key="internet_map")
        uploaded_region_map = st.file_uploader("3. Mapeo de Regiones", type="xlsx", key="region_map")
        uploaded_empresa_map = st.file_uploader("4. Mapeo de Nombres de Empresas", type="xlsx", key="empresa_map")
        st.divider()
        process_button = st.button("🚀 Analizar y Depurar Archivos", type="primary", use_container_width=True)

    # --- Inicialización del estado de la sesión ---
    if 'deduplication_complete' not in st.session_state:
        st.session_state.deduplication_complete = False
    if 'ai_analysis_complete' not in st.session_state:
        st.session_state.ai_analysis_complete = False

    all_files_uploaded = (uploaded_main_file and uploaded_internet_map and 
                          uploaded_region_map and uploaded_empresa_map)

    # --- Lógica del botón de deduplicación ---
    if process_button:
        st.session_state.deduplication_complete = False
        st.session_state.ai_analysis_complete = False 
        if all_files_uploaded:
            with st.status("Iniciando proceso de deduplicación... ⏳", expanded=True) as status:
                try:
                    status.write("Cargando archivos y creando diccionarios de mapeo...")
                    wb_main = openpyxl.load_workbook(uploaded_main_file)
                    internet_dict = {str(r[0].value).lower().strip(): str(r[1].value) for r in openpyxl.load_workbook(uploaded_internet_map, data_only=True).active.iter_rows(min_row=2) if r[0].value and r[1].value}
                    region_dict = {str(r[0].value).lower().strip(): str(r[1].value) for r in openpyxl.load_workbook(uploaded_region_map, data_only=True).active.iter_rows(min_row=2) if r[0].value and r[1].value}
                    empresa_dict = {str(r[0].value).lower().strip(): str(r[1].value) for r in openpyxl.load_workbook(uploaded_empresa_map, data_only=True).active.iter_rows(min_row=2) if r[0].value and r[1].value}

                    status.write("🧠 Ejecutando deduplicación...")
                    final_wb, nissan_wb, summary = run_deduplication_process(wb_main, empresa_dict, internet_dict, region_dict)
                    
                    status.update(label="✅ ¡Deduplicación completada!", state="complete", expanded=False)

                    st.session_state.summary = summary
                    
                    main_stream = io.BytesIO()
                    final_wb.save(main_stream)
                    st.session_state.main_stream = main_stream

                    # Guardar el DataFrame depurado y el workbook de nissan para el análisis IA
                    main_stream.seek(0)
                    st.session_state.df_depurado = pd.read_excel(main_stream)
                    st.session_state.nissan_wb = nissan_wb
                    
                    st.session_state.deduplication_complete = True

                except Exception as e:
                    status.update(label="❌ Error en el proceso de deduplicación", state="error", expanded=True)
                    st.error(f"Ha ocurrido un error inesperado: {e}")
                    st.exception(e)
                    st.session_state.deduplication_complete = False
        else:
            st.warning("⚠️ Por favor, asegúrate de cargar los cuatro archivos requeridos.")

    # --- Bloque para mostrar resultados de deduplicación y lanzar análisis IA ---
    if st.session_state.deduplication_complete:
        st.header("Resultados de la Deduplicación")
        st.subheader("📊 Resumen del Proceso")
        
        summary = st.session_state.summary
        col1, col2, col3 = st.columns(3)
        col1.metric("Filas Totales Procesadas", summary.get('total_rows', 0))
        col2.metric("👍 Filas para Conservar", summary.get('to_conserve', 0))
        col3.metric("🗑️ Filas para Eliminar", summary.get('to_eliminate', 0))
        
        st.download_button(
            label="1. Descargar Informe Principal Depurado", 
            data=st.session_state.main_stream.getvalue(), 
            file_name=f"Informe_Depurado_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
            use_container_width=True
        )
        st.divider()
        
        # --- SECCIÓN PARA ANÁLISIS IA ---
        st.header("🤖 Paso 2: Análisis con IA (Opcional)")
        st.info("Utiliza el poder de la IA para clasificar Tono, Tema y generar un diagnóstico estratégico.")

        with st.form("ai_analysis_form"):
            try:
                api_key = st.secrets["OPENAI_API_KEY"]
                st.success("API Key cargada desde Secrets.", icon="✅")
            except (KeyError, AttributeError):
                st.warning("No se encontraron Secrets. Ingresa la API Key manualmente.", icon="⚠️")
                api_key = st.text_input("Ingresa tu API Key de OpenAI", type="password")

            marca = st.text_input("➡️ Ingresa la Marca o Cliente para el análisis", placeholder="Ej: Coca-Cola")
            ai_submit_button = st.form_submit_button("🧠 Analizar y Diagnosticar con IA", type="primary", use_container_width=True)

        if ai_submit_button:
            if not api_key or not marca:
                st.error("Es necesario proporcionar la API Key de OpenAI y el nombre de la Marca.")
            else:
                try:
                    nissan_wb_obj = st.session_state.nissan_wb
                    stream = io.BytesIO()
                    nissan_wb_obj.save(stream)
                    stream.seek(0)
                    df_for_ai = pd.read_excel(stream)

                    if df_for_ai.empty:
                        st.warning("No hay datos para analizar con IA (el archivo de resúmenes está vacío).")
                    else:
                        client = openai.OpenAI(api_key=api_key)
                        with st.status("Realizando análisis con IA...", expanded=True) as status:
                            status.write("Clasificando Tono...")
                            clasificador_tono = ClasificadorTonoNoticias(marca, client)
                            df_con_tono = clasificador_tono.procesar_dataframe(df_for_ai.copy(), 'resumen')
                            
                            status.write("Identificando Temas...")
                            clasificador_tema = ClasificadorTemasAvanzado(client)
                            df_analizado_ia = clasificador_tema.procesar_dataframe(df_con_tono.copy(), 'resumen')
                            
                            # Combinar resultados del AI con el dataframe depurado original
                            df_final_completo = st.session_state.df_depurado.copy()
                            if len(df_final_completo) == len(df_analizado_ia):
                                df_final_completo['Tono_IA'] = df_analizado_ia['Tono'].values
                                df_final_completo['Tema_IA'] = df_analizado_ia['Tema'].values
                            else:
                                st.error("Error: el número de filas del análisis no coincide con el informe depurado. No se puede combinar.")
                                raise Exception("Error de coincidencia de filas.")

                            status.write("Generando Resumen Ejecutivo...")
                            resumen_texto = generar_resumen_estrategico(client, df_final_completo, marca)
                            st.session_state.resumen_ejecutivo = resumen_texto

                            status.update(label="¡Análisis IA completado!", state="complete", expanded=False)

                        output_stream = io.BytesIO()
                        with pd.ExcelWriter(output_stream, engine='openpyxl') as writer:
                            df_final_completo.to_excel(writer, index=False, sheet_name='Resultados_IA')
                        st.session_state.ai_analyzed_stream = output_stream
                        st.session_state.ai_analysis_complete = True
                        st.balloons()
                
                except openai.AuthenticationError:
                    st.error("Error de autenticación con OpenAI. Verifica que tu API Key sea correcta.")
                    st.session_state.ai_analysis_complete = False
                except Exception as e:
                    st.error(f"Ocurrió un error inesperado durante el análisis IA: {e}")
                    st.exception(e)
                    st.session_state.ai_analysis_complete = False

    # --- Bloque para mostrar resultados del análisis IA y Diagnóstico ---
    if st.session_state.get('ai_analysis_complete', False):
        st.header("✅ Resultados del Análisis IA")
        
        st.subheader("📝 Resumen Ejecutivo Estratégico")
        st.text_area("Diagnóstico General", value=st.session_state.resumen_ejecutivo, height=250, disabled=True)
        st.markdown("---")

        st.subheader("📥 Descargar Reporte Final")
        st.info("Este archivo contiene los datos depurados con las columnas 'Tono_IA' y 'Tema_IA' añadidas.")
        st.download_button(
            label="2. Descargar Reporte Completo Analizado con IA",
            data=st.session_state.ai_analyzed_stream.getvalue(),
            file_name=f"Reporte_Analizado_IA_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    # --- Mensaje inicial ---
    elif not st.session_state.deduplication_complete and not process_button:
        st.info("Carga los archivos en el menú de la izquierda y haz clic en 'Analizar y Depurar' para comenzar.")
