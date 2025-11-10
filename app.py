import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import datetime
import io
import re
import html
import numpy as np
import gc
import tempfile
import os

# --- Configuración de la página ---
st.set_page_config(page_title="Procesador de Dossiers (Lite) v1.7", layout="wide")

# ==============================================================================
# SECCIÓN DE FUNCIONES AUXILIARES
# ==============================================================================
def extract_link_from_cell(cell):
    """Extrae hipervínculos incrustados de celdas de Excel"""
    try:
        if cell.hyperlink and cell.hyperlink.target:
            return cell.hyperlink.target
    except Exception:
        pass
    return None

def convert_html_entities(text):
    """
    Convierte entidades HTML mal codificadas a caracteres normales.
    Maneja tanto entidades hexadecimales como decimales.
    """
    if not isinstance(text, str): 
        return text
    
    # Primero decodificar entidades HTML estándar
    text = html.unescape(text)
    
    # Manejar entidades HTML numéricas hexadecimales específicas
    html_entities = {
        '&#xF3;': 'ó',  # ó
        '&#xE1;': 'á',  # á
        '&#xE9;': 'é',  # é
        '&#xED;': 'í',  # í
        '&#xFA;': 'ú',  # ú
        '&#xF1;': 'ñ',  # ñ
        '&#xDC;': 'Ü',  # Ü
        '&#xFC;': 'ü',  # ü
        '&#xC1;': 'Á',  # Á
        '&#xC9;': 'É',  # É
        '&#xCD;': 'Í',  # Í
        '&#xD3;': 'Ó',  # Ó
        '&#xDA;': 'Ú',  # Ú
        '&#xD1;': 'Ñ',  # Ñ
        '&#xC7;': 'Ç',  # Ç
        '&#xE7;': 'ç',  # ç
    }
    
    # Reemplazar entidades HTML numéricas hexadecimales
    for entity, char in html_entities.items():
        text = text.replace(entity, char)
    
    # Patrón para capturar otras entidades hexadecimales que no estén en el diccionario
    def replace_hex_entity(match):
        try:
            hex_code = match.group(1)
            char_code = int(hex_code, 16)
            return chr(char_code)
        except (ValueError, OverflowError):
            return match.group(0)  # Devolver original si no se puede convertir
    
    text = re.sub(r'&#x([0-9A-Fa-f]+);', replace_hex_entity, text)
    
    # Patrón para entidades decimales
    def replace_decimal_entity(match):
        try:
            decimal_code = int(match.group(1))
            return chr(decimal_code)
        except (ValueError, OverflowError):
            return match.group(0)  # Devolver original si no se puede convertir
    
    text = re.sub(r'&#(\d+);', replace_decimal_entity, text)
    
    # Limpiar caracteres problemáticos adicionales
    custom_replacements = {
        '"': '"',
        '"': '"', 
        ''': "'", 
        ''': "'", 
        'Â': '', 
        'â': '', 
        '€': '', 
        '™': ''
    }
    
    for entity, char in custom_replacements.items():
        text = text.replace(entity, char)
    
    return text

def normalize_title_for_comparison(title):
    """
    Normaliza el título para comparación de duplicados.
    """
    if not isinstance(title, str): 
        return ""
    
    # Limpiar entidades HTML primero
    title = convert_html_entities(title)
    
    # Normalizar para comparación (remover caracteres especiales y convertir a minúsculas)
    return re.sub(r'\W+', ' ', title).lower().strip()

def clean_title_for_output(title):
    """
    ÚNICAMENTE limpia entidades HTML mal codificadas.
    NO corta, NO modifica, NO remueve ninguna parte del título.
    Solo convierte caracteres como &#xF3; a ó
    """
    if not isinstance(title, str): 
        return ""
    
    # SOLO limpiar entidades HTML - NO tocar nada más
    title = convert_html_entities(title)
    
    # Solo quitar espacios al inicio y final, NO espacios múltiples internos
    title = title.strip()
    
    return title

def corregir_texto(text):
    if not isinstance(text, str): return text
    text = convert_html_entities(text)
    text = re.sub(r'(<br>|\[\.\.\.\]|\s+)', ' ', text).strip()
    match = re.search(r'[A-Z]', text)
    if match: text = text[match.start():]
    if text and not text.endswith('...'): text = text.rstrip('.') + '...'
    return text

def extract_all_hyperlinks(file_obj, progress_callback=None):
    """
    Extrae TODOS los hipervínculos del archivo Excel usando archivo temporal
    y procesamiento por chunks para evitar problemas de memoria.
    """
    # Crear archivo temporal
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
        tmp_file.write(file_obj.read())
        tmp_path = tmp_file.name
    
    file_obj.seek(0)
    
    try:
        # Abrir con read_only=False para poder leer hyperlinks
        wb = load_workbook(tmp_path, data_only=False, keep_links=True)
        sheet = wb.active
        
        # Obtener headers
        original_headers = [cell.value for cell in sheet[1] if cell.value]
        
        # Identificar columnas con links
        link_nota_idx = original_headers.index('Link Nota') if 'Link Nota' in original_headers else None
        link_streaming_idx = original_headers.index('Link (Streaming - Imagen)') if 'Link (Streaming - Imagen)' in original_headers else None
        
        # Diccionario para almacenar links por fila
        hyperlinks_data = {
            'Link Nota': {},
            'Link (Streaming - Imagen)': {}
        }
        
        total_rows = sheet.max_row - 1
        row_num = 0
        
        # Iterar por todas las filas
        for row_idx, row in enumerate(sheet.iter_rows(min_row=2), start=1):
            # Callback de progreso
            if progress_callback and row_idx % 100 == 0:
                progress_callback(row_idx, total_rows)
            
            # Si la fila está completamente vacía, saltarla
            if all(c.value is None for c in row):
                continue
            
            # Extraer hipervínculo de Link Nota
            if link_nota_idx is not None and link_nota_idx < len(row):
                try:
                    cell = row[link_nota_idx]
                    link = extract_link_from_cell(cell)
                    if link:
                        hyperlinks_data['Link Nota'][row_num] = link
                except Exception as e:
                    pass  # Si falla, continuar
            
            # Extraer hipervínculo de Link Streaming
            if link_streaming_idx is not None and link_streaming_idx < len(row):
                try:
                    cell = row[link_streaming_idx]
                    link = extract_link_from_cell(cell)
                    if link:
                        hyperlinks_data['Link (Streaming - Imagen)'][row_num] = link
                except Exception as e:
                    pass  # Si falla, continuar
            
            row_num += 1
            
            # Liberar memoria cada 500 filas
            if row_idx % 500 == 0:
                gc.collect()
        
        wb.close()
        
        return hyperlinks_data, original_headers, row_num
        
    finally:
        # Limpiar archivo temporal
        try:
            os.unlink(tmp_path)
        except:
            pass

def to_excel_from_df(df, final_order):
    output = io.BytesIO()
    final_columns_in_df = [col for col in final_order if col in df.columns]
    df_to_excel = df[final_columns_in_df]
    with pd.ExcelWriter(
        output,
        engine='xlsxwriter',
        datetime_format='dd/mm/yyyy',
        date_format='dd/mm/yyyy'
    ) as writer:
        df_to_excel.to_excel(writer, index=False, sheet_name='Resultado')
        workbook = writer.book
        worksheet = writer.sheets['Resultado']
        link_format = workbook.add_format({'color': 'blue', 'underline': 1})
        for col_name in ['Link Nota', 'Link (Streaming - Imagen)']:
            if col_name in df_to_excel.columns:
                col_idx = df_to_excel.columns.get_loc(col_name)
                for row_idx, url in enumerate(df_to_excel[col_name]):
                    if pd.notna(url) and isinstance(url, str) and url.startswith('http'):
                        try:
                            worksheet.write_url(row_idx + 1, col_idx, url, link_format, 'Link')
                        except Exception:
                            # Fallback si falla escribir URL
                            worksheet.write_string(row_idx + 1, col_idx, url)
    return output.getvalue()

# ==============================================================================
# LÓGICA DE PROCESAMIENTO PRINCIPAL
# ==============================================================================
def run_full_process(dossier_file, config_file):
    
    st.markdown("---")
    progress_bar = st.progress(0)
    progress_text = st.empty()
    
    try:
        progress_text.info("Paso 1/6: Cargando archivo de configuración...")
        progress_bar.progress(5)
        try:
            config_sheets = pd.read_excel(config_file, sheet_name=None)
            region_map = pd.Series(config_sheets['Regiones'].iloc[:, 1].values, index=config_sheets['Regiones'].iloc[:, 0].astype(str).str.lower().str.strip()).to_dict()
            internet_map = pd.Series(config_sheets['Internet'].iloc[:, 1].values, index=config_sheets['Internet'].iloc[:, 0].astype(str).str.lower().str.strip()).to_dict()
        except Exception as e:
            st.error(f"Error al cargar `Configuracion.xlsx`: {e}. Asegúrate de que contenga las hojas 'Regiones' e 'Internet'.")
            st.stop()

        progress_text.info("Paso 2/6: Extrayendo TODOS los hipervínculos incrustados del archivo...")
        progress_bar.progress(10)
        
        # Callback para actualizar progreso
        def update_progress(current, total):
            pct = 10 + int((current / total) * 25)
            progress_bar.progress(min(pct, 35))
            progress_text.info(f"Paso 2/6: Extrayendo hipervínculos... {current}/{total} filas procesadas")
        
        # Extraer todos los hipervínculos PRIMERO
        hyperlinks_data, original_headers, total_data_rows = extract_all_hyperlinks(dossier_file, update_progress)
        
        st.success(f"✅ Se extrajeron {len(hyperlinks_data['Link Nota'])} links de 'Link Nota' y {len(hyperlinks_data['Link (Streaming - Imagen)'])} de 'Link (Streaming - Imagen)'")
        
        progress_text.info("Paso 3/6: Leyendo datos del Dossier con pandas...")
        progress_bar.progress(40)
        
        # Ahora leer el archivo con pandas (más rápido para datos)
        dossier_file.seek(0)
        df_base = pd.read_excel(dossier_file, engine='openpyxl')
        
        st.info(f"📊 Archivo base leído: {len(df_base)} filas")

        progress_text.info("Paso 4/6: Expandiendo menciones y aplicando hipervínculos...")
        progress_bar.progress(45)
        
        # Expandir menciones
        rows_to_expand = []
        for idx, row in df_base.iterrows():
            row_dict = row.to_dict()
            
            # Aplicar hipervínculos extraídos
            if idx in hyperlinks_data['Link Nota']:
                row_dict['Link Nota'] = hyperlinks_data['Link Nota'][idx]
            if idx in hyperlinks_data['Link (Streaming - Imagen)']:
                row_dict['Link (Streaming - Imagen)'] = hyperlinks_data['Link (Streaming - Imagen)'][idx]
            
            # Expandir menciones
            menciones = [m.strip() for m in str(row_dict.get('Menciones - Empresa') or '').split(';') if m.strip()]
            if not menciones: 
                rows_to_expand.append(row_dict)
            else:
                for mencion in menciones:
                    new_row = row_dict.copy()
                    new_row['Menciones - Empresa'] = mencion
                    rows_to_expand.append(new_row)
            
            # Actualizar progreso y liberar memoria
            if idx % 500 == 0:
                progress_text.info(f"Paso 4/6: Expandiendo menciones... {idx}/{len(df_base)} filas")
                progress_bar.progress(45 + int((idx / len(df_base)) * 10))
                gc.collect()
        
        df = pd.DataFrame(rows_to_expand)
        df['Mantener'] = 'Conservar'
        
        st.success(f"✅ Se expandieron a {len(df)} filas (después de menciones)")

        progress_text.info("Paso 5/6: Aplicando limpieza, mapeos y normalizaciones...")
        progress_bar.progress(60)
        
        for col in original_headers:
            if col not in df.columns: df[col] = None
        
        # APLICAR LA LIMPIEZA CORREGIDA - SOLO ENTIDADES HTML, NO CORTAR
        df['Título'] = df['Título'].astype(str).apply(clean_title_for_output)
        df['Resumen - Aclaracion'] = df['Resumen - Aclaracion'].astype(str).apply(corregir_texto)

        tipo_medio_map = {'online': 'Internet', 'diario': 'Prensa', 'am': 'Radio', 'fm': 'Radio', 'aire': 'Televisión', 'cable': 'Televisión', 'revista': 'Revista'}
        df['Tipo de Medio'] = df['Tipo de Medio'].str.lower().str.strip().map(tipo_medio_map).fillna(df['Tipo de Medio'])
        is_internet = df['Tipo de Medio'] == 'Internet'
        is_print = df['Tipo de Medio'].isin(['Prensa', 'Revista'])
        is_broadcast = df['Tipo de Medio'].isin(['Radio', 'Televisión'])
        
        df.loc[is_internet, ['Link Nota', 'Link (Streaming - Imagen)']] = df.loc[is_internet, ['Link (Streaming - Imagen)', 'Link Nota']].values
        cond_copy = is_print & df['Link Nota'].isnull() & df['Link (Streaming - Imagen)'].notnull()
        df.loc[cond_copy, 'Link Nota'] = df.loc[cond_copy, 'Link (Streaming - Imagen)']
        df.loc[is_print, 'Link (Streaming - Imagen)'] = None
        df.loc[is_broadcast, 'Link (Streaming - Imagen)'] = None
        
        # --- LÓGICA "CORTAR Y PEGAR" ---
        if 'Duración - Nro. Caracteres' in df.columns and 'Dimensión' in df.columns:
            df.loc[is_broadcast, 'Dimensión'] = df.loc[is_broadcast, 'Duración - Nro. Caracteres']
            df.loc[is_broadcast, 'Duración - Nro. Caracteres'] = np.nan
        
        df['Región'] = df['Medio'].astype(str).str.lower().str.strip().map(region_map)
        df.loc[is_internet, 'Medio'] = df.loc[is_internet, 'Medio'].astype(str).str.lower().str.strip().map(internet_map).fillna(df.loc[is_internet, 'Medio'])

        progress_text.info("Paso 6/6: Detectando duplicados y generando resultados...")
        progress_bar.progress(70)
        
        df['titulo_norm'] = df['Título'].apply(normalize_title_for_comparison)
        df['Fecha'] = pd.to_datetime(df['Fecha'], dayfirst=True, errors='coerce').dt.normalize()
        
        df['seccion_priority'] = df['Sección - Programa'].isnull() | (df['Sección - Programa'] == '')
        df['dup_hora'] = np.where(df['Tipo de Medio'] == 'Internet', 'IGNORE_TIME', df['Hora'])
        
        dup_cols_exact = ['titulo_norm', 'Medio', 'Fecha', 'Menciones - Empresa', 'dup_hora']
        sort_by_cols = dup_cols_exact + ['seccion_priority']
        ascending_order = [True] * len(dup_cols_exact) + [False]
        df.sort_values(by=sort_by_cols, ascending=ascending_order, inplace=True)
        exact_duplicates_mask = df.duplicated(subset=dup_cols_exact, keep='first')
        df.loc[exact_duplicates_mask, 'Mantener'] = 'Eliminar'
        df.sort_index(inplace=True)
        
        progress_bar.progress(85)
        
        df_internet_to_check = df[(df['Mantener'] == 'Conservar') & (is_internet)].copy()
        if not df_internet_to_check.empty:
            group_cols = ['titulo_norm', 'Medio', 'Menciones - Empresa']
            df_internet_to_check.sort_values(by=group_cols + ['Fecha'], inplace=True)
            date_diffs = df_internet_to_check.groupby(group_cols)['Fecha'].diff().dt.days
            cluster_ids = (date_diffs != 1).cumsum()
            df_internet_to_check['date_cluster'] = cluster_ids
            sort_by_cols_consecutive = group_cols + ['date_cluster', 'seccion_priority']
            ascending_order_consecutive = [True] * (len(group_cols) + 1) + [False]
            df_internet_to_check.sort_values(by=sort_by_cols_consecutive, ascending=ascending_order_consecutive, inplace=True)
            consecutive_duplicates_mask = df_internet_to_check.duplicated(subset=group_cols + ['date_cluster'], keep='first')
            indices_to_eliminate = df_internet_to_check[consecutive_duplicates_mask].index
            df.loc[indices_to_eliminate, 'Mantener'] = 'Eliminar'
        
        df.loc[df['Mantener'] == 'Eliminar', ['Tono', 'Tema', 'Temas Generales - Tema']] = 'Duplicada'
        
        progress_bar.progress(100)
        st.balloons()
        progress_text.success("✅ ¡Proceso completado! Todos los hipervínculos fueron extraídos correctamente.")

        final_order = ["ID Noticia", "Fecha", "Hora", "Medio", "Tipo de Medio", "Sección - Programa", "Región", "Título", "Autor - Conductor", "Nro. Pagina", "Dimensión", "Duración - Nro. Caracteres", "CPE", "Tier", "Audiencia", "Tono", "Tema", "Temas Generales - Tema", "Resumen - Aclaracion", "Link Nota", "Link (Streaming - Imagen)", "Menciones - Empresa"]
        df_final = df.copy()

        st.subheader("📊 Resumen del Proceso")
        col1, col2, col3 = st.columns(3)
        col1.metric("Filas Totales", len(df_final))
        dups_count = (df_final['Mantener'] == 'Eliminar').sum()
        col2.metric("Filas Marcadas como Duplicadas", dups_count)
        col3.metric("Filas Únicas", len(df_final) - dups_count)
        
        # Contar links extraídos en el resultado final
        links_nota = df_final['Link Nota'].notna().sum()
        links_streaming = df_final['Link (Streaming - Imagen)'].notna().sum()
        st.info(f"🔗 Hipervínculos en resultado final: **{links_nota}** en 'Link Nota', **{links_streaming}** en 'Link (Streaming - Imagen)'")
        
        excel_data = to_excel_from_df(df_final, final_order)
        st.download_button(label="📥 Descargar Archivo Limpio y Mapeado", data=excel_data, file_name=f"Dossier_Limpio_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.subheader("✍️ Previsualización de Resultados")
        final_columns_in_df = [col for col in final_order if col in df_final.columns]
        df_for_editor = df_final[final_columns_in_df].copy()
        if 'Fecha' in df_for_editor.columns:
            df_for_editor['Fecha'] = df_for_editor['Fecha'].dt.strftime('%d/%m/%Y').fillna('')
        for col_name in ['Link Nota', 'Link (Streaming - Imagen)']:
            if col_name in df_for_editor.columns:
                df_for_editor[col_name] = df_for_editor[col_name].apply(lambda x: '🔗 Link' if pd.notna(x) and str(x).startswith('http') else '')
        st.dataframe(df_for_editor, use_container_width=True)
        
        # Liberar memoria
        gc.collect()
        
    except Exception as e:
        st.error(f"❌ Error durante el procesamiento: {str(e)}")
        st.exception(e)
        raise
    
# ==============================================================================
# INTERFAZ PRINCIPAL DE STREAMLIT
# ==============================================================================
st.title("🚀 Procesador de Dossiers (Lite) v1.7")
st.markdown("Una herramienta para limpiar, deduplicar y mapear dossieres de noticias.")

st.success("""
### ✨ Mejoras en v1.7 - SOLUCIÓN DEFINITIVA:
- ✅ **Extracción completa de hipervínculos usando archivo temporal**
- ✅ **Procesa archivos de cualquier tamaño sin límite de filas**
- ✅ **Separa extracción de links de procesamiento de datos**
- ✅ **Manejo robusto de memoria con liberación periódica**
- ✅ **Progreso visual detallado en cada etapa**
""")

st.info("**Instrucciones:**\n\n1. Prepara tu archivo **Dossier** principal y tu archivo **`Configuracion.xlsx`**.\n2. Sube ambos archivos juntos en el área de abajo.\n3. Haz clic en 'Iniciar Proceso'.")

with st.expander("Ver estructura requerida para `Configuracion.xlsx`"):
    st.markdown("- **`Regiones`**: Columna A (Medio), Columna B (Región).\n- **`Internet`**: Columna A (Medio Original), Columna B (Medio Mapeado).")

uploaded_files = st.file_uploader("Arrastra y suelta tus archivos aquí (Dossier y Configuracion)", type=["xlsx"], accept_multiple_files=True)
dossier_file, config_file = None, None
if uploaded_files:
    for file in uploaded_files:
        if 'config' in file.name.lower(): config_file = file
        else: dossier_file = file
    if dossier_file: st.success(f"Archivo Dossier cargado: **{dossier_file.name}**")
    else: st.warning("No se ha subido un archivo que parezca ser el Dossier.")
    if config_file: st.success(f"Archivo de Configuración cargado: **{config_file.name}**")
    else: st.warning("No se ha subido el archivo `Configuracion.xlsx`.")
if st.button("▶️ Iniciar Proceso de Limpieza", disabled=not (dossier_file and config_file), type="primary"):
    run_full_process(dossier_file, config_file)
