import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import datetime
import io
import re
import html
import numpy as np

# --- Configuración de la página ---
st.set_page_config(page_title="Procesador de Dossiers (Eficiente) v1.9", layout="wide")

# ==============================================================================
# SECCIÓN DE FUNCIONES AUXILIARES (Sin cambios)
# ==============================================================================
def extract_link_from_cell(cell):
    if cell.hyperlink and cell.hyperlink.target:
        return cell.hyperlink.target
    return cell.value

def convert_html_entities(text):
    if not isinstance(text, str): return text
    text = html.unescape(text)
    html_entities = {
        '&#xF3;': 'ó', '&#xE1;': 'á', '&#xE9;': 'é', '&#xED;': 'í', '&#xFA;': 'ú', '&#xF1;': 'ñ',
        '&#xDC;': 'Ü', '&#xFC;': 'ü', '&#xC1;': 'Á', '&#xC9;': 'É', '&#xCD;': 'Í', '&#xD3;': 'Ó',
        '&#xDA;': 'Ú', '&#xD1;': 'Ñ', '&#xC7;': 'Ç', '&#xE7;': 'ç',
    }
    for entity, char in html_entities.items(): text = text.replace(entity, char)
    def replace_hex(match):
        try: return chr(int(match.group(1), 16))
        except (ValueError, OverflowError): return match.group(0)
    text = re.sub(r'&#x([0-9A-Fa-f]+);', replace_hex, text)
    def replace_dec(match):
        try: return chr(int(match.group(1)))
        except (ValueError, OverflowError): return match.group(0)
    text = re.sub(r'&#(\d+);', replace_dec, text)
    custom_replacements = {'"': '"', "''": "'", 'Â': '', 'â': '', '€': '', '™': ''}
    for entity, char in custom_replacements.items(): text = text.replace(entity, char)
    return text

def normalize_title_for_comparison(title):
    if not isinstance(title, str): return ""
    return re.sub(r'\W+', ' ', convert_html_entities(title)).lower().strip()

def clean_title_for_output(title):
    if not isinstance(title, str): return ""
    return convert_html_entities(title).strip()

def corregir_texto(text):
    if not isinstance(text, str): return text
    text = convert_html_entities(text)
    text = re.sub(r'(<br>|\[\.\.\.\]|\s+)', ' ', text).strip()
    match = re.search(r'[A-Z]', text)
    if match: text = text[match.start():]
    if text and not text.endswith('...'): text = text.rstrip('.') + '...'
    return text

def to_excel_from_df(df, final_order):
    output = io.BytesIO()
    final_columns_in_df = [col for col in final_order if col in df.columns]
    df_to_excel = df[final_columns_in_df]
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_to_excel.to_excel(writer, index=False, sheet_name='Resultado')
        workbook = writer.book
        worksheet = writer.sheets['Resultado']
        link_format = workbook.add_format({'color': 'blue', 'underline': 1})
        link_columns = ['Link Nota', 'Link (Streaming - Imagen)']
        for col_name in link_columns:
            if col_name in df_to_excel.columns:
                col_idx = df_to_excel.columns.get_loc(col_name)
                for row_idx, url in enumerate(df_to_excel[col_name]):
                    if pd.notna(url) and isinstance(url, str) and url.startswith('http'):
                        safe_url = url.replace('"', '""')
                        formula = f'=HYPERLINK("{safe_url}", "Link")'
                        worksheet.write_formula(row_idx + 1, col_idx, formula, link_format)
    return output.getvalue()

# ==============================================================================
# LÓGICA DE PROCESAMIENTO PRINCIPAL (MODIFICADA PARA EFICIENCIA Y DUPLICADOS)
# ==============================================================================
def run_full_process(dossier_file, config_file):
    
    st.markdown("---")
    progress_text = st.empty()
    
    progress_text.info("Paso 1/4: Cargando archivo de configuración...")
    try:
        config_sheets = pd.read_excel(config_file, sheet_name=None)
        region_map = pd.Series(config_sheets['Regiones'].iloc[:, 1].values, index=config_sheets['Regiones'].iloc[:, 0].astype(str).str.lower().str.strip()).to_dict()
        internet_map = pd.Series(config_sheets['Internet'].iloc[:, 1].values, index=config_sheets['Internet'].iloc[:, 0].astype(str).str.lower().str.strip()).to_dict()
    except Exception as e:
        st.error(f"Error al cargar `Configuracion.xlsx`: {e}. Asegúrate de que contenga las hojas 'Regiones' e 'Internet'.")
        st.stop()

    # --- CAMBIO: LECTURA RÁPIDA CON PYARROW Y EXTRACCIÓN HÍBRIDA DE LINKS ---
    progress_text.info("Paso 2/4: Leyendo datos del Dossier (método rápido con PyArrow)...")
    try:
        df = pd.read_excel(dossier_file, engine='pyarrow')
        original_headers = df.columns.tolist()
        
        # Extracción de hipervínculos por separado
        wb = load_workbook(dossier_file)
        sheet = wb.active
        
        links = {'Link Nota': {}, 'Link (Streaming - Imagen)': {}}
        link_cols_indices = {h: i for i, h in enumerate(original_headers) if h in links}

        for row_idx, row in enumerate(sheet.iter_rows(min_row=2, max_row=df.shape[0] + 1)):
            for col_name, col_idx in link_cols_indices.items():
                link = extract_link_from_cell(row[col_idx])
                if link:
                    links[col_name][row_idx] = link
        
        # Mapear los links extraídos al DataFrame
        df['Link Nota'] = df.index.map(links['Link Nota']).fillna(df['Link Nota'])
        df['Link (Streaming - Imagen)'] = df.index.map(links['Link (Streaming - Imagen)']).fillna(df['Link (Streaming - Imagen)'])

    except Exception as e:
        st.error(f"Error al leer el Dossier con PyArrow: {e}. Asegúrate de que `pyarrow` esté instalado (`pip install pyarrow`).")
        st.stop()

    # Expansión de filas por "Menciones - Empresa"
    df['Menciones - Empresa'] = df['Menciones - Empresa'].astype(str).str.split(';')
    df = df.explode('Menciones - Empresa')
    df['Menciones - Empresa'] = df['Menciones - Empresa'].str.strip()
    df.reset_index(drop=True, inplace=True)
    
    # --- FIN DEL CAMBIO ---
    
    df['Mantener'] = 'Conservar'

    progress_text.info("Paso 3/4: Aplicando limpieza, mapeos y normalizaciones...")
    for col in original_headers:
        if col not in df.columns: df[col] = np.nan
    
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
    df.loc[is_print, 'Link (Streaming - Imagen)'] = np.nan
    df.loc[is_broadcast, 'Link (Streaming - Imagen)'] = np.nan
    
    if 'Duración - Nro. Caracteres' in df.columns and 'Dimensión' in df.columns:
        df.loc[is_broadcast, 'Dimensión'] = df.loc[is_broadcast, 'Duración - Nro. Caracteres']
        df.loc[is_broadcast, 'Duración - Nro. Caracteres'] = np.nan
    
    df['Región'] = df['Medio'].astype(str).str.lower().str.strip().map(region_map)
    df.loc[is_internet, 'Medio'] = df.loc[is_internet, 'Medio'].astype(str).str.lower().str.strip().map(internet_map).fillna(df.loc[is_internet, 'Medio'])

    progress_text.info("Paso 4/4: Detectando duplicados y generando resultados...")
    df['titulo_norm'] = df['Título'].apply(normalize_title_for_comparison)
    df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce').dt.normalize()
    
    # --- CAMBIO: LÓGICA DE DUPLICADOS MEJORADA ---
    df['seccion_priority'] = df['Sección - Programa'].isnull() | (df['Sección - Programa'] == '')
    df['dup_hora'] = np.where(df['Tipo de Medio'] == 'Internet', 'IGNORE_TIME', df['Hora'])
    
    dup_cols_exact = ['titulo_norm', 'Medio', 'Fecha', 'Menciones - Empresa', 'dup_hora']
    sort_by_cols = dup_cols_exact + ['seccion_priority']
    df.sort_values(by=sort_by_cols, ascending=True, na_position='last', inplace=True)
    
    # Identificar duplicados exactos y guardar el ID del original
    original_id = df.groupby(dup_cols_exact)['ID Noticia'].transform('first')
    exact_duplicates_mask = df.duplicated(subset=dup_cols_exact, keep='first')
    df.loc[exact_duplicates_mask, 'Mantener'] = "Duplicado del ID: " + original_id[exact_duplicates_mask].astype(str)

    # Lógica para duplicados consecutivos en Internet
    df_internet_to_check = df[(df['Mantener'] == 'Conservar') & (is_internet)].copy()
    if not df_internet_to_check.empty:
        group_cols = ['titulo_norm', 'Medio', 'Menciones - Empresa']
        df_internet_to_check.sort_values(by=group_cols + ['Fecha'], inplace=True)
        date_diffs = df_internet_to_check.groupby(group_cols)['Fecha'].diff().dt.days.fillna(0)
        cluster_ids = (date_diffs != 1).cumsum()
        df_internet_to_check['date_cluster'] = cluster_ids
        
        # Identificar duplicados y guardar el ID del original en el subconjunto
        original_id_consecutive = df_internet_to_check.groupby(group_cols + ['date_cluster'])['ID Noticia'].transform('first')
        consecutive_duplicates_mask = df_internet_to_check.duplicated(subset=group_cols + ['date_cluster'], keep='first')
        
        indices_to_update = df_internet_to_check[consecutive_duplicates_mask].index
        ids_to_reference = original_id_consecutive[consecutive_duplicates_mask]
        
        # Actualizar el DataFrame principal
        df.loc[indices_to_update, 'Mantener'] = "Duplicado del ID: " + ids_to_reference.astype(str)

    # Restablecer el orden original para que el archivo de salida sea coherente
    df.sort_index(inplace=True)
    df.loc[df['Mantener'].str.startswith('Duplicado'), ['Tono', 'Tema', 'Temas Generales - Tema']] = 'Duplicada'
    # --- FIN DEL CAMBIO ---

    st.balloons()
    progress_text.success("¡Proceso de limpieza completado!")

    final_order = ["ID Noticia", "Fecha", "Hora", "Medio", "Tipo de Medio", "Sección - Programa", "Región", "Título", "Autor - Conductor", "Nro. Pagina", "Dimensión", "Duración - Nro. Caracteres", "CPE", "Tier", "Audiencia", "Tono", "Tema", "Temas Generales - Tema", "Resumen - Aclaracion", "Link Nota", "Link (Streaming - Imagen)", "Menciones - Empresa", "Mantener"]
    
    st.subheader("📊 Resumen del Proceso")
    col1, col2, col3 = st.columns(3)
    col1.metric("Filas Totales", len(df))
    dups_count = df['Mantener'].str.startswith('Duplicado').sum()
    col2.metric("Filas Marcadas como Duplicadas", dups_count)
    col3.metric("Filas Únicas", len(df) - dups_count)
    
    excel_data = to_excel_from_df(df, final_order)
    st.download_button(label="📥 Descargar Archivo Limpio y Mapeado", data=excel_data, file_name=f"Dossier_Limpio_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ==============================================================================
# INTERFAZ PRINCIPAL DE STREAMLIT
# ==============================================================================
st.title("🚀 Procesador de Dossiers (Eficiente) v1.9")
st.markdown("Una herramienta para limpiar, deduplicar y mapear dossieres de noticias.")
st.info("**Instrucciones:**\n\n1. Prepara tu archivo **Dossier** y tu archivo **`Configuracion.xlsx`**.\n2. Sube ambos archivos juntos.\n3. Haz clic en 'Iniciar Proceso'.")

st.success("✅ **NUEVO (v1.9)**: ¡Lectura de archivos más rápida con PyArrow! La columna 'Mantener' ahora especifica el ID de la noticia original en los duplicados.")
st.warning("⚠️ **Nota**: La previsualización de datos ha sido eliminada para maximizar la velocidad.")

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
