import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import datetime
import io
import re
import html
import numpy as np

# --- Configuración de la página ---
st.set_page_config(page_title="Procesador de Dossiers (Lite) v1.3", layout="wide")

# ==============================================================================
# SECCIÓN DE FUNCIONES AUXILIARES
# ==============================================================================
def extract_link_from_cell(cell):
    if cell.hyperlink and cell.hyperlink.target:
        return cell.hyperlink.target
    return None

def convert_html_entities(text):
    if not isinstance(text, str): return text
    text = html.unescape(text)
    custom_replacements = { '“': '\"', '”': '\"', '‘': "'", '’': "'", 'Â': '', 'â': '', '€': '', '™': '' }
    for entity, char in custom_replacements.items():
        text = text.replace(entity, char)
    return text

def normalize_title_for_comparison(title):
    if not isinstance(title, str): return ""
    title = convert_html_entities(title)
    return re.sub(r'\W+', ' ', title).lower().strip()

def clean_title_for_output(title):
    if not isinstance(title, str): return ""
    title = convert_html_entities(title)
    title = re.sub(r'\s*[|-]\s*[\w\s]+$', '', title).strip()
    return title

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
                        worksheet.write_url(row_idx + 1, col_idx, url, link_format, 'Link')
    return output.getvalue()

# ==============================================================================
# LÓGICA DE PROCESAMIENTO PRINCIPAL
# ==============================================================================
def run_full_process(dossier_file, config_file):
    
    st.markdown("---")
    progress_text = st.empty()
    
    progress_text.info("Paso 1/4: Cargando archivo de configuración...")
    try:
        config_sheets = pd.read_excel(config_file.read(), sheet_name=None)
        region_map = pd.Series(config_sheets['Regiones'].iloc[:, 1].values, index=config_sheets['Regiones'].iloc[:, 0].astype(str).str.lower().str.strip()).to_dict()
        internet_map = pd.Series(config_sheets['Internet'].iloc[:, 1].values, index=config_sheets['Internet'].iloc[:, 0].astype(str).str.lower().str.strip()).to_dict()
    except Exception as e:
        st.error(f"Error al cargar `Configuracion.xlsx`: {e}. Asegúrate de que contenga las hojas 'Regiones' e 'Internet'.")
        st.stop()

    progress_text.info("Paso 2/4: Leyendo Dossier y expandiendo filas...")
    wb = load_workbook(dossier_file)
    sheet = wb.active
    original_headers = [cell.value for cell in sheet[1] if cell.value]
    rows_to_expand = []
    for row in sheet.iter_rows(min_row=2):
        if all(c.value is None for c in row): continue
        row_values = [c.value for c in row]
        row_data = dict(zip(original_headers, row_values))
        if 'Link Nota' in original_headers: row_data['Link Nota'] = extract_link_from_cell(row[original_headers.index('Link Nota')])
        if 'Link (Streaming - Imagen)' in original_headers: row_data['Link (Streaming - Imagen)'] = extract_link_from_cell(row[original_headers.index('Link (Streaming - Imagen)')])
        menciones = [m.strip() for m in str(row_data.get('Menciones - Empresa') or '').split(';') if m.strip()]
        if not menciones: rows_to_expand.append(row_data)
        else:
            for mencion in menciones:
                new_row = row_data.copy()
                new_row['Menciones - Empresa'] = mencion
                rows_to_expand.append(new_row)
    df = pd.DataFrame(rows_to_expand)
    df['Mantener'] = 'Conservar'

    progress_text.info("Paso 3/4: Aplicando limpieza, mapeos y normalizaciones...")
    for col in original_headers:
        if col not in df.columns: df[col] = None
    
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
    df['Región'] = df['Medio'].astype(str).str.lower().str.strip().map(region_map)
    df.loc[is_internet, 'Medio'] = df.loc[is_internet, 'Medio'].astype(str).str.lower().str.strip().map(internet_map).fillna(df.loc[is_internet, 'Medio'])

    progress_text.info("Paso 4/4: Detectando duplicados y generando resultados...")
    df['titulo_norm'] = df['Título'].apply(normalize_title_for_comparison)
    # --- MANEJO DE FECHA DEFINITIVO ---
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
    
    st.balloons()
    progress_text.success("¡Proceso de limpieza completado!")

    final_order = ["ID Noticia", "Fecha", "Hora", "Medio", "Tipo de Medio", "Sección - Programa", "Región", "Título", "Autor - Conductor", "Nro. Pagina", "Dimensión", "Duración - Nro. Caracteres", "CPE", "Tier", "Audiencia", "Tono", "Tema", "Temas Generales - Tema", "Resumen - Aclaracion", "Link Nota", "Link (Streaming - Imagen)", "Menciones - Empresa"]
    df_final = df.copy()

    st.subheader("📊 Resumen del Proceso")
    col1, col2, col3 = st.columns(3)
    col1.metric("Filas Totales", len(df_final))
    dups_count = (df_final['Mantener'] == 'Eliminar').sum()
    col2.metric("Filas Marcadas como Duplicadas", dups_count)
    col3.metric("Filas Únicas", len(df_final) - dups_count)
    
    excel_data = to_excel_from_df(df_final, final_order)
    st.download_button(label="📥 Descargar Archivo Limpio y Mapeado", data=excel_data, file_name=f"Dossier_Limpio_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx", mime="application/vnd.openxmlformats-officedocument.sheet")

    st.subheader("✍️ Previsualización de Resultados")
    final_columns_in_df = [col for col in final_order if col in df_final.columns]
    df_for_editor = df_final[final_columns_in_df].copy()
    if 'Fecha' in df_for_editor.columns:
        df_for_editor['Fecha'] = df_for_editor['Fecha'].dt.strftime('%d/%m/%Y').fillna('')
    for col_name in ['Link Nota', 'Link (Streaming - Imagen)']:
        if col_name in df_for_editor.columns:
            df_for_editor[col_name] = df_for_editor[col_name].apply(lambda x: 'Link' if pd.notna(x) else '')
    st.dataframe(df_for_editor, use_container_width=True)
    
# ==============================================================================
# INTERFAZ PRINCIPAL DE STREAMLIT
# ==============================================================================
st.title("🚀 Procesador de Dossiers (Lite) v1.3")
st.markdown("Una herramienta para limpiar, deduplicar y mapear dossieres de noticias.")
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
