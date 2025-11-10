import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import datetime
import io
import re
import html
import numpy as np
from typing import Any

# --- Configuración de la página ---
st.set_page_config(page_title="Procesador de Dossiers (Lite) v2.0", layout="wide")

# --- Constantes ---
FINAL_COLUMN_ORDER = [
    "ID Noticia", "Fecha", "Hora", "Medio", "Tipo de Medio", "Sección - Programa", 
    "Región", "Título", "Autor - Conductor", "Nro. Pagina", "Dimensión", 
    "Duración - Nro. Caracteres", "CPE", "Tier", "Audiencia", "Tono", "Tema", 
    "Temas Generales - Tema", "Resumen - Aclaracion", "Link Nota", 
    "Link (Streaming - Imagen)", "Menciones - Empresa"
]

# ==============================================================================
# SECCIÓN DE FUNCIONES AUXILIARES
# ==============================================================================

def convert_html_entities(text: Any) -> Any:
    if not isinstance(text, str): return text
    text = html.unescape(text)
    def replace_hex(match):
        try: return chr(int(match.group(1), 16))
        except (ValueError, OverflowError): return match.group(0)
    def replace_dec(match):
        try: return chr(int(match.group(1)))
        except (ValueError, OverflowError): return match.group(0)
    text = re.sub(r'&#x([0-9A-Fa-f]+);', replace_hex, text)
    text = re.sub(r'&#(\d+);', replace_dec, text)
    return text.replace('Â', '')

def normalize_title_for_comparison(title: Any) -> str:
    if not isinstance(title, str): return ""
    title = convert_html_entities(title)
    normalized_title = re.sub(r'[^\w\s]', ' ', title, flags=re.UNICODE)
    return ' '.join(normalized_title.lower().split())

def clean_title_for_output(title: Any) -> str:
    if not isinstance(title, str): return ""
    return convert_html_entities(title).strip()

def clean_summary_text(text: Any) -> Any:
    if not isinstance(text, str): return text
    text = convert_html_entities(text)
    text = re.sub(r'(<br>|\[\.\.\.\]|\s+)', ' ', text).strip()
    match = re.search(r'[A-ZÁÉÍÓÚÜÑ]', text)
    if match: text = text[match.start():]
    if text and not text.endswith('...'): text = text.rstrip('.') + '...'
    return text

def read_and_expand_dossier(dossier_file: io.BytesIO) -> pd.DataFrame:
    """
    Lee el archivo en modo estándar (no read-only) para garantizar que los
    hipervínculos se puedan leer directamente de cada celda. Este es el método
    más robusto y compatible.
    """
    # Cargar en modo estándar (read_only=False) y con data_only=True para obtener valores de celda
    wb = load_workbook(dossier_file, read_only=False, data_only=True)
    sheet = wb.active

    headers = [cell.value for cell in sheet[1] if cell.value]
    
    link_nota_idx = headers.index('Link Nota') if 'Link Nota' in headers else -1
    link_streaming_idx = headers.index('Link (Streaming - Imagen)') if 'Link (Streaming - Imagen)' in headers else -1
    
    expanded_rows = []
    # Iterar sobre las filas (a partir de la 2) usando el objeto sheet
    for row_cells in sheet.iter_rows(min_row=2):
        if all(cell.value is None for cell in row_cells): continue

        row_data = {headers[i]: cell.value for i, cell in enumerate(row_cells) if i < len(headers)}

        # --- LÓGICA DE EXTRACCIÓN DE LINKS DIRECTA Y ROBUSTA ---
        # Acceder al atributo .hyperlink directamente de la celda.
        if link_nota_idx != -1:
            cell = row_cells[link_nota_idx]
            if cell.hyperlink:
                row_data['Link Nota'] = cell.hyperlink.target
        
        if link_streaming_idx != -1:
            cell = row_cells[link_streaming_idx]
            if cell.hyperlink:
                row_data['Link (Streaming - Imagen)'] = cell.hyperlink.target
        # --- FIN DE LA LÓGICA ROBUSTA ---

        menciones_str = str(row_data.get('Menciones - Empresa') or '')
        menciones = [m.strip() for m in menciones_str.split(';') if m.strip()]
        
        if not menciones:
            expanded_rows.append(row_data)
        else:
            for mencion in menciones:
                new_row = row_data.copy()
                new_row['Menciones - Empresa'] = mencion
                expanded_rows.append(new_row)

    return pd.DataFrame(expanded_rows)

def to_excel_output(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    final_columns_in_df = [col for col in FINAL_COLUMN_ORDER if col in df.columns]
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
                        worksheet.write_url(row_idx + 1, col_idx, url, link_format, string='Link')
    return output.getvalue()

# ==============================================================================
# LÓGICA DE PROCESAMIENTO PRINCIPAL
# ==============================================================================
def run_full_process(dossier_file: io.BytesIO, config_file: io.BytesIO):
    st.markdown("---")
    progress_bar = st.progress(0)
    progress_text = st.empty()

    progress_text.info("Paso 1/5: Cargando archivo de configuración...")
    try:
        config_sheets = pd.read_excel(config_file, sheet_name=None)
        region_map = pd.Series(
            config_sheets['Regiones'].iloc[:, 1].values, 
            index=config_sheets['Regiones'].iloc[:, 0].astype(str).str.lower().str.strip()
        ).to_dict()
        internet_map = pd.Series(
            config_sheets['Internet'].iloc[:, 1].values, 
            index=config_sheets['Internet'].iloc[:, 0].astype(str).str.lower().str.strip()
        ).to_dict()
    except Exception as e:
        st.error(f"Error al cargar `Configuracion.xlsx`. Revisa que contenga las hojas 'Regiones' e 'Internet'. Detalle: {e}")
        st.stop()
    progress_bar.progress(20)

    progress_text.info("Paso 2/5: Leyendo Dossier y extrayendo links (modo robusto)...")
    try:
        df = read_and_expand_dossier(dossier_file)
        df['Mantener'] = 'Conservar'
    except Exception as e:
        st.error(f"Error crítico al leer el archivo Dossier. Revisa el formato del archivo. Detalle: {e}")
        st.exception(e)
        st.stop()
    progress_bar.progress(40)

    progress_text.info("Paso 3/5: Aplicando limpieza, mapeos y transformaciones...")
    # (El resto de la lógica no necesita cambios)
    df['Título'] = df['Título'].apply(clean_title_for_output)
    df['Resumen - Aclaracion'] = df['Resumen - Aclaracion'].apply(clean_summary_text)

    tipo_medio_map = {'online': 'Internet', 'diario': 'Prensa', 'am': 'Radio', 'fm': 'Radio', 'aire': 'Televisión', 'cable': 'Televisión', 'revista': 'Revista'}
    df['Tipo de Medio'] = df['Tipo de Medio'].astype(str).str.lower().str.strip().map(tipo_medio_map).fillna(df['Tipo de Medio'])
    
    is_internet = df['Tipo de Medio'] == 'Internet'
    is_print = df['Tipo de Medio'].isin(['Prensa', 'Revista'])
    is_broadcast = df['Tipo de Medio'].isin(['Radio', 'Televisión'])

    df.loc[is_internet, ['Link Nota', 'Link (Streaming - Imagen)']] = df.loc[is_internet, ['Link (Streaming - Imagen)', 'Link Nota']].values
    cond_copy_link = is_print & df['Link Nota'].isnull() & df['Link (Streaming - Imagen)'].notnull()
    df.loc[cond_copy_link, 'Link Nota'] = df.loc[cond_copy_link, 'Link (Streaming - Imagen)']
    df.loc[is_print | is_broadcast, 'Link (Streaming - Imagen)'] = None

    if 'Duración - Nro. Caracteres' in df.columns and 'Dimensión' in df.columns:
        df.loc[is_broadcast, 'Dimensión'] = df.loc[is_broadcast, 'Duración - Nro. Caracteres']
        df.loc[is_broadcast, 'Duración - Nro. Caracteres'] = np.nan
    
    df['Región'] = df['Medio'].astype(str).str.lower().str.strip().map(region_map)
    df.loc[is_internet, 'Medio'] = df.loc[is_internet, 'Medio'].astype(str).str.lower().str.strip().map(internet_map).fillna(df.loc[is_internet, 'Medio'])
    progress_bar.progress(60)

    progress_text.info("Paso 4/5: Detectando duplicados...")
    df['titulo_norm'] = df['Título'].apply(normalize_title_for_comparison)
    df['Fecha'] = pd.to_datetime(df['Fecha'], dayfirst=True, errors='coerce').dt.normalize()
    df['seccion_priority'] = df['Sección - Programa'].isnull() | (df['Sección - Programa'] == '')
    df['dup_hora'] = np.where(df['Tipo de Medio'] == 'Internet', 'IGNORE_TIME', df['Hora'])
    dup_cols_exact = ['titulo_norm', 'Medio', 'Fecha', 'Menciones - Empresa', 'dup_hora']
    df.sort_values(by=dup_cols_exact + ['seccion_priority'], ascending=True, inplace=True)
    exact_duplicates_mask = df.duplicated(subset=dup_cols_exact, keep='first')
    df.loc[exact_duplicates_mask, 'Mantener'] = 'Eliminar'
    
    df_internet_to_check = df[(df['Mantener'] == 'Conservar') & is_internet].copy()
    if not df_internet_to_check.empty:
        group_cols = ['titulo_norm', 'Medio', 'Menciones - Empresa']
        df_internet_to_check.sort_values(by=group_cols + ['Fecha'], inplace=True)
        date_diffs = df_internet_to_check.groupby(group_cols)['Fecha'].diff().dt.days
        cluster_ids = (date_diffs.fillna(0) != 1).cumsum()
        df_internet_to_check['date_cluster'] = cluster_ids
        consecutive_duplicates_mask = df_internet_to_check.duplicated(subset=group_cols + ['date_cluster'], keep='first')
        indices_to_eliminate = df_internet_to_check[consecutive_duplicates_mask].index
        df.loc[indices_to_eliminate, 'Mantener'] = 'Eliminar'
    
    df.sort_index(inplace=True)
    df.loc[df['Mantener'] == 'Eliminar', ['Tono', 'Tema', 'Temas Generales - Tema']] = 'Duplicada'
    progress_bar.progress(80)

    progress_text.info("Paso 5/5: Generando resultados y archivo de descarga...")
    st.balloons()
    progress_text.success("¡Proceso completado con éxito!")

    st.subheader("📊 Resumen del Proceso")
    col1, col2, col3 = st.columns(3)
    col1.metric("Filas Totales Procesadas", len(df))
    dups_count = (df['Mantener'] == 'Eliminar').sum()
    col2.metric("Filas Marcadas como Duplicadas", dups_count)
    col3.metric("Filas Únicas Conservadas", len(df) - dups_count)
    
    excel_data = to_excel_output(df)
    st.download_button(
        label="📥 Descargar Archivo Limpio y Mapeado",
        data=excel_data,
        file_name=f"Dossier_Limpio_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.sheet"
    )

    st.subheader("✍️ Previsualización de Resultados")
    final_cols_in_df = [col for col in FINAL_COLUMN_ORDER if col in df.columns]
    df_preview = df[final_cols_in_df].copy()
    if 'Fecha' in df_preview.columns:
        df_preview['Fecha'] = df_preview['Fecha'].dt.strftime('%d/%m/%Y').fillna('')
    for col_name in ['Link Nota', 'Link (Streaming - Imagen)']:
        if col_name in df_preview.columns:
            df_preview[col_name] = df_preview[col_name].apply(lambda x: '🔗 Link' if pd.notna(x) else '')
    st.dataframe(df_preview, use_container_width=True)
    progress_bar.progress(100)

# ==============================================================================
# INTERFAZ PRINCIPAL DE STREAMLIT
# ==============================================================================
st.title("🚀 Procesador de Dossiers (Lite) v2.0")
st.markdown("Herramienta para limpiar, mapear y deduplicar dossieres de noticias de forma rápida y eficiente.")

st.info(
    "**Instrucciones:**\n\n"
    "1. Prepara tu archivo **Dossier** principal en formato `.xlsx`.\n"
    "2. Asegúrate de tener tu archivo `Configuracion.xlsx` con las hojas requeridas.\n"
    "3. Sube ambos archivos juntos en el área de abajo y haz clic en 'Iniciar Proceso'."
)
st.success("✅ **MÉTODO ROBUSTO (v2.0)**: Se utiliza el método de lectura más compatible para garantizar que todos los links se extraigan correctamente, priorizando la fiabilidad sobre la optimización extrema.")

with st.expander("Ver estructura requerida para `Configuracion.xlsx`", expanded=False):
    st.markdown("""
    El archivo debe contener dos hojas de cálculo con los siguientes nombres y estructuras:
    - **`Regiones`**:
        - Columna A: Nombre del Medio (tal como aparece en el dossier).
        - Columna B: Región a la que pertenece (ej. CABA, GBA, Córdoba, etc.).
    - **`Internet`**:
        - Columna A: Nombre del Medio de internet (ej. infobae.com).
        - Columna B: Nombre Mapeado del Medio (ej. Infobae).
    """)

uploaded_files = st.file_uploader(
    "Arrastra y suelta tus archivos aquí (Dossier y Configuracion)", 
    type=["xlsx"], 
    accept_multiple_files=True
)

dossier_file, config_file = None, None

if uploaded_files:
    if len(uploaded_files) != 2:
        st.warning("Por favor, sube exactamente dos archivos: el Dossier y el de Configuración.")
    else:
        for file in uploaded_files:
            if 'config' in file.name.lower():
                config_file = file
            else:
                dossier_file = file
        if dossier_file and config_file:
            st.success(f"✔️ Archivo Dossier cargado: **{dossier_file.name}**")
            st.success(f"✔️ Archivo de Configuración cargado: **{config_file.name}**")
        else:
            st.error("No se pudo identificar cuál es el archivo Dossier y cuál el de Configuración. Asegúrate de que uno contenga 'config' en su nombre.")

if st.button("▶️ Iniciar Proceso de Limpieza", disabled=not (dossier_file and config_file), type="primary"):
    run_full_process(dossier_file, config_file)
