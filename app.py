import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import datetime
import io
import re
import html
import numpy as np

# --- Configuración de la página ---
st.set_page_config(page_title="Procesador de Dossiers (Lite) v1.8", layout="wide")

# ==============================================================================
# SECCIÓN DE FUNCIONES AUXILIARES
# ==============================================================================
def extract_link_from_cell(cell):
    """Extrae el hipervínculo de una celda si existe."""
    if cell.hyperlink and cell.hyperlink.target:
        return cell.hyperlink.target
    # Si no hay hipervínculo, devuelve el valor de la celda por si el link está como texto
    return cell.value

def convert_html_entities(text):
    """
    Convierte entidades HTML mal codificadas a caracteres normales.
    Maneja tanto entidades hexadecimales como decimales.
    """
    if not isinstance(text, str):
        return text
    
    text = html.unescape(text)
    
    html_entities = {
        '&#xF3;': 'ó', '&#xE1;': 'á', '&#xE9;': 'é', '&#xED;': 'í', '&#xFA;': 'ú',
        '&#xF1;': 'ñ', '&#xDC;': 'Ü', '&#xFC;': 'ü', '&#xC1;': 'Á', '&#xC9;': 'É',
        '&#xCD;': 'Í', '&#xD3;': 'Ó', '&#xDA;': 'Ú', '&#xD1;': 'Ñ', '&#xC7;': 'Ç',
        '&#xE7;': 'ç',
    }
    
    for entity, char in html_entities.items():
        text = text.replace(entity, char)
    
    def replace_hex_entity(match):
        try:
            return chr(int(match.group(1), 16))
        except (ValueError, OverflowError):
            return match.group(0)
    
    text = re.sub(r'&#x([0-9A-Fa-f]+);', replace_hex_entity, text)
    
    def replace_decimal_entity(match):
        try:
            return chr(int(match.group(1)))
        except (ValueError, OverflowError):
            return match.group(0)
    
    text = re.sub(r'&#(\d+);', replace_decimal_entity, text)
    
    custom_replacements = {'"': '"', "''": "'", 'Â': '', 'â': '', '€': '', '™': ''}
    
    for entity, char in custom_replacements.items():
        text = text.replace(entity, char)
    
    return text

def normalize_title_for_comparison(title):
    """Normaliza el título para comparación de duplicados."""
    if not isinstance(title, str): 
        return ""
    title = convert_html_entities(title)
    return re.sub(r'\W+', ' ', title).lower().strip()

def clean_title_for_output(title):
    """Limpia entidades HTML y espacios en los extremos del título."""
    if not isinstance(title, str): 
        return ""
    title = convert_html_entities(title)
    return title.strip()

def corregir_texto(text):
    """Limpia y formatea el texto de resumen."""
    if not isinstance(text, str): return text
    text = convert_html_entities(text)
    text = re.sub(r'(<br>|\[\.\.\.\]|\s+)', ' ', text).strip()
    match = re.search(r'[A-Z]', text)
    if match: text = text[match.start():]
    if text and not text.endswith('...'): text = text.rstrip('.') + '...'
    return text

# ==============================================================================
# --- FUNCIÓN to_excel_from_df (CORREGIDA) ---
# ==============================================================================
def to_excel_from_df(df, final_order):
    """
    Convierte un DataFrame a un archivo Excel en memoria.
    1. Escribe el DataFrame completo.
    2. Sobrescribe las celdas de link con fórmulas HYPERLINK saneadas para evitar corrupción.
       Esto respeta la lógica de negocio ya aplicada al DataFrame.
    """
    output = io.BytesIO()
    final_columns_in_df = [col for col in final_order if col in df.columns]
    df_to_excel = df[final_columns_in_df]

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Paso 1: Escribir todo el DataFrame. Las columnas de link se escribirán como texto plano por ahora.
        df_to_excel.to_excel(writer, index=False, sheet_name='Resultado')

        workbook = writer.book
        worksheet = writer.sheets['Resultado']
        link_format = workbook.add_format({'color': 'blue', 'underline': 1})
        
        # Paso 2: Iterar solo sobre las columnas de link y sobrescribir con fórmulas donde corresponda.
        # Esto respeta el DataFrame que ya ha sido procesado (links intercambiados/borrados).
        link_columns_to_process = ['Link Nota', 'Link (Streaming - Imagen)']
        
        for col_name in link_columns_to_process:
            if col_name in df_to_excel.columns:
                col_idx = df_to_excel.columns.get_loc(col_name)
                
                # Iteramos sobre los valores de la columna en el DataFrame
                for row_idx, url in enumerate(df_to_excel[col_name]):
                    # La fila en Excel es el índice + 1 (por la cabecera)
                    excel_row = row_idx + 1
                    
                    # Condición: Solo actuar si hay una URL válida
                    if pd.notna(url) and isinstance(url, str) and url.startswith('http'):
                        # --- SOLUCIÓN A LA CORRUPCIÓN ---
                        # Escapar comillas dobles en la URL, que es la causa principal de corrupción.
                        safe_url = url.replace('"', '""')
                        
                        # Crear la fórmula de Excel
                        formula = f'=HYPERLINK("{safe_url}", "Link")'
                        
                        # Escribir la fórmula en la celda correcta, sobrescribiendo el texto plano
                        worksheet.write_formula(excel_row, col_idx, formula, link_format)

    return output.getvalue()


# ==============================================================================
# LÓGICA DE PROCESAMIENTO PRINCIPAL
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

    progress_text.info("Paso 2/4: Leyendo Dossier y expandiendo filas...")
    wb = load_workbook(dossier_file)
    sheet = wb.active
    original_headers = [cell.value for cell in sheet[1] if cell.value]
    rows_to_expand = []
    
    # Se itera con enumerate para obtener el número de fila para una extracción de links más precisa
    for row_num, row in enumerate(sheet.iter_rows(min_row=2), start=2):
        if all(c.value is None for c in row): continue
        row_values = []
        # Extraer valores y links de forma segura
        for col_num, cell in enumerate(row, start=1):
            header = sheet.cell(row=1, column=col_num).value
            if header in ['Link Nota', 'Link (Streaming - Imagen)']:
                # Usamos nuestra función mejorada para obtener el link
                row_values.append(extract_link_from_cell(cell))
            else:
                row_values.append(cell.value)

        row_data = dict(zip(original_headers, row_values))

        menciones = [m.strip() for m in str(row_data.get('Menciones - Empresa') or '').split(';') if m.strip()]
        if not menciones:
            rows_to_expand.append(row_data)
        else:
            for mencion in menciones:
                new_row = row_data.copy()
                new_row['Menciones - Empresa'] = mencion
                rows_to_expand.append(new_row)

    df = pd.DataFrame(rows_to_expand)
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
    
    # --- LÓGICA DE NEGOCIO ORIGINAL (Y CORRECTA) PARA LOS LINKS ---
    # Esta sección ahora se ejecutará sin problemas antes de la exportación a Excel
    df.loc[is_internet, ['Link Nota', 'Link (Streaming - Imagen)']] = df.loc[is_internet, ['Link (Streaming - Imagen)', 'Link Nota']].values
    cond_copy = is_print & df['Link Nota'].isnull() & df['Link (Streaming - Imagen)'].notnull()
    df.loc[cond_copy, 'Link Nota'] = df.loc[cond_copy, 'Link (Streaming - Imagen)']
    df.loc[is_print, 'Link (Streaming - Imagen)'] = np.nan
    df.loc[is_broadcast, 'Link (Streaming - Imagen)'] = np.nan
    
    if 'Duración - Nro. Caracteres' in df.columns and 'Dimensión' in df.columns:
        df.loc[is_broadcast, 'Dimensión'] = df.loc[is_broadcast, 'Duración - Nro. Caracteres']
        df.loc[is_broadcast, 'Duración - Nro. Caracteres)'] = np.nan
    
    df['Región'] = df['Medio'].astype(str).str.lower().str.strip().map(region_map)
    df.loc[is_internet, 'Medio'] = df.loc[is_internet, 'Medio'].astype(str).str.lower().str.strip().map(internet_map).fillna(df.loc[is_internet, 'Medio'])

    progress_text.info("Paso 4/4: Detectando duplicados y generando resultados...")
    df['titulo_norm'] = df['Título'].apply(normalize_title_for_comparison)
    df['Fecha'] = pd.to_datetime(df['Fecha'], dayfirst=True, errors='coerce').dt.normalize()
    
    df['seccion_priority'] = df['Sección - Programa'].isnull() | (df['Sección - Programa'] == '')
    df['dup_hora'] = np.where(df['Tipo de Medio'] == 'Internet', 'IGNORE_TIME', df['Hora'])
    
    dup_cols_exact = ['titulo_norm', 'Medio', 'Fecha', 'Menciones - Empresa', 'dup_hora']
    sort_by_cols = dup_cols_exact + ['seccion_priority']
    ascending_order = [True] * len(dup_cols_exact) + [False]
    df.sort_values(by=sort_by_cols, ascending=ascending_order, inplace=True, na_position='last')
    exact_duplicates_mask = df.duplicated(subset=dup_cols_exact, keep='first')
    df.loc[exact_duplicates_mask, 'Mantener'] = 'Eliminar'
    df.sort_index(inplace=True)
    
    df_internet_to_check = df[(df['Mantener'] == 'Conservar') & (is_internet)].copy()
    if not df_internet_to_check.empty:
        group_cols = ['titulo_norm', 'Medio', 'Menciones - Empresa']
        df_internet_to_check.sort_values(by=group_cols + ['Fecha'], inplace=True)
        date_diffs = df_internet_to_check.groupby(group_cols)['Fecha'].diff().dt.days
        cluster_ids = (date_diffs.notna() & (date_diffs != 1)).cumsum()
        df_internet_to_check['date_cluster'] = cluster_ids
        sort_by_cols_consecutive = group_cols + ['date_cluster', 'seccion_priority']
        ascending_order_consecutive = [True] * (len(group_cols) + 1) + [False]
        df_internet_to_check.sort_values(by=sort_by_cols_consecutive, ascending=ascending_order_consecutive, inplace=True, na_position='last')
        consecutive_duplicates_mask = df_internet_to_check.duplicated(subset=group_cols + ['date_cluster'], keep='first')
        indices_to_eliminate = df_internet_to_check[consecutive_duplicates_mask].index
        df.loc[indices_to_eliminate, 'Mantener'] = 'Eliminar'
    
    df.loc[df['Mantener'] == 'Eliminar', ['Tono', 'Tema', 'Temas Generales - Tema']] = 'Duplicada'
    
    st.balloons()
    progress_text.success("¡Proceso de limpieza completado! Los títulos se mantienen completos.")

    final_order = ["ID Noticia", "Fecha", "Hora", "Medio", "Tipo de Medio", "Sección - Programa", "Región", "Título", "Autor - Conductor", "Nro. Pagina", "Dimensión", "Duración - Nro. Caracteres", "CPE", "Tier", "Audiencia", "Tono", "Tema", "Temas Generales - Tema", "Resumen - Aclaracion", "Link Nota", "Link (Streaming - Imagen)", "Menciones - Empresa", "Mantener"]
    df_final = df.copy()

    st.subheader("📊 Resumen del Proceso")
    col1, col2, col3 = st.columns(3)
    col1.metric("Filas Totales", len(df_final))
    dups_count = (df_final['Mantener'] == 'Eliminar').sum()
    col2.metric("Filas Marcadas como Duplicadas", dups_count)
    col3.metric("Filas Únicas", len(df_final) - dups_count)
    
    excel_data = to_excel_from_df(df_final, final_order)
    st.download_button(label="📥 Descargar Archivo Limpio y Mapeado", data=excel_data, file_name=f"Dossier_Limpio_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.subheader("✍️ Previsualización de Resultados")
    final_columns_in_df = [col for col in final_order if col in df_final.columns]
    df_for_editor = df_final[final_columns_in_df].copy()
    if 'Fecha' in df_for_editor.columns:
        df_for_editor['Fecha'] = df_for_editor['Fecha'].dt.strftime('%d/%m/%Y').fillna('')
    for col_name in ['Link Nota', 'Link (Streaming - Imagen)']:
        if col_name in df_for_editor.columns:
            df_for_editor[col_name] = df_for_editor[col_name].apply(lambda x: 'Link' if pd.notna(x) and str(x).startswith('http') else '')
    st.dataframe(df_for_editor, use_container_width=True)
    
# ==============================================================================
# INTERFAZ PRINCIPAL DE STREAMLIT
# ==============================================================================
st.title("🚀 Procesador de Dossiers (Lite) v1.8")
st.markdown("Una herramienta para limpiar, deduplicar y mapear dossieres de noticias.")
st.info("**Instrucciones:**\n\n1. Prepara tu archivo **Dossier** principal y tu archivo **`Configuracion.xlsx`**.\n2. Sube ambos archivos juntos en el área de abajo.\n3. Haz clic en 'Iniciar Proceso'.")

st.success("✅ **CORREGIDO (v1.8)**: Solucionado el error de corrupción de archivos Excel al generar hipervínculos. Se respeta la lógica de columnas de links.")
st.success("✅ **MEJORADO**: Los títulos ahora se mantienen completos. Solo se limpian entidades HTML como &#xF3; → ó")

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
