import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import datetime
import io
import re
import html
import numpy as np
from typing import Dict, List, Any
import gc

# --- Configuración de la página ---
st.set_page_config(page_title="Procesador de Dossiers (Lite) v1.7", layout="wide")

# ==============================================================================
# SECCIÓN DE FUNCIONES AUXILIARES
# ==============================================================================
def extract_link_from_cell(cell):
    """Extrae hipervínculos de celdas de Excel"""
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
        '&#xF3;': 'ó', '&#xE1;': 'á', '&#xE9;': 'é', '&#xED;': 'í', '&#xFA;': 'ú',
        '&#xF1;': 'ñ', '&#xDC;': 'Ü', '&#xFC;': 'ü', '&#xC1;': 'Á', '&#xC9;': 'É',
        '&#xCD;': 'Í', '&#xD3;': 'Ó', '&#xDA;': 'Ú', '&#xD1;': 'Ñ', '&#xC7;': 'Ç',
        '&#xE7;': 'ç',
    }
    
    for entity, char in html_entities.items():
        text = text.replace(entity, char)
    
    # Patrón para capturar otras entidades hexadecimales
    def replace_hex_entity(match):
        try:
            hex_code = match.group(1)
            char_code = int(hex_code, 16)
            return chr(char_code)
        except (ValueError, OverflowError):
            return match.group(0)
    
    text = re.sub(r'&#x([0-9A-Fa-f]+);', replace_hex_entity, text)
    
    # Patrón para entidades decimales
    def replace_decimal_entity(match):
        try:
            decimal_code = int(match.group(1))
            return chr(decimal_code)
        except (ValueError, OverflowError):
            return match.group(0)
    
    text = re.sub(r'&#(\d+);', replace_decimal_entity, text)
    
    # Limpiar caracteres problemáticos adicionales
    custom_replacements = {
        '"': '"', '"': '"', ''': "'", ''': "'", 
        'Â': '', 'â': '', '€': '', '™': ''
    }
    
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
    """
    ÚNICAMENTE limpia entidades HTML mal codificadas.
    NO corta, NO modifica, NO remueve ninguna parte del título.
    """
    if not isinstance(title, str): 
        return ""
    
    title = convert_html_entities(title)
    title = title.strip()
    
    return title

def corregir_texto(text):
    """Corrige formato de resúmenes"""
    if not isinstance(text, str): 
        return text
    text = convert_html_entities(text)
    text = re.sub(r'(<br>|\[\.\.\.\]|\s+)', ' ', text).strip()
    match = re.search(r'[A-Z]', text)
    if match: 
        text = text[match.start():]
    if text and not text.endswith('...'): 
        text = text.rstrip('.') + '...'
    return text

def read_excel_with_links_optimized(file_obj, chunk_size=1000):
    """
    Lee un archivo Excel de forma optimizada, procesando por chunks
    y extrayendo hipervínculos correctamente.
    """
    # Guardar el archivo temporalmente para poder leerlo múltiples veces
    file_bytes = file_obj.read()
    file_obj.seek(0)
    
    # Leer estructura básica con pandas (rápido)
    df_base = pd.read_excel(io.BytesIO(file_bytes), engine='openpyxl')
    
    # Cargar workbook para extraer links
    wb = load_workbook(io.BytesIO(file_bytes), data_only=False, read_only=True)
    sheet = wb.active
    
    # Obtener headers
    headers = [cell.value for cell in sheet[1] if cell.value]
    
    # Identificar columnas con links
    link_columns = ['Link Nota', 'Link (Streaming - Imagen)']
    link_col_indices = {}
    for col_name in link_columns:
        if col_name in headers:
            link_col_indices[col_name] = headers.index(col_name)
    
    # Extraer links por chunks para no sobrecargar memoria
    all_links = {col: {} for col in link_columns if col in headers}
    
    row_count = 0
    for row_idx, row in enumerate(sheet.iter_rows(min_row=2), start=0):
        if all(c.value is None for c in row):
            continue
            
        # Extraer links de esta fila
        for col_name, col_idx in link_col_indices.items():
            if col_idx < len(row):
                link = extract_link_from_cell(row[col_idx])
                if link:
                    all_links[col_name][row_count] = link
        
        row_count += 1
        
        # Liberar memoria cada 1000 filas
        if row_count % 1000 == 0:
            gc.collect()
    
    wb.close()
    
    # Aplicar links al dataframe
    for col_name, links_dict in all_links.items():
        if col_name in df_base.columns:
            for idx, link in links_dict.items():
                if idx < len(df_base):
                    df_base.at[idx, col_name] = link
    
    return df_base, headers

def expand_mentions(df, original_headers):
    """
    Expande filas por menciones de empresa de forma optimizada
    """
    rows_expanded = []
    
    for idx, row in df.iterrows():
        menciones_str = str(row.get('Menciones - Empresa', ''))
        menciones = [m.strip() for m in menciones_str.split(';') if m.strip()]
        
        if not menciones:
            rows_expanded.append(row.to_dict())
        else:
            for mencion in menciones:
                new_row = row.to_dict()
                new_row['Menciones - Empresa'] = mencion
                rows_expanded.append(new_row)
        
        # Liberar memoria cada 1000 filas
        if len(rows_expanded) % 1000 == 0:
            gc.collect()
    
    return pd.DataFrame(rows_expanded)

def to_excel_from_df(df, final_order):
    """Exporta DataFrame a Excel con formato de links"""
    output = io.BytesIO()
    final_columns_in_df = [col for col in final_order if col in df.columns]
    df_to_excel = df[final_columns_in_df].copy()
    
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
        
        # Aplicar formato a links
        for col_name in ['Link Nota', 'Link (Streaming - Imagen)']:
            if col_name in df_to_excel.columns:
                col_idx = df_to_excel.columns.get_loc(col_name)
                for row_idx, url in enumerate(df_to_excel[col_name]):
                    if pd.notna(url) and isinstance(url, str) and url.startswith('http'):
                        try:
                            worksheet.write_url(row_idx + 1, col_idx, url, link_format, 'Link')
                        except Exception:
                            # Si falla escribir URL, al menos poner el texto
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
        # PASO 1: Cargar configuración
        progress_text.info("Paso 1/5: Cargando archivo de configuración...")
        progress_bar.progress(10)
        
        config_sheets = pd.read_excel(config_file, sheet_name=None)
        region_map = pd.Series(
            config_sheets['Regiones'].iloc[:, 1].values, 
            index=config_sheets['Regiones'].iloc[:, 0].astype(str).str.lower().str.strip()
        ).to_dict()
        internet_map = pd.Series(
            config_sheets['Internet'].iloc[:, 1].values, 
            index=config_sheets['Internet'].iloc[:, 0].astype(str).str.lower().str.strip()
        ).to_dict()
        
        # PASO 2: Leer Dossier con links
        progress_text.info("Paso 2/5: Leyendo Dossier y extrayendo hipervínculos...")
        progress_bar.progress(20)
        
        df, original_headers = read_excel_with_links_optimized(dossier_file)
        
        st.info(f"✅ Se leyeron {len(df)} filas del archivo original")
        
        # PASO 3: Expandir menciones
        progress_text.info("Paso 3/5: Expandiendo menciones de empresa...")
        progress_bar.progress(35)
        
        df = expand_mentions(df, original_headers)
        df['Mantener'] = 'Conservar'
        
        st.info(f"✅ Después de expandir menciones: {len(df)} filas")
        
        # Asegurar que todas las columnas estén presentes
        for col in original_headers:
            if col not in df.columns:
                df[col] = None
        
        # PASO 4: Limpieza y mapeos
        progress_text.info("Paso 4/5: Aplicando limpieza, mapeos y normalizaciones...")
        progress_bar.progress(50)
        
        # Limpieza de texto
        df['Título'] = df['Título'].astype(str).apply(clean_title_for_output)
        df['Resumen - Aclaracion'] = df['Resumen - Aclaracion'].astype(str).apply(corregir_texto)
        
        # Mapeo de tipo de medio
        tipo_medio_map = {
            'online': 'Internet', 'diario': 'Prensa', 'am': 'Radio', 
            'fm': 'Radio', 'aire': 'Televisión', 'cable': 'Televisión', 
            'revista': 'Revista'
        }
        df['Tipo de Medio'] = df['Tipo de Medio'].str.lower().str.strip().map(tipo_medio_map).fillna(df['Tipo de Medio'])
        
        # Máscaras para diferentes tipos de medios
        is_internet = df['Tipo de Medio'] == 'Internet'
        is_print = df['Tipo de Medio'].isin(['Prensa', 'Revista'])
        is_broadcast = df['Tipo de Medio'].isin(['Radio', 'Televisión'])
        
        # Intercambio de links para Internet
        df.loc[is_internet, ['Link Nota', 'Link (Streaming - Imagen)']] = \
            df.loc[is_internet, ['Link (Streaming - Imagen)', 'Link Nota']].values
        
        # Copia de links para medios impresos
        cond_copy = is_print & df['Link Nota'].isnull() & df['Link (Streaming - Imagen)'].notnull()
        df.loc[cond_copy, 'Link Nota'] = df.loc[cond_copy, 'Link (Streaming - Imagen)']
        df.loc[is_print, 'Link (Streaming - Imagen)'] = None
        df.loc[is_broadcast, 'Link (Streaming - Imagen)'] = None
        
        # Lógica de cortar y pegar para Duración/Dimensión
        if 'Duración - Nro. Caracteres' in df.columns and 'Dimensión' in df.columns:
            df.loc[is_broadcast, 'Dimensión'] = df.loc[is_broadcast, 'Duración - Nro. Caracteres']
            df.loc[is_broadcast, 'Duración - Nro. Caracteres'] = np.nan
        
        # Mapeos de región y medio
        df['Región'] = df['Medio'].astype(str).str.lower().str.strip().map(region_map)
        df.loc[is_internet, 'Medio'] = \
            df.loc[is_internet, 'Medio'].astype(str).str.lower().str.strip().map(internet_map).fillna(df.loc[is_internet, 'Medio'])
        
        # PASO 5: Detección de duplicados
        progress_text.info("Paso 5/5: Detectando duplicados...")
        progress_bar.progress(70)
        
        df['titulo_norm'] = df['Título'].apply(normalize_title_for_comparison)
        df['Fecha'] = pd.to_datetime(df['Fecha'], dayfirst=True, errors='coerce').dt.normalize()
        
        df['seccion_priority'] = df['Sección - Programa'].isnull() | (df['Sección - Programa'] == '')
        df['dup_hora'] = np.where(df['Tipo de Medio'] == 'Internet', 'IGNORE_TIME', df['Hora'])
        
        # Duplicados exactos
        dup_cols_exact = ['titulo_norm', 'Medio', 'Fecha', 'Menciones - Empresa', 'dup_hora']
        sort_by_cols = dup_cols_exact + ['seccion_priority']
        ascending_order = [True] * len(dup_cols_exact) + [False]
        
        df.sort_values(by=sort_by_cols, ascending=ascending_order, inplace=True)
        exact_duplicates_mask = df.duplicated(subset=dup_cols_exact, keep='first')
        df.loc[exact_duplicates_mask, 'Mantener'] = 'Eliminar'
        df.sort_index(inplace=True)
        
        progress_bar.progress(85)
        
        # Duplicados consecutivos de Internet
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
        
        # Marcar duplicadas
        df.loc[df['Mantener'] == 'Eliminar', ['Tono', 'Tema', 'Temas Generales - Tema']] = 'Duplicada'
        
        progress_bar.progress(100)
        progress_text.success("✅ ¡Proceso completado exitosamente!")
        st.balloons()
        
        # Limpiar columnas temporales
        cols_to_drop = ['titulo_norm', 'seccion_priority', 'dup_hora']
        df.drop(columns=[col for col in cols_to_drop if col in df.columns], inplace=True)
        
        # Orden final
        final_order = [
            "ID Noticia", "Fecha", "Hora", "Medio", "Tipo de Medio", "Sección - Programa", 
            "Región", "Título", "Autor - Conductor", "Nro. Pagina", "Dimensión", 
            "Duración - Nro. Caracteres", "CPE", "Tier", "Audiencia", "Tono", "Tema", 
            "Temas Generales - Tema", "Resumen - Aclaracion", "Link Nota", 
            "Link (Streaming - Imagen)", "Menciones - Empresa"
        ]
        
        df_final = df.copy()
        
        # Resumen
        st.subheader("📊 Resumen del Proceso")
        col1, col2, col3 = st.columns(3)
        col1.metric("Filas Totales", len(df_final))
        dups_count = (df_final['Mantener'] == 'Eliminar').sum()
        col2.metric("Filas Duplicadas", dups_count)
        col3.metric("Filas Únicas", len(df_final) - dups_count)
        
        # Generar Excel
        st.info("Generando archivo Excel...")
        excel_data = to_excel_from_df(df_final, final_order)
        
        st.download_button(
            label="📥 Descargar Archivo Procesado",
            data=excel_data,
            file_name=f"Dossier_Limpio_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # Previsualización
        st.subheader("✍️ Previsualización (primeras 1000 filas)")
        final_columns_in_df = [col for col in final_order if col in df_final.columns]
        df_preview = df_final[final_columns_in_df].head(1000).copy()
        
        if 'Fecha' in df_preview.columns:
            df_preview['Fecha'] = df_preview['Fecha'].dt.strftime('%d/%m/%Y').fillna('')
        
        for col_name in ['Link Nota', 'Link (Streaming - Imagen)']:
            if col_name in df_preview.columns:
                df_preview[col_name] = df_preview[col_name].apply(
                    lambda x: '🔗 Link' if pd.notna(x) and str(x).startswith('http') else ''
                )
        
        st.dataframe(df_preview, use_container_width=True, height=400)
        
        # Liberar memoria
        gc.collect()
        
    except Exception as e:
        st.error(f"❌ Error durante el procesamiento: {str(e)}")
        st.exception(e)
        raise

# ==============================================================================
# INTERFAZ PRINCIPAL
# ==============================================================================
st.title("🚀 Procesador de Dossiers (Lite) v1.7")
st.markdown("Una herramienta optimizada para limpiar, deduplicar y mapear dossieres de noticias.")

st.success("""
### ✨ Mejoras en v1.7:
- ✅ Procesamiento optimizado para archivos grandes (+2000 filas)
- ✅ Extracción correcta de hipervínculos en todas las filas
- ✅ Mejor gestión de memoria
- ✅ Manejo robusto de errores
- ✅ Títulos completos preservados (solo limpieza de entidades HTML)
""")

st.info("""
**📋 Instrucciones:**
1. Prepara tu archivo **Dossier** principal y tu archivo **`Configuracion.xlsx`**
2. Sube ambos archivos en el área de abajo
3. Haz clic en 'Iniciar Proceso'
""")

with st.expander("📖 Ver estructura de Configuracion.xlsx"):
    st.markdown("""
    - **Hoja `Regiones`**: 
        - Columna A: Nombre del Medio
        - Columna B: Región
    - **Hoja `Internet`**: 
        - Columna A: Medio Original
        - Columna B: Medio Mapeado
    """)

# Subida de archivos
uploaded_files = st.file_uploader(
    "📁 Arrastra tus archivos aquí (Dossier y Configuracion)",
    type=["xlsx"],
    accept_multiple_files=True
)

dossier_file, config_file = None, None

if uploaded_files:
    for file in uploaded_files:
        if 'config' in file.name.lower():
            config_file = file
        else:
            dossier_file = file
    
    if dossier_file:
        st.success(f"✅ Dossier: **{dossier_file.name}**")
    else:
        st.warning("⚠️ No se detectó el archivo Dossier")
    
    if config_file:
        st.success(f"✅ Configuración: **{config_file.name}**")
    else:
        st.warning("⚠️ No se detectó Configuracion.xlsx")

if st.button(
    "▶️ Iniciar Proceso",
    disabled=not (dossier_file and config_file),
    type="primary",
    use_container_width=True
):
    run_full_process(dossier_file, config_file)

# Footer
st.markdown("---")
st.markdown("💡 **Tip**: Para archivos muy grandes, el proceso puede tardar varios minutos.")
