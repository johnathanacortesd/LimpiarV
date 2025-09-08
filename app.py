import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import datetime
import io
import re
import html
import numpy as np
from urllib.parse import urlparse

# --- Configuración de la página ---
st.set_page_config(page_title="Procesador de Dossiers (Lite) v1.8", layout="wide")

# ==============================================================================
# NUEVAS FUNCIONES PARA EXTRACCIÓN DE DOMINIO Y REGIÓN
# ==============================================================================
def format_domain_name(url):
    """
    Extrae y formatea el nombre del dominio de una URL
    Ejemplo: https://www.lavibrante.com -> Lavibrante.com
    """
    if not url:
        return ''
    
    try:
        # Parsear la URL
        parsed = urlparse(url)
        domain = parsed.netloc
        
        # Remover 'www.' si está presente
        if domain.startswith('www.'):
            domain = domain[4:]
        
        # Capitalizar la primera letra
        if domain:
            # Separar el nombre del dominio de la extensión
            parts = domain.split('.')
            if len(parts) >= 2:
                # Capitalizar solo la primera parte
                parts[0] = parts[0].capitalize()
                domain = '.'.join(parts)
            else:
                domain = domain.capitalize()
        
        return domain
    except:
        return ''

def get_country_from_domain(url):
    """
    Determina el país/región basándose en el dominio de la URL
    """
    if not url:
        return ''
    
    try:
        # Parsear la URL
        parsed = urlparse(url)
        domain = parsed.netloc.lower()
        
        # Remover www. si está presente
        if domain.startswith('www.'):
            domain = domain[4:]
        
        # Diccionario de extensiones de dominio por país
        domain_countries = {
            # América Latina
            '.ar': 'Argentina',
            '.bo': 'Bolivia',
            '.br': 'Brasil',
            '.cl': 'Chile',
            '.co': 'Colombia',
            '.cr': 'Costa Rica',
            '.cu': 'Cuba',
            '.do': 'República Dominicana',
            '.ec': 'Ecuador',
            '.sv': 'El Salvador',
            '.gt': 'Guatemala',
            '.hn': 'Honduras',
            '.mx': 'México',
            '.ni': 'Nicaragua',
            '.pa': 'Panamá',
            '.py': 'Paraguay',
            '.pe': 'Perú',
            '.pr': 'Puerto Rico',
            '.uy': 'Uruguay',
            '.ve': 'Venezuela',
            
            # Europa
            '.es': 'España',
            '.fr': 'Francia',
            '.de': 'Alemania',
            '.it': 'Italia',
            '.pt': 'Portugal',
            '.uk': 'Reino Unido',
            '.gb': 'Reino Unido',
            '.nl': 'Países Bajos',
            '.be': 'Bélgica',
            '.ch': 'Suiza',
            '.at': 'Austria',
            '.se': 'Suecia',
            '.no': 'Noruega',
            '.dk': 'Dinamarca',
            '.fi': 'Finlandia',
            '.pl': 'Polonia',
            '.ru': 'Rusia',
            '.gr': 'Grecia',
            '.ie': 'Irlanda',
            
            # Asia
            '.cn': 'China',
            '.jp': 'Japón',
            '.kr': 'Corea del Sur',
            '.in': 'India',
            '.sg': 'Singapur',
            '.hk': 'Hong Kong',
            '.tw': 'Taiwán',
            '.th': 'Tailandia',
            '.my': 'Malasia',
            '.id': 'Indonesia',
            '.ph': 'Filipinas',
            '.vn': 'Vietnam',
            
            # Oceanía
            '.au': 'Australia',
            '.nz': 'Nueva Zelanda',
            
            # África
            '.za': 'Sudáfrica',
            '.eg': 'Egipto',
            '.ng': 'Nigeria',
            '.ke': 'Kenia',
            '.ma': 'Marruecos',
            
            # Norteamérica
            '.ca': 'Canadá',
            '.us': 'Estados Unidos',
            
            # Medio Oriente
            '.ae': 'Emiratos Árabes Unidos',
            '.sa': 'Arabia Saudita',
            '.il': 'Israel',
            '.tr': 'Turquía',
            '.ir': 'Irán'
        }
        
        # Primero verificar extensiones de dominio específicas de país
        for ext, country in domain_countries.items():
            if domain.endswith(ext):
                return country
        
        # Verificar subdominios .com.xx
        com_extensions = {
            '.com.ar': 'Argentina',
            '.com.mx': 'México',
            '.com.br': 'Brasil',
            '.com.co': 'Colombia',
            '.com.pe': 'Perú',
            '.com.ve': 'Venezuela',
            '.com.ec': 'Ecuador',
            '.com.uy': 'Uruguay',
            '.com.py': 'Paraguay',
            '.com.bo': 'Bolivia',
            '.com.cl': 'Chile',
            '.com.gt': 'Guatemala',
            '.com.do': 'República Dominicana',
            '.com.pa': 'Panamá',
            '.com.ni': 'Nicaragua',
            '.com.sv': 'El Salvador',
            '.com.hn': 'Honduras'
        }
        
        for ext, country in com_extensions.items():
            if ext in domain:
                return country
        
        # Si es .com, .org, .net sin indicador de país específico
        if any(domain.endswith(ext) for ext in ['.com', '.org', '.net', '.info', '.biz']):
            return 'Internacional'
        
        # Si es .edu o .gov generalmente es de Estados Unidos
        if domain.endswith('.edu') or domain.endswith('.gov'):
            return 'Estados Unidos'
        
        return 'No identificado'
        
    except:
        return ''

def clean_and_format_online_media(medio_name, link_url):
    """
    Limpia y formatea los nombres de medios online
    Ejemplo: "Ser Peruano (Online)" -> "Serperuano.com" si tiene link válido
    """
    if not medio_name:
        return medio_name
    
    medio_str = str(medio_name).strip()
    
    # Si tiene "(Online)" al final y tenemos un link válido
    if '(online)' in medio_str.lower() and link_url:
        domain_name = format_domain_name(link_url)
        if domain_name:
            # Remover "(Online)" y cualquier espacio extra
            base_name = re.sub(r'\s*\(online\)\s*', '', medio_str, flags=re.IGNORECASE).strip()
            # Si el dominio ya está en el nombre, no duplicarlo
            if domain_name.lower() not in base_name.lower():
                return f"{base_name} - {domain_name}"
            else:
                return base_name
    
    return medio_str

# ==============================================================================
# SECCIÓN DE FUNCIONES AUXILIARES (EXISTENTES)
# ==============================================================================
def extract_link_from_cell(cell):
    if cell.hyperlink and cell.hyperlink.target:
        return cell.hyperlink.target
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
# LÓGICA DE PROCESAMIENTO PRINCIPAL (MEJORADA)
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
    
    # APLICAR LA LIMPIEZA CORREGIDA - SOLO ENTIDADES HTML, NO CORTAR
    df['Título'] = df['Título'].astype(str).apply(clean_title_for_output)
    df['Resumen - Aclaracion'] = df['Resumen - Aclaracion'].astype(str).apply(corregir_texto)

    tipo_medio_map = {'online': 'Internet', 'diario': 'Prensa', 'am': 'Radio', 'fm': 'Radio', 'aire': 'Televisión', 'cable': 'Televisión', 'revista': 'Revista'}
    df['Tipo de Medio'] = df['Tipo de Medio'].str.lower().str.strip().map(tipo_medio_map).fillna(df['Tipo de Medio'])
    is_internet = df['Tipo de Medio'] == 'Internet'
    is_print = df['Tipo de Medio'].isin(['Prensa', 'Revista'])
    is_broadcast = df['Tipo de Medio'].isin(['Radio', 'Televisión'])
    
    # Intercambiar links para medios de Internet
    df.loc[is_internet, ['Link Nota', 'Link (Streaming - Imagen)']] = df.loc[is_internet, ['Link (Streaming - Imagen)', 'Link Nota']].values
    cond_copy = is_print & df['Link Nota'].isnull() & df['Link (Streaming - Imagen)'].notnull()
    df.loc[cond_copy, 'Link Nota'] = df.loc[cond_copy, 'Link (Streaming - Imagen)']
    df.loc[is_print, 'Link (Streaming - Imagen)'] = None
    df.loc[is_broadcast, 'Link (Streaming - Imagen)'] = None
    
    # LÓGICA "CORTAR Y PEGAR" para Duración - Nro. Caracteres
    if 'Duración - Nro. Caracteres' in df.columns and 'Dimensión' in df.columns:
        df.loc[is_broadcast, 'Dimensión'] = df.loc[is_broadcast, 'Duración - Nro. Caracteres']
        df.loc[is_broadcast, 'Duración - Nro. Caracteres'] = np.nan
    
    # === NUEVA LÓGICA MEJORADA PARA REGIÓN Y MEDIO ===
    # Primero intentar mapear región con el archivo de configuración
    df['Región'] = df['Medio'].astype(str).str.lower().str.strip().map(region_map)
    
    # Para medios de Internet, aplicar lógica adicional
    internet_indices = df.index[is_internet]
    for idx in internet_indices:
        medio_original = str(df.loc[idx, 'Medio'])
        link_url = df.loc[idx, 'Link Nota'] if pd.notna(df.loc[idx, 'Link Nota']) else df.loc[idx, 'Link (Streaming - Imagen)']
        
        # Intentar mapear el medio usando el archivo de configuración
        medio_mapped = internet_map.get(medio_original.lower().strip())
        
        if medio_mapped:
            # Si encontramos un mapeo, usarlo
            df.loc[idx, 'Medio'] = medio_mapped
        else:
            # Si no hay mapeo y tiene "(Online)", formatear usando el dominio
            if link_url:
                df.loc[idx, 'Medio'] = clean_and_format_online_media(medio_original, link_url)
            
        # Si no se encontró región en el mapeo y tenemos un link, intentar detectarla del dominio
        if pd.isna(df.loc[idx, 'Región']) and link_url:
            detected_region = get_country_from_domain(link_url)
            if detected_region and detected_region not in ['No identificado', '']:
                df.loc[idx, 'Región'] = detected_region

    progress_text.info("Paso 4/4: Detectando duplicados y generando resultados...")
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
    progress_text.success("¡Proceso de limpieza completado! Los títulos se mantienen completos y los medios online están formateados.")

    final_order = ["ID Noticia", "Fecha", "Hora", "Medio", "Tipo de Medio", "Sección - Programa", "Región", "Título", "Autor - Conductor", "Nro. Pagina", "Dimensión", "Duración - Nro. Caracteres", "CPE", "Tier", "Audiencia", "Tono", "Tema", "Temas Generales - Tema", "Resumen - Aclaracion", "Link Nota", "Link (Streaming - Imagen)", "Menciones - Empresa"]
    df_final = df.copy()

    st.subheader("📊 Resumen del Proceso")
    col1, col2, col3 = st.columns(3)
    col1.metric("Filas Totales", len(df_final))
    dups_count = (df_final['Mantener'] == 'Eliminar').sum()
    col2.metric("Filas Marcadas como Duplicadas", dups_count)
    col3.metric("Filas Únicas", len(df_final) - dups_count)
    
    # Mostrar estadísticas de detección de región
    with st.expander("📍 Ver estadísticas de detección de región"):
        # Estadísticas para medios Internet
        internet_df = df_final[df_final['Tipo de Medio'] == 'Internet']
        if not internet_df.empty:
            st.write("**Medios de Internet:**")
            total_internet = len(internet_df)
            with_region = internet_df['Región'].notna().sum()
            without_region = total_internet - with_region
            
            col1, col2 = st.columns(2)
            col1.metric("Con región detectada", f"{with_region} ({with_region/total_internet*100:.1f}%)")
            col2.metric("Sin región", f"{without_region} ({without_region/total_internet*100:.1f}%)")
        
        # Estadísticas para TODOS los medios
        st.write("**Todos los medios:**")
        total_all = len(df_final)
        with_region_all = df_final['Región'].notna().sum()
        without_region_all = total_all - with_region_all
        
        col3, col4 = st.columns(2)
        col3.metric("Total con región", f"{with_region_all} ({with_region_all/total_all*100:.1f}%)")
        col4.metric("Total sin región", f"{without_region_all} ({without_region_all/total_all*100:.1f}%)")
        
        # Mostrar distribución de regiones
        if with_region_all > 0:
            st.write("**Top 10 regiones detectadas:**")
            region_counts = df_final['Región'].value_counts().head(10)
            st.bar_chart(region_counts)
            
        # Mostrar algunos ejemplos de medios sin región para debugging
        if without_region_all > 0:
            st.write("**Ejemplos de medios sin región detectada:**")
            no_region_df = df_final[df_final['Región'].isna()][['Medio', 'Tipo de Medio', 'Link (Streaming - Imagen)', 'Link Nota']].head(5)
            st.dataframe(no_region_df)
    
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
st.title("🚀 Procesador de Dossiers (Lite) v1.7")
st.markdown("Una herramienta para limpiar, deduplicar y mapear dossieres de noticias.")
st.info("**Instrucciones:**\n\n1. Prepara tu archivo **Dossier** principal y tu archivo **`Configuracion.xlsx`**.\n2. Sube ambos archivos juntos en el área de abajo.\n3. Haz clic en 'Iniciar Proceso'.")

# Información adicional sobre las mejoras
st.success("""✅ **MEJORAS v1.7**: 
- Los títulos se mantienen completos (solo se limpian entidades HTML)
- Detección automática de región para medios online usando dominios de Link Nota
- Formateo automático de medios online: 'Medio (Online)' → 'Medio - Dominio.com'
- Si no encuentra región en el mapeo, intenta detectarla del dominio del link
- Para medios Internet, usa Link Nota (donde están los enlaces web después del procesamiento)""")

with st.expander("Ver estructura requerida para `Configuracion.xlsx`"):
    st.markdown("""
    - **`Regiones`**: Columna A (Medio), Columna B (Región).
    - **`Internet`**: Columna A (Medio Original), Columna B (Medio Mapeado).
    
    **Nota:** Si un medio online no está en estos archivos, el sistema intentará:
    1. Formatear el nombre usando el dominio del link
    2. Detectar la región automáticamente basándose en la extensión del dominio (.pe = Perú, .cl = Chile, etc.)
    """)

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
