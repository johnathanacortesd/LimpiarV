# app.py (versión mejorada con mapeo de Menciones-Empresa)

import streamlit as st
import openpyxl
import io
import datetime
from deduplicator import run_deduplication_process # Asumiendo que tu lógica está en este archivo

# --- Configuración de la Página ---
st.set_page_config(
    page_title="Intelli-Clean | Depurador de Noticias IA",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- LÓGICA DE AUTENTICACIÓN ---
def check_password():
    def password_entered():
        try:
            if st.session_state["password"] == st.secrets.password.password:
                st.session_state["password_correct"] = True
                del st.session_state["password"]
            else:
                st.session_state["password_correct"] = False
        except (AttributeError, KeyError):
            st.session_state["password_correct"] = False

    try:
        _ = st.secrets.password.password
    except (AttributeError, KeyError):
        st.error("🚨 ¡Error de configuración! La contraseña no está definida en los 'Secrets' de la aplicación.")
        st.info("""
            Por favor, ve a la configuración de tu app en Streamlit Cloud y añade lo siguiente en la sección 'Secrets':
            ```toml
            [password]
            password = "TU_CONTRASEÑA_AQUI"
            ```
        """)
        return False

    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if not st.session_state["password_correct"]:
        c1, c2, c3 = st.columns([1, 1, 1])
        with c2:
            # --- INTERFAZ DE LOGIN MINIMALISTA CON EMOJI ---
            st.markdown("<h1 style='text-align: center;'>🤖</h1>", unsafe_allow_html=True)
            st.markdown("<h3 style='text-align: center;'>Intelli-Clean Access</h3>", unsafe_allow_html=True)
            
            st.text_input(
                "Contraseña", 
                type="password", 
                on_change=password_entered, 
                key="password",
                placeholder="Introduce la contraseña para continuar",
                label_visibility="collapsed" # Oculta la etiqueta "Contraseña"
            )
            if 'password' in st.session_state and st.session_state.password != "" and not st.session_state.password_correct:
                 st.error("😕 Contraseña incorrecta. Por favor, inténtalo de nuevo.")
        return False
    return True


# --- FLUJO PRINCIPAL DE LA APLICACIÓN ---
if check_password():

    st.title("✨ Intelli-Clean: Depurador de Noticias")
    st.caption("Una herramienta inteligente para mapear, limpiar y deduplicar tus informes de noticias con precisión.")
    st.divider()

    with st.sidebar:
        st.header("📂 Carga tus Archivos")
        uploaded_main_file = st.file_uploader(
            "1. Informe Principal de Noticias", 
            type="xlsx", 
            help="El archivo Excel con todas las noticias a procesar."
        )
        uploaded_internet_map = st.file_uploader(
            "2. Mapeo de Medios de Internet", 
            type="xlsx", 
            help="Un archivo con dos columnas: 'Medio' (nombre antiguo) e 'Internet' (nombre nuevo)."
        )
        uploaded_region_map = st.file_uploader(
            "3. Mapeo de Regiones", 
            type="xlsx", 
            help="Un archivo con dos columnas: 'Medio' y 'Región'."
        )
        uploaded_mentions_map = st.file_uploader(
            "4. Mapeo de Menciones - Empresa", 
            type="xlsx", 
            help="Un archivo con dos columnas: 'Mencion_Original' y 'Mencion_Normalizada' para estandarizar nombres de empresas."
        )
        
        st.divider()
        
        # Configuraciones adicionales
        st.subheader("⚙️ Configuraciones")
        enable_mentions_mapping = st.checkbox(
            "Aplicar mapeo de menciones", 
            value=True if uploaded_mentions_map else False,
            help="Activa esta opción para aplicar el mapeo de menciones-empresa"
        )
        
        process_button = st.button("🚀 Analizar y Depurar Archivos", type="primary", use_container_width=True)

    st.header("Resultados del Análisis")
    
    if process_button:
        # Validar archivos mínimos requeridos
        required_files = [uploaded_main_file, uploaded_internet_map, uploaded_region_map]
        required_names = ["Informe Principal", "Mapeo de Internet", "Mapeo de Regiones"]
        
        if all(required_files):
            with st.status("Iniciando proceso... ⏳", expanded=True) as status:
                try:
                    status.write("📥 Cargando archivos en memoria...")
                    
                    # Cargar archivos principales
                    wb_main = openpyxl.load_workbook(uploaded_main_file)
                    ws_main = wb_main.active
                    wb_internet = openpyxl.load_workbook(uploaded_internet_map, data_only=True)
                    ws_internet = wb_internet.active
                    wb_region = openpyxl.load_workbook(uploaded_region_map, data_only=True)
                    ws_region = wb_region.active

                    status.write("📋 Creando diccionarios de mapeo...")
                    # Crear diccionarios de mapeo
                    internet_dict = {
                        str(row[0].value).lower().strip(): str(row[1].value) 
                        for row in ws_internet.iter_rows(min_row=2) 
                        if row[0].value and len(row) >= 2 and row[1].value
                    }
                    
                    region_dict = {
                        str(row[0].value).lower().strip(): str(row[1].value) 
                        for row in ws_region.iter_rows(min_row=2) 
                        if row[0].value and len(row) >= 2 and row[1].value
                    }
                    
                    # Cargar mapeo de menciones si está disponible
                    mentions_dict = {}
                    if uploaded_mentions_map and enable_mentions_mapping:
                        status.write("🏢 Cargando mapeo de menciones-empresa...")
                        wb_mentions = openpyxl.load_workbook(uploaded_mentions_map, data_only=True)
                        ws_mentions = wb_mentions.active
                        
                        mentions_dict = {
                            str(row[0].value).lower().strip(): str(row[1].value) 
                            for row in ws_mentions.iter_rows(min_row=2) 
                            if row[0].value and len(row) >= 2 and row[1].value
                        }
                        
                        st.info(f"📊 Mapeo de menciones cargado: {len(mentions_dict)} registros")

                    status.write("🗺️ Aplicando mapeo de Internet y Regiones...")
                    headers = [cell.value for cell in ws_main[1]]
                    
                    # Obtener índices de columnas
                    medio_idx = headers.index("Medio")
                    tipo_medio_idx = headers.index("Tipo de Medio")
                    
                    # Agregar columna Región si no existe
                    if "Región" not in headers:
                        seccion_idx = headers.index("Sección - Programa")
                        insert_col_idx = seccion_idx + 2 
                        ws_main.insert_cols(insert_col_idx)
                        ws_main.cell(row=1, column=insert_col_idx, value="Región")
                        region_idx = insert_col_idx - 1
                        headers = [cell.value for cell in ws_main[1]]  # Actualizar headers
                    else:
                        region_idx = headers.index("Región")
                    
                    # Obtener índice de menciones
                    menciones_idx = headers.index("Menciones - Empresa") if "Menciones - Empresa" in headers else None
                    
                    status.write("🔄 Procesando filas del archivo principal...")
                    processed_count = 0
                    mapping_stats = {
                        'internet_mapped': 0,
                        'regions_mapped': 0,
                        'mentions_mapped': 0
                    }
                    
                    for row in ws_main.iter_rows(min_row=2):
                        processed_count += 1
                        
                        # Mapeo de Internet
                        if str(row[tipo_medio_idx].value).lower().strip() == 'internet':
                            medio_val = str(row[medio_idx].value).lower().strip()
                            if nuevo_medio := internet_dict.get(medio_val):
                                row[medio_idx].value = nuevo_medio
                                mapping_stats['internet_mapped'] += 1
                        
                        # Mapeo de Regiones
                        medio_actual_val = str(row[medio_idx].value).lower().strip()
                        if nueva_region := region_dict.get(medio_actual_val):
                            row[region_idx].value = nueva_region
                            mapping_stats['regions_mapped'] += 1
                        else:
                            row[region_idx].value = "No Asignada"
                        
                        # Mapeo de Menciones - Empresa
                        if menciones_idx is not None and mentions_dict and enable_mentions_mapping:
                            menciones_value = row[menciones_idx].value
                            if menciones_value:
                                menciones_str = str(menciones_value)
                                # Dividir por punto y coma si hay múltiples menciones
                                menciones_list = [m.strip() for m in menciones_str.split(';') if m.strip()]
                                menciones_mapped = []
                                
                                for mencion in menciones_list:
                                    mencion_lower = mencion.lower().strip()
                                    # Buscar coincidencia exacta primero
                                    if mencion_lower in mentions_dict:
                                        menciones_mapped.append(mentions_dict[mencion_lower])
                                        mapping_stats['mentions_mapped'] += 1
                                    else:
                                        # Buscar coincidencia parcial
                                        found_partial = False
                                        for original_key, mapped_value in mentions_dict.items():
                                            if original_key in mencion_lower or mencion_lower in original_key:
                                                menciones_mapped.append(mapped_value)
                                                mapping_stats['mentions_mapped'] += 1
                                                found_partial = True
                                                break
                                        
                                        if not found_partial:
                                            menciones_mapped.append(mencion)  # Mantener original si no hay mapeo
                                
                                # Actualizar la celda con las menciones mapeadas
                                row[menciones_idx].value = '; '.join(menciones_mapped)
                    
                    status.write("🧠 Iniciando detección inteligente de duplicados...")
                    final_wb, summary = run_deduplication_process(wb_main, mentions_dict if enable_mentions_mapping else {})
                    
                    status.update(label="✅ ¡Análisis completado!", state="complete", expanded=False)

                    # Mostrar estadísticas detalladas
                    st.subheader("📊 Resumen del Proceso")
                    
                    col1, col2, col3, col4 = st.columns(4)
                    col1.metric("Filas Procesadas", processed_count)
                    col2.metric("👍 Para Conservar", summary['to_conserve'])
                    col3.metric("🗑️ Para Eliminar", summary['to_eliminate'])
                    col4.metric("💾 Total Final", summary['total_rows'])
                    
                    # Estadísticas de mapeo
                    st.subheader("🔄 Estadísticas de Mapeo")
                    col1, col2, col3 = st.columns(3)
                    col1.metric("🌐 Medios Internet", mapping_stats['internet_mapped'])
                    col2.metric("📍 Regiones Asignadas", mapping_stats['regions_mapped'])
                    col3.metric("🏢 Menciones Mapeadas", mapping_stats['mentions_mapped'])
                    
                    with st.expander("📋 Ver detalles de duplicados"):
                         st.write(f"**Duplicados exactos identificados:** {summary['exact_duplicates']}")
                         st.write(f"**Posibles duplicados identificados:** {summary['possible_duplicates']}")
                         if enable_mentions_mapping and mentions_dict:
                             st.write(f"**Registros de mapeo de menciones disponibles:** {len(mentions_dict)}")

                    # Generar archivo de descarga
                    stream = io.BytesIO()
                    final_wb.save(stream)
                    stream.seek(0)
                    output_filename = f"Informe_Depurado_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
                    
                    st.download_button(
                        label="📥 Descargar Informe Final Depurado",
                        data=stream,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
                except Exception as e:
                    status.update(label="❌ Error en el proceso", state="error", expanded=True)
                    st.error(f"Ha ocurrido un error inesperado: {e}")
                    st.exception(e)
        else:
            missing_files = [name for file, name in zip(required_files, required_names) if not file]
            st.warning(f"⚠️ Faltan archivos requeridos: {', '.join(missing_files)}")
            st.info("Los primeros 3 archivos son obligatorios. El mapeo de menciones es opcional.")
    else:
        st.info("Carga los archivos en el menú de la izquierda y haz clic en 'Analizar y Depurar' para comenzar.")
        
        # Mostrar información de ayuda
        with st.expander("ℹ️ Información sobre los archivos de mapeo"):
            st.markdown("""
            ### Estructura de archivos requerida:
            
            **1. Mapeo de Medios de Internet**
            - Columna A: `Medio` (nombre actual)
            - Columna B: `Internet` (nombre nuevo)
            
            **2. Mapeo de Regiones**
            - Columna A: `Medio` (nombre del medio)
            - Columna B: `Región` (región asignada)
            
            **3. Mapeo de Menciones - Empresa (Opcional)**
            - Columna A: `Mencion_Original` (nombre actual de la empresa)
            - Columna B: `Mencion_Normalizada` (nombre estandarizado)
            
            ### Notas importantes:
            - Los archivos deben tener encabezados en la primera fila
            - Las coincidencias se buscan ignorando mayúsculas/minúsculas
            - Para menciones, se buscan coincidencias exactas y parciales
            - Múltiples menciones se separan con punto y coma (;)
            """)
