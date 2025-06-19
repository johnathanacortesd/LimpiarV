# app.py (versión con interfaz moderna "IA 2025")

import streamlit as st
import openpyxl
import io
import datetime
from deduplicator import run_deduplication_process # Asumiendo que tu lógica está en este archivo

# --- Configuración de la Página ---
st.set_page_config(
    page_title="Intelli-Clean | Depurador de Noticias IA",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="expanded" # Menú de carga visible por defecto
)

# --- LÓGICA DE AUTENTICACIÓN (sin cambios, pero con un estilo más centrado) ---
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
        c1, c2, c3 = st.columns([1, 2, 1])
        with c2:
            st.image("https://raw.githubusercontent.com/streamlit/brand/main/logo/primary/horizontal/streamlit-logo-primary-lockup-horizontal-colormark-dark-background.png") # Un logo para darle un toque
            st.text_input("Contraseña de Acceso", type="password", on_change=password_entered, key="password")
            if st.session_state.get('password', '') != '': # Para mostrar el error solo después de un intento
                 st.error("😕 Contraseña incorrecta. Por favor, inténtalo de nuevo.")
        return False
    return True


# --- FLUJO PRINCIPAL DE LA APLICACIÓN ---
if check_password():

    # --- Encabezado ---
    st.title("✨ Intelli-Clean: Depurador de Noticias")
    st.caption("Una herramienta inteligente para mapear, limpiar y deduplicar tus informes de noticias con precisión.")
    st.divider()

    # --- Sección de Carga de Archivos (en la barra lateral) ---
    with st.sidebar:
        st.header("📂 Carga tus Archivos")
        uploaded_main_file = st.file_uploader("1. Informe Principal de Noticias", type="xlsx", help="El archivo Excel con todas las noticias a procesar.")
        uploaded_internet_map = st.file_uploader("2. Mapeo de Medios de Internet", type="xlsx", help="Un archivo con dos columnas: 'Medio' (nombre antiguo) e 'Internet' (nombre nuevo).")
        uploaded_region_map = st.file_uploader("3. Mapeo de Regiones", type="xlsx", help="Un archivo con dos columnas: 'Medio' y 'Región'.")
        
        st.divider()
        process_button = st.button("🚀 Analizar y Depurar Archivos", type="primary", use_container_width=True)

    # --- Área de Trabajo Principal ---
    st.header("Resultados del Análisis")
    
    if process_button:
        if uploaded_main_file and uploaded_internet_map and uploaded_region_map:
            # Usamos st.status para un feedback de proceso más detallado y moderno
            with st.status("Iniciando proceso... ⏳", expanded=True) as status:
                try:
                    status.write("Cargando archivos en memoria...")
                    wb_main = openpyxl.load_workbook(uploaded_main_file)
                    ws_main = wb_main.active
                    wb_internet = openpyxl.load_workbook(uploaded_internet_map, data_only=True)
                    ws_internet = wb_internet.active
                    wb_region = openpyxl.load_workbook(uploaded_region_map, data_only=True)
                    ws_region = wb_region.active

                    status.write("Creando diccionarios de mapeo...")
                    internet_dict = {str(row[0].value).lower().strip(): str(row[1].value) for row in ws_internet.iter_rows(min_row=2) if row[0].value}
                    region_dict = {str(row[0].value).lower().strip(): str(row[1].value) for row in ws_region.iter_rows(min_row=2) if row[0].value}

                    status.write("🗺️ Aplicando mapeo de Internet y Regiones...")
                    headers = [cell.value for cell in ws_main[1]]
                    medio_idx = headers.index("Medio"); tipo_medio_idx = headers.index("Tipo de Medio")
                    if "Región" not in headers:
                        seccion_idx = headers.index("Sección - Programa")
                        insert_col_idx = seccion_idx + 2 
                        ws_main.insert_cols(insert_col_idx)
                        ws_main.cell(row=1, column=insert_col_idx, value="Región")
                        region_idx = insert_col_idx - 1
                    else:
                        region_idx = headers.index("Región")
                    
                    for row in ws_main.iter_rows(min_row=2):
                        if str(row[tipo_medio_idx].value).lower().strip() == 'internet':
                            medio_val = str(row[medio_idx].value).lower().strip()
                            if nuevo_medio := internet_dict.get(medio_val): row[medio_idx].value = nuevo_medio
                        medio_actual_val = str(row[medio_idx].value).lower().strip()
                        row[region_idx].value = region_dict.get(medio_actual_val, "No Asignada")
                    
                    status.write("🧠 Iniciando detección inteligente de duplicados...")
                    final_wb, summary = run_deduplication_process(wb_main)
                    
                    status.update(label="✅ ¡Análisis completado!", state="complete", expanded=False)

                    # --- Mostrar Resultados ---
                    st.subheader("📊 Resumen del Proceso")
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Filas Totales Procesadas", summary['total_rows'])
                    col2.metric("👍 Filas para Conservar", summary['to_conserve'])
                    col3.metric("🗑️ Filas para Eliminar", summary['to_eliminate'])
                    
                    with st.expander("Ver detalles de duplicados"):
                         st.write(f"**Duplicados exactos identificados:** {summary['exact_duplicates']}")
                         st.write(f"**Posibles duplicados identificados:** {summary['possible_duplicates']}")

                    # --- Descarga ---
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
            st.warning("⚠️ Por favor, asegúrate de cargar los tres archivos en la barra lateral antes de continuar.")
    else:
        st.info("Carga los archivos en el menú de la izquierda y haz clic en 'Analizar y Depurar' para comenzar.")
