# app.py

import streamlit as st
import openpyxl
import io
import datetime
from deduplicator import run_deduplication_process

# --- Configuración y Autenticación ---
st.set_page_config(page_title="Intelli-Clean | Depurador IA", page_icon="🤖", layout="wide")

def check_password():
    """Retorna True si la contraseña es correcta."""
    def password_entered():
        """Valida la contraseña ingresada."""
        session_password = st.session_state.get("password", "")
        correct_password = st.secrets.get("password", {}).get("password")
        
        if correct_password and session_password == correct_password:
            st.session_state["password_correct"] = True
            del st.session_state["password"] # Eliminar contraseña de la sesión por seguridad
        else:
            st.session_state["password_correct"] = False

    # Verificar que la contraseña esté configurada en st.secrets
    if not st.secrets.get("password", {}).get("password"):
        st.error("🚨 ¡Error de configuración! La contraseña no está definida en los 'Secrets' de Streamlit.")
        return False

    if st.session_state.get("password_correct", False):
        return True

    # Mostrar formulario de login si no está autenticado
    c1, c2, c3 = st.columns([1, 1, 1])
    with c2:
        st.markdown("<h1 style='text-align: center;'>🤖</h1>", unsafe_allow_html=True)
        st.markdown("<h3 style='text-align: center;'>Intelli-Clean Access</h3>", unsafe_allow_html=True)
        st.text_input(
            "Contraseña", 
            type="password", 
            on_change=password_entered, 
            key="password", 
            placeholder="Introduce la contraseña", 
            label_visibility="collapsed"
        )
        if "password_correct" in st.session_state and not st.session_state.password_correct:
             st.error("😕 Contraseña incorrecta. Inténtalo de nuevo.")
    return False

def load_mapping_dict(uploaded_file):
    """Carga un archivo Excel de mapeo y lo convierte en un diccionario robusto."""
    if not uploaded_file:
        return {}
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    sheet = wb.active
    # Se asegura de que ni la clave ni el valor sean nulos antes de agregarlos.
    return {
        str(row[0].value).lower().strip(): str(row[1].value) 
        for row in sheet.iter_rows(min_row=2) 
        if row[0].value is not None and row[1].value is not None
    }

# --- FLUJO PRINCIPAL DE LA APLICACIÓN ---
if check_password():
    st.title("✨ Intelli-Clean: Depurador de Noticias")
    st.caption("Una herramienta inteligente para mapear, limpiar y deduplicar tus informes con precisión.")
    st.divider()

    with st.sidebar:
        st.header("📂 Carga tus Archivos")
        uploaded_main_file = st.file_uploader("1. Informe Principal de Noticias", type="xlsx")
        uploaded_internet_map = st.file_uploader("2. Mapeo de Medios de Internet", type="xlsx")
        uploaded_region_map = st.file_uploader("3. Mapeo de Regiones", type="xlsx")
        uploaded_empresa_map = st.file_uploader("4. Mapeo de Nombres de Empresas", type="xlsx")
        st.divider()
        process_button = st.button("🚀 Analizar y Depurar Archivos", type="primary", use_container_width=True)

    st.header("Resultados del Análisis")
    
    if process_button:
        all_files_uploaded = (uploaded_main_file and uploaded_internet_map and 
                              uploaded_region_map and uploaded_empresa_map)
        
        if all_files_uploaded:
            with st.status("Procesando archivos... ⏳", expanded=True) as status:
                try:
                    status.write("Cargando archivo principal...")
                    wb_main = openpyxl.load_workbook(uploaded_main_file)

                    status.write("Cargando y creando diccionarios de mapeo...")
                    internet_dict = load_mapping_dict(uploaded_internet_map)
                    region_dict = load_mapping_dict(uploaded_region_map)
                    empresa_dict = load_mapping_dict(uploaded_empresa_map)
                    
                    # El código que insertaba la columna 'Región' se ha eliminado.
                    # El script `deduplicator.py` ahora maneja toda la estructura del archivo final.
                    # Esto hace que el proceso sea más simple y robusto.

                    status.write("🧠 Iniciando proceso de expansión, mapeo y deduplicación...")
                    final_wb, summary = run_deduplication_process(
                        wb_main, empresa_dict, internet_dict, region_dict
                    )
                    
                    status.update(label="✅ ¡Análisis completado!", state="complete", expanded=False)
                    
                    st.subheader("📊 Resumen del Proceso")
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Filas Totales Procesadas", summary['total_rows'])
                    col2.metric("👍 Filas para Conservar", summary['to_conserve'])
                    col3.metric("🗑️ Filas para Eliminar", summary['to_eliminate'])
                    
                    with st.expander("Ver detalles de duplicados detectados"):
                         st.write(f"**Duplicados exactos:** {summary['exact_duplicates']}")
                         st.write(f"**Posibles duplicados (por similitud):** {summary['possible_duplicates']}")

                    # Preparar archivo para descarga
                    stream = io.BytesIO()
                    final_wb.save(stream)
                    stream.seek(0)
                    output_filename = f"Informe_Depurado_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
                    
                    st.download_button(
                        label="📥 Descargar Informe Final (Ordenado y Limpio)", 
                        data=stream, 
                        file_name=output_filename, 
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                        use_container_width=True
                    )
                
                except Exception as e:
                    status.update(label="❌ Error en el proceso", state="error", expanded=True)
                    st.error(f"Ha ocurrido un error inesperado durante el procesamiento.")
                    st.exception(e) # Muestra el traceback completo para depuración
        else:
            st.warning("⚠️ Por favor, asegúrate de cargar los cuatro archivos requeridos en la barra lateral.")
    else:
        st.info("Carga los archivos en el menú de la izquierda y haz clic en 'Analizar y Depurar' para comenzar.")
