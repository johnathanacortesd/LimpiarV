# app.py

import streamlit as st
import openpyxl
import io
import datetime
from deduplicator import run_deduplication_process

# --- Configuración y Autenticación (sin cambios) ---
st.set_page_config(page_title="Intelli-Clean | Depurador IA", page_icon="🤖", layout="wide", initial_sidebar_state="expanded")
def check_password():
    def password_entered():
        try:
            if st.session_state["password"] == st.secrets.password.password:
                st.session_state["password_correct"] = True; del st.session_state["password"]
            else: st.session_state["password_correct"] = False
        except (AttributeError, KeyError): st.session_state["password_correct"] = False
    try: _ = st.secrets.password.password
    except (AttributeError, KeyError):
        st.error("🚨 ¡Error de configuración! Contraseña no definida en 'Secrets'."); return False
    if "password_correct" not in st.session_state: st.session_state["password_correct"] = False
    if not st.session_state["password_correct"]:
        c1, c2, c3 = st.columns([1, 1, 1]);
        with c2:
            st.markdown("<h1 style='text-align: center;'>🤖</h1>", unsafe_allow_html=True)
            st.markdown("<h3 style='text-align: center;'>Intelli-Clean Access</h3>", unsafe_allow_html=True)
            st.text_input("Contraseña", type="password", on_change=password_entered, key="password", placeholder="Introduce la contraseña", label_visibility="collapsed")
            if 'password' in st.session_state and st.session_state.password != "" and not st.session_state.password_correct:
                 st.error("😕 Contraseña incorrecta.")
        return False
    return True

# --- FLUJO PRINCIPAL DE LA APLICACIÓN ---
if check_password():
    st.title("✨ Intelli-Clean: Depurador de Noticias")
    st.caption("Una herramienta inteligente para mapear, limpiar y deduplicar tus informes con precisión.")
    st.divider()

    with st.sidebar:
        st.header("📂 Carga tus Archivos")
        uploaded_main_file = st.file_uploader("1. Informe Principal de Noticias", type="xlsx", key="main_file")
        uploaded_internet_map = st.file_uploader("2. Mapeo de Medios de Internet", type="xlsx", key="internet_map")
        uploaded_region_map = st.file_uploader("3. Mapeo de Regiones", type="xlsx", key="region_map")
        uploaded_empresa_map = st.file_uploader("4. Mapeo de Nombres de Empresas", type="xlsx", key="empresa_map")
        st.divider()
        process_button = st.button("🚀 Analizar y Depurar Archivos", type="primary", use_container_width=True)

    # <<< --- INICIO DE LA LÓGICA CON SESSION STATE --- >>>

    # Inicializamos las variables en session_state si no existen
    if 'processing_complete' not in st.session_state:
        st.session_state.processing_complete = False
    if 'summary' not in st.session_state:
        st.session_state.summary = {}
    if 'main_stream' not in st.session_state:
        st.session_state.main_stream = None
    if 'nissan_stream' not in st.session_state:
        st.session_state.nissan_stream = None

    all_files_uploaded = (uploaded_main_file and uploaded_internet_map and 
                          uploaded_region_map and uploaded_empresa_map)

    # El bloque `if process_button` ahora solo se encarga de INICIAR el proceso y GUARDAR los resultados.
    if process_button:
        if all_files_uploaded:
            with st.status("Iniciando proceso... ⏳", expanded=True) as status:
                try:
                    status.write("Cargando archivos y creando diccionarios de mapeo...")
                    wb_main = openpyxl.load_workbook(uploaded_main_file)
                    internet_dict = {str(r[0].value).lower().strip(): str(r[1].value) for r in openpyxl.load_workbook(uploaded_internet_map, data_only=True).active.iter_rows(min_row=2) if r[0].value and r[1].value}
                    region_dict = {str(r[0].value).lower().strip(): str(r[1].value) for r in openpyxl.load_workbook(uploaded_region_map, data_only=True).active.iter_rows(min_row=2) if r[0].value and r[1].value}
                    empresa_dict = {str(r[0].value).lower().strip(): str(r[1].value) for r in openpyxl.load_workbook(uploaded_empresa_map, data_only=True).active.iter_rows(min_row=2) if r[0].value and r[1].value}

                    status.write("🧠 Iniciando proceso de expansión, mapeo y deduplicación...")
                    final_wb, nissan_wb, summary = run_deduplication_process(wb_main, empresa_dict, internet_dict, region_dict)
                    
                    status.update(label="✅ ¡Análisis completado!", state="complete", expanded=False)

                    # Guardamos el resumen en session_state
                    st.session_state.summary = summary

                    # Convertimos los workbooks a streams y los guardamos en session_state
                    main_stream = io.BytesIO()
                    final_wb.save(main_stream)
                    st.session_state.main_stream = main_stream

                    nissan_stream = io.BytesIO()
                    nissan_wb.save(nissan_stream)
                    st.session_state.nissan_stream = nissan_stream
                    
                    # Activamos la bandera para indicar que el proceso terminó con éxito
                    st.session_state.processing_complete = True

                except Exception as e:
                    status.update(label="❌ Error en el proceso", state="error", expanded=True)
                    st.error(f"Ha ocurrido un error inesperado: {e}")
                    st.exception(e)
                    # Nos aseguramos de que no se muestren resultados si hay un error
                    st.session_state.processing_complete = False
        else:
            st.warning("⚠️ Por favor, asegúrate de cargar los cuatro archivos requeridos en la barra lateral.")

    # Este bloque ahora se encarga de MOSTRAR los resultados si la bandera está activa.
    # Se ejecutará después del procesamiento y también en cada recarga (como al descargar un archivo).
    if st.session_state.processing_complete:
        st.header("Resultados del Análisis")
        st.subheader("📊 Resumen del Proceso")
        
        summary = st.session_state.summary
        col1, col2, col3 = st.columns(3)
        col1.metric("Filas Totales Procesadas", summary.get('total_rows', 0))
        col2.metric("👍 Filas para Conservar", summary.get('to_conserve', 0))
        col3.metric("🗑️ Filas para Eliminar", summary.get('to_eliminate', 0))
        
        with st.expander("Ver detalles de duplicados"):
             st.write(f"**Duplicados exactos:** {summary.get('exact_duplicates', 0)}")
             st.write(f"**Posibles duplicados:** {summary.get('possible_duplicates', 0)}")

        st.divider()
        st.subheader("📥 Archivos para Descargar")
        
        # Botón de descarga para el informe principal
        main_filename = f"Informe_Depurado_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        st.download_button(
            label="1. Descargar Informe Principal Depurado", 
            data=st.session_state.main_stream, 
            file_name=main_filename, 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
            use_container_width=True
        )
    
        # Botón de descarga para el informe Nissan Test
        nissan_filename = "nissan_test.xlsx"
        st.download_button(
            label="2. Descargar Reporte 'Nissan Test' (Resumen)",
            data=st.session_state.nissan_stream,
            file_name=nissan_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    elif not process_button:
        # Mensaje inicial si no se ha procesado nada todavía
        st.info("Carga los archivos en el menú de la izquierda y haz clic en 'Analizar y Depurar' para comenzar.")

    # <<< --- FIN DE LA LÓGICA CON SESSION STATE --- >>>
