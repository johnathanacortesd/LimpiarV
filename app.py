# app.py

import streamlit as st
import openpyxl
import io
import datetime
import re
from deduplicator import run_deduplication_process

# --- Funciones Auxiliares para el Mapeo en app.py ---
def get_url_from_cell(cell):
    if cell.hyperlink: return cell.hyperlink.target
    if cell.value and isinstance(cell.value, str):
        if match := re.search(r'=HYPERLINK\("([^"]+)"', cell.value): return match.group(1)
    return None
def extract_root_domain(url):
    if not url: return None
    try:
        cleaned_url = re.sub(r'^https?://', '', url).lower().replace('www.', ''); domain = cleaned_url.split('/')[0]
        return domain.capitalize()
    except Exception: return None

# --- Configuración y Autenticación ---
st.set_page_config(page_title="Intelli-Clean | Depurador IA", page_icon="🤖", layout="wide", initial_sidebar_state="expanded")
def check_password():
    # ... (código de contraseña idéntico y funcional) ...
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
        uploaded_main_file = st.file_uploader("1. Informe Principal de Noticias", type="xlsx")
        uploaded_internet_map = st.file_uploader("2. Mapeo de Medios de Internet", type="xlsx")
        uploaded_region_map = st.file_uploader("3. Mapeo de Regiones", type="xlsx")
        uploaded_empresa_map = st.file_uploader("4. Mapeo de Nombres de Empresas", type="xlsx")
        st.divider()
        process_button = st.button("🚀 Analizar y Depurar Archivos", type="primary", use_container_width=True)

    st.header("Resultados del Análisis")
    
    if process_button:
        if uploaded_main_file and uploaded_internet_map and uploaded_region_map and uploaded_empresa_map:
            with st.status("Iniciando proceso... ⏳", expanded=True) as status:
                try:
                    status.write("Cargando archivos y creando diccionarios de mapeo...")
                    wb_main = openpyxl.load_workbook(uploaded_main_file)
                    internet_dict = {str(r[0].value).lower().strip(): str(r[1].value) for r in openpyxl.load_workbook(uploaded_internet_map, data_only=True).active.iter_rows(min_row=2) if r[0].value is not None}
                    region_dict = {str(r[0].value).lower().strip(): str(r[1].value) for r in openpyxl.load_workbook(uploaded_region_map, data_only=True).active.iter_rows(min_row=2) if r[0].value is not None}
                    empresa_dict = {str(r[0].value).lower().strip(): str(r[1].value) for r in openpyxl.load_workbook(uploaded_empresa_map, data_only=True).active.iter_rows(min_row=2) if r[0].value is not None}

                    status.write("🗺️ Aplicando mapeos previos (Internet y Región)...")
                    ws_main = wb_main.active
                    headers = [cell.value for cell in ws_main[1]]
                    
                    try:
                        medio_idx = headers.index("Medio"); tipo_medio_idx = headers.index("Tipo de Medio"); link_nota_idx = headers.index("Link Nota")
                    except ValueError as e:
                        st.error(f"Error Crítico: La columna '{e.args[0].split(' ')[0]}' no se encontró."); st.stop()
                    
                    if "Región" not in headers:
                        try: seccion_idx = headers.index("Sección - Programa"); insert_col_idx = seccion_idx + 2
                        except ValueError: insert_col_idx = len(headers) + 1
                        ws_main.insert_cols(insert_col_idx); ws_main.cell(row=1, column=insert_col_idx, value="Región")
                        headers = [cell.value for cell in ws_main[1]] # Refrescar encabezados
                    region_idx = headers.index("Región")

                    # Aplicar mapeos de Internet y Región ANTES de la deduplicación
                    for row in ws_main.iter_rows(min_row=2):
                        # Lógica de Internet (CORREGIDA según tus instrucciones)
                        if str(row[tipo_medio_idx].value).lower().strip() == 'internet':
                            medio_val = str(row[medio_idx].value).lower().strip()
                            # Prioridad 1: Buscar en el archivo de mapeo.
                            if medio_val in internet_dict:
                                row[medio_idx].value = internet_dict[medio_val]
                            else:
                                # Prioridad 2 (Plan B): Si NO se encuentra, extraer de la URL.
                                if root_domain := extract_root_domain(get_url_from_cell(row[link_nota_idx])):
                                    row[medio_idx].value = root_domain
                        
                        # Lógica de Región
                        medio_actual_val = str(row[medio_idx].value).lower().strip()
                        row[region_idx].value = region_dict.get(medio_actual_val, "Online")

                    status.write("🧠 Iniciando proceso de expansión y deduplicación...")
                    final_wb, summary = run_deduplication_process(wb_main, empresa_dict)
                    
                    status.update(label="✅ ¡Análisis completado!", state="complete", expanded=False)
                    st.subheader("📊 Resumen del Proceso")
                    col1, col2, col3 = st.columns(3); col1.metric("Filas Totales", summary['total_rows'])
                    col2.metric("👍 Filas para Conservar", summary['to_conserve'])
                    col3.metric("🗑️ Filas para Eliminar", summary['to_eliminate'])
                    with st.expander("Ver detalles de duplicados"):
                         st.write(f"**Duplicados exactos:** {summary['exact_duplicates']}")
                         st.write(f"**Posibles duplicados:** {summary['possible_duplicates']}")

                    stream = io.BytesIO(); final_wb.save(stream); stream.seek(0)
                    output_filename = f"Informe_Depurado_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
                    st.download_button("📥 Descargar Informe Final Depurado", stream, output_filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                
                except Exception as e:
                    status.update(label="❌ Error en el proceso", state="error", expanded=True)
                    st.error(f"Ha ocurrido un error inesperado: {e}"); st.exception(e)
        else:
            st.warning("⚠️ Por favor, asegúrate de cargar los cuatro archivos requeridos en la barra lateral.")
    else:
        st.info("Carga los archivos en el menú de la izquierda y haz clic en 'Analizar y Depurar' para comenzar.")
