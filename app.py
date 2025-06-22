# app.py (versión final con mapeo de empresas y fallback de URL)

import streamlit as st
import openpyxl
import io
import datetime
import re
from deduplicator import run_deduplication_process

# --- Configuración de la Página ---
st.set_page_config(
    page_title="Intelli-Clean | Depurador de Noticias IA",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Funciones Auxiliares para el Mapeo ---
def get_url_from_cell(cell):
    """Extrae una URL de una celda, ya sea de un hipervínculo o de una fórmula HYPERLINK."""
    if cell.hyperlink:
        return cell.hyperlink.target
    if cell.value and isinstance(cell.value, str):
        match = re.search(r'=HYPERLINK\("([^"]+)"', cell.value)
        if match:
            return match.group(1)
    return None

def extract_root_domain(url):
    """Extrae y formatea el dominio raíz de una URL."""
    if not url:
        return None
    try:
        # Limpia el prefijo http/https y www.
        cleaned_url = re.sub(r'^https?://', '', url).lower()
        cleaned_url = cleaned_url.replace('www.', '')
        # Obtiene la parte antes de la primera barra
        domain = cleaned_url.split('/')[0]
        # Capitaliza la primera letra
        return domain.capitalize()
    except Exception:
        return None

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
    
    # (El resto de la función de contraseña no cambia...)
    try: _ = st.secrets.password.password
    except (AttributeError, KeyError):
        st.error("🚨 ¡Error de configuración! Contraseña no definida en 'Secrets'.")
        return False
    if "password_correct" not in st.session_state: st.session_state["password_correct"] = False
    if not st.session_state["password_correct"]:
        c1, c2, c3 = st.columns([1, 1, 1])
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
                    status.write("Cargando archivos en memoria...")
                    wb_main = openpyxl.load_workbook(uploaded_main_file)
                    ws_main = wb_main.active
                    internet_dict = {str(r[0].value).lower().strip(): str(r[1].value) for r in openpyxl.load_workbook(uploaded_internet_map, data_only=True).active.iter_rows(min_row=2) if r[0].value}
                    region_dict = {str(r[0].value).lower().strip(): str(r[1].value) for r in openpyxl.load_workbook(uploaded_region_map, data_only=True).active.iter_rows(min_row=2) if r[0].value}
                    empresa_dict = {str(r[0].value).lower().strip(): str(r[1].value) for r in openpyxl.load_workbook(uploaded_empresa_map, data_only=True).active.iter_rows(min_row=2) if r[0].value}

                    status.write("🗺️ Aplicando mapeos inteligentes...")
                    headers = [cell.value for cell in ws_main[1]]
                    # Validar y obtener índices de todas las columnas necesarias
                    try:
                        medio_idx = headers.index("Medio")
                        tipo_medio_idx = headers.index("Tipo de Medio")
                        empresa_idx = headers.index("Menciones - Empresa")
                        link_nota_idx = headers.index("Link Nota")
                    except ValueError as e:
                        st.error(f"Error Crítico: La columna '{e.args[0].split(' ')[0]}' no se encontró en el archivo principal.")
                        st.stop()

                    if "Región" not in headers:
                        seccion_idx = headers.index("Sección - Programa")
                        insert_col_idx = seccion_idx + 2; ws_main.insert_cols(insert_col_idx)
                        ws_main.cell(row=1, column=insert_col_idx, value="Región")
                        region_idx = insert_col_idx - 1
                    else:
                        region_idx = headers.index("Región")
                    
                    for row in ws_main.iter_rows(min_row=2):
                        # 1. Mapeo de Empresas
                        if row[empresa_idx].value:
                            empresa_val = str(row[empresa_idx].value).lower().strip()
                            if nuevo_nombre := empresa_dict.get(empresa_val): row[empresa_idx].value = nuevo_nombre
                        
                        # 2. Mapeo de Internet con fallback a URL
                        if str(row[tipo_medio_idx].value).lower().strip() == 'internet':
                            medio_val = str(row[medio_idx].value).lower().strip()
                            nuevo_medio = internet_dict.get(medio_val)
                            if nuevo_medio:
                                row[medio_idx].value = nuevo_medio
                            else: # Fallback: si no está en el dict, intentar extraer de la URL
                                url = get_url_from_cell(row[link_nota_idx])
                                if root_domain := extract_root_domain(url):
                                    row[medio_idx].value = root_domain
                        
                        # 3. Mapeo de Región (se usa el valor del medio ya actualizado)
                        medio_actual_val = str(row[medio_idx].value).lower().strip()
                        row[region_idx].value = region_dict.get(medio_actual_val, "Online") # Default "Online"
                    
                    status.write("🧠 Iniciando detección inteligente de duplicados...")
                    final_wb, summary = run_deduplication_process(wb_main)
                    
                    status.update(label="✅ ¡Análisis completado!", state="complete", expanded=False)
                    st.subheader("📊 Resumen del Proceso")
                    col1, col2, col3 = st.columns(3); col1.metric("Filas Totales", summary['total_rows'])
                    col2.metric("👍 Filas para Conservar", summary['to_conserve'])
                    col3.metric("🗑️ Filas para Eliminar", summary['to_eliminate'])
                    with st.expander("Ver detalles de duplicados"):
                         st.write(f"**Duplicados exactos:** {summary['exact_duplicates']}")
                         st.write(f"**Posibles duplicados:** {summary['possible_duplicates']}")

                    stream = io.BytesIO()
                    final_wb.save(stream); stream.seek(0)
                    output_filename = f"Informe_Depurado_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
                    st.download_button("📥 Descargar Informe Final Depurado", stream, output_filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                
                except Exception as e:
                    status.update(label="❌ Error en el proceso", state="error", expanded=True)
                    st.error(f"Ha ocurrido un error inesperado: {e}")
                    st.exception(e)
        else:
            st.warning("⚠️ Por favor, asegúrate de cargar los cuatro archivos requeridos en la barra lateral.")
    else:
        st.info("Carga los archivos en el menú de la izquierda y haz clic en 'Analizar y Depurar' para comenzar.")
