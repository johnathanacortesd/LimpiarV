# app.py (versión final con contraseña)

import streamlit as st
import pandas as pd
import openpyxl
import io
import datetime
from deduplicator import run_deduplication_process, norm_key

st.set_page_config(page_title="Depurador de Noticias", layout="wide")

# --- LÓGICA DE AUTENTICACIÓN ---
def check_password():
    """Returns `True` if the user had a correct password."""

    def password_entered():
        """Checks whether a password entered by the user is correct."""
        if st.session_state["password"] == st.secrets["password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # No almacenar la contraseña en texto plano
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # Primera ejecución, mostrar formulario de contraseña.
        st.text_input(
            "Contraseña", type="password", on_change=password_entered, key="password"
        )
        return False
    elif not st.session_state["password_correct"]:
        # Contraseña incorrecta, mostrar formulario de nuevo con mensaje de error.
        st.text_input(
            "Contraseña", type="password", on_change=password_entered, key="password"
        )
        st.error("😕 Contraseña incorrecta. Por favor, inténtalo de nuevo.")
        return False
    else:
        # Contraseña correcta.
        return True

# --- CONFIGURACIÓN DE LA CONTRASEÑA EN STREAMLIT ---
# Para que esto funcione, debes configurar un "Secret" en Streamlit Community Cloud.
# 1. Ve a la configuración de tu app en Streamlit (el pequeño engranaje).
# 2. Ve a la sección "Secrets".
# 3. Pega esto en el cuadro de texto y guarda:
#
# [password]
# password = "TU_CONTRASEÑA_AQUI"
#
# Reemplaza "TU_CONTRASEÑA_AQUI" con la clave que desees.

# --- FLUJO PRINCIPAL DE LA APLICACIÓN ---
if check_password():

    st.title("🚀 Depurador y Mapeador de Informes de Noticias")
    st.write("""
        Esta herramienta automatiza la limpieza de informes de noticias. Sube tus archivos para:
        1.  **Mapear Nombres de Internet**: Actualiza los nombres de los medios de tipo 'Internet'.
        2.  **Asignar Regiones**: Añade o actualiza la columna 'Región' basada en el medio.
        3.  **Detectar y Marcar Duplicados**: Aplica una lógica avanzada para encontrar noticias duplicadas.
    """)

    # --- Sección de Carga de Archivos en la Barra Lateral ---
    st.sidebar.header("📂 Carga tus Archivos")

    uploaded_main_file = st.sidebar.file_uploader(
        "1. Sube el Informe Principal de Noticias (.xlsx)",
        type="xlsx"
    )
    uploaded_internet_map = st.sidebar.file_uploader(
        "2. Sube el Mapeo de Medios de Internet (.xlsx)",
        type="xlsx"
    )
    uploaded_region_map = st.sidebar.file_uploader(
        "3. Sube el Mapeo de Regiones (.xlsx)",
        type="xlsx"
    )

    if st.sidebar.button("✨ Procesar Archivos"):
        if uploaded_main_file and uploaded_internet_map and uploaded_region_map:
            with st.spinner("Procesando... Este proceso puede tardar unos momentos."):
                try:
                    # El resto del código de procesamiento es idéntico
                    wb_main = openpyxl.load_workbook(uploaded_main_file)
                    ws_main = wb_main.active
                    wb_internet = openpyxl.load_workbook(uploaded_internet_map, data_only=True)
                    ws_internet = wb_internet.active
                    wb_region = openpyxl.load_workbook(uploaded_region_map, data_only=True)
                    ws_region = wb_region.active

                    internet_dict = {str(row[0].value).lower().strip(): str(row[1].value) for row in ws_internet.iter_rows(min_row=2) if row[0].value}
                    region_dict = {str(row[0].value).lower().strip(): str(row[1].value) for row in ws_region.iter_rows(min_row=2) if row[0].value}

                    headers = [cell.value for cell in ws_main[1]]
                    
                    try:
                        medio_idx = headers.index("Medio")
                        tipo_medio_idx = headers.index("Tipo de Medio")
                    except ValueError as e:
                        st.error(f"Error: La columna '{e.args[0].split(' ')[0]}' no se encontró en el archivo principal. Revisa las cabeceras.")
                        st.stop()

                    if "Región" not in headers:
                        try:
                            seccion_idx = headers.index("Sección - Programa")
                            insert_col_idx = seccion_idx + 2 
                            ws_main.insert_cols(insert_col_idx)
                            ws_main.cell(row=1, column=insert_col_idx, value="Región")
                            region_idx = insert_col_idx -1
                            st.info("Columna 'Región' creada automáticamente.")
                        except ValueError:
                            insert_col_idx = len(headers) + 1
                            ws_main.cell(row=1, column=insert_col_idx, value="Región")
                            region_idx = insert_col_idx - 1
                            st.warning("Columna 'Sección - Programa' no encontrada. 'Región' se añadió al final.")
                    else:
                        region_idx = headers.index("Región")
                    
                    for row in ws_main.iter_rows(min_row=2):
                        tipo_medio_val = str(row[tipo_medio_idx].value).lower().strip()
                        if tipo_medio_val == 'internet':
                            medio_val = str(row[medio_idx].value).lower().strip()
                            nuevo_medio = internet_dict.get(medio_val)
                            if nuevo_medio: row[medio_idx].value = nuevo_medio
                        
                        medio_actual_val = str(row[medio_idx].value).lower().strip()
                        nueva_region = region_dict.get(medio_actual_val, "No Asignada")
                        row[region_idx].value = nueva_region

                    st.info("✅ Mapeo completado. Iniciando deduplicación...")

                    final_wb, summary = run_deduplication_process(wb_main)
                    
                    st.success("🎉 ¡Procesamiento completado!")

                    st.subheader("📊 Resumen del Proceso")
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Filas Totales", summary['total_rows'])
                    col2.metric("👍 Filas para Conservar", summary['to_conserve'], delta=f"-{summary['to_eliminate']} eliminadas", delta_color="inverse")
                    col3.metric("🗑️ Filas para Eliminar", summary['to_eliminate'], delta_color="inverse")
                    st.write(f"**Duplicados exactos:** {summary['exact_duplicates']}")
                    st.write(f"**Posibles duplicados:** {summary['possible_duplicates']}")

                    stream = io.BytesIO()
                    final_wb.save(stream)
                    stream.seek(0)
                    output_filename = f"Informe_Depurado_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

                    st.download_button(
                        label="📥 Descargar Archivo Final",
                        data=stream,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                except Exception as e:
                    st.error(f"Ha ocurrido un error: {e}")
                    st.exception(e)
                    st.error("Por favor, verifica el formato y las columnas de los archivos.")

        else:
            st.warning("⚠️ Por favor, sube los tres archivos requeridos.")
