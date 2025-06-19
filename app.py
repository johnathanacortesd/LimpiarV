# app.py

import streamlit as st
import pandas as pd
import openpyxl
import io
import datetime
from deduplicator import run_deduplication_process, norm_key

st.set_page_config(page_title="Depurador de Noticias", layout="wide")

st.title("🚀 Depurador y Mapeador de Informes de Noticias")
st.write("""
    Esta herramienta automatiza la limpieza de informes de noticias. Sube tus archivos para:
    1.  **Mapear Nombres de Internet**: Actualiza los nombres de los medios de tipo 'Internet'.
    2.  **Asignar Regiones**: Añade la columna 'Región' basada en el medio.
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
                # --- Paso 1: Cargar todos los libros de trabajo ---
                wb_main = openpyxl.load_workbook(uploaded_main_file)
                ws_main = wb_main.active

                wb_internet = openpyxl.load_workbook(uploaded_internet_map, data_only=True)
                ws_internet = wb_internet.active

                wb_region = openpyxl.load_workbook(uploaded_region_map, data_only=True)
                ws_region = wb_region.active

                # --- Paso 2: Crear diccionarios de mapeo ---
                internet_dict = {
                    str(row[0].value).lower().strip(): str(row[1].value)
                    for row in ws_internet.iter_rows(min_row=2) if row[0].value
                }
                region_dict = {
                    str(row[0].value).lower().strip(): str(row[1].value)
                    for row in ws_region.iter_rows(min_row=2) if row[0].value
                }

                # --- Paso 3: Aplicar los Mapeos directamente con Openpyxl para preservar todo ---
                headers = [cell.value for cell in ws_main[1]]
                # Encontrar los índices de las columnas importantes
                try:
                    medio_idx = headers.index("Medio")
                    tipo_medio_idx = headers.index("Tipo de Medio")
                    region_idx = headers.index("Región")
                except ValueError as e:
                    st.error(f"Error: La columna '{e.args[0].split(' ')[0]}' no se encontró en el archivo principal. Por favor, revisa las cabeceras.")
                    st.stop()


                for row in ws_main.iter_rows(min_row=2):
                    # Mapeo de Internet
                    tipo_medio_val = str(row[tipo_medio_idx].value).lower().strip()
                    if tipo_medio_val == 'internet':
                        medio_val = str(row[medio_idx].value).lower().strip()
                        nuevo_medio = internet_dict.get(medio_val)
                        if nuevo_medio:
                            row[medio_idx].value = nuevo_medio
                    
                    # Mapeo de Región (se usa el valor ya actualizado del medio)
                    medio_actual_val = str(row[medio_idx].value).lower().strip()
                    nueva_region = region_dict.get(medio_actual_val, "Error") # Default a "Error" si no se encuentra
                    row[region_idx].value = nueva_region

                st.info("✅ Mapeo de Internet y Regiones completado.")
                st.info("⚙️ Iniciando la detección de duplicados...")

                # --- Paso 4: Ejecutar el proceso de deduplicación ---
                final_wb, summary = run_deduplication_process(wb_main)
                
                st.success("🎉 ¡Procesamiento completado con éxito!")

                # --- Paso 5: Mostrar Resumen y Ofrecer Descarga ---
                st.subheader("📊 Resumen del Proceso")
                col1, col2, col3 = st.columns(3)
                col1.metric("Filas Totales Procesadas", summary['total_rows'])
                col2.metric("👍 Filas para Conservar", summary['to_conserve'], delta=f"-{summary['to_eliminate']} eliminadas", delta_color="inverse")
                col3.metric("🗑️ Filas para Eliminar", summary['to_eliminate'], delta_color="inverse")
                
                st.write(f"**Duplicados exactos identificados:** {summary['exact_duplicates']}")
                st.write(f"**Posibles duplicados identificados:** {summary['possible_duplicates']}")

                # Crear el archivo en memoria para la descarga
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
                st.error(f"Ha ocurrido un error inesperado durante el procesamiento: {e}")
                st.error("Por favor, verifica que los archivos tengan el formato y las columnas correctas.")

    else:
        st.warning("⚠️ Por favor, sube los tres archivos requeridos para iniciar el procesamiento.")
