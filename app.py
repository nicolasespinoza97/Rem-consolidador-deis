import streamlit as st
import openpyxl
import pandas as pd
import io
import os
import zipfile
from core.detector import detect_rem_info
from core.dictionary_loader import get_master_template_path
from core.engine import consolidate_rem_files

st.set_page_config(page_title="Consolidador Universal REM DEIS", layout="wide", page_icon="🏥")

st.title("🏥 Consolidador Universal de Planillas REM DEIS")
st.subheader("Procesamiento Inteligente Multi-Serie (SA, SBM, SBS, SD, SP) con Protección Estricta por Color (#FFFFCC / #FFCC99)")

st.markdown("""
Esta aplicación permite cargar, clasificar y consolidar de forma automatizada los archivos mensuales de todas las familias estadísticas REM DEIS (`.xlsm` o `.xlsx`).

### 🛡️ Reglas de Seguridad y Funcionamiento:
1. 🎨 **Filtrado Estricto por Color de Ingreso (#FFFFCC / #FFCC99):** ÚNICAMENTE se modifican las celdas cuyo fondo corresponda a los colores oficiales de ingreso de datos: **amarillo `#FFFFCC`** o **naranjo suave `#FFCC99`** (RGB 255, 204, 153).
2. 🚫 **Protección Absoluta:** Encabezados, títulos, glosas, códigos, fórmulas `=SUM(...)`, totales, subtotales o celdas blancas quedan **100% INTOCADAS y PROTEGIDAS**.
3. 📁 **Auto-Clasificación por Serie:** Detección automática de **SA** (REM A), **SBM** (REM BM), **SBS** (REM BS), **SD** (REM D) y **SP** (REM P).
4. ⚙️ **Preservación de Macros:** Mantiene intactas las macros VBA (`.xlsm`) y la hoja de `NOMBRE` del establecimiento.
""")

uploaded_files = st.file_uploader(
    "Seleccione o arrastre aquí todos los archivos mensuales de cualquier serie REM (.xlsm / .xlsx):",
    type=["xlsm", "xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    st.info(f"📁 Archivos cargados: **{len(uploaded_files)}**")
    
    file_summary = []
    series_buckets = {"SA": [], "SBM": [], "SBS": [], "SD": [], "SP": [], "DESCONOCIDO": []}
    
    with st.spinner("Analizando y clasificando archivos REM..."):
        for f in uploaded_files:
            bytes_data = f.getvalue()
            try:
                wb_data = openpyxl.load_workbook(io.BytesIO(bytes_data), data_only=True)
                info = detect_rem_info(wb_data, f.name)
            except Exception as e:
                info = {"series": "DESCONOCIDO", "version": "Error", "target_sheets": []}
                wb_data = None

            series = info["series"]
            item_data = {
                "name": f.name,
                "version": info["version"],
                "series": series,
                "bytes": bytes_data,
                "wb_data": wb_data,
                "sheets_count": len(info.get("target_sheets", []))
            }
            
            if series in series_buckets:
                series_buckets[series].append(item_data)
            else:
                series_buckets["DESCONOCIDO"].append(item_data)
                
            file_summary.append({
                "Archivo": f.name,
                "Serie Detectada": series,
                "Versión Detectada": info["version"],
                "Hojas de Datos": len(info.get("target_sheets", [])),
                "Tamaño (KB)": round(f.size/1024, 1)
            })

    # Mostrar métricas por serie
    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("REM SA (Atenciones)", len(series_buckets["SA"]))
    col2.metric("REM SBM (Salud Bucal/Ment.)", len(series_buckets["SBM"]))
    col3.metric("REM SBS (Salud Base)", len(series_buckets["SBS"]))
    col4.metric("REM SD (Diagnóstico/Apoyo)", len(series_buckets["SD"]))
    col5.metric("REM SP (Población)", len(series_buckets["SP"]))

    st.dataframe(pd.DataFrame(file_summary), use_container_width=True)
    
    if st.button("🚀 Procesar y Generar Consolidados REM", type="primary"):
        results = []
        errors = []
        
        progress_bar = st.progress(0)
        active_series = [s for s, items in series_buckets.items() if s != "DESCONOCIDO" and len(items) > 0]
        
        if not active_series:
            st.warning("⚠️ No se detectaron archivos válidos pertenecientes a las series SA, SBM, SBS, SD o SP.")
        else:
            total_steps = len(active_series)
            for idx, series_name in enumerate(active_series):
                st.write(f"🔄 Procesando serie **REM {series_name}** ({len(series_buckets[series_name])} archivos)...")
                master_path = get_master_template_path(series_name)
                
                if not master_path:
                    errors.append(f"No se encontró la plantilla maestra para la serie {series_name} en la carpeta 'Versiones Rem'.")
                    continue
                    
                try:
                    res = consolidate_rem_files(series_buckets[series_name], master_path, series_name)
                    results.append(res)
                except Exception as e:
                    errors.append(f"Error procesando la serie {series_name}: {str(e)}")
                    
                progress_bar.progress((idx + 1) / total_steps)

            if errors:
                for err in errors:
                    st.error(f"❌ {err}")
                    
            if results:
                st.success(f"🎉 ¡Procesamiento completado con éxito! Se generaron **{len(results)}** consolidado(s).")
                st.markdown("### 📥 Área de Descarga")
                
                # Descarga en ZIP si hay múltiples archivos consolidados
                if len(results) > 1:
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        for res in results:
                            zip_file.writestr(res["filename"], res["bytes"])
                    zip_buffer.seek(0)
                    
                    st.download_button(
                        label="📦 Descargar Paquete Completo (.ZIP con todos los Consolidados)",
                        data=zip_buffer,
                        file_name="Consolidados_REM_DEIS_2026.zip",
                        mime="application/zip",
                        type="primary"
                    )
                    st.markdown("---")
                
                # Botones individuales por archivo
                for res in results:
                    st.download_button(
                        label=f"📥 Descargar {res['filename']} ({res['series']})",
                        data=res["bytes"],
                        file_name=res["filename"],
                        mime=res["mime"]
                    )
