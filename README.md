# 🏥 Consolidador Universal de Planillas REM DEIS (SA, SBM, SBS, SD, SP)

Aplicación web desarrollada en Streamlit para la clasificación automática, homologación y consolidación anual de todas las series estadísticas de planillas REM DEIS (`.xlsm` / `.xlsx`): **SA** (REM A), **SBM** (REM BM), **SBS** (REM BS), **SD** (REM D) y **SP** (REM P).

## 🚀 Características
- **Detección Automática de Serie y Versión:** Identifica la familia de REM (`SA`, `SBM`, `SBS`, `SD`, `SP`) y su versión a partir del nombre o la hoja `NOMBRE`.
- **Protección Absoluta por Color de Fondo (#FFFFCC):** Modifica **ÚNICAMENTE** las celdas amarillas editables. Títulos, encabezados, glosas, fórmulas `=SUM(...)`, totales y celdas sin fondo amarillo permanecen **100% intactos**.
- **Soporte Multi-Serie en Lote:** Carga simultánea de múltiples archivos de distintas series REM.
- **Soporte de Macros VBA (.xlsm):** Mantiene intacto el proyecto de macros y formatos de las plantillas oficiales.
- **Descarga Flexible:** Botones de descarga individuales por serie o descarga de paquete completo comprimido `.zip`.

## 📁 Estructura del Proyecto
- `app.py`: Interfaz principal en Streamlit (Dashboard Multi-REM).
- `Versiones Rem/`: Contiene las plantillas maestras oficiales de referencia (`SA_26_V1.2.xlsm`, `SBM_26_V1.1.xlsm`, `SBS_26_V1.3.xlsm`, `SD_26_V1.1.xlsm`, `SP_26_V1.2.xlsm`).
- `Diccionario de Codigos/`: Diccionarios oficiales de prestaciones por serie y versión.
- `core/`:
  - `detector.py`: Módulo de detección de tipo de REM y versión.
  - `dictionary_loader.py`: Gestor de plantillas maestras y diccionarios.
  - `engine.py`: Motor genérico de consolidación por celda amarilla.

## 🛠️ Cómo Ejecutar

```bash
uv run --with streamlit --with openpyxl --with pandas streamlit run app.py
```
O con python tradicional:
```bash
pip install -r requirements.txt
streamlit run app.py
```
