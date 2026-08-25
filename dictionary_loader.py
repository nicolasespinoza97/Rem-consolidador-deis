import os
import glob
import openpyxl

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
VERSIONES_REM_DIR = os.path.join(BASE_DIR, "Versiones Rem")
DICCIONARIOS_DIR = os.path.join(BASE_DIR, "Diccionario de Codigos")

SERIES_MASTER_MAP = {
    "SA": "SA_26_V1.2.xlsm",
    "SBM": "SBM_26_V1.1.xlsm",
    "SBS": "SBS_26_V1.3.xlsm",
    "SD": "SD_26_V1.1.xlsm",
    "SP": "SP_26_V1.2.xlsm"
}

def get_master_template_path(series):
    """
    Busca la plantilla maestra oficial para una serie específica (SA, SBM, SBS, SD, SP)
    en la carpeta 'Versiones Rem' o en el directorio base.
    """
    expected_filename = SERIES_MASTER_MAP.get(series)
    if expected_filename:
        target = os.path.join(VERSIONES_REM_DIR, expected_filename)
        if os.path.exists(target):
            return target
        # Búsqueda fallback en directorio raíz
        target_root = os.path.join(BASE_DIR, expected_filename)
        if os.path.exists(target_root):
            return target_root
            
    # Búsqueda por patrón en Versiones Rem
    candidates = glob.glob(os.path.join(VERSIONES_REM_DIR, f"*{series}*.xls*"))
    if candidates:
        return candidates[0]
        
    return None

def get_dictionary_path(series, version=None):
    """
    Busca el archivo de diccionario correspondiente a la serie y versión.
    """
    if not os.path.exists(DICCIONARIOS_DIR):
        return None
        
    candidates = glob.glob(os.path.join(DICCIONARIOS_DIR, f"*{series}*.xls*"))
    if not candidates:
        return None
        
    if version:
        for c in candidates:
            if version.lower() in c.lower():
                return c
                
    return candidates[0]
