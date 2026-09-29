#file_manager.py

#====================================
# 1. Librerías y Variables
#====================================

from pathlib import Path
import pandas as pd
import datetime
import hashlib


fecha_actual = datetime.datetime.now()
EXECUTION_DATE = fecha_actual.strftime("%d/%m/%Y %H:%M:%S")

#====================================
# 2. Funciones Privadas (Uso Interno)
#====================================

def _get_excel_files(directory):
    #Esta función busca todos los archivos de Excel en una carpeta y devuelve una lista con sus rutas

    folder = Path(directory)
    if not folder.exists():
        raise FileNotFoundError(f"La carpeta {directory} no existe.")

    if not folder.is_dir():
        raise FileNotFoundError(f"La ruta {directory} no es una carpeta")
    
    excel_files = [file for file in folder.glob("*.xlsx") if file.name != "Caratula de estimaciones.xlsx"]

    return excel_files

def _get_excel_sheets(ruta):
    #Esta función toma una ruta de un archivo excel y devuelve una lista con el nombre de las hojas que contiene
    xl = pd.ExcelFile(ruta)
    sheets = xl.sheet_names
    return sheets

def _get_load_hash(ruta,name_sheets):
    #Esta función toma la ruta y el nombre de las hojas que un excel 
    #crea una cadena de texto con ellos, calcula y devuelve un hash
    
    path_n = Path(ruta).resolve()
    sheets = name_sheets
    #Creación de hash de carga
    ids_loads = f"{str(path_n)}|{','.join(sheets)}|{EXECUTION_DATE}"
    hash_load = hashlib.md5(str(ids_loads).encode("utf-8")).hexdigest()
    #Creación de hash de archivo
    ids_files = f"{str(path_n)}|{','.join(sheets)}|{Path(ruta).name}"
    hash_file = hashlib.md5(str(ids_files).encode("utf-8")).hexdigest()
    return hash_load, hash_file

def _get_excel_schema(list_path):
    #Esta función crea un diccionario con el nombre del archivo y las hojas que contiene
    #la ruta, el hash de carga y la fecha de ejecución. Devuelve una lista de diccionarios con esta información.

    metadata_list = []

    for count, i in enumerate(list_path, start=1):
        sheets = _get_excel_sheets(i)
        hash_load, hash_file = _get_load_hash(i,sheets)
        schema_file = {
            "file_id" : count,
            "file_name" : Path(i).name,
            "path" : Path(i),
            "sheets" : sheets,
            "hash_file" : hash_file,
            "hash_load" : hash_load,
            "load_time" : EXECUTION_DATE
        }
        metadata_list.append(schema_file)

    return metadata_list

#====================================
# 3. Función Pública (Interfaz)
#====================================

def excel_schema_manager(path):
    #Esta función toma una ruta, obtiene todos los excel en ella y devuelve un diccionario con las caracteristicas del archivo
    list_path = _get_excel_files(path)
    schema = _get_excel_schema(list_path)
    return schema

__all__ = ["excel_schema_manager"]
