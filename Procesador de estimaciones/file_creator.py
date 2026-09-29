#file_creator.py

#====================================
# 1. Librerías
#====================================

import pandas as pd
from pathlib import Path

#====================================
# 2. Funciones Privadas (Uso Interno)
#====================================

def _sheet_instructions(writer):
    #Esta función construye una hoja de excel con las instrucciones de ejecución del programa

    message = [
        "",
        "IMPORTANTE : ",
        "",
        "El programa contiene un registro de archivos procesados que no verifica el contenido del archivo.",
        "Si necesita hacer alguna corrección, haga lo siguiente: ",
        "",
        "Paso 1: Haga las correciones necesarias",
        "",
        "Paso 2: Modifique el nombre del archivo como se muestra a continuación, ejemplo: ",
        "Anterior: Destajo 1.xlsx ---> Final: Destajo 1 CORREGIDO.xslx",
        "se agrego la palabra CORREGIDO al final del nombre del archivo",
        "",
        "Paso 3: Vuelva a correr el programa",
    ]

    df = pd.DataFrame(message, columns=["Instrucciones de Ejecución"])
    df.to_excel(writer, sheet_name="Instrucciones", index=False)


def _excel_file(writer):
    #Esta función crea un archivo en excel, utilizando el nombre de columnas predefinidas

    interest_columns = ["ID", "CLAVE", "CONCEPTO", "UNIDAD", "CANT_CONTRATO", "P_UNITARIO", "IMPORTE_CONTRATO", "AVANCE_ACUM_CANTIDAD",
                        "AVANCE_ACUM_PCT", "AVANCE_ACUM_IMPORTE", "POR_EJERCER_CANTIDAD", "POR_EJERCER_IMPORTE","NOMBRE_OBRA", "NO_ESTIMACION",
                        "PERIODO_ESTIMACION", "CONTRATISTA", "TRABAJOS_A_REALIZAR", "NO_OBRA", "NO_CONTRATO", "FECHA_ELABORACION",
                        "IMPORTE_CONTRATO", ]
    df = pd.DataFrame(columns=interest_columns)
    df.to_excel(writer, sheet_name="Registro", index=False)

def _exist_excel_file(path):
    #Esta función comprueba si el archivo de registro de destajos existe, si existe regresa true
    file_path = Path(path)/"Caratula de estimaciones.xlsx"
    exist = file_path.exists()

    return exist

#====================================
# 3. Función Pública (Interfaz)
#====================================

def create_excel_file(path):
    #Esta función crea el archivo de registro en excel si no existe, si existe regresa true
    if not _exist_excel_file(path):
        file_path = Path(path)/"Caratula de estimaciones.xlsx"

        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
            _sheet_instructions(writer)
            _excel_file(writer)
        
    return True

__all__ = ["create_excel_file"]
