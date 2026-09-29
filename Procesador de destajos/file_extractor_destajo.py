#file_extractor_destajo.py
#Extractor para el "FORMATO DE DESTAJO C.O." -> hoja CT_EST

#====================================
# 1. Librerías y Variables
#====================================

import openpyxl
import pandas as pd

SHEET_NAME = "CT_EST"

interest_cells = {
    "FOLIO": "M2",
    "CLIENTE": "B3",
    "OBRA_NO": "M3",
    "OBRA": "F4",
    "ACTIVIDAD_CONTRATADA": "F5",
    "DESTAJO_NO": "N5",
    "DESTAJISTA": "F6",
    "PERIODO": "N6",
    "RESIDENTE": "F7",
    "FECHA": "N7",
}

#La tabla de conceptos no tiene un número fijo de filas (varía por archivo:
#renglones eliminados, texto que hace wrap a 2+ filas, etc.), así que se lee
#dinámicamente desde TABLE_START hasta encontrar la palabra STOP_MARKER.
TABLE_START = 13
STOP_MARKER = "TOTAL"

TABLE_COLUMNS = {
    "B": "CLAVE",
    "C": "CONCEPTO",
    "D": "UNIDAD",
    "E": "P_UNITARIO",
    "F": "CANT_CONTRATADO",
    "G": "IMPORTE_CONTRATADO",
    "H": "CANT_ACUM_ANTERIOR",
    "I": "IMPORTE_ACUM_ANTERIOR",
    "J": "CANT_ESTA_ESTIMACION",
    "K": "IMPORTE_ESTA_ESTIMACION",
    "L": "CANT_ACUM_ACTUAL",
    "M": "IMPORTE_ACUM_ACTUAL",
    "N": "CANT_POR_EJERCER",
    "O": "IMPORTE_POR_EJERCER",
}

#====================================
# 2. Funciones Privadas (Uso Interno)
#====================================

def _extract_header_data(sheet):
    #Extrae los datos de la carátula usando las celdas fijas definidas en interest_cells
    header_data = {}
    for column, cell in interest_cells.items():
        header_data[column] = sheet[cell].value
    return header_data


def _is_stop_row(sheet, row):
    #Revisa si en la fila actual aparece la palabra STOP_MARKER (fin de tabla)
    for col in TABLE_COLUMNS:
        value = sheet[f"{col}{row}"].value
        if isinstance(value, str) and STOP_MARKER in value.upper():
            return True
    return False


def _extract_table_data(sheet):
    #Recorre la tabla desde TABLE_START fila por fila, tomando solo las filas
    #donde exista CONCEPTO (columna C), hasta toparse con la fila de TOTAL
    rows = []
    row = TABLE_START

    while row <= sheet.max_row:
        if _is_stop_row(sheet, row):
            break

        #Se exige CLAVE (col B) y CONCEPTO (col C) para evitar falsas filas de
        #datos (notas sueltas, celdas de wrap-text vacías, etc.)
        clave = sheet[f"B{row}"].value
        concepto = sheet[f"C{row}"].value
        if clave is not None and concepto is not None:
            registro = {
                name: sheet[f"{col}{row}"].value
                for col, name in TABLE_COLUMNS.items()
            }
            rows.append(registro)

        row += 1

    return pd.DataFrame(rows)

#====================================
# 3. Función Pública (Interfaz)
#====================================

def extracted_data_destajo(path):
    #Función principal: abre el archivo, extrae encabezado y tabla de la hoja CT_EST
    wb = openpyxl.load_workbook(path, data_only=True)
    sheet = wb[SHEET_NAME]

    header_data = _extract_header_data(sheet)
    table_data = _extract_table_data(sheet)

    wb.close()
    return header_data, table_data


__all__ = ["extracted_data_destajo"]
