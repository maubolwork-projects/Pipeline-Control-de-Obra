#db_schema.py

#====================================
# 1. Librerías
#====================================

import sqlite3
from pathlib import Path

#====================================
# 2. Funciones Privadas (Uso Interno)
#====================================

def _db_columns(dictionary):
    #Esta función crea las columnas que se cargaran en la db
    columns = ["id INTEGER PRIMARY KEY AUTOINCREMENT"]

    for key, value in dictionary.items():
        if key == "file_id":
            columns.append(f"{key} INTEGER NOT NULL")
        else:
            columns.append(f"{key} TEXT NOT NULL")

    columns_fields = ",".join(columns)
    return columns_fields

def _db_init(path: str, dictionary: dict):
    #Crea una base de datos en un archivo .db y define su estructura

    db_name = "file_control.db"
    complete_path = Path(path) / db_name
    #Asegura la existencia de la tabla
    complete_path.parent.mkdir(parents=True, exist_ok=True)

    #Conectamos o creamos el archivo .db
    conection = sqlite3.connect(complete_path)
    cursor = conection.cursor()

    columns_fields = _db_columns(dictionary)

    querry = f"""CREATE TABLE IF NOT EXISTS file_control (
                    {columns_fields});"""

    cursor.execute(querry)
    conection.commit()
    conection.close()


def _prepare_values(register, columns):
    #Esta función convierte los datos que puedan estar en formato lista a str uniendolos con una ","
    #los demás datos tambien los convierte a str
    
    row = []
    for col in columns:
        value = register[col]

        if isinstance(value, list):
            value = ",".join(map(str, value))
        elif not isinstance(value, (int, float, str, type(None))):
            value = str(value)

        row.append(value)
    insert_values = tuple(row)
        
    return insert_values


def _insert_metadata(path: str, metadata: dict):
    #Inserta un diccionario de metadatos en la tabla de SQLite

    db_name = "file_control.db"
    complete_path = Path(path) / db_name
    #Se obtienen las columnas 
    columns = list(metadata.keys())

    conection = sqlite3.connect(complete_path)
    cursor = conection.cursor()

    marks = ",".join(["?"] * len(columns))
    columns_str = ",".join(columns)
    querry = f"INSERT INTO file_control ({columns_str}) VALUES ({marks})"

    insert_values = _prepare_values(metadata, columns)

    try:
        cursor.execute(querry, insert_values)
        conection.commit()
    except sqlite3.Error as e:
        print(f"Error al insertar los datos: {e}")
        conection.rollback()
    finally:
        conection.close()

#====================================
# 3. Función Pública (Interfaz)
#====================================

def db_extractor(path: str):
    #Esta función extrae la columna hash_file de una db de SQLite
    db_name = "file_control.db"
    complete_path = Path(path) / db_name
    conection = sqlite3.connect(complete_path)
    conection.row_factory = sqlite3.Row
    cursor = conection.cursor()

    try: 
        cursor.execute("SELECT hash_file FROM file_control")
        rows = cursor.fetchall()

        result = [row[0] for row in rows]
        return result

    except sqlite3.Error as e:
        print(f"Error ala extraer datos: {e}")
        return []

    finally:
        conection.close()

def db_creator(path, metadata):
    #Esta función crea una database en un archivo .db y llena con los metadatos

    if metadata:
        _db_init(path, metadata)

    _insert_metadata(path, metadata)


__all__ = ["db_creator"]