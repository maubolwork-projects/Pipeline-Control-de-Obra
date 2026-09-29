#pipeline_presupuestos.py

#====================================
# 1. Librerías
#====================================
from pathlib import Path
import file_manager as fm
import file_creator as fc
import table_constructor as tc
import db_schema as db
import sys

#====================================
# 2. Función Publica
#====================================

def run_pipeline(path):
    #Se crea el catalogo de archivos contenidos en la ruta asignada

    print("=" * 40)
    print("       PROCESADOR DE DESTAJO")
    print("=" * 40)
    print()

    catalog = fm.excel_schema_manager(path)

    if not catalog:
        print("No se encontraron archivos para procesar.")
        return None

    print(f"Archivos encontrados: {len(catalog)}")
    print()
    
    #Se crea el archivo destino del registro
    fc.create_excel_file(path)
    #Se crea el archivo de auditoria 
    db._db_init(path, catalog[0])
    #Se obtienen los hahses de los archivos procesados
    hashes_db = db.db_extractor(path)

    nuevos = 0
    procesados = 0

    for file in catalog :
        if file["hash_file"] not in hashes_db:

            print(f"Procesando: {file['file_name']}")

            tc.process_excel_file(path, file["file_name"])
            #Registro del archivo procesado 
            db.db_creator(path, file)

            nuevos += 1
            procesados += 1

        else:

            print(f"Ya procesado: {file['file_name']}")

    print()
    print("-" * 40)
    print("Proceso terminado correctamente.")
    print(f"Archivos nuevos procesados: {nuevos}")
    print(f"Archivos omitidos: {len(catalog) - nuevos}")
    print("-" * 40)

#====================================
# 3. Función Principal
#====================================

def main():

    if getattr(sys, "frozen", False):
        ruta = Path(sys.executable).resolve().parent
    else:
        ruta = Path(__file__).resolve().parent

    run_pipeline(ruta)

    input("\nPresione ENTER para cerrar...")

if __name__ == "__main__":
    main()