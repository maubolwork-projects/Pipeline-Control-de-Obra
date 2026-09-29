# pdf_converter_validator.py

#====================================
# 1. Librerías
#====================================

import pdf_extractor_destajo as ext
import seal_creator as sc
import excel_validator as ev
import file_manager as fm
import datetime
from tempfile import NamedTemporaryFile
from pathlib import Path
import sys

#====================================
# 2. Variables globales
#====================================

fecha_actual = datetime.datetime.now()
EXECUTION_DATE = fecha_actual.strftime("%d/%m/%Y")

#====================================
# 3. Función orquestadora
#====================================

def excel_converter_validator(path):
    #Esta función agrega un sello de validación a la caratula de destajos y estimaciones
    #Convierte la hoja de interes de un excel a PDF

    print("=" * 40)
    print("       VALIDADOR DE DESTAJOS Y CONVERTIDOR DE PDF")
    print("=" * 40)
    print()

    #Creamos el catálogo de archivos a validar y convertir
    catalog = fm.excel_schema_manager(path)

    if not catalog:
            print("No se encontraron archivos para procesar.")
            return None

    print(f"Archivos encontrados: {len(catalog)}")
    print()

    nuevos = 0
    procesados = 0

    
    for revision, file in enumerate(catalog, start=1):
        print(f"Procesando: {file['file_name']}")
        nuevos += 1
        procesados += 1
        temp_seal = None
        try:
            with NamedTemporaryFile(suffix=".png", delete=False) as temp:
                temp_seal = temp.name

            #Se crea el sello de verificación para cada archivo procesado
            sc.crear_sello(EXECUTION_DATE, revision, temp_seal)

            #Se incerta el sello al archivo 
            ev.insert_seal(file["path"], temp_seal, file["sheets"][0])

            #Se exporta la caratula de excel a pdf
            ext.export_excel_pdf(file["path"], file["sheets"][0])

        finally:

            if temp_seal:
                Path(temp_seal).unlink(
                    missing_ok=True
                )
    print()
    print("-" * 40)
    print("Proceso terminado correctamente.")
    print(f"Archivos nuevos procesados: {nuevos}")
    print(f"Archivos omitidos: {len(catalog) - nuevos}")
    print("-" * 40)

#====================================
# 4. Función principal
#====================================

def main():

    if getattr(sys, "frozen", False):
        ruta = Path(sys.executable).resolve().parent
    else:
        ruta = Path(__file__).resolve().parent

    excel_converter_validator(ruta)

    input("\nPresione ENTER para cerrar...")

if __name__ == "__main__":
    main()