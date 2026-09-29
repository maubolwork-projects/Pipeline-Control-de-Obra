# excel_validator.py

#====================================
# 1. Librerías
#====================================

import win32com.client as win32
from pathlib import Path

#====================================
# 2. Variables globales
#====================================

# Medidas aproximadas de 5cm

SEAL_WIDE = 142
SEAL_HIGH = 61

#====================================
# 3. Función principal
#====================================

def insert_seal(xl_path, seal_path, sheet):
    # Inserta el sello de validación en el pie de página derecho en una hoja de excel

    xl_path = Path(xl_path).resolve()
    seal_path = Path(seal_path).resolve()

    xl = None
    wb = None

    try:
        # Abrir Excel
        xl = win32.DispatchEx("Excel.Application")

        wb = xl.Workbooks.Open(
            str(xl_path)
        )

        ws = wb.Worksheets(sheet)

        # Configurar pie de página derecho

        ws.PageSetup.RightFooter = "&G"

        # &G indica a Excel que debe mostrar la imagen al pie de página

        footer = ws.PageSetup.RightFooterPicture

        # Cargar imagen

        footer.Filename = str(seal_path)

        # Tamaño del sello

        footer.Width = SEAL_WIDE
        footer.Height = SEAL_HIGH

        # Guardar

        wb.Save()

    finally:

        if wb is not None:
            try:
                wb.Close(SaveChanges=True)
            except Exception:
                pass

        if xl is not None:
            try:
                xl.Quit()
            except Exception:
                pass

__all__ = ["insert_seal"]