#pdf_extractor_destajo.py

#====================================
# 1. Librerías
#====================================

import win32com.client
from pathlib import Path

def _create_pdf_folder(xl_path):
    pdf_folder = xl_path.parent / "PDFs CREADOS"
    pdf_folder.mkdir(exist_ok=True)

    return pdf_folder

#====================================
# 2. Función principal
#====================================

def export_excel_pdf(path, sheet_name):
    # Esta función convierte una hoja de excel a pdf

    # Ruta de excel fuente y pdf resultado
    xl_path = Path(path).resolve()
    pdf_folder = _create_pdf_folder(xl_path)
    pdf_path = pdf_folder / f"{xl_path.stem}.pdf"

    if not xl_path.exists():
        raise FileNotFoundError(f"No se encontró archivo en: {xl_path}")
    
    xl = win32com.client.DispatchEx("Excel.Application")

    try: 
        wb = xl.Workbooks.Open (str(xl_path))

        ws = wb.Worksheets(sheet_name)

        ws.ExportAsFixedFormat(
            Type=0, #PDF
            Filename=str(pdf_path)
        )

        wb.Close(SaveChanges=False)

    except Exception as e:
        print(f"Ocurrió un error durante la conversión: {e}")
        raise

    finally:
        if xl is not None:
            try:
                xl.Quit()
            except Exception:
                pass
            del xl

__all__ = ["export_excel_pdf"]