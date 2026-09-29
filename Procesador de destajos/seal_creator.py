#seal_creator.py

#====================================
# 1. Librerías
#====================================

from PIL import Image, ImageDraw, ImageFont
from pathlib import Path
import datetime

#====================================
# 2. Variables globales
#====================================

COLOR = "#20A9E8"
COLOR_OSCURO = "#168DCC"
fecha_actual = datetime.datetime.now()
EXECUTION_DATE = fecha_actual.strftime("%d/%m/%Y")

#====================================
# 3. Funciones internas
#====================================

def _cargar_fuente(tamano, negrita=False):
# Asigna el tipo de fuente que se usa en el sello, puede ser negrita
    if negrita:
        ruta = "C:/Windows/Fonts/arialbd.ttf"
    else:
        ruta = "C:/Windows/Fonts/arial.ttf"

    return ImageFont.truetype(ruta, tamano)

#====================================
# 4. Función principal
#====================================

def crear_sello(fecha, folio, salida):
# Crea la imagen que será utilizada para el sello de verificación
    ancho = 1200
    alto = 650

    imagen = Image.new(
        "RGBA",
        (ancho, alto),
        (255, 255, 255, 0)
    )

    draw = ImageDraw.Draw(imagen)

    # --------------------------------------------------
    # Borde exterior

    draw.rounded_rectangle(
        (20, 20, ancho - 20, alto - 20),
        radius=35,
        outline=COLOR,
        width=8
    )

    # --------------------------------------------------
    # Logo Infinity

    fuente_logo = _cargar_fuente(120, negrita=True)

    draw.text(
        (80, 75),
        "∞",
        font=fuente_logo,
        fill=COLOR
    )

    fuente_empresa = _cargar_fuente(38, negrita=True)

    draw.text(
        (80, 195),
        "INFINITY",
        font=fuente_empresa,
        fill=COLOR
    )

    draw.text(
        (80, 240),
        "CONSTRUCTION",
        font=_cargar_fuente(25, negrita=True),
        fill=COLOR
    )

    # --------------------------------------------------
    # Separador

    draw.line(
        (400, 70, 400, 300),
        fill=COLOR,
        width=5
    )

    # --------------------------------------------------
    # VALIDADO

    draw.text(
        (450, 65),
        "VALIDADO",
        font=_cargar_fuente(80, negrita=True),
        fill=COLOR
    )

    draw.text(
        (455, 155),
        "POR CONTROL DE OBRA",
        font=_cargar_fuente(40, negrita=True),
        fill=COLOR
    )

    # --------------------------------------------------
    # Línea divisoria

    draw.line(
        (70, 320, ancho - 70, 320),
        fill=COLOR,
        width=4
    )

    # --------------------------------------------------
    # Fecha

    draw.text(
        (100, 365),
        "FECHA:",
        font=_cargar_fuente(30, negrita=True),
        fill=COLOR
    )

    draw.text(
        (100, 410),
        fecha,
        font=_cargar_fuente(45, negrita=True),
        fill=COLOR
    )

    # --------------------------------------------------
    # Separador central

    draw.line(
        (600, 350, 600, 500),
        fill=COLOR,
        width=4
    )

    # --------------------------------------------------
    # Folio

    draw.text(
        (680, 365),
        "FOLIO:",
        font=_cargar_fuente(30, negrita=True),
        fill=COLOR
    )

    draw.text(
        (680, 410),
        f"{EXECUTION_DATE}-{folio}",
        font=_cargar_fuente(45, negrita=True),
        fill=COLOR
    )

    # --------------------------------------------------
    # Línea inferior

    draw.line(
        (70, 525, ancho - 70, 525),
        fill=COLOR,
        width=4
    )

    # --------------------------------------------------
    # Empresa

    fuente_footer = _cargar_fuente(32, negrita=True)

    texto = "INFINITY CONSTRUCTION"

    bbox = draw.textbbox(
        (0, 0),
        texto,
        font=fuente_footer
    )

    texto_ancho = bbox[2] - bbox[0]

    draw.text(
        ((ancho - texto_ancho) / 2, 550),
        texto,
        font=fuente_footer,
        fill=COLOR
    )

    # --------------------------------------------------
    # Guardar

    salida = Path(salida)

    salida.parent.mkdir(
        parents=True,
        exist_ok=True
    )

    imagen.save(
        salida,
        format="PNG"
    ) 
    return salida

__all__ = ["crear_sello"]