# Validador de Destajos y Convertidor de PDF

Aplicación desarrollada en Python para automatizar el proceso de **validación y conversión de formatos de destajos y estimaciones en Excel a PDF**.

El proceso identifica los archivos Excel disponibles en una carpeta, genera un sello de validación, lo inserta en el pie de página de la hoja correspondiente y posteriormente genera una versión PDF del documento.

Los archivos PDF generados se almacenan automáticamente en una carpeta independiente llamada `PDFs CREADOS`.

---

## 1. Objetivo

El objetivo del proyecto es automatizar un proceso que anteriormente requería realizar manualmente varias actividades:

* Identificar los archivos Excel pendientes de revisión.
* Generar un sello de validación.
* Insertar el sello en el documento Excel.
* Guardar los cambios realizados.
* Convertir el documento a PDF.
* Organizar los archivos PDF generados.
* Eliminar los archivos temporales utilizados durante el proceso.

La aplicación busca reducir la intervención manual y estandarizar la generación de documentos validados para el área de **Control de Obra**.

---

## 2. Flujo general

El proceso completo se ejecuta de manera secuencial:

```text
                    ARCHIVOS EXCEL
                          │
                          ▼
                 ┌─────────────────┐
                 │  file_manager   │
                 │                 │
                 │ Detecta Excel   │
                 │ Obtiene hojas   │
                 │ Genera catálogo │
                 └────────┬────────┘
                          │
                          ▼
                ┌───────────────────┐
                │   seal_creator    │
                │                   │
                │ Genera sello PNG  │
                └─────────┬─────────┘
                          │
                          ▼
                ┌───────────────────┐
                │  excel_validator  │
                │                   │
                │ Inserta el sello  │
                │ en el pie de      │
                │ página de Excel   │
                └─────────┬─────────┘
                          │
                          ▼
                ┌───────────────────┐
                │ pdf_extractor_    │
                │     destajo       │
                │                   │
                │ Exporta la hoja   │
                │ seleccionada a PDF│
                └─────────┬─────────┘
                          │
                          ▼
                    PDFs CREADOS/
                          │
                          ▼
                    PDF FINAL
```

Durante todo el proceso, `pdf_converter_validator.py` funciona como **orquestador principal**.

---

## 3. Arquitectura

El proyecto está dividido en módulos con responsabilidades específicas.

```text
Pipeline_destajo/
│
├── scripts/
│   │
│   ├── file_manager.py
│   ├── seal_creator.py
│   ├── excel_validator.py
│   ├── pdf_extractor_destajo.py
│   └── pdf_converter_validator.py
│
└── PDFs CREADOS/
```

### Responsabilidad de cada módulo

| Módulo                       | Responsabilidad                                                     |
| ---------------------------- | ------------------------------------------------------------------- |
| `file_manager.py`            | Identificar archivos Excel y construir el catálogo de procesamiento |
| `seal_creator.py`            | Generar el sello de validación en formato PNG                       |
| `excel_validator.py`         | Insertar el sello en el pie de página del Excel                     |
| `pdf_extractor_destajo.py`   | Convertir la hoja de Excel a PDF                                    |
| `pdf_converter_validator.py` | Coordinar y ejecutar todo el flujo                                  |

---

# 4. Descripción de los módulos

## 4.1 `file_manager.py`

Este módulo administra la identificación de los archivos Excel que serán procesados.

### Funciones principales

#### `_get_excel_files(directory)`

Busca los archivos `.xlsx` dentro de la carpeta proporcionada.

Excluye:

* `Caratula de destajos.xlsx`
* Archivos temporales de Excel que comienzan con `~$`

También valida que la ruta proporcionada exista y corresponda a una carpeta.

---

#### `_get_excel_sheets(ruta)`

Obtiene los nombres de las hojas existentes dentro de un archivo Excel.

Utiliza `pandas.ExcelFile` para realizar la lectura de la estructura del libro.

---

#### `_get_load_hash(ruta, name_sheets)`

Genera dos identificadores mediante MD5:

* `hash_load`: identifica la ejecución considerando la ruta, hojas y fecha de ejecución.
* `hash_file`: identifica el archivo considerando la ruta, nombre y hojas.

Estos valores forman parte de la metadata generada para cada archivo.

---

#### `_get_excel_schema(list_path)`

Construye el catálogo de archivos que serán procesados.

Cada archivo contiene información como:

```python
{
    "file_id": ...,
    "file_name": ...,
    "path": ...,
    "sheets": ...,
    "hash_file": ...,
    "hash_load": ...,
    "load_time": ...
}
```

---

#### `excel_schema_manager(path)`

Es la interfaz pública del módulo.

Recibe una ruta y devuelve el catálogo de archivos Excel encontrados.

```python
catalog = excel_schema_manager(path)
```

---

## 4.2 `seal_creator.py`

Este módulo genera la imagen que funciona como **sello de validación**.

El sello se genera como una imagen PNG con fondo transparente para posteriormente poder insertarlo en el pie de página del documento Excel.

### Características del sello

Incluye:

* Identidad visual de INFINITY CONSTRUCTION.
* Texto `VALIDADO`.
* Texto `POR CONTROL DE OBRA`.
* Fecha de validación.
* Folio de validación.
* Identificador de la empresa.

El folio se construye utilizando la fecha de ejecución y el identificador asignado al archivo durante el procesamiento.

### Funciones

#### `_cargar_fuente(tamano, negrita=False)`

Función interna encargada de seleccionar la fuente Arial normal o negrita.

---

#### `crear_sello(fecha, folio, salida)`

Genera el sello y lo guarda en la ubicación especificada.

Recibe:

| Parámetro | Descripción                                   |
| --------- | --------------------------------------------- |
| `fecha`   | Fecha de validación                           |
| `folio`   | Identificador del archivo procesado           |
| `salida`  | Ruta donde se almacenará temporalmente el PNG |

Devuelve la ruta del archivo generado.

---

## 4.3 `excel_validator.py`

Este módulo automatiza Excel mediante `win32com`.

Su función es insertar el sello de validación generado previamente en el pie de página derecho de la hoja de Excel.

### `insert_seal(xl_path, seal_path, sheet)`

Recibe:

* Ruta del archivo Excel.
* Ruta de la imagen del sello.
* Nombre de la hoja donde se insertará.

El proceso realiza:

1. Apertura de una instancia independiente de Excel.
2. Apertura del libro.
3. Selección de la hoja.
4. Configuración del pie de página derecho.
5. Carga de la imagen.
6. Ajuste de dimensiones.
7. Guardado del libro.
8. Cierre de Excel.

La imagen se inserta mediante:

```python
ws.PageSetup.RightFooter = "&G"
```

`&G` indica a Excel que debe mostrar una imagen en el pie de página.

Las dimensiones utilizadas actualmente son:

```python
SEAL_WIDE = 142
SEAL_HIGH = 61
```

Estas dimensiones corresponden aproximadamente a un sello de 5 cm.

---

## 4.4 `pdf_extractor_destajo.py`

Este módulo se encarga exclusivamente de generar el PDF a partir del archivo Excel validado.

### `_create_pdf_folder(xl_path)`

Crea la carpeta:

```text
PDFs CREADOS
```

en el mismo directorio donde se encuentra el archivo Excel.

Si la carpeta ya existe, no se genera un error.

---

### `export_excel_pdf(path, sheet_name)`

Convierte la hoja especificada del archivo Excel a PDF.

El proceso:

1. Obtiene la ruta del Excel.
2. Crea la carpeta `PDFs CREADOS`.
3. Construye el nombre del PDF.
4. Abre Excel mediante `win32com`.
5. Abre el archivo.
6. Selecciona la hoja indicada.
7. Utiliza `ExportAsFixedFormat`.
8. Guarda el PDF.
9. Cierra Excel.

El nombre del PDF conserva el nombre del archivo Excel original:

```text
FORMATO DE DESTAJO C.O_.xlsx
                ↓
FORMATO DE DESTAJO C.O_.pdf
```

El PDF se almacena en:

```text
PDFs CREADOS/
```

---

# 5. `pdf_converter_validator.py`

Este archivo es el **punto de entrada principal y orquestador del proyecto**.

Es el encargado de coordinar todos los módulos.

### `excel_converter_validator(path)`

Ejecuta el flujo completo para cada archivo encontrado.

Para cada Excel:

```text
1. Identificar archivo
       ↓
2. Crear sello temporal
       ↓
3. Insertar sello en Excel
       ↓
4. Convertir Excel a PDF
       ↓
5. Eliminar sello temporal
```

El sello PNG se crea mediante `NamedTemporaryFile`, por lo que no forma parte de los archivos finales generados por el proceso.

Al finalizar el procesamiento se muestra un resumen en consola.

---

## 6. Ejecución

El programa puede ejecutarse directamente desde Python.

Desde la carpeta `scripts`:

```bash
python pdf_converter_validator.py
```

También puede ejecutarse utilizando el launcher de Python en Windows:

```bash
py pdf_converter_validator.py
```

El programa determina automáticamente la ubicación desde la cual está siendo ejecutado.

---

## 7. Ejecución como `.exe`

El proyecto puede empaquetarse como ejecutable utilizando **PyInstaller**.

Ejemplo:

```bash
py -m PyInstaller --onefile --name "Procesador_Destajos" pdf_converter_validator.py
```

El programa utiliza:

```python
if getattr(sys, "frozen", False):
    ruta = Path(sys.executable).resolve().parent
else:
    ruta = Path(__file__).resolve().parent
```

Esto permite que el programa determine correctamente su ubicación tanto cuando se ejecuta como:

* archivo `.py`
* ejecutable `.exe`

Por lo tanto, el usuario final no necesita ejecutar individualmente los módulos internos.

---

# 8. Requisitos

El proyecto requiere un entorno Windows debido a la automatización de Microsoft Excel mediante COM.

### Software

* Windows
* Microsoft Excel
* Python 3.x

### Librerías Python

```text
pandas
openpyxl
Pillow
pywin32
```

Las librerías pueden instalarse mediante:

```bash
pip install pandas openpyxl Pillow pywin32
```

---

# 9. Dependencias principales

### Pandas

Utilizado para consultar la estructura de los archivos Excel y obtener sus hojas.

### Pillow

Utilizado para generar la imagen del sello de validación.

### pathlib

Utilizado para manejar rutas y archivos de manera estructurada.

### pywin32

Utilizado para automatizar Microsoft Excel mediante COM.

### tempfile

Utilizado para crear el archivo PNG temporal que contiene el sello de validación.

### hashlib

Utilizado para generar los identificadores MD5 asociados a los archivos y ejecuciones.

---

# 10. Entradas

La aplicación espera encontrar archivos Excel `.xlsx` dentro de la carpeta desde la que se ejecuta.

Los archivos temporales de Excel que comienzan con:

```text
~$
```

son ignorados.

También se excluye:

```text
Caratula de destajos.xlsx
```

---

# 11. Salidas

Por cada archivo Excel procesado se generan dos tipos de resultado.

### Excel validado

El archivo Excel original es modificado para incorporar el sello de validación en el pie de página.

### PDF

Se genera un archivo PDF dentro de:

```text
PDFs CREADOS/
```

Ejemplo:

```text
FORMATO DE DESTAJO C.O_.xlsx
        │
        ├── Excel actualizado con sello
        │
        └── PDFs CREADOS/
                └── FORMATO DE DESTAJO C.O_.pdf
```

El archivo PNG utilizado para insertar el sello es temporal y se elimina al finalizar el procesamiento.

---

# 12. Consideraciones

## Sistema operativo

El proyecto está diseñado actualmente para Windows debido al uso de:

```python
win32com.client
```

y la automatización directa de Microsoft Excel.

## Microsoft Excel

Es necesario tener Microsoft Excel instalado para que `win32com` pueda abrir, modificar y exportar los archivos.

## Formato de Excel

El proceso está diseñado para trabajar con un formato de documento conocido y estable.

Actualmente se utiliza:

```python
file["sheets"][0]
```

por lo que se procesa la primera hoja del libro Excel.

## Fuentes

El sello utiliza actualmente las fuentes Arial instaladas en Windows:

```text
C:/Windows/Fonts/arial.ttf
C:/Windows/Fonts/arialbd.ttf
```

Por lo tanto, el diseño actual depende de la disponibilidad de estas fuentes.

---

# 13. Manejo de archivos temporales

Durante el proceso se genera temporalmente una imagen PNG con el sello.

El flujo es:

```text
Crear PNG temporal
       ↓
Insertar en Excel
       ↓
Guardar Excel
       ↓
Generar PDF
       ↓
Eliminar PNG temporal
```

El sello no se conserva como archivo independiente después de finalizar el proceso.

---

# 14. Manejo de errores

El proceso utiliza validaciones y bloques `try/finally` para controlar recursos de Excel.

Entre las validaciones implementadas se encuentran:

* Existencia de la carpeta de entrada.
* Validación de que la ruta corresponda a una carpeta.
* Existencia del archivo Excel.
* Cierre del libro de Excel.
* Cierre de la instancia de Excel.
* Eliminación del archivo temporal del sello.

En caso de producirse un error durante el procesamiento, la excepción se propaga para permitir identificar el punto donde ocurrió el problema.

---

# 15. Estructura conceptual del proyecto

La separación de responsabilidades permite mantener el proyecto modular:

```text
file_manager
     │
     │ Catálogo de archivos
     ▼
seal_creator
     │
     │ Sello temporal
     ▼
excel_validator
     │
     │ Excel validado
     ▼
pdf_extractor_destajo
     │
     │ PDF
     ▼
PDFs CREADOS
```

El archivo:

```text
pdf_converter_validator.py
```

coordina todas estas etapas.

---

# 16. Mejoras futuras

Algunas mejoras que pueden incorporarse posteriormente:

* Control de archivos ya procesados.
* Registro histórico de ejecuciones.
* Identificación de archivos nuevos y previamente procesados.
* Selección de la hoja mediante configuración en lugar de utilizar siempre la primera hoja.
* Archivo de configuración para parámetros del proceso.
* Registro de errores en un archivo de log.
* Interfaz gráfica para usuarios no técnicos.
* Generación de un ejecutable distribuible.
* Validación de la estructura del Excel antes de procesarlo.
* Incorporación de más formatos de documentos.
* Automatización de la distribución de los PDF generados.

---

# 17. Estado actual

**Versión:** 1.0

### Flujo implementado

* [ ] Detección automática de archivos Excel.
* [ ] Identificación de hojas.
* [ ] Generación de metadata de archivos.
* [ ] Generación de sello de validación.
* [ ] Inserción del sello en Excel.
* [ ] Guardado del Excel validado.
* [ ] Conversión de Excel a PDF.
* [ ] Creación automática de la carpeta `PDFs CREADOS`.
* [ ] Eliminación del sello temporal.
* [ ] Ejecución mediante Python.
* [ ] Preparación para distribución mediante `.exe`.

---

## Autor

Proyecto desarrollado para la automatización del proceso de validación y generación de documentos del área de **Control de Obra – INFINITY CONSTRUCTION**.
