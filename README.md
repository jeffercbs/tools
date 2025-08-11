## Herramientas del proyecto (autorename, PDF→Word, imágenes)

Colección de utilidades que he ido creando para automatizar tareas comunes:

- Autorename de reportes a partir de un Excel (Python)
- Convertidor de PDF a Word (.docx) (Python)
- Scripts para optimizar imágenes y generar metadatos de galería (TypeScript con Bun)

Estructura relevante:

- `autorename-reports/` — Renombrado masivo según Excel
- `pdf_to_word/` — Conversión PDF→DOCX + script `.bat` para Windows
- `bin/` — Scripts TS para optimización y metadatos de imágenes

Requisitos generales:

- Windows con PowerShell (pwsh)
- Python 3.10+ (para scripts Python)
- Bun 1.0+ (para scripts TypeScript)


## 1) Autorename de reportes (Python)

Renombra archivos PDF ubicados en una carpeta según un Excel de referencia.

- Carpeta: `autorename-reports/`
- Entrada: `soportes.xlsx` con columnas exactas: `Nº documento`, `Doc.compensación`
- Carpeta de origen: `soports` (coloca aquí los PDFs a renombrar; el nombre base debe coincidir con los códigos de la columna `Doc.compensación`)
- Salida: `renombrados/`

Instalación y ejecución rápida:

```powershell
# Desde la raíz del repo
pwsh .\autorename-reports\Init.ps1
```

Ejecución manual (opcional):

```powershell
cd .\autorename-reports
python -m venv venv
. .\venv\Scripts\Activate.ps1
pip install -r .\requirements.txt
python .\main.py
```

Qué hace:

- Lee `soportes.xlsx` y busca coincidencias con los nombres de los PDFs en `soports/` (sin extensión)
- Renombra a `soporte de pago {Nº documento}.pdf` y mueve a `renombrados/`
- Muestra conteo de renombrados, no encontrados y tiempo de ejecución


## 2) Convertidor PDF → Word (.docx) (Python)

Conversión individual o por carpeta usando `pdf2docx`.

- Carpeta: `pdf_to_word/`
- Script: `pdf_to_word.py`
- Requisitos: `pip install -r pdf_to_word/requirements.txt`
- Alternativa Windows: `pdf_to_word.bat` (interactivo)

Instalación (una vez):

```powershell
cd .\pdf_to_word
python -m venv venv
. .\venv\Scripts\Activate.ps1
pip install -r .\requirements.txt
```

Uso básico:

```powershell
# Archivo único
python .\pdf_to_word.py .\documento.pdf

# Archivo único con salida personalizada
python .\pdf_to_word.py .\documento.pdf -o .\resultado.docx

# Carpeta completa (todos los .pdf)
python .\pdf_to_word.py -d .\pdfs

# Lanzador en Windows (interactivo)
.\pdf_to_word.bat
```

Notas:

- Si falta la dependencia, verás un mensaje pidiendo instalar `pdf2docx`.
- El script gestiona códigos de salida: 0 (OK), 1 (dependencia faltante), 2 (rutas inválidas), 3 (error inesperado).


## 3) Scripts de imágenes y metadatos (Bun + TypeScript)

Dos scripts:

1) `bin/gallery-optimization.ts` — Convierte/optimiza imágenes a `.webp` con `sharp`.
	- Entrada por defecto: `src/assets/*.{webp,jpg,jpeg,png}`
	- Salida: `public/_thumbnail/`
	- Registra tamaños antes/después

2) `bin/gallery-metadata.ts` — Lee imágenes de `public/_thumbnail/` y genera `src/data/meta-gallery.json` con `{ width, height, src, alt }` usando `image-meta`.

Instalación de dependencias (Bun):

```powershell
# Desde la raíz del repo
bun add sharp image-meta
```

Ejecutar scripts:

```powershell
# Optimización (creará/actualizará public/_thumbnail)
bun .\bin\gallery-optimization.ts

# Metadatos (leerá public/_thumbnail y escribirá src/data/meta-gallery.json)
bun .\bin\gallery-metadata.ts
```

Notas:

- Ajusta los patrones/globs dentro de los scripts si tu proyecto usa rutas distintas.
- Si `sharp` falla al instalar en Windows, asegúrate de tener una versión reciente de Bun y Visual C++ Runtime actualizado.


## Solución de problemas

- pdf2docx no encontrado: activa el venv y ejecuta `pip install -r pdf_to_word/requirements.txt`.
- Columnas del Excel: verifica nombres exactos (`Nº documento`, `Doc.compensación`).
- Bun/TS: si no reconoce imports, ejecuta con Bun (no Node). Reinstala deps con `bun add ...`.


## Licencia

Este proyecto se distribuye bajo la licencia incluida en `LICENSE`.

