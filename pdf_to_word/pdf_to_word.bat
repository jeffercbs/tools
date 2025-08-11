@echo off
REM Script para convertir PDF a Word usando Python
REM Autor: jeffercbs
REM Fecha: 2025

echo ========================================
echo    CONVERTIDOR PDF A WORD
echo ========================================
echo.

REM Verificar si Python está instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Error: Python no está instalado o no está en el PATH
    echo 💡 Descarga Python desde: https://www.python.org/downloads/
    pause
    exit /b 1
)

REM Verificar si el script Python existe
if not exist "%~dp0pdf_to_word.py" (
    echo ❌ Error: No se encuentra el archivo pdf_to_word.py
    echo 💡 Asegúrate de que ambos archivos estén en la misma carpeta
    pause
    exit /b 1
)

REM Verificar si pdf2docx está instalado
python -c "import pdf2docx" >nul 2>&1
if errorlevel 1 (
    echo ❗ La librería pdf2docx no está instalada
    echo 💡 ¿Deseas instalarla ahora? (s/n^)
    set /p install_lib=
    if /i "!install_lib!"=="s" (
        echo 📦 Instalando pdf2docx...
        pip install pdf2docx
        if errorlevel 1 (
            echo ❌ Error al instalar pdf2docx
            pause
            exit /b 1
        )
        echo ✅ pdf2docx instalado correctamente
    ) else (
        echo ❌ No se puede continuar sin pdf2docx
        pause
        exit /b 1
    )
)

REM Si no hay argumentos, mostrar menú interactivo
if "%~1"=="" (
    echo 📋 Selecciona una opción:
    echo.
    echo 1. Convertir un archivo PDF específico
    echo 2. Convertir todos los PDFs de una carpeta
    echo 3. Mostrar ayuda
    echo.
    set /p choice=Ingresa tu opción (1-3^): 
    
    if "!choice!"=="1" (
        set /p pdf_file=📄 Ingresa la ruta del archivo PDF: 
        if "!pdf_file!"=="" (
            echo ❌ Error: Debes especificar un archivo PDF
            pause
            exit /b 1
        )
        python "%~dp0pdf_to_word.py" "!pdf_file!"
    ) else if "!choice!"=="2" (
        set /p pdf_dir=📂 Ingresa la ruta de la carpeta: 
        if "!pdf_dir!"=="" (
            echo ❌ Error: Debes especificar una carpeta
            pause
            exit /b 1
        )
        python "%~dp0pdf_to_word.py" -d "!pdf_dir!"
    ) else if "!choice!"=="3" (
        python "%~dp0pdf_to_word.py" --help
    ) else (
        echo ❌ Opción inválida
        pause
        exit /b 1
    )
) else (
    REM Pasar todos los argumentos al script Python
    python "%~dp0pdf_to_word.py" %*
)

echo.
echo 🎉 Proceso completado
pause
