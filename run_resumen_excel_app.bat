@echo off
setlocal
cd /d "%~dp0"

if not exist app.py (
    echo No se encontro app.py en esta carpeta.
    echo Pon este .bat dentro de la carpeta del programa.
    pause
    exit /b 1
)

where py >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON_CMD=py"
) else (
    where python >nul 2>nul
    if %errorlevel%==0 (
        set "PYTHON_CMD=python"
    ) else (
        echo Python no esta instalado o no esta en el PATH.
        echo Instala Python y vuelve a intentarlo.
        pause
        exit /b 1
    )
)

if not exist .venv (
    echo Creando entorno virtual...
    %PYTHON_CMD% -m venv .venv
    if errorlevel 1 (
        echo No se pudo crear el entorno virtual.
        pause
        exit /b 1
    )
)

call .venv\Scripts\activate.bat
if errorlevel 1 (
    echo No se pudo activar el entorno virtual.
    pause
    exit /b 1
)

if exist requirements.txt (
    echo Instalando dependencias...
    python -m pip install --upgrade pip
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo Error instalando dependencias.
        pause
        exit /b 1
    )
)

echo Abriendo la aplicacion en el navegador...
start http://127.0.0.1:5000

echo Ejecutando servidor...
python app.py

pause
