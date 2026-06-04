@echo off
:: setup.bat — Create virtualenv and install dependencies on Windows
:: Corresponds to setup.sh

setlocal EnableExtensions EnableDelayedExpansion

set "SCRIPT_DIR=%~dp0"
if "!SCRIPT_DIR:~-1!"=="\" set "SCRIPT_DIR=!SCRIPT_DIR:~0,-1!"
cd /d "!SCRIPT_DIR!"

echo [MD2PPTX] Windows setup
echo Project directory: !SCRIPT_DIR!
echo.

:: Prefer the Python launcher if available; otherwise use python on PATH.
set "PY_CMD=py -3"
%PY_CMD% -c "import sys; sys.exit(0)" >nul 2>nul
if errorlevel 1 (
    set "PY_CMD=python"
)

:: Python version check: require Python 3.10+
%PY_CMD% -c "import sys; print('Python %d.%d.%d' %% sys.version_info[:3]); sys.exit(0 if sys.version_info >= (3,10) else 1)"
if errorlevel 1 (
    echo [ERROR] Python 3.10 or newer is required.
    echo Install Python from https://www.python.org/downloads/ and enable "Add python.exe to PATH",
    echo or install it from the Microsoft Store / winget.
    exit /b 1
)

echo Creating virtual environment: env
if exist "!SCRIPT_DIR!\env\Scripts\python.exe" (
    echo Existing env found. Reusing it.
) else (
    %PY_CMD% -m venv env
    if errorlevel 1 (
        echo [ERROR] Failed to create virtual environment.
        exit /b 1
    )
)

set "VENV_PY=!SCRIPT_DIR!\env\Scripts\python.exe"
set "VENV_PIP=!SCRIPT_DIR!\env\Scripts\pip.exe"

echo Upgrading pip...
"!VENV_PY!" -m pip install --upgrade pip
if errorlevel 1 (
    echo [ERROR] Failed to upgrade pip.
    exit /b 1
)

echo Installing dependencies from requirements.txt...
if exist "!SCRIPT_DIR!\requirements.txt" (
    "!VENV_PIP!" install -r "!SCRIPT_DIR!\requirements.txt"
) else (
    "!VENV_PIP!" install python-pptx lxml chardet
)
if errorlevel 1 (
    echo [ERROR] Failed to install dependencies.
    exit /b 1
)

echo.
echo [OK] Setup completed.
echo To convert the sample file:
echo     run.bat sample.md
echo.
endlocal
