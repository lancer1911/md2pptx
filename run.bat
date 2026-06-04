@echo off
:: run.bat — Convert Markdown to PPTX on Windows
:: Corresponds to run.sh
::
:: Usage:
::   run.bat input.md
::   run.bat input.md output.pptx
::   run.bat input.md output.pptx template.pptx
::
:: Positional arguments:
::   %1  input .md file        required
::   %2  output .pptx file     optional; defaults to <input>.pptx in the same directory
::   %3  template .pptx file   optional; defaults to template.pptx next to this script

setlocal EnableExtensions EnableDelayedExpansion

set "SCRIPT_DIR=%~dp0"
if "!SCRIPT_DIR:~-1!"=="\" set "SCRIPT_DIR=!SCRIPT_DIR:~0,-1!"
set "CALL_DIR=%CD%"

if "%~1"=="" (
    echo Usage: run.bat input.md [output.pptx] [template.pptx]
    exit /b 1
)

:: Resolve input Markdown path relative to the caller's current directory.
set "MD_ARG=%~1"
if "!MD_ARG:~1,1!"==":" (
    set "MD_FILE=!MD_ARG!"
) else if "!MD_ARG:~0,1!"=="\" (
    set "MD_FILE=%CALL_DIR:~0,2%!MD_ARG!"
) else (
    set "MD_FILE=%CALL_DIR%\!MD_ARG!"
)

if not exist "!MD_FILE!" (
    echo [ERROR] Input file not found: !MD_FILE!
    exit /b 1
)

:: Resolve output path.
if not "%~2"=="" (
    set "OUT_ARG=%~2"
    if "!OUT_ARG:~1,1!"==":" (
        set "OUTPUT=!OUT_ARG!"
    ) else if "!OUT_ARG:~0,1!"=="\" (
        set "OUTPUT=%CALL_DIR:~0,2%!OUT_ARG!"
    ) else (
        set "OUTPUT=%CALL_DIR%\!OUT_ARG!"
    )
) else (
    for %%F in ("!MD_FILE!") do (
        set "MD_DIR=%%~dpF"
        set "MD_BASE=%%~nF"
    )
    if "!MD_DIR:~-1!"=="\" set "MD_DIR=!MD_DIR:~0,-1!"
    set "OUTPUT=!MD_DIR!\!MD_BASE!.pptx"
)

:: Resolve template path.
set "USING_CUSTOM_TEMPLATE=false"
if not "%~3"=="" (
    set "TPL_ARG=%~3"
    if "!TPL_ARG:~1,1!"==":" (
        set "TEMPLATE=!TPL_ARG!"
    ) else if "!TPL_ARG:~0,1!"=="\" (
        set "TEMPLATE=%CALL_DIR:~0,2%!TPL_ARG!"
    ) else (
        set "TEMPLATE=%CALL_DIR%\!TPL_ARG!"
    )
    if not exist "!TEMPLATE!" (
        echo [ERROR] Template file not found: !TEMPLATE!
        exit /b 1
    )
    set "USING_CUSTOM_TEMPLATE=true"
) else (
    if exist "!SCRIPT_DIR!\template.pptx" (
        set "TEMPLATE=!SCRIPT_DIR!\template.pptx"
    ) else if exist "%CALL_DIR%\template.pptx" (
        set "TEMPLATE=%CALL_DIR%\template.pptx"
    ) else (
        echo [ERROR] template.pptx was not found.
        echo Put template.pptx next to run.bat or in the current directory,
        echo or pass a template as the third argument.
        exit /b 1
    )
)

:: Use the virtualenv Python if setup.bat has been run.
set "PYTHON_EXE=!SCRIPT_DIR!\env\Scripts\python.exe"
if not exist "!PYTHON_EXE!" (
    set "PYTHON_EXE=py -3"
    py -3 -c "import sys; sys.exit(0)" >nul 2>nul
    if errorlevel 1 (
        set "PYTHON_EXE=python"
    )
)

:: Auto-generate a template JSON config for custom templates.
for %%T in ("!TEMPLATE!") do set "TEMPLATE_JSON=%%~dpnT.json"

if "!USING_CUSTOM_TEMPLATE!"=="true" (
    if not exist "!TEMPLATE_JSON!" (
        echo Detecting template layout and generating config...
        !PYTHON_EXE! "!SCRIPT_DIR!\inspect_template.py" "!TEMPLATE!" --json "!TEMPLATE_JSON!"
        if errorlevel 1 (
            echo [ERROR] inspect_template.py failed.
            exit /b 1
        )
        echo Config saved to: !TEMPLATE_JSON!
    ) else (
        echo Using existing config: !TEMPLATE_JSON!
    )
)

echo Input:    !MD_FILE!
echo Template: !TEMPLATE!
echo Output:   !OUTPUT!
echo.

if exist "!TEMPLATE_JSON!" (
    !PYTHON_EXE! "!SCRIPT_DIR!\md_to_pptx.py" "!MD_FILE!" "!TEMPLATE!" "!OUTPUT!" --config "!TEMPLATE_JSON!"
) else (
    !PYTHON_EXE! "!SCRIPT_DIR!\md_to_pptx.py" "!MD_FILE!" "!TEMPLATE!" "!OUTPUT!"
)

if errorlevel 1 (
    echo.
    echo [ERROR] Conversion failed.
    exit /b 1
)

echo.
echo [OK] Completed: !OUTPUT!
endlocal
