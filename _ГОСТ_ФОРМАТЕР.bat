@echo off
setlocal enabledelayedexpansion
title Ultimate GOST-Formatter

set "PYTHON_EXEC=C:\Users\New-life\AppData\Local\Programs\Python\Python313\python.exe"
:: Fallback to python in PATH
if not exist "%PYTHON_EXEC%" (
    set "PYTHON_EXEC=python"
)

set "BUILD_SCRIPT=%~dp0WORD\execution\build_docx.py"
set "FORMAT_SCRIPT=%~dp0WORD\execution\format_docx.py"

echo ===================================================
echo             ULTIMATE GOST-FORMATTER
echo ===================================================
echo.

if "%~1"=="" (
    echo [INFO] No input file detected.
    echo In the future, you can simply DRAG AND DROP your
    echo .md or .docx file directly onto this .bat icon!
    echo.
    echo What do you want to do right now?
    echo 1 - Build GOST document from Markdown ^(.md^)
    echo 2 - Format and fix existing Word document ^(.docx^)
    echo.
    set /p choice="Enter 1 or 2 and press Enter: "
    echo.
    if "!choice!"=="1" (
        set /p filepath="Enter or drag-and-drop the path to .md file: "
        set "filepath=!filepath:"=!"
        "%PYTHON_EXEC%" "%BUILD_SCRIPT%" "!filepath!"
    ) else if "!choice!"=="2" (
        set /p filepath="Enter or drag-and-drop the path to .docx file: "
        set "filepath=!filepath:"=!"
        "%PYTHON_EXEC%" "%FORMAT_SCRIPT%" "!filepath!"
    ) else (
        echo Invalid choice!
    )
    echo.
    pause
    exit /b
)

:process_files
if "%~1"=="" goto end

set "FILE_EXT=%~x1"

if /i "%FILE_EXT%"==".md" (
    echo [~] Building Markdown: %~n1%FILE_EXT%
    "%PYTHON_EXEC%" "%BUILD_SCRIPT%" "%~1"
) else if /i "%FILE_EXT%"==".docx" (
    echo [~] Formatting Word: %~n1%FILE_EXT%
    "%PYTHON_EXEC%" "%FORMAT_SCRIPT%" "%~1"
) else (
    echo [!] SKIPPED: Unknown format %FILE_EXT% ^(%~nx1^)
)

shift
goto process_files

:end
echo.
echo ===================================================
echo ALL FILES PROCESSED! Documents successfully generated.
pause
