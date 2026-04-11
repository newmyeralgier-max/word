@echo off
setlocal enabledelayedexpansion
title Ultimate GOST-Formatter

:: Устанавливаем путь к Python. Если он не найдется по этому пути, скрипт будет искать его в системных переменных PATH
set "PYTHON_EXEC=C:\Users\New-life\AppData\Local\Programs\Python\Python313\python.exe"
if not exist "%PYTHON_EXEC%" (
    set "PYTHON_EXEC=python"
)

:: Указываем пути к нашим Python-скриптам относительно расположения этого .bat файла
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
    echo 1 - Build GOST document (Modern 7.32-2017)
    echo 2 - Build GOST document (Legacy / Professor style)
    echo 3 - Format existing Word document ^(.docx^)
    echo.
    set /p choice="Enter 1, 2, or 3 and press Enter: "
    echo.
    if "!choice!"=="1" (
        set /p filepath="Enter or drag-and-drop the path to .md file: "
        set "filepath=!filepath:"=!"
        "%PYTHON_EXEC%" "%BUILD_SCRIPT%" -i "!filepath!"
    ) else if "!choice!"=="2" (
        set /p filepath="Enter or drag-and-drop the path to .md file: "
        set "filepath=!filepath:"=!"
        "%PYTHON_EXEC%" "%BUILD_SCRIPT%" -i "!filepath!" --legacy
    ) else if "!choice!"=="3" (
        set /p filepath="Enter or drag-and-drop the path to .docx file: "
        set "filepath=!filepath:"=!"
        "%PYTHON_EXEC%" "%FORMAT_SCRIPT%" -i "!filepath!"
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
    :: ИСПРАВЛЕНИЕ: Добавлен флаг -i для Drag-and-Drop
    "%PYTHON_EXEC%" "%BUILD_SCRIPT%" -i "%~1"
) else if /i "%FILE_EXT%"==".docx" (
    echo [~] Formatting Word: %~n1%FILE_EXT%
    :: ИСПРАВЛЕНИЕ: Добавлен флаг -i для Drag-and-Drop
    "%PYTHON_EXEC%" "%FORMAT_SCRIPT%" -i "%~1"
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