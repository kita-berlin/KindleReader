@echo off
REM MIT License
REM Copyright (c) 2025 Quantrosoft
REM See LICENSE file for full license text.
REM
REM Kindle Book Capture, PDF Creation & Markdown for AI
REM Usage: Run from the book folder (e.g., "Order Flow & Volume Profile")
REM Skips steps if output already exists

REM Get the directory where this batch file is located
set BOOKREADER=%~dp0

REM --- Python im PATH sicherstellen ---
set "PYTHON_HOME=C:\Users\hmunz\AppData\Local\Programs\Python\Python314"
if exist "%PYTHON_HOME%\python.exe" (
    set "PATH=%PYTHON_HOME%;%PYTHON_HOME%\Scripts;%PATH%"
)

REM --- SCHRITT 0: ABHAENGIGKEITEN ---
echo [INFO] Pruefe Python-Abhaengigkeiten...
pip install -r "%BOOKREADER%requirements.txt"
if errorlevel 1 (
    echo [FEHLER] pip install fehlgeschlagen!
    echo [INFO] Bitte manuell ausfuehren: pip install -r "%BOOKREADER%requirements.txt"
    pause
    exit /b 1
)
echo [OK] Alle Abhaengigkeiten installiert.
echo.

REM Get folder name for PDF filename
for %%I in (.) do set BOOKNAME=%%~nxI

REM Check what already exists
set HAS_PAGES=0
set HAS_PDF=0
set HAS_MD=0

if exist "pages\page_*.png" set HAS_PAGES=1
if exist "%BOOKNAME%.pdf" set HAS_PDF=1
if exist "markdown\book.md" set HAS_MD=1

if %HAS_PAGES%==1 if %HAS_PDF%==1 if %HAS_MD%==1 (
    echo [INFO] Alles vorhanden - nichts zu tun.
    echo   Pages: pages\
    echo   PDF:   %BOOKNAME%.pdf
    echo   MD:    markdown\book.md
    pause
    exit /b 0
)

REM --- SCHRITT 1: CAPTURE ---

if %HAS_PAGES%==1 (
    echo [SKIP] Pages existieren bereits - ueberspringe Scan.
    echo.
) else (
    echo ============================================================
    echo   KINDLE BUCH ERFASSUNG
    echo ============================================================
    echo.

    python "%BOOKREADER%\kindle_capture.py"

    if errorlevel 1 (
        echo.
        echo [INFO] Erfassung wurde gestoppt.
        pause
        exit /b 1
    )
    echo.
)

REM --- SCHRITT 2: PDF ---

if %HAS_PDF%==1 (
    echo [SKIP] PDF existiert bereits - ueberspringe PDF-Erstellung.
    echo.
) else (
    echo ============================================================
    echo   ERSTELLE DURCHSUCHBARES PDF
    echo ============================================================
    echo.

    python "%BOOKREADER%\create_pdf.py"

    if errorlevel 1 (
        echo.
        echo [FEHLER] PDF-Erstellung fehlgeschlagen!
        pause
        exit /b 1
    )
    echo.
)

REM --- SCHRITT 3: MARKDOWN ---

if %HAS_MD%==1 (
    echo [SKIP] Markdown existiert bereits - ueberspringe Markdown-Erstellung.
    echo.
) else (
    echo ============================================================
    echo   ERSTELLE MARKDOWN FUER KI
    echo ============================================================
    echo.

    python "%BOOKREADER%\create_markdown.py"

    if errorlevel 1 (
        echo.
        echo [FEHLER] Markdown-Erstellung fehlgeschlagen!
        pause
        exit /b 1
    )
    echo.
)

echo ============================================================
echo   FERTIG!
echo ============================================================
pause
