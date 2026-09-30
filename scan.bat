@echo off
REM MIT License
REM Copyright (c) 2025 Quantrosoft
REM See LICENSE file for full license text.
REM
REM Kindle Book Capture, PDF Creation & Markdown for AI
REM Usage: scan.bat [book folder]
REM   No argument (e.g. double-click): the title of the book currently open in Kindle
REM   is read from Kindle's own log (kindle_book_title.py) and the book folder
REM   KINDLE_ROOT\<title> is created if missing.
REM   With argument: that existing book folder is used (e.g. to redo PDF/markdown).
REM Output (pages\, PDF, markdown\) always lands in the book folder, no matter from
REM which directory scan.bat is started. Skips steps if output already exists.

REM Get the directory where this batch file is located
set "BOOKREADER=%~dp0"

REM Root under which the book folder is created when no argument is given
set "KINDLE_ROOT=D:\GoogleDriveData\ShareFile\eBooks\Trading\_Kindle"

REM --- Make sure Python is on PATH (before the book-folder step, which runs Python) ---
set "PYTHON_HOME=%LOCALAPPDATA%\Programs\Python\Python313"
if exist "%PYTHON_HOME%\python.exe" (
    set "PATH=%PYTHON_HOME%;%PYTHON_HOME%\Scripts;%PATH%"
)

REM --- Book folder (output) ---
REM Every path/name echoed below is QUOTED. Kindle titles contain ( ) and & (e.g.
REM "... (Wiley Trading)"); unquoted inside a ( ... ) block the ')' ends the block.
REM Measured 29.09.2026: folder "X (Wiley Trading)" -> '".pdf" kann syntaktisch an
REM dieser Stelle nicht verarbeitet werden.', exit 255; "X Wiley Trading" -> ran through.
if "%~1"=="" goto AUTO_TARGET
cd /d "%~1"
if errorlevel 1 ( echo [FEHLER] Ordner nicht gefunden: "%~1" & pause & exit /b 1 )
goto TARGET_OK
:AUTO_TARGET
echo [INFO] Kein Buch-Ordner uebergeben - lese Titel des offenen Buchs aus dem Kindle-Log...
REM kindle_book_title.py prints the folder path as UTF-8 on stdout, and for /f decodes
REM it with the console code page. Measured 29.09.2026 in a real console: under the
REM default code page 850 a path with an umlaut came back broken ('cd' failed), under
REM 65001 it was intact. So switch to 65001 for the capture only, then restore.
for /f "tokens=2 delims=:." %%C in ('chcp') do set "OLDCP=%%C"
chcp 65001 >nul
set "TARGET="
for /f "usebackq delims=" %%I in (`python "%BOOKREADER%kindle_book_title.py" "%KINDLE_ROOT%"`) do set "TARGET=%%I"
chcp %OLDCP% >nul
if not defined TARGET ( echo [FEHLER] Kein Buch-Ordner - siehe Fehlermeldung oben. & pause & exit /b 1 )
cd /d "%TARGET%"
if errorlevel 1 ( echo [FEHLER] Ordner ungueltig: "%TARGET%" & pause & exit /b 1 )
:TARGET_OK
echo [INFO] Buch-Ordner (Ausgabe): "%CD%"
echo.

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
if exist "markdown\%BOOKNAME%.md" set HAS_MD=1

if %HAS_PAGES%==1 if %HAS_PDF%==1 if %HAS_MD%==1 (
    echo [INFO] Alles vorhanden - nichts zu tun.
    echo   Pages: pages\
    echo   PDF:   "%BOOKNAME%.pdf"
    echo   MD:    "markdown\%BOOKNAME%.md"
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
