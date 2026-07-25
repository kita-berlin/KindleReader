@echo off
REM MIT License
REM Copyright (c) 2025 Quantrosoft
REM See LICENSE file for full license text.
REM
REM Batch PDF -> Markdown (KindleReader-Methode, ohne Kindle-Erfassung).
REM Aufruf:  pdf2md.bat <Wurzelordner>
REM Konvertiert rekursiv alle *.pdf unter dem Wurzelordner nach
REM <pdf_ordner>\markdown\<pdfname>\<pdfname>.md (+ Chart-Seiten als JPG).
REM Log geht per tee nach <Wurzelordner>\pdf2md.log.

set "BOOKREADER=%~dp0"

if "%~1"=="" (
    echo [FEHLER] Wurzelordner fehlt. Aufruf: pdf2md.bat ^<Wurzelordner^>
    pause
    exit /b 1
)
if not exist "%~1\" (
    echo [FEHLER] Ordner nicht gefunden: %~1
    pause
    exit /b 1
)

set "PYTHON_HOME=%LOCALAPPDATA%\Programs\Python\Python313"
if exist "%PYTHON_HOME%\python.exe" (
    set "PATH=%PYTHON_HOME%;%PYTHON_HOME%\Scripts;%PATH%"
)

echo [INFO] Pruefe Python-Abhaengigkeiten...
pip install -r "%BOOKREADER%requirements.txt" >nul
if errorlevel 1 (
    echo [FEHLER] pip install fehlgeschlagen!
    pause
    exit /b 1
)

REM pwsh (PowerShell 7) statt Windows-PowerShell 5.1: dessen Tee-Object schreibt
REM das Log als UTF-16 und macht es fuer grep & Co. unbrauchbar; pwsh schreibt UTF-8.
where pwsh >nul 2>&1
if errorlevel 1 (
    echo [FEHLER] pwsh ^(PowerShell 7^) nicht gefunden - wird fuer UTF-8-Log gebraucht!
    pause
    exit /b 1
)

title pdf2md %~1
pwsh -NoProfile -Command "& python -u '%BOOKREADER%batch_pdf2md.py' '%~1' 2>&1 | Tee-Object -FilePath '%~1\pdf2md.log'"
echo.
echo [INFO] Fertig. Log: %~1\pdf2md.log
pause
