#!/usr/bin/env python3
# MIT License
# Copyright (c) 2025 Quantrosoft
# See LICENSE file for full license text.

"""
Build script for Kindle Capture Tools using Nuitka.

Creates standalone executables:
- kindle_capture.exe - Captures Kindle pages as PNG images
- create_pdf.exe - Creates searchable PDF from captured pages

Build time: 5-10 minutes per EXE
Output: dist/

Author: Claude
"""

import os
import shutil
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

# Force UTF-8 output for Windows console
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')

def log_progress(message, status='INFO'):
    """Log progress with timestamp"""
    timestamp = datetime.now().strftime('%H:%M:%S')
    print(f"[{timestamp}] [{status}] {message}")

def log_section(title):
    """Print a section header"""
    print('\n' + '=' * 60)
    print(f"  {title}")
    print('=' * 60 + '\n')

def build_exe(project_root, script_name, exe_name, packages):
    """Build a single executable with Nuitka"""
    dist_dir = project_root / "dist"
    script_path = project_root / script_name

    if not script_path.exists():
        log_progress(f"FEHLER: Script nicht gefunden: {script_path}", 'ERROR')
        return False

    log_progress(f"Starte Nuitka Build fuer {exe_name}...", 'PROGRESS')
    log_progress("Dies kann einige Minuten dauern (Python -> C Kompilierung)...", 'INFO')

    nuitka_start_time = time.time()

    cmd = [
        sys.executable, "-m", "nuitka",
        "--standalone",
        "--onefile",
        "--jobs=24",
        "--assume-yes-for-downloads",
        "--remove-output",
        "--output-dir=" + str(dist_dir),
        f"--output-filename={exe_name}",
        "--company-name=Quantrosoft Pte. Ltd.",
        "--product-name=Kindle Book Capture Tools",
        "--file-version=1.1.0.0",
        "--product-version=1.1.0.0",
        f"--file-description=Quantrosoft {exe_name.replace('.exe', '')} - Kindle eBook Tool",
        "--copyright=Copyright (c) 2025 Quantrosoft Pte. Ltd.",
        "--trademarks=Quantrosoft is a trademark of Quantrosoft Pte. Ltd.",
    ]

    for pkg in packages:
        cmd.append(f"--include-package={pkg}")

    cmd.append(str(script_path))

    try:
        subprocess.check_call(cmd)
    except subprocess.CalledProcessError as e:
        log_progress(f"Nuitka Build fehlgeschlagen: {e}", 'ERROR')
        return False

    nuitka_duration = time.time() - nuitka_start_time
    exe_file = dist_dir / exe_name

    if exe_file.exists():
        exe_size = exe_file.stat().st_size / (1024 * 1024)
        log_progress(f"Build abgeschlossen: {exe_name} ({exe_size:.2f} MB, {nuitka_duration:.1f}s)", 'OK')
        return True
    else:
        log_progress(f"Build fehlgeschlagen: {exe_name}", 'ERROR')
        return False

def main():
    print('\n' * 10)
    print('=' * 60)
    print('   KINDLE TO PDF - BUILD PROCESS')
    print('=' * 60 + '\n')

    build_start_time = time.time()
    project_root = Path(__file__).parent
    dist_dir = project_root / "dist"
    build_dir = project_root / "build"

    log_section('Kindle to PDF - EXE Build (Nuitka)')
    log_progress('Build gestartet', 'INFO')

    # Kill any running processes
    log_progress('Pruefe laufende Prozesse...', 'PROGRESS')
    if sys.platform == 'win32':
        for exe in ['kindle_capture.exe', 'create_pdf.exe']:
            try:
                result = subprocess.run(
                    ['taskkill', '/F', '/IM', exe],
                    capture_output=True,
                    text=True
                )
                if result.returncode == 0:
                    log_progress(f'{exe} beendet', 'OK')
                    time.sleep(1)
            except Exception:
                pass

    # Clean previous builds
    log_progress('Bereinige alte Build-Artefakte...', 'PROGRESS')

    if dist_dir.exists():
        for exe_name in ['kindle_capture.exe', 'create_pdf.exe']:
            exe_file = dist_dir / exe_name
            if exe_file.exists():
                old_size = exe_file.stat().st_size / (1024 * 1024)
                log_progress(f"Loesche alte {exe_name} ({old_size:.2f} MB)...", 'PROGRESS')
                exe_file.unlink()
                time.sleep(1)
                if exe_file.exists():
                    log_progress(f"FEHLER: Konnte {exe_name} nicht loeschen!", 'ERROR')
                    sys.exit(1)

    if build_dir.exists():
        log_progress("Bereinige build/ Verzeichnis...", 'PROGRESS')
        shutil.rmtree(build_dir)

    dist_dir.mkdir(exist_ok=True)
    log_progress("dist/ Verzeichnis bereit", 'OK')

    # Setup ccache if available
    ccache_dir = project_root.parent.parent / "MCPServer" / "tools" / "ccache-4.10.2-windows-x86_64"
    if ccache_dir.exists():
        ccache_exe = ccache_dir / "ccache.exe"
        os.environ["PATH"] = str(ccache_dir) + os.pathsep + os.environ["PATH"]
        os.environ["NUITKA_CCACHE_BINARY"] = str(ccache_exe)
        log_progress(f"ccache konfiguriert: {ccache_exe}", 'OK')
    else:
        log_progress("ccache nicht gefunden (Build wird langsamer sein)", 'INFO')

    # Install Nuitka if not available
    log_progress("Pruefe Nuitka Installation...", 'PROGRESS')
    try:
        import nuitka
        log_progress("Nuitka bereits installiert", 'OK')
    except ImportError:
        log_progress("Installiere Nuitka & zstandard...", 'PROGRESS')
        subprocess.check_call([sys.executable, "-m", "pip", "install", "nuitka", "zstandard"])
        log_progress("Nuitka installiert", 'OK')

    # Build kindle_capture.exe
    log_section('Build: kindle_capture.exe')
    capture_packages = [
        "pyautogui",
        "pygetwindow",
        "pywinauto",
        "PIL",
        "numpy",
        "pynput",
        "winsdk",
    ]
    success1 = build_exe(project_root, "kindle_capture.py", "kindle_capture.exe", capture_packages)

    # Build create_pdf.exe
    log_section('Build: create_pdf.exe')
    pdf_packages = [
        "PIL",
        "reportlab",
        "winsdk",
    ]
    success2 = build_exe(project_root, "create_pdf.py", "create_pdf.exe", pdf_packages)

    success = success1 and success2

    # Create batch files
    log_progress("Erstelle Batch-Dateien...", 'PROGRESS')

    # Capture batch - uses %~dp0 for portable relative path to batch file's directory
    capture_batch = dist_dir / "capture_book.bat"
    capture_batch.write_text('''@echo off
REM Kindle Book Capture
REM Usage: Run from the target folder where pages should be saved

"%~dp0kindle_capture.exe"

pause
''')
    log_progress(f"capture_book.bat erstellt", 'OK')

    # PDF batch
    pdf_batch = dist_dir / "create_pdf.bat"
    pdf_batch.write_text('''@echo off
REM Create PDF from captured pages
REM Usage: Run from the folder containing page_*.png files

"%~dp0create_pdf.exe"

pause
''')
    log_progress(f"create_pdf.bat erstellt", 'OK')

    # Summary
    total_duration = time.time() - build_start_time

    log_section('Build-Ergebnis')

    for exe_name in ['kindle_capture.exe', 'create_pdf.exe']:
        exe = dist_dir / exe_name
        if exe.exists():
            size = exe.stat().st_size / (1024 * 1024)
            print(f"  {exe_name}: {size:.2f} MB")
        else:
            print(f"  {exe_name}: FEHLER")

    print(f"\n  Output: {dist_dir}")
    print(f"  Gesamtzeit: {total_duration:.1f}s")

    if success:
        print('\n' + '=' * 60)
        print('  BUILD ERFOLGREICH ABGESCHLOSSEN!')
        print('=' * 60 + '\n')
        print("Verwendung:")
        print("  1. Wechsle in den Zielordner fuer das Buch")
        print("  2. Oeffne Kindle mit dem gewuenschten Buch")
        print("  3. Fuehre capture_book.bat aus")
        print("  4. Fuehre create_pdf.bat aus")
        print()
    else:
        print('\n' + '=' * 60)
        print('  BUILD FEHLGESCHLAGEN!')
        print('=' * 60 + '\n')
        sys.exit(1)

if __name__ == "__main__":
    main()
