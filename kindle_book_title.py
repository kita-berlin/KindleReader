# MIT License
# Copyright (c) 2025 Quantrosoft
# See LICENSE file for full license text.

"""
Kindle Open-Book Title
======================
Determines the title of the book that is currently open in Kindle for PC and
creates the book folder for it (folder name = book title).

Source is Kindle's own log - the Kindle window title is just 'Kindle' and carries
no book name:
    %LOCALAPPDATA%\\Packages\\AMZNKindle.AmazonKindleReadingApp_*\\LocalState\\logs\\kindle.log

What Kindle writes there (measured in kindle.log, app 1.0.25218.0):
    [2026-09-29 17:50:03.055] [info] mp.15.5                         <- every app start
    [...] [YJLocalBookItem] [loadMetadata] ASIN: B0096CCRG8           <- metadata of a book;
    [...] [YJLocalBookItem] [loadMetadata] title: Cybernetic ...     <- title on the NEXT line
    [...] [Renderer] Open book C:/.../Content/B0096CCRG8_EBOK/B0096CCRG8_EBOK.azw
                                                                      <- book opened in the reader
    [...] InBookSearchWorker: Worker thread exiting: uid_1            <- reader left for the
    [...] [GetLibraryController] Loading library started:            <- library (both lines)
The OPEN marker is the Renderer line, NOT '[LibraryItemOpener] openItem': openItem is
only written when the user opens a book from the library. When Kindle restarts and
restores the open book itself, there is no openItem (measured: session 18:23:49, book
restored at 18:23:53 - 0 openItem lines, run aborted with 'kein Buch geoeffnet').
The Renderer line was present at all 4 opens in the log: 3 user opens (2-3 s after
their openItem) and that restore.
The two close lines were measured 2026-07-19 12:07:30-31, between the opens of two
different books; while a book stayed open (12:07:42-16:33:28) neither line occurred,
nor in the restored session (0 of 1122 lines).
loadMetadata is also written for books that are NOT opened (measured: two such books
at 17:50:07, right after the app start), so the title is looked up via the ASIN of the
last Renderer open, never taken from the last title line. Timestamps are local time
(log 18:00:38 = breadcrumb 16:00:38Z). Lines of 2026-07-19 carry an extra
'[MazamaReader] ' before the component tag, lines of 2026-09-29 do not; the patterns
below are not anchored, so both forms match.

Checks - each one fails loud (exit 1), nothing is guessed:
  1. exactly one Kindle.exe is running;
  2. the log's last app start lies within SESSION_START_TOLERANCE_S of that process's
     start time, i.e. the log belongs to the running instance (measured: 0.9 s and 0.5 s);
  3. a book was opened in this session (Renderer open after the last app start);
  4. it was not closed again (no close marker after that open);
  5. a loadMetadata ASIN/title pair exists for its ASIN in this session.

Usage:
    python kindle_book_title.py <root folder>

Creates <root folder>\\<title> if missing and prints its absolute path as the ONLY
line on stdout (UTF-8; scan.bat captures it). All messages go to stderr.
"""

import ctypes
import re
import sys
from ctypes import wintypes
from datetime import datetime
from pathlib import Path

SESSION_START_TOLERANCE_S = 60

LOG_GLOB = "Packages/AMZNKindle.AmazonKindleReadingApp_*/LocalState/logs/kindle.log"

RE_TIMESTAMP = re.compile(r"^\[(\d{4}-\d\d-\d\d \d\d:\d\d:\d\d\.\d{3})\]")
RE_APP_START = re.compile(r"^\[[^\]]+\] \[info\] mp\.\d")
RE_META_ASIN = re.compile(r"\[loadMetadata\] ASIN: (\S+)")
RE_META_TITLE = re.compile(r"\[loadMetadata\] title: (.+)$")
RE_OPEN_BOOK = re.compile(r"\[Renderer\] Open book .*/Content/([^/_]+)_[^/]*/")
CLOSE_MARKERS = ("InBookSearchWorker: Worker thread exiting",
                 "[GetLibraryController] Loading library started")


def fail(msg):
    print(f"[FEHLER] {msg}", file=sys.stderr)
    sys.exit(1)


def info(msg):
    print(f"[INFO] {msg}", file=sys.stderr)


# ============================================================
# Kindle process (stdlib ctypes - runs before scan.bat installs requirements)
# ============================================================

TH32CS_SNAPPROCESS = 0x00000002
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
INVALID_HANDLE_VALUE = wintypes.HANDLE(-1).value


class PROCESSENTRY32W(ctypes.Structure):
    _fields_ = [("dwSize", wintypes.DWORD),
                ("cntUsage", wintypes.DWORD),
                ("th32ProcessID", wintypes.DWORD),
                ("th32DefaultHeapID", ctypes.c_size_t),
                ("th32ModuleID", wintypes.DWORD),
                ("cntThreads", wintypes.DWORD),
                ("th32ParentProcessID", wintypes.DWORD),
                ("pcPriClassBase", ctypes.c_long),
                ("dwFlags", wintypes.DWORD),
                ("szExeFile", wintypes.WCHAR * 260)]


_k32 = ctypes.WinDLL("kernel32", use_last_error=True)
_k32.CreateToolhelp32Snapshot.restype = wintypes.HANDLE
_k32.CreateToolhelp32Snapshot.argtypes = [wintypes.DWORD, wintypes.DWORD]
_k32.Process32FirstW.argtypes = [wintypes.HANDLE, ctypes.POINTER(PROCESSENTRY32W)]
_k32.Process32NextW.argtypes = [wintypes.HANDLE, ctypes.POINTER(PROCESSENTRY32W)]
_k32.OpenProcess.restype = wintypes.HANDLE
_k32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
_k32.GetProcessTimes.argtypes = [wintypes.HANDLE] + [ctypes.POINTER(wintypes.FILETIME)] * 4
_k32.CloseHandle.argtypes = [wintypes.HANDLE]


def kindle_pids():
    """PIDs of all running processes named Kindle.exe."""
    snap = _k32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    if snap == INVALID_HANDLE_VALUE:
        fail(f"CreateToolhelp32Snapshot fehlgeschlagen (WinError {ctypes.get_last_error()}).")
    try:
        entry = PROCESSENTRY32W()
        entry.dwSize = ctypes.sizeof(PROCESSENTRY32W)
        pids = []
        ok = _k32.Process32FirstW(snap, ctypes.byref(entry))
        while ok:
            if entry.szExeFile.lower() == "kindle.exe":
                pids.append(entry.th32ProcessID)
            ok = _k32.Process32NextW(snap, ctypes.byref(entry))
        return pids
    finally:
        _k32.CloseHandle(snap)


def process_start_local(pid):
    """Start time of a process as naive local datetime (same form as the log)."""
    handle = _k32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not handle:
        fail(f"OpenProcess({pid}) fehlgeschlagen (WinError {ctypes.get_last_error()}).")
    try:
        created, exited, kernel, user = (wintypes.FILETIME() for _ in range(4))
        if not _k32.GetProcessTimes(handle, ctypes.byref(created), ctypes.byref(exited),
                                    ctypes.byref(kernel), ctypes.byref(user)):
            fail(f"GetProcessTimes({pid}) fehlgeschlagen (WinError {ctypes.get_last_error()}).")
    finally:
        _k32.CloseHandle(handle)
    ticks = (created.dwHighDateTime << 32) | created.dwLowDateTime  # 100 ns since 1601 UTC
    return datetime.fromtimestamp(ticks / 1e7 - 11644473600)


# ============================================================
# Kindle log
# ============================================================

def find_log():
    local = Path.home() / "AppData" / "Local"
    logs = sorted(local.glob(LOG_GLOB))
    if len(logs) != 1:
        fail(f"Erwartet genau ein Kindle-Log unter {local}\\{LOG_GLOB}, gefunden {len(logs)}: "
             f"{[str(p) for p in logs]}")
    return logs[0]


def log_time(line):
    m = RE_TIMESTAMP.match(line)
    return datetime.strptime(m.group(1), "%Y-%m-%d %H:%M:%S.%f") if m else None


def open_book_from_lines(lines):
    """Parse the log lines. Returns (session_start, asin, title) of the book open in
    the LAST session, or fails loud. Pure function - no process/file access."""
    starts = [i for i, line in enumerate(lines) if RE_APP_START.match(line)]
    if not starts:
        fail("Kindle-Log enthaelt keine App-Start-Zeile ('[...] [info] mp.<version>').")
    session = lines[starts[-1]:]
    session_start = log_time(session[0])

    opens = [(i, m.group(1)) for i, line in enumerate(session)
             for m in [RE_OPEN_BOOK.search(line)] if m]
    if not opens:
        fail(f"Seit dem Kindle-Start ({session_start:%Y-%m-%d %H:%M:%S}) wurde kein Buch "
             "geoeffnet. Buch in Kindle oeffnen, dann neu starten.")
    open_idx, asin = opens[-1]

    for line in session[open_idx + 1:]:
        for marker in CLOSE_MARKERS:
            if marker in line:
                fail(f"Das zuletzt geoeffnete Buch (ASIN {asin}) wurde wieder geschlossen "
                     f"('{marker}' um {log_time(line):%H:%M:%S}). "
                     "Buch in Kindle oeffnen, dann neu starten.")

    title = None
    for i, line in enumerate(session[:open_idx]):
        m = RE_META_ASIN.search(line)
        if m and m.group(1) == asin and i + 1 < len(session):
            t = RE_META_TITLE.search(session[i + 1])
            if t:
                title = t.group(1).strip()
    if not title:
        fail(f"Keine Zeile '[loadMetadata] ASIN: {asin}' mit folgender Titel-Zeile in dieser "
             "Kindle-Sitzung.")
    return session_start, asin, title


def folder_name_from_title(title):
    """Title -> Windows folder name. ':' becomes ' - ' (e.g. 'Algorithmic Trading: 625'
    -> 'Algorithmic Trading - 625'); other characters Windows forbids are replaced."""
    name = re.sub(r"\s*:\s*", " - ", title)
    name = name.replace('"', "'")
    name = re.sub(r"[/\\|]", "-", name)
    name = re.sub(r"[<>*?\x00-\x1f]", "", name)
    name = re.sub(r"\s+", " ", name).strip().rstrip(". ")
    if not name:
        fail(f"Titel {title!r} ergibt einen leeren Ordnernamen.")
    return name


# ============================================================
# Main
# ============================================================

def main():
    if len(sys.argv) != 2:
        fail("Aufruf: python kindle_book_title.py <Wurzelordner>")
    root = Path(sys.argv[1])
    if not root.is_dir():
        fail(f"Wurzelordner nicht gefunden: {root}")

    pids = kindle_pids()
    if len(pids) != 1:
        fail(f"Erwartet genau ein laufendes Kindle.exe, gefunden {len(pids)} (PIDs {pids}). "
             "Kindle muss LAUFEN und das Buch geoeffnet sein.")
    proc_start = process_start_local(pids[0])

    log = find_log()
    lines = log.read_text(encoding="utf-8").splitlines()
    session_start, asin, title = open_book_from_lines(lines)

    delta = (session_start - proc_start).total_seconds()
    if abs(delta) > SESSION_START_TOLERANCE_S:
        fail(f"Kindle-Log gehoert nicht zum laufenden Kindle: letzter App-Start im Log "
             f"{session_start:%Y-%m-%d %H:%M:%S}, Kindle.exe (PID {pids[0]}) gestartet "
             f"{proc_start:%Y-%m-%d %H:%M:%S} ({delta:+.1f} s, Toleranz "
             f"{SESSION_START_TOLERANCE_S} s). Log: {log}")

    folder = (root / folder_name_from_title(title)).resolve()
    info(f"Offenes Buch: {title} (ASIN {asin})")
    if folder.is_dir():
        info(f"Buch-Ordner existiert - vorhandene Ausgaben werden uebersprungen: {folder}")
    else:
        folder.mkdir()
        info(f"Buch-Ordner angelegt: {folder}")

    sys.stdout.reconfigure(encoding="utf-8")
    print(folder)


if __name__ == "__main__":
    main()
