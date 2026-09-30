# MIT License
# Copyright (c) 2025 Quantrosoft
# See LICENSE file for full license text.

"""
Kindle Book Capture Tool
========================
Captures all pages from a Kindle book as PNG images.

Works with the new WinUI-3 Kindle for PC, which has NO menu bar - navigation is
done via hotkeys. The reader renders inside a WinUI content bridge child window
that only responds to F11 / the page keys when it has keyboard focus, and the only
reliable way to focus it is a single click in the reading area. So the tool clicks
ONCE to focus the reader, presses F11 (fullscreen resets to a clean page - the
click's toolbar chrome does not carry in), then navigates with keys only while the
mouse stays parked in a neutral spot. Ctrl+G is avoided (its dialog steals the
reader's focus); the cover is reached by pressing PageUp until the page stops
changing. Pages are captured with PrintWindow, which works even when Kindle's
fullscreen blocks the normal GDI screen grab.

Usage:
1. Open Kindle app with the book you want to capture (windowed, book loaded)
2. Change to the book folder (used as output folder)
3. Run: kindle_capture.exe
4. The script will automatically:
   - Find and activate the Kindle window
   - Click once to focus the reader, then F11 -> clean fullscreen (fullscreen and
     mouse parking are computed against the monitor the KINDLE WINDOW is on, so
     multi-monitor setups work)
   - PROVE the page keys reach the reader (a PageDown or PageUp must visibly turn
     the page; if neither does, fail loud with a _debug_fokus.png evidence shot
     instead of misreading 'no change' as 'cover reached')
   - Navigate to the very beginning / cover (PageUp until the page stops changing)
   - Detect the page format from the title page: the page is what lies between the
     uniform black/white letterbox bars left and right, over the full window height
     (only left/right are searched; letterbox-coloured print ON the page does not
     split it - see the function)
   - Capture every page CROPPED to that title-page format (PrintWindow), paging
     forward with PageDown - so all pages have the same format as the cover
   - End of book = the content no longer changes when paging forward, decided by an
     exact pixel comparison without any threshold

Author: Claude
"""

import pyautogui
pyautogui.FAILSAFE = False  # Disable fail-safe (mouse in corner)
import pygetwindow as gw
import time
import sys
import signal
from PIL import Image
import numpy as np
from pathlib import Path
from pynput import keyboard

# pywin32 for PrintWindow-based window capture. ESSENTIAL, not optional: the new
# Kindle's fullscreen can enter an exclusive/protected mode where GDI screen grab
# (PIL.ImageGrab) fails or returns black. PrintWindow(PW_RENDERFULLCONTENT) reads
# the window's own rendering (WinUI + WebView2 content) and works regardless.
try:
    import win32gui
    import win32ui
    import win32api
    import win32con
except ImportError as e:
    print("[FEHLER] PYWIN32 NICHT INSTALLIERT!")
    print("[FEHLER] BEFEHL: pip install pywin32")
    print(f"[FEHLER] Details: {e}")
    sys.exit(1)

# ============================================================
# Configuration
# ============================================================
WAIT_AFTER_PAGE = 0.5  # Seconds to wait after a page-turn keypress (also the render poll interval)
RENDER_POLL_TRIES = 5  # Max polls to wait for a page to render/advance before concluding "no advance"

# Global flag for immediate stop
STOP_FLAG = False
keyboard_listener = None

# ============================================================
# Keyboard and Signal Handling
# ============================================================

# Only these keys will stop the script (normal typing keys)
STOP_KEYS = {keyboard.Key.esc, keyboard.Key.space, keyboard.Key.enter}

def on_key_press(key):
    """Global keyboard hook - only Esc/Space/Enter or letter/number keys stop the script."""
    global STOP_FLAG
    # Character keys (letters, numbers, punctuation) -> stop
    if isinstance(key, keyboard.KeyCode):
        STOP_FLAG = True
        print("\n[!] Taste gedrueckt - stoppe...")
        return False
    # Specific stop keys -> stop
    if key in STOP_KEYS:
        STOP_FLAG = True
        print("\n[!] Taste gedrueckt - stoppe...")
        return False
    # Everything else (F-keys, media keys, modifiers, etc.) -> ignore
    return True

def start_keyboard_listener():
    """Start global keyboard listener."""
    global keyboard_listener
    keyboard_listener = keyboard.Listener(on_press=on_key_press)
    keyboard_listener.start()

def stop_keyboard_listener():
    """Stop global keyboard listener."""
    global keyboard_listener
    if keyboard_listener:
        keyboard_listener.stop()
        keyboard_listener = None

def signal_handler(signum, frame):
    """Handle Ctrl+C signal."""
    global STOP_FLAG
    STOP_FLAG = True
    print("\n[!] Ctrl+C empfangen - stoppe...")

signal.signal(signal.SIGINT, signal_handler)

def check_stop():
    """Check if stop was requested."""
    return STOP_FLAG

def check_stop_and_exit():
    """Check if stop was requested and exit if so."""
    if STOP_FLAG:
        stop_keyboard_listener()
        print("\n[GESTOPPT] Script vom Benutzer gestoppt.")
        sys.exit(1)

# ============================================================
# Kindle Window Control
# ============================================================

def _get_window_class(hwnd):
    """Get Win32 window class name for a hwnd."""
    import ctypes
    buf = ctypes.create_unicode_buffer(256)
    ctypes.windll.user32.GetClassNameW(hwnd, buf, 256)
    return buf.value

# Window classes to exclude (Explorer, WebView2, etc.)
_EXCLUDED_CLASSES = {'CabinetWClass', 'ExplorerWClass', 'Shell_TrayWnd', 'Progman'}

def get_kindle_window():
    """Find main Kindle window by filtering out Explorer and WebView2 windows.
    The Kindle title contains 'Kindle' - but so does an Explorer window showing
    a folder path with 'Kindle' in it. We use the window class to distinguish."""
    windows = gw.getWindowsWithTitle('Kindle')
    if not windows:
        return None

    # Filter by window class: exclude Explorer, keep Kindle app windows
    kindle_windows = [w for w in windows if _get_window_class(w._hWnd) not in _EXCLUDED_CLASSES]
    if not kindle_windows:
        return None

    # Prefer non-minimized windows, but accept minimized if it's the only one
    non_minimized = [w for w in kindle_windows if not w.isMinimized]
    if non_minimized:
        return max(non_minimized, key=lambda w: w.width * w.height)

    return kindle_windows[0]


def activate_and_get_kindle():
    """Activate Kindle window and return (left, top, width, height)."""
    kindle = get_kindle_window()
    if kindle:
        try:
            kindle.activate()
        except Exception:
            pass
        time.sleep(0.1)
        return (kindle.left, kindle.top, kindle.width, kindle.height)
    return None


def exit_fullscreen_and_minimize():
    """Exit fullscreen (F11) and minimize Kindle."""
    kindle = get_kindle_window()
    if kindle:
        try:
            kindle.activate()
            time.sleep(0.3)
            pyautogui.press('f11')  # leave fullscreen reading mode
            time.sleep(0.5)
            kindle.minimize()
            print("[INFO] Kindle minimiert")
        except Exception as e:
            print(f"[WARNUNG] Konnte Kindle nicht minimieren: {e}")


# ============================================================
# Kindle Hotkey Navigation (new WinUI Kindle - no menu bar)
# ============================================================
# The new Kindle for PC is a WinUI-3 app; the reader lives inside a
# Microsoft.UI.Content.DesktopChildSiteBridge child window. Both F11 (fullscreen)
# and the page keys (PageDown/PageUp) only work when that reader has keyboard
# focus, and the only reliable way to give it focus is a single left-click in the
# reading area. So the flow is: click ONCE to focus the reader, then F11 to go
# fullscreen. Entering fullscreen resets to a clean page - the click's toolbar
# chrome does NOT carry into fullscreen (only a transient 'Drücke F11' hint shows,
# which fades) - so captures via PrintWindow are clean. After that we navigate
# with keys only and keep the mouse parked in a neutral spot, so no further chrome
# appears. We do NOT use Ctrl+G to jump to a page: its dialog steals the reader's
# focus, which then kills the page keys. All verified live 2026-07-19.

def _monitor_size_of(hwnd):
    """Width/height of the monitor the window is on. Multi-monitor safe: with two
    monitors, pyautogui.size() reports only the PRIMARY monitor - comparing a
    Kindle window on the second monitor against that gives wrong fullscreen
    verdicts (seen in the field: Kindle at (-2560,0) was never 'fullscreen')."""
    mon = win32api.MonitorFromWindow(hwnd, win32con.MONITOR_DEFAULTTONEAREST)
    left, top, right, bottom = win32api.GetMonitorInfo(mon)['Monitor']
    return right - left, bottom - top


def park_mouse_center():
    """Move (NOT click) the mouse to a neutral spot in the middle of the KINDLE
    WINDOW's reading area, away from the side arrows / top toolbar / bottom
    slider, so no hover-chrome appears. Window-relative, NOT primary-screen
    center: with two monitors the primary center can lie on a different monitor
    over some other app entirely. Moving the mouse does not steal keyboard
    focus."""
    kindle = get_kindle_window()
    if not kindle:
        return
    pyautogui.moveTo(kindle.left + kindle.width // 2,
                     kindle.top + kindle.height // 2, duration=0.1)


def _is_fullscreen():
    """True if the Kindle window currently covers (almost) the whole monitor IT
    IS ON (not the primary monitor - see _monitor_size_of)."""
    kindle = get_kindle_window()
    if not kindle:
        return False
    mon_width, mon_height = _monitor_size_of(kindle._hWnd)
    return kindle.width >= mon_width * 0.98 and kindle.height >= mon_height * 0.98


def _click_reader_margin():
    """Left-click the LEFT black margin of the reading pane (~15% of the width,
    mid-height) to give the WinUI reader keyboard focus (required for both F11 and
    the page keys). We deliberately do NOT click the center: with a table of links
    (e.g. the book's resources page) the center can land on a hyperlink, which
    Kindle then opens in the browser. With the single-column layout the content is
    a centred column, so ~15% from the left is always empty margin (no link, past
    the '<' arrow at the very edge). Side effect: this may page one step back
    (left tap zone) - harmless, go_to_book_start() pages back to the cover anyway."""
    kindle = get_kindle_window()
    if not kindle:
        print("[FEHLER] Kindle-Fenster fuer Fokus-Klick nicht gefunden!")
        sys.exit(1)
    pyautogui.click(kindle.left + int(kindle.width * 0.15),
                    kindle.top + kindle.height // 2)
    time.sleep(0.5)


def enter_fullscreen():
    """Enter fullscreen reading mode. F11 only toggles fullscreen when the reader
    has keyboard focus, so we click the reader first, then press F11. If Kindle is
    ALREADY fullscreen we normalize to windowed first (click + F11 to exit), so the
    enter is the clean click->F11 path. Retries because a just-activated window can
    swallow the first F11 press."""
    print("[INFO] Aktiviere Vollbildmodus (F11)...")

    activate_and_get_kindle()
    time.sleep(0.4)

    # If already fullscreen, leave it first so the enter below is the clean,
    # focus-granting click->F11 path (F11-exit also needs the click for focus).
    if _is_fullscreen():
        _click_reader_margin()
        pyautogui.press('f11')
        time.sleep(1.5)

    for attempt in range(1, 4):
        activate_and_get_kindle()
        time.sleep(0.4)
        if _is_fullscreen():
            park_mouse_center()
            print(f"[OK] Vollbildmodus aktiv (Versuch {attempt})")
            return
        _click_reader_margin()   # give the reader focus so F11 is delivered
        pyautogui.press('f11')
        time.sleep(1.6)
        if _is_fullscreen():
            park_mouse_center()
            print(f"[OK] Vollbildmodus aktiviert (Versuch {attempt})")
            return

    print("[FEHLER] Vollbildmodus konnte nicht aktiviert werden (F11)!")
    sys.exit(1)


STILL_SAMPLES = 4      # consecutive pixel-identical grabs that count as "the screen is still"
STILL_TIMEOUT_S = 20   # max wait for the screen to come to rest before failing loud

_still_frames_verified = False


def wait_until_still(label):
    """Wait until STILL_SAMPLES consecutive grabs are PIXEL-IDENTICAL: the screen has
    come to rest (fullscreen transition, 'F11' hint, click chrome all done). No
    threshold - a single differing pixel restarts the count.

    Replaces the former mean-difference check ('stable' = mean |diff| <= 1.0), which
    let a late change through: measured 29.09.2026, right after it reported 'stable',
    4 grabs of the still page differed by [0, 0, 22073] px.

    It is also the precondition of the exact page comparison (images_are_similar):
    if the screen never comes to rest within STILL_TIMEOUT_S, 'the content no longer
    changes' could never be decided -> fail loud, saving the last differing pair as
    _debug_unruhe_1.png / _debug_unruhe_2.png."""
    global _still_frames_verified
    print(f"[INFO] Warte auf Stillstand ({label})...")
    prev = grab_kindle_screenshot()
    if prev is None:
        print("[FEHLER] Konnte Screenshot fuer die Stillstands-Pruefung nicht erstellen!")
        sys.exit(1)
    counts = []
    identical = 1
    last_pair = None
    start = time.time()
    while identical < STILL_SAMPLES:
        check_stop_and_exit()
        if time.time() - start > STILL_TIMEOUT_S:
            print(f"[FEHLER] Bildschirm kommt nicht zur Ruhe ({STILL_TIMEOUT_S}s). "
                  f"Pixel-Unterschiede: {counts}")
            print("[FEHLER] 'Inhalt aendert sich nicht mehr' waere so nie feststellbar.")
            if last_pair is not None:
                last_pair[0].save(Path.cwd() / "_debug_unruhe_1.png")
                last_pair[1].save(Path.cwd() / "_debug_unruhe_2.png")
                print("[FEHLER] Beweisbilder: _debug_unruhe_1.png, _debug_unruhe_2.png")
            sys.exit(1)
        time.sleep(WAIT_AFTER_PAGE)
        cur = grab_kindle_screenshot()
        if cur is None:
            print("[FEHLER] Konnte Screenshot fuer die Stillstands-Pruefung nicht erstellen!")
            sys.exit(1)
        n = changed_pixels(prev, cur) if prev.size == cur.size else f"Groesse {prev.size}->{cur.size}"
        counts.append(n)
        if n == 0:
            identical += 1
        else:
            identical = 1
            last_pair = (prev, cur)
        prev = cur
    print(f"[OK] Stillstand nach {time.time() - start:.1f}s (Pixel-Unterschiede: {counts})")
    _still_frames_verified = True


def go_to_book_start():
    """Navigate to the very beginning of the book (the cover) by pressing PageUp
    until the page stops changing.

    Relies on the reader already having keyboard focus (from the click in
    enter_fullscreen), so PageUp is delivered. We deliberately do NOT use Ctrl+G
    here: its dialog steals that focus. PageUp walks back through any front matter
    to the cover; since the whole book is paged through afterwards anyway, the
    extra presses are cheap."""
    print("[INFO] Navigiere zum Buchanfang (Cover) per PageUp...")
    park_mouse_center()

    # PROOF that page keys reach the reader. Without this, "PageUp changes
    # nothing" is ambiguous: already at the cover - OR the reader has no keyboard
    # focus / PrintWindow returns stale frames. Seen in the field: 3x PageUp with
    # zero effect was reported as '[OK] Cover erreicht' while Kindle sat
    # untouched on page 3. A page turn in SOME direction must provably work
    # before "no change" may be read as "cover reached".
    print("[INFO] Pruefe Seitentasten-Reaktion (Tastaturfokus)...")
    before = grab_kindle_screenshot()
    pyautogui.press('pagedown')
    time.sleep(WAIT_AFTER_PAGE)
    after = grab_kindle_screenshot()
    if before is not None and after is not None and not images_are_similar(after, before):
        print("  [OK] Seitentasten wirken (PageDown hat geblaettert)")
        pyautogui.press('pageup')  # go back to where we started
        time.sleep(WAIT_AFTER_PAGE)
    else:
        # PageDown did nothing - legitimate only at the very last page. PageUp
        # must then work; if NEITHER key changes the page, fail loud.
        pyautogui.press('pageup')
        time.sleep(WAIT_AFTER_PAGE)
        after2 = grab_kindle_screenshot()
        if before is None or after2 is None or images_are_similar(after2, before):
            debug_path = Path.cwd() / "_debug_fokus.png"
            if after2 is not None:
                after2.save(debug_path)
            print("[FEHLER] Reader reagiert weder auf PageDown noch auf PageUp!")
            print("[FEHLER] Ursache: kein Tastaturfokus im Reader ODER PrintWindow liefert")
            print("[FEHLER] eingefrorene Frames. Kindle-Fenster manuell anklicken und neu starten.")
            print(f"[FEHLER] Beweis-Screenshot: {debug_path}")
            sys.exit(1)
        print("  [OK] Seitentasten wirken (PageUp hat geblaettert - Buch stand am Ende)")

    last = grab_kindle_screenshot()
    no_change = 0
    MAX_PAGEUP = 600
    for i in range(MAX_PAGEUP):
        check_stop_and_exit()
        pyautogui.press('pageup')
        time.sleep(WAIT_AFTER_PAGE)
        cur = grab_kindle_screenshot()
        if images_are_similar(cur, last):
            no_change += 1
            if no_change >= 3:
                print(f"  [OK] Cover erreicht (nach {i + 1} PageUp)")
                return
        else:
            no_change = 0
        last = cur

    print("[WARNUNG] Cover nach max. PageUp nicht sicher erreicht - fahre fort")


def press_next_page():
    """Turn to the next page (PageDown). Relies on the reader keeping the keyboard
    focus it got from the click in enter_fullscreen."""
    pyautogui.press('pagedown')


def press_prev_page():
    """Turn to the previous page (PageUp). Relies on the reader keeping the keyboard
    focus it got from the click in enter_fullscreen."""
    pyautogui.press('pageup')


# ============================================================
# Kindle Preparation (Find, Navigate, Fullscreen)
# ============================================================

def find_and_activate_kindle():
    """Find the Kindle window and bring it to the foreground.

    Kindle must ALREADY be running with the book loaded - this tool does NOT
    launch Kindle itself (see the requirements in CLAUDE.md). Returns False (which
    aborts the run) if no Kindle window is found."""
    import ctypes

    print("[INFO] Suche Kindle-Fenster...")
    kindle = get_kindle_window()

    if not kindle:
        print("[FEHLER] Kein Kindle-Fenster gefunden!")
        print("[FEHLER] Kindle muss LAUFEN und das Buch GELADEN sein - das Tool")
        print("[FEHLER] startet Kindle NICHT selbst. Bitte Kindle oeffnen, das Buch")
        print("[FEHLER] laden und erneut ausfuehren.")
        return False

    try:
        hwnd = kindle._hWnd

        if kindle.isMinimized:
            print("[INFO] Kindle ist minimiert - stelle wieder her...")
            ctypes.windll.user32.ShowWindow(hwnd, 9)  # SW_RESTORE
            time.sleep(1.0)

        # Bring to foreground reliably using Win32 API
        # Trick: simulate Alt key press to allow SetForegroundWindow from background process
        ctypes.windll.user32.keybd_event(0x12, 0, 0, 0)  # Alt down
        ctypes.windll.user32.keybd_event(0x12, 0, 2, 0)  # Alt up
        ctypes.windll.user32.SetForegroundWindow(hwnd)
        time.sleep(0.5)

        # Verify and retry if needed
        if ctypes.windll.user32.GetForegroundWindow() != hwnd:
            ctypes.windll.user32.BringWindowToTop(hwnd)
            ctypes.windll.user32.SetForegroundWindow(hwnd)
            time.sleep(0.5)
    except Exception as e:
        print(f"[WARNUNG] Konnte Kindle nicht aktivieren: {e}")
        try:
            kindle.activate()
            time.sleep(0.5)
        except Exception:
            pass

    bounds = activate_and_get_kindle()
    if bounds:
        print(f"[OK] Kindle-Fenster: {bounds[2]}x{bounds[3]} at ({bounds[0]},{bounds[1]})")
    return True


def prepare_kindle_for_capture():
    """Complete preparation sequence (hotkey-based, new WinUI Kindle):
    find window -> fullscreen (F11) -> go to cover -> detect the page format from
    the title page. Returns the crop region (the title-page bounding box) or None
    on failure. Every page is cropped to this region so all pages have the same
    format as the title page (not the mostly-empty full screen)."""
    print()
    print("[SCHRITT 1/4] Kindle-Fenster finden...")
    if not find_and_activate_kindle():
        return None

    time.sleep(0.5)

    print()
    print("[SCHRITT 2/4] Vollbildmodus aktivieren (F11)...")
    enter_fullscreen()  # Bricht bei Fehler mit sys.exit(1) ab
    # Every page-change decision (go_to_book_start, capture) is an exact pixel
    # comparison; this also proves that still frames are pixel-identical.
    wait_until_still("nach F11")

    print()
    print("[SCHRITT 3/4] Zum Buchanfang (Cover) navigieren...")
    go_to_book_start()
    wait_until_still("Titelseite")

    # We are on the title page now -> derive the page crop region from it.
    print()
    print("[SCHRITT 4/4] Seitenformat von der Titelseite bestimmen...")
    screenshot = grab_kindle_screenshot()
    if screenshot is None:
        print("[FEHLER] Konnte Screenshot nicht erstellen!")
        sys.exit(1)

    book_region = detect_page_region_from_cover(screenshot)
    if book_region is None:
        # Fail loud WITH evidence: save the grab the detection ran on, so a
        # failing setup (other machine, other monitor, stale rendering) can be
        # diagnosed from the file instead of guessed at.
        debug_path = Path.cwd() / "_debug_titelseite.png"
        screenshot.save(debug_path)
        print(f"[FEHLER] Beweis-Screenshot der fehlgeschlagenen Erkennung: {debug_path}")
        sys.exit(1)
    left, top, right, bottom = book_region
    print(f"[OK] Seitenformat (von Titelseite): {right - left} x {bottom - top} Pixel")
    # Always keep the UNCROPPED title-page grab as evidence: the crop region is
    # derived from it, so a wrong crop can be re-checked offline. Measured 29.09.2026
    # (Cybernetic Analysis ...): region (554,157)-(1364,1080), yet text pages had ink
    # in the crop's first pixel row and page 7 (text at the top only) came out pure
    # white - without this file the fullscreen geometry could not be inspected.
    evidence = Path.cwd() / "_debug_titelseite_voll.png"
    screenshot.save(evidence)
    print(f"[INFO] Titelseite ungeschnitten gespeichert: {evidence}")

    print()
    print("[OK] Kindle bereit fuer Erfassung!")
    return book_region

def _grab_window_printwindow(hwnd):
    """Capture a window's pixels via PrintWindow(PW_RENDERFULLCONTENT=2). This
    reads the window's own rendering (WinUI + WebView2 content), so it works even
    when the Kindle fullscreen is in an exclusive/protected mode that makes GDI
    screen grab fail or return black. Returns a PIL Image, or None on failure."""
    import ctypes

    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width, height = right - left, bottom - top
    if width <= 0 or height <= 0:
        return None

    hwnd_dc = win32gui.GetWindowDC(hwnd)
    mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
    save_dc = mfc_dc.CreateCompatibleDC()
    bitmap = win32ui.CreateBitmap()
    bitmap.CreateCompatibleBitmap(mfc_dc, width, height)
    save_dc.SelectObject(bitmap)
    try:
        # PW_RENDERFULLCONTENT = 2 -> include DirectComposition / WebView2 content
        result = ctypes.windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), 2)
        info = bitmap.GetInfo()
        bits = bitmap.GetBitmapBits(True)
        img = Image.frombuffer('RGB', (info['bmWidth'], info['bmHeight']),
                               bits, 'raw', 'BGRX', 0, 1)
    finally:
        win32gui.DeleteObject(bitmap.GetHandle())
        save_dc.DeleteDC()
        mfc_dc.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwnd_dc)

    return img if result == 1 else None


def grab_kindle_screenshot(retries=4, delay=0.2):
    """Capture the current Kindle window (the reading page) via PrintWindow.
    Returns a PIL Image, or None if the window is gone / capture keeps failing."""
    kindle = get_kindle_window()
    if not kindle:
        return None
    hwnd = kindle._hWnd
    for _ in range(retries):
        img = _grab_window_printwindow(hwnd)
        if img is not None:
            return img
        time.sleep(delay)
    return None

# ============================================================
# Image Analysis
# ============================================================

# Cover detection tuning
BG_TOL = 10           # Per-channel difference from the letterbox colour that counts as page content
MIN_COVERAGE = 0.5    # Content fraction: above it an edge-strip row is window chrome;
                      # below it the page area is no letterboxed title page
LETTERBOX_MAX = 0.02  # Content fraction up to which a column still counts as pure letterbox
EDGE_STRIP = 10       # Width (px) of the far-left/far-right strips that define the letterbox colour


def _page_span(is_letterbox):
    """(start, end) of the page between the two letterbox bars along one axis.

    Walks inward from each edge: past a stray non-letterbox strip AT the edge (the
    1px window border), across the letterbox bar, up to the first index that is not
    letterbox. Everything between those two stops is page - INCLUDING interior
    stretches that look like letterbox (e.g. a black bar printed on the cover).
    A side without a bar (no letterbox index in that half) -> the page reaches that
    edge. Returns None if a bar runs to the middle (no page found)."""
    n = len(is_letterbox)
    half = n // 2
    lb = np.flatnonzero(is_letterbox)

    left_bar = lb[lb < half]
    if left_bar.size == 0:
        start = 0
    else:
        after = np.flatnonzero(~is_letterbox[left_bar[0]:half])
        if after.size == 0:
            return None
        start = left_bar[0] + after[0]

    right_bar = lb[lb >= half]
    if right_bar.size == 0:
        end = n
    else:
        before = np.flatnonzero(~is_letterbox[half:right_bar[-1] + 1])
        if before.size == 0:
            return None
        end = half + before[-1] + 1

    return start, end


def detect_page_region_from_cover(cover_img):
    """Detect the book PAGE region (bounding box) from the cover / title page and
    return (left, top, right, bottom). ALL pages are cropped to this region, so
    every captured page has the same format as the title page - not the full
    (mostly empty) screen.

    Definition: the page is what lies BETWEEN the uniform letterbox bars on the
    left and right (black in dark mode, white in light mode). Only the LEFT and
    RIGHT edge are searched; vertically the page is the full window height - in
    Kindle's fullscreen the text runs from the very top to the very bottom of the
    screen (user decision 29.09.2026). The page may itself contain the letterbox
    colour - e.g. a black bar printed on a white title page - so nothing here
    requires the page to be 'mostly different' from the letterbox.

    Measured 29.09.2026 (Cybernetic Analysis ..., fullscreen): the former
    longest-run-of-coverage>50% approach returned top=157 instead of 0 - rows
    141-156 (the black bar on the cover) had only 44% non-black pixels, split the
    page in two runs and the longer one (157-1078) won; every text page then lost
    its top lines, and page 7 (text only at the top) came out pure white.

    Steps:
      reference rows - rows whose far-left/far-right edge strips are letterbox.
                A title bar spans the full width, so its rows drop out here; that
                keeps chrome from making letterbox columns look like content.
      columns - over the reference rows a letterbox column is (almost) pure
                letterbox colour; _page_span walks from both edges to the page
                (past a stray 1px edge column - measured: x=1919 white on the
                1920x1080 fullscreen title page).
    We do NOT use pixel variance: window chrome has variance too, so a variance scan
    returned the whole window (measured: 1443x834 landscape instead of the 516x804
    cover)."""
    a = np.asarray(cover_img.convert('RGB')).astype(np.int16)
    height, width, _ = a.shape

    # Letterbox colour: the far-left/far-right edge strips over the vertical middle
    # are always uniform margin, because the page is centred horizontally.
    band = a[int(height * 0.25):int(height * 0.75)]
    bg = np.median(np.concatenate([band[:, :EDGE_STRIP], band[:, -EDGE_STRIP:]], axis=1).reshape(-1, 3), axis=0)
    mask = np.abs(a - bg).max(axis=2) > BG_TOL  # True = differs from letterbox colour

    edges = np.concatenate([mask[:, :EDGE_STRIP], mask[:, -EDGE_STRIP:]], axis=1)
    ref_rows = edges.mean(axis=1) < MIN_COVERAGE
    if not ref_rows.any():
        print("[FEHLER] Kein Seitenrand gefunden - die Fensterraender sind in keiner Zeile einfarbig!")
        return None

    span = _page_span(mask[ref_rows].mean(axis=0) <= LETTERBOX_MAX)
    if span is None:
        print("[FEHLER] Keine Seite zwischen den Randbalken gefunden!")
        return None
    left, right = span
    top, bottom = 0, height  # only left/right are searched (see docstring)

    pw, ph = right - left, bottom - top
    print(f"  Titelseiten-Format: ({left},{top})-({right},{bottom}) = {pw} x {ph} Pixel "
          f"(Seitenverhaeltnis {pw / ph:.3f})")

    # No letterbox found means we are not looking at a letterboxed title page
    # (wrong page, or chrome swallowed everything) - cropping would be wrong.
    # Return None; the caller fails loud WITH the offending screenshot as
    # evidence (we cannot save it here - only the caller knows the book folder).
    if pw >= width * 0.98:
        print("[FEHLER] Kein Seitenrand gefunden - das ist keine letterboxte Titelseite!")
        print("[FEHLER] Steht Kindle auf der Titelseite und ist Layout 'Einzelne Spalte' gesetzt?")
        return None
    if pw < width * 0.15:
        print(f"[FEHLER] Titelseiten-Format nicht erkennbar ({pw}x{ph} zu schmal)!")
        return None
    # Plausibility check only (it never moves the edges): a letterboxed title page
    # differs from the letterbox colour over most of its area, a plain text page on a
    # letterbox-coloured background does not - then left/right would just be the
    # text block. Measured 29.09.2026: title page 97.5% of its area, text pages
    # 7 and 8 (white, a few lines of text) 1.0% within their ink columns.
    area = float(mask[:, left:right].mean())
    if area < MIN_COVERAGE:
        print(f"[FEHLER] Keine Titelseite: nur {area:.1%} der Flaeche zwischen den Raendern "
              f"weichen von der Randfarbe ab (Titelseite: >{MIN_COVERAGE:.0%})!")
        return None

    return (left, top, right, bottom)


# Page-change detection - NO threshold (user order 29.09.2026): the content "did not
# change" only if the two grabs are PIXEL-IDENTICAL; any difference is a change.
# End of book = the content no longer changes when paging forward; nothing else.
# The former 'changed < 0.6% = no turn' missed a real turn - measured 29.09.2026,
# pages 7->8 (two short front-matter pages) changed 3095 px = 0.354% of the 810x1080
# crop (0.15% of the full frame), were taken for 'no turn', and the capture ended
# after 7 of 326 pages.


def changed_pixels(img1, img2):
    """Number of pixels that differ at all between the two images."""
    a = np.asarray(img1)
    b = np.asarray(img2)
    diff = (a != b).any(axis=2) if a.ndim == 3 else (a != b)
    return int(np.count_nonzero(diff))


def images_are_similar(img1, img2):
    """True if the content did NOT change: the two images are pixel-identical.
    Any difference - however small - counts as a change (no threshold). Requires a
    prior wait_until_still(), which proves that still frames are identical."""
    if not _still_frames_verified:
        print("[FEHLER] Interner Fehler: Stillstands-Pruefung wurde nicht ausgefuehrt!")
        sys.exit(1)
    if img1 is None or img2 is None:
        return False
    if img1.size != img2.size:
        return False
    return changed_pixels(img1, img2) == 0

# ============================================================
# Main Functions
# ============================================================

def clear_output_folder(folder):
    """Delete all existing PNG files in output folder."""
    png_files = list(folder.glob("page_*.png"))
    if png_files:
        print(f"[INFO] Loesche {len(png_files)} existierende Seitendateien...")
        for f in png_files:
            f.unlink()
        print(f"[OK] Ausgabeordner geleert")

def _save_page(output_folder, page_num, image):
    """Save one captured page image and wait until it is on disk."""
    filename = f"page_{page_num:04d}.png"
    filepath = output_folder / filename
    image.save(filepath, "PNG")
    while not filepath.exists():
        check_stop_and_exit()
        time.sleep(0.05)
    print(f"[OK] Gespeichert: {filename}")


def capture_pages(output_folder, book_region):
    """Capture all pages: save the current page, PageDown, repeat until the book
    no longer advances.

    A page turn is verified by watching for the page image to change. Crucially,
    on a 'no change' we do NOT turn again while waiting - a page that simply
    renders slowly would otherwise be skipped. Only after the page fails to change
    across RENDER_POLL_TRIES polls (and one re-focus click + retry) do we conclude
    the end of the book has been reached."""
    global STOP_FLAG

    # Window size the crop region was measured on. book_region is fixed for the
    # whole run, so if the window is resized (or drops out of fullscreen) midway,
    # every following crop would silently cut the wrong part of the page.
    expected_size = None
    # Uncropped grab behind the most recent grab_page() result - kept so that the
    # end-of-book decision can be saved as uncropped evidence (see below).
    last_shot = None

    def grab_page():
        nonlocal expected_size, last_shot
        shot = grab_kindle_screenshot()
        if shot is None:
            return None
        last_shot = shot
        if expected_size is None:
            expected_size = shot.size
            if book_region[2] > shot.size[0] or book_region[3] > shot.size[1]:
                print(f"[FEHLER] Titelseiten-Format {book_region} passt nicht ins "
                      f"Fenster {shot.size}!")
                sys.exit(1)
        elif shot.size != expected_size:
            print(f"[FEHLER] Fenstergroesse hat sich waehrend der Erfassung geaendert: "
                  f"{expected_size} -> {shot.size}")
            print("[FEHLER] Der Zuschnitt auf das Titelseiten-Format waere ab hier falsch.")
            print("[FEHLER] Kindle-Fenster waehrend des Laufs NICHT veraendern!")
            sys.exit(1)
        return shot.crop(book_region)

    def wait_for_new_page(reference):
        """Poll (without turning the page) until the page differs from `reference`,
        giving a slow render time to appear. Returns the new page image, or None if
        it never changes (book did not advance)."""
        for _ in range(RENDER_POLL_TRIES):
            check_stop_and_exit()
            time.sleep(WAIT_AFTER_PAGE)
            cur = grab_page()
            if cur is not None and not images_are_similar(cur, reference):
                return cur
        return None

    page_num = 1
    try:
        # Save the first (current) page = cover
        current = grab_page()
        if current is None:
            print("[FEHLER] Kindle-Fenster verloren!")
            return 0
        _save_page(output_folder, page_num, current)
        page_num += 1
        last_saved = current
        last_saved_full = last_shot

        while True:
            check_stop_and_exit()
            press_next_page()
            new_page = wait_for_new_page(last_saved)

            if new_page is None:
                # No advance: might be end of book, or the reader lost keyboard
                # focus (e.g. foreground was stolen). Re-activate and re-focus the
                # reader with a click, wait until the screen is still again (the
                # click's chrome done), then give it one more chance before
                # concluding "end".
                print("[INFO] Keine Aenderung - pruefe Buchende / Fokus...")
                find_and_activate_kindle()
                _click_reader_margin()
                park_mouse_center()
                wait_until_still("nach Fokus-Klick")
                press_next_page()
                new_page = wait_for_new_page(last_saved)
                if new_page is None:
                    # Evidence for the end-of-book decision, both UNCROPPED: the
                    # last saved page and what the window shows now. The decision
                    # compares CROPPED images only, so content outside the crop is
                    # invisible to it - these files show whether that happened.
                    last_saved_full.save(Path.cwd() / "_debug_letzte_seite_voll.png")
                    final = grab_kindle_screenshot()
                    if final is not None:
                        final.save(Path.cwd() / "_debug_buchende_voll.png")
                    print("[INFO] Beweisbilder gespeichert: _debug_letzte_seite_voll.png, "
                          "_debug_buchende_voll.png")
                    print("[OK] Buchende erreicht.")
                    break

            _save_page(output_folder, page_num, new_page)
            page_num += 1
            last_saved = new_page
            last_saved_full = last_shot

    except KeyboardInterrupt:
        print("\n[INFO] Erfassung vom Benutzer gestoppt.")
    except SystemExit:
        raise

    return page_num - 1

def keep_session_awake(enable=True):
    """Prevent the display from sleeping / the screensaver-lock from kicking in
    during a long capture run. Synthetic pyautogui input does NOT reset Windows'
    idle timer, so without this a multi-minute capture can end up on the lock
    screen - where keyboard input no longer reaches Kindle and paging dies."""
    import ctypes
    ES_CONTINUOUS = 0x80000000
    ES_SYSTEM_REQUIRED = 0x00000001
    ES_DISPLAY_REQUIRED = 0x00000002
    if enable:
        ctypes.windll.kernel32.SetThreadExecutionState(
            ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED)
    else:
        ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)


def main():
    """Main function - capture book pages."""
    global STOP_FLAG
    STOP_FLAG = False

    print("=" * 60)
    print("  KINDLE BUCH ERFASSUNG")
    print("=" * 60)
    print()
    print("Stelle sicher:")
    print("  - Kindle-App ist geoeffnet mit dem gewuenschten Buch")
    print()
    print(">>> Druecke EINE BELIEBIGE TASTE zum Stoppen <<<")
    print()

    output_folder = Path.cwd() / "pages"
    output_folder.mkdir(exist_ok=True)
    print(f"[INFO] Ausgabeordner: {output_folder}")
    print()

    # Keep the session awake for the whole (multi-minute) run so a screensaver /
    # display timeout can't lock the desktop mid-capture (which would kill paging).
    keep_session_awake(True)

    # Prepare Kindle: find window, click-focus + F11 fullscreen, go to cover
    # NOTE: keyboard listener starts AFTER preparation to avoid accidental stops
    book_region = prepare_kindle_for_capture()
    if book_region is None:
        print("[FEHLER] Kindle-Vorbereitung fehlgeschlagen!")
        sys.exit(1)

    # Now start keyboard listener for stop during capture
    start_keyboard_listener()

    clear_output_folder(output_folder)

    # Capture
    print()
    print(f"[INFO] Buchbereich: {book_region}")
    print("[INFO] Starte Erfassung...")
    print()
    time.sleep(1)

    captured_pages = 0
    try:
        captured_pages = capture_pages(output_folder, book_region)
    except SystemExit:
        pass
    finally:
        stop_keyboard_listener()
        keep_session_awake(False)

    # Exit fullscreen and minimize Kindle
    exit_fullscreen_and_minimize()

    print()
    print("=" * 60)
    if STOP_FLAG:
        print("  ERFASSUNG ABGEBROCHEN")
        print("=" * 60)
        print(f"  Erfasste Seiten: {captured_pages}")
        sys.exit(1)
    else:
        print("  ERFASSUNG ABGESCHLOSSEN")
        print("=" * 60)
        print(f"  Erfasste Seiten: {captured_pages}")
        sys.exit(0)

if __name__ == "__main__":
    main()
