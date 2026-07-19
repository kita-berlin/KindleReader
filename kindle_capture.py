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
   - Click once to focus the reader, then F11 -> clean fullscreen
   - Navigate to the very beginning / cover (PageUp until the page stops changing)
   - Capture every fullscreen page (PrintWindow), paging forward with PageDown

Author: Claude
"""

import pyautogui
pyautogui.FAILSAFE = False  # Disable fail-safe (mouse in corner)
import pygetwindow as gw
import subprocess
import os
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

def park_mouse_center():
    """Move (NOT click) the mouse to a neutral spot in the middle of the reading
    area, away from the side arrows / top toolbar / bottom slider, so no
    hover-chrome appears. Moving the mouse does not steal keyboard focus."""
    screen_width, screen_height = pyautogui.size()
    pyautogui.moveTo(screen_width // 2, int(screen_height * 0.5), duration=0.1)


def _is_fullscreen():
    """True if the Kindle window currently covers (almost) the whole screen."""
    kindle = get_kindle_window()
    if not kindle:
        return False
    screen_width, screen_height = pyautogui.size()
    return kindle.width >= screen_width * 0.98 and kindle.height >= screen_height * 0.98


def _click_reader_center():
    """Left-click the center of the current Kindle window to give the WinUI reader
    keyboard focus (required for both F11 and the page keys). The center is
    neutral - it does not turn a page. In windowed mode this briefly shows the
    toolbar chrome, but entering fullscreen (F11) resets to a clean page."""
    kindle = get_kindle_window()
    if not kindle:
        print("[FEHLER] Kindle-Fenster fuer Fokus-Klick nicht gefunden!")
        sys.exit(1)
    pyautogui.click(kindle.left + kindle.width // 2, kindle.top + kindle.height // 2)
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
        _click_reader_center()
        pyautogui.press('f11')
        time.sleep(1.5)

    for attempt in range(1, 4):
        activate_and_get_kindle()
        time.sleep(0.4)
        if _is_fullscreen():
            park_mouse_center()
            print(f"[OK] Vollbildmodus aktiv (Versuch {attempt})")
            return
        _click_reader_center()   # give the reader focus so F11 is delivered
        pyautogui.press('f11')
        time.sleep(1.6)
        if _is_fullscreen():
            park_mouse_center()
            print(f"[OK] Vollbildmodus aktiviert (Versuch {attempt})")
            return

    print("[FEHLER] Vollbildmodus konnte nicht aktiviert werden (F11)!")
    sys.exit(1)


def wait_until_screen_stable(max_wait=8, interval=0.5, stable_needed=2, threshold=1.0):
    """Wait until the screen stops changing (fullscreen transition + the transient
    'Drücke F11 zum Beenden' hint fading). Replaces the old OCR-based hint wait
    with a UI-independent screenshot-difference check."""
    print("[INFO] Warte bis Bildschirm stabil...")
    prev = grab_kindle_screenshot()
    stable = 0
    waited = 0.0
    while waited < max_wait:
        check_stop_and_exit()
        time.sleep(interval)
        waited += interval
        cur = grab_kindle_screenshot()
        if prev is not None and cur is not None and prev.size == cur.size:
            a = np.asarray(prev, dtype=np.float32)
            b = np.asarray(cur, dtype=np.float32)
            d = float(np.mean(np.abs(a - b)))
        else:
            d = 999.0
        prev = cur
        if d <= threshold:
            stable += 1
            if stable >= stable_needed:
                print(f"[OK] Bildschirm stabil nach {waited:.1f}s")
                return
        else:
            stable = 0
    print("[WARNUNG] Timeout beim Warten auf stabilen Bildschirm - fahre fort")


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

def start_kindle_app():
    """Start the Kindle application."""
    kindle_paths = [
        Path(os.environ.get("LOCALAPPDATA", "")) / "Amazon" / "Kindle" / "application" / "Kindle.exe",
        Path("C:/Program Files/Amazon/Kindle/Kindle.exe"),
        Path("C:/Program Files (x86)/Amazon/Kindle/Kindle.exe"),
    ]

    for kindle_path in kindle_paths:
        if kindle_path.exists():
            print(f"[INFO] Starte Kindle: {kindle_path}")
            subprocess.Popen([str(kindle_path)], shell=False)
            return True

    print("[FEHLER] Kindle.exe nicht gefunden!")
    return False


def find_and_activate_kindle():
    """Find Kindle window, start app if needed, and bring to foreground."""
    import ctypes

    print("[INFO] Suche Kindle-Fenster...")
    kindle = get_kindle_window()

    if not kindle:
        print("[INFO] Kindle nicht offen - starte Kindle-App...")
        if not start_kindle_app():
            return False

        # Wait for Kindle to start
        for i in range(30):  # Max 30 seconds
            time.sleep(1)
            kindle = get_kindle_window()
            if kindle:
                print(f"[OK] Kindle gestartet nach {i+1}s")
                break
            if i % 5 == 4:
                print(f"  Warte auf Kindle... ({i+1}s)")

        if not kindle:
            print("[FEHLER] Kindle konnte nicht gestartet werden!")
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
    find window -> fullscreen (F11) -> go to cover. Returns the capture region
    (the full fullscreen page) or None on failure.

    We capture the whole fullscreen page rather than trying to crop the content:
    in clean fullscreen the page has only a small, uniform margin, the cover is
    letterboxed, and per-page margin detection proved unreliable. The downstream
    OCR (create_pdf / create_markdown) ignores the margins anyway."""
    print()
    print("[SCHRITT 1/3] Kindle-Fenster finden...")
    if not find_and_activate_kindle():
        return None

    time.sleep(0.5)

    print()
    print("[SCHRITT 2/3] Vollbildmodus aktivieren (F11)...")
    enter_fullscreen()  # Bricht bei Fehler mit sys.exit(1) ab
    wait_until_screen_stable()

    print()
    print("[SCHRITT 3/3] Zum Buchanfang (Cover) navigieren...")
    go_to_book_start()
    wait_until_screen_stable(max_wait=4)

    screenshot = grab_kindle_screenshot()
    if screenshot is None:
        print("[FEHLER] Konnte Screenshot nicht erstellen!")
        sys.exit(1)

    screen_width, screen_height = screenshot.size
    book_region = (0, 0, screen_width, screen_height)
    print(f"[OK] Erfassungsbereich (Vollbild): {screen_width} x {screen_height} Pixel")

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

def images_are_similar(img1, img2, change_threshold=0.006, pixel_diff=24):
    """True if the two page images are essentially identical (i.e. the page did NOT
    turn). Compares the FRACTION of pixels that changed noticeably - NOT the mean
    difference. With a single-column page the content is a small region on a large,
    identical (black) background, so a mean-difference metric is dominated by that
    background and wrongly reports 'no change' for two clearly different sparse pages.
    Counting changed pixels is robust: measured on real pages a page turn changes
    ~5-37% of pixels, while an unchanged page changes 0%. So anything below ~0.6% is
    treated as 'no turn'."""
    if img1 is None or img2 is None:
        return False

    if img1.size != img2.size:
        return False

    a = np.asarray(img1).astype(np.int16)
    b = np.asarray(img2).astype(np.int16)
    if a.ndim == 3:
        diff = np.abs(a - b).max(axis=2)
    else:
        diff = np.abs(a - b)

    changed_fraction = float(np.mean(diff > pixel_diff))
    return changed_fraction < change_threshold

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

    def grab_page():
        shot = grab_kindle_screenshot()
        return shot.crop(book_region) if shot is not None else None

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

        while True:
            check_stop_and_exit()
            press_next_page()
            new_page = wait_for_new_page(last_saved)

            if new_page is None:
                # No advance: might be end of book, or the reader lost keyboard
                # focus (e.g. foreground was stolen). Re-activate and re-focus the
                # reader with a click, wait for the click's chrome to fade again,
                # then give it one more chance before concluding "end".
                print("[INFO] Keine Aenderung - pruefe Buchende / Fokus...")
                find_and_activate_kindle()
                _click_reader_center()
                park_mouse_center()
                wait_until_screen_stable(max_wait=6)
                press_next_page()
                new_page = wait_for_new_page(last_saved)
                if new_page is None:
                    print("[OK] Buchende erreicht.")
                    break

            _save_page(output_folder, page_num, new_page)
            page_num += 1
            last_saved = new_page

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
