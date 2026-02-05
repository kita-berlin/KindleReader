"""
Kindle Book Capture Tool
========================
Captures all pages from a Kindle book as PNG images.

Usage:
1. Open Kindle app with the book you want to capture
2. Change to the book folder (used as output folder)
3. Run: kindle_capture.exe
4. The script will automatically:
   - Find the Kindle window
   - Navigate to the title page
   - Enter fullscreen mode
   - Capture all pages as PNG files

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
import asyncio
from PIL import Image, ImageGrab
import numpy as np
from pathlib import Path
from pynput import keyboard

# pywinauto für zuverlässige Fenster-Erkennung
try:
    from pywinauto import Application, Desktop
    from pywinauto.timings import Timings
    Timings.after_click_wait = 0.1
    Timings.after_setcursorpos_wait = 0.01
    PYWINAUTO_AVAILABLE = True
except ImportError:
    PYWINAUTO_AVAILABLE = False

# Windows OCR für Menü-Erkennung
try:
    from winsdk.windows.media.ocr import OcrEngine
    from winsdk.windows.globalization import Language
    from winsdk.windows.graphics.imaging import BitmapDecoder
    from winsdk.windows.storage import StorageFile
    WINDOWS_OCR_AVAILABLE = True
except ImportError:
    WINDOWS_OCR_AVAILABLE = False

# Optional: Tkinter for click indicator
try:
    import tkinter as tk
    TKINTER_AVAILABLE = True
except ImportError:
    TKINTER_AVAILABLE = False

# Click indicator overlay
current_overlay = None

# ============================================================
# Configuration
# ============================================================
WAIT_AFTER_CLICK = 0.5  # Seconds to wait after click
MAX_NO_CHANGE_COUNT = 3  # Stop after this many pages without change
FULLSCREEN_MSG_TIMEOUT = 10  # Max seconds to wait for fullscreen message to disappear

# Global flag for immediate stop
STOP_FLAG = False
keyboard_listener = None

# ============================================================
# Keyboard and Signal Handling
# ============================================================

def on_key_press(key):
    """Global keyboard hook - any key stops the script."""
    global STOP_FLAG
    STOP_FLAG = True
    print("\n[!] Taste gedrueckt - stoppe...")
    return False  # Stop listener

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
        hide_click_indicator()
        print("\n[GESTOPPT] Script vom Benutzer gestoppt.")
        sys.exit(1)

# ============================================================
# Click Indicator (Red Dot)
# ============================================================

def show_click_indicator(x, y):
    """Show visual indicator at click position."""
    global current_overlay

    if not TKINTER_AVAILABLE:
        return

    hide_click_indicator()

    try:
        import win32gui
        import win32con

        overlay = tk.Tk()
        overlay.overrideredirect(True)
        overlay.attributes('-topmost', True)
        overlay.attributes('-alpha', 0.9)

        size = 30
        target_x = x - size // 2
        target_y = y - size // 2

        overlay.geometry(f"{size}x{size}+{target_x}+{target_y}")

        canvas = tk.Canvas(overlay, width=size, height=size, bg='black', highlightthickness=0)
        canvas.pack()

        center = size // 2
        radius = 12
        canvas.create_oval(
            center - radius, center - radius,
            center + radius, center + radius,
            fill='red', outline='red', width=2
        )

        try:
            overlay.update()
            overlay.update_idletasks()
        except Exception:
            pass

        try:
            hwnd = int(overlay.winfo_id())
            win32gui.SetWindowPos(
                hwnd, win32con.HWND_TOPMOST,
                target_x, target_y, size, size,
                win32con.SWP_SHOWWINDOW | win32con.SWP_NOACTIVATE
            )
        except Exception:
            pass

        current_overlay = overlay

    except Exception as e:
        pass  # Silently fail if overlay can't be shown


def hide_click_indicator():
    """Remove visual overlay."""
    global current_overlay
    if current_overlay:
        try:
            current_overlay.quit()
            current_overlay.destroy()
        except:
            pass
        current_overlay = None

# ============================================================
# Windows OCR Functions
# ============================================================

# Global OCR engine
_ocr_engine = None

def get_ocr_engine():
    """Initialize and return Windows OCR engine."""
    global _ocr_engine
    if _ocr_engine is not None:
        return _ocr_engine

    if not WINDOWS_OCR_AVAILABLE:
        return None

    for lang_tag in ['de-DE', 'de', 'en-US', 'en']:
        try:
            lang = Language(lang_tag)
            if OcrEngine.is_language_supported(lang):
                engine = OcrEngine.try_create_from_language(lang)
                if engine:
                    _ocr_engine = engine
                    return engine
        except:
            continue
    return None


async def ocr_image_async(engine, img_path):
    """Run OCR on image and return word list with positions."""
    try:
        abs_path = str(Path(img_path).resolve())
        storage_file = await StorageFile.get_file_from_path_async(abs_path)
        stream = await storage_file.open_read_async()
        decoder = await BitmapDecoder.create_async(stream)
        bitmap = await decoder.get_software_bitmap_async()
        result = await engine.recognize_async(bitmap)

        words = []
        if result and result.lines:
            for line in result.lines:
                for word in line.words:
                    text = word.text.strip()
                    if text:
                        rect = word.bounding_rect
                        words.append({
                            'text': text,
                            'x': int(rect.x),
                            'y': int(rect.y),
                            'width': int(rect.width),
                            'height': int(rect.height),
                            'center_x': int(rect.x + rect.width / 2),
                            'center_y': int(rect.y + rect.height / 2),
                        })
        return words
    except Exception as e:
        print(f"  [DEBUG] OCR-Fehler: {e}")
        return []


def ocr_image(engine, img_path):
    """Synchronous wrapper for OCR."""
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    try:
        return loop.run_until_complete(ocr_image_async(engine, img_path))
    finally:
        loop.close()


def find_text_in_region(engine, region_bbox, search_text, take_topmost=True, scale_factor=2):
    """
    Search for text in a region using OCR.
    Returns (abs_x, abs_y, word_info) or None.
    """
    left, top, right, bottom = region_bbox

    # Screenshot the region
    img = ImageGrab.grab(bbox=(left, top, right, bottom))

    # Scale up for better OCR recognition
    if scale_factor > 1:
        new_width = img.width * scale_factor
        new_height = img.height * scale_factor
        img = img.resize((new_width, new_height), resample=Image.LANCZOS)

    temp_path = Path.cwd() / "_temp_ocr.png"
    img.save(temp_path)

    # Run OCR
    words = ocr_image(engine, temp_path)

    # Search for text
    search_lower = search_text.lower()
    matches = [w for w in words if search_lower in w['text'].lower()]

    # Cleanup
    try:
        temp_path.unlink()
    except:
        pass

    if not matches:
        return None

    # Sort by Y position (topmost first)
    matches_sorted = sorted(matches, key=lambda w: w['y'])

    if take_topmost:
        best_match = matches_sorted[0]
    else:
        best_match = matches_sorted[-1]

    # Calculate absolute coordinates (scale back)
    abs_x = left + int(best_match['center_x'] / scale_factor)
    abs_y = top + int(best_match['center_y'] / scale_factor)

    return (abs_x, abs_y, best_match)


# ============================================================
# Kindle Window Control
# ============================================================

# Global pywinauto window reference
_kindle_pywinauto = None

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


def get_kindle_window_pywinauto():
    """Find main Kindle window with pywinauto for reliable clicking.
    Skips WebView2 child windows by picking the largest visible window."""
    global _kindle_pywinauto

    if not PYWINAUTO_AVAILABLE:
        return None, None

    try:
        import ctypes

        # Find the Kindle window hwnd via get_kindle_window()
        kindle_gw = get_kindle_window()
        if not kindle_gw:
            return None, None

        kindle_hwnd = kindle_gw._hWnd

        # Get the process ID from the hwnd
        pid = ctypes.c_ulong()
        ctypes.windll.user32.GetWindowThreadProcessId(kindle_hwnd, ctypes.byref(pid))
        kindle_pid = pid.value

        # Connect pywinauto via process ID (avoids ambiguity error)
        app = Application(backend='uia').connect(process=kindle_pid, timeout=5)

        # Find the main window matching our hwnd
        all_windows = app.windows()
        kindle_window = None
        for w in all_windows:
            try:
                if w.handle == kindle_hwnd:
                    kindle_window = w
                    break
            except Exception:
                continue

        if not kindle_window:
            # Fallback: pick the largest
            best = None
            best_area = 0
            for w in all_windows:
                try:
                    rect = w.rectangle()
                    area = rect.width() * rect.height()
                    if area > best_area:
                        best = w
                        best_area = area
                except Exception:
                    continue
            kindle_window = best

        if not kindle_window:
            return None, None

        kindle_window.set_focus()
        time.sleep(0.3)

        rect = kindle_window.rectangle()
        bounds = (rect.left, rect.top, rect.width(), rect.height())

        _kindle_pywinauto = kindle_window
        return kindle_window, bounds
    except Exception as e:
        print(f"  [WARNUNG] pywinauto: {e}")
        return None, None


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
    """Focus Kindle, send ESC to exit fullscreen, then minimize."""
    kindle = get_kindle_window()
    if kindle:
        try:
            kindle.activate()
            time.sleep(0.3)
            pyautogui.press('escape')
            time.sleep(0.5)
            kindle.minimize()
            print("[INFO] Kindle minimiert")
        except Exception as e:
            print(f"[WARNUNG] Konnte Kindle nicht minimieren: {e}")

def click_at(x, y, show_indicator=True):
    """Click at absolute position with optional visual indicator.
    Uses moveTo + click pattern for reliable menu clicks."""
    if show_indicator:
        show_click_indicator(x, y)
        time.sleep(0.3)

    # Move mouse first, then click (more reliable for menus)
    pyautogui.moveTo(x, y, duration=0.01)
    time.sleep(0.1)

    if show_indicator:
        hide_click_indicator()
        time.sleep(0.1)

    pyautogui.click(x, y)
    time.sleep(0.3)


def click_kindle(x_percent, y_percent, show_indicator=True):
    """Click at position relative to Kindle window (0-1 for x and y)."""
    bounds = activate_and_get_kindle()
    if bounds:
        left, top, width, height = bounds
        abs_x = left + int(width * x_percent)
        abs_y = top + int(height * y_percent)
        click_at(abs_x, abs_y, show_indicator)
        return True
    return False


# Global arrow positions (detected once from title page)
ARROW_RIGHT_POS = None
ARROW_LEFT_POS = None


def detect_arrow_positions(book_region=None):
    """Detect arrow positions by hovering over margins to make arrows visible."""
    global ARROW_RIGHT_POS, ARROW_LEFT_POS

    print("[INFO] Suche Navigationspfeile...")

    screen_width, screen_height = pyautogui.size()

    # Calculate positions to hover based on book region or screen
    if book_region:
        book_left, book_top, book_right, book_bottom = book_region
        # Right margin: between book right edge and screen edge
        right_hover_x = book_right + (screen_width - book_right) // 2
        right_hover_y = (book_top + book_bottom) // 2
        # Left margin: between screen left and book left edge
        left_hover_x = book_left // 2
        left_hover_y = (book_top + book_bottom) // 2
    else:
        # Fallback: use 95% / 5% of screen width
        right_hover_x = int(screen_width * 0.95)
        right_hover_y = screen_height // 2
        left_hover_x = int(screen_width * 0.05)
        left_hover_y = screen_height // 2

    # Find right arrow: move mouse to right margin, wait for arrow to appear
    print(f"  Bewege Maus zum rechten Rand ({right_hover_x}, {right_hover_y})...")
    pyautogui.moveTo(right_hover_x, right_hover_y, duration=0.1)
    time.sleep(0.5)  # Wait for arrow to appear

    # Take screenshot and find arrow
    screenshot = ImageGrab.grab()
    right_pos = find_arrow_button(screenshot, side='right')
    if right_pos:
        ARROW_RIGHT_POS = right_pos
        print(f"  Rechter Pfeil (weiter): ({right_pos[0]}, {right_pos[1]})")
    else:
        # Use hover position as fallback
        ARROW_RIGHT_POS = (right_hover_x, right_hover_y)
        print(f"  [INFO] Verwende Hover-Position als Pfeilposition: ({right_hover_x}, {right_hover_y})")

    # Find left arrow: move mouse to left margin, wait for arrow to appear
    print(f"  Bewege Maus zum linken Rand ({left_hover_x}, {left_hover_y})...")
    pyautogui.moveTo(left_hover_x, left_hover_y, duration=0.1)
    time.sleep(0.5)  # Wait for arrow to appear

    # Take screenshot and find arrow
    screenshot = ImageGrab.grab()
    left_pos = find_arrow_button(screenshot, side='left')
    if left_pos:
        ARROW_LEFT_POS = left_pos
        print(f"  Linker Pfeil (zurueck): ({left_pos[0]}, {left_pos[1]})")
    else:
        # Use hover position as fallback
        ARROW_LEFT_POS = (left_hover_x, left_hover_y)
        print(f"  [INFO] Verwende Hover-Position als Pfeilposition: ({left_hover_x}, {left_hover_y})")

    # Move mouse away from margins
    pyautogui.moveTo(screen_width // 2, screen_height // 2, duration=0.1)

    return ARROW_RIGHT_POS is not None


def click_next_page():
    """Click to go to next page."""
    global ARROW_RIGHT_POS

    if FULLSCREEN_MODE:
        if ARROW_RIGHT_POS:
            # Use detected arrow position
            click_at(ARROW_RIGHT_POS[0], ARROW_RIGHT_POS[1], show_indicator=True)
        else:
            # Fallback: click on the right side of the screen
            screen_width, screen_height = pyautogui.size()
            click_at(int(screen_width * 0.95), int(screen_height * 0.5), show_indicator=True)
        return True
    return click_kindle(0.95, 0.5, show_indicator=True)


def click_prev_page():
    """Click to go to previous page."""
    global ARROW_LEFT_POS

    if FULLSCREEN_MODE:
        if ARROW_LEFT_POS:
            # Use detected arrow position
            click_at(ARROW_LEFT_POS[0], ARROW_LEFT_POS[1], show_indicator=True)
        else:
            # Fallback: click on the left side of the screen
            screen_width, screen_height = pyautogui.size()
            click_at(int(screen_width * 0.05), int(screen_height * 0.5), show_indicator=True)
        return True
    return click_kindle(0.05, 0.5, show_indicator=True)

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


def click_menu_item(menu_name, item_text, timeout=5):
    """Click a menu item in Kindle by searching for text."""
    bounds = activate_and_get_kindle()
    if not bounds:
        return False

    left, top, width, height = bounds

    # Menu bar is at the top of the window
    # "Gehe zu" is typically around 150-200px from left
    menu_positions = {
        'Gehe zu': 0.15,  # ~15% from left
    }

    if menu_name in menu_positions:
        menu_x = menu_positions[menu_name]
        # Click on menu (menu bar is about 60px from top of window)
        abs_x = left + int(width * menu_x)
        abs_y = top + 62  # Menu bar height
        pyautogui.click(abs_x, abs_y)
        time.sleep(0.5)
        return True
    return False


def open_goto_menu():
    """Open the 'Gehe zu' menu in Kindle."""
    print("[INFO] Oeffne 'Gehe zu' Menu...")
    bounds = activate_and_get_kindle()
    if not bounds:
        return False

    left, top, width, height = bounds

    # "Gehe zu" menu position (approximately)
    # Based on screenshot: around 300px from left on a typical window
    abs_x = left + 300
    abs_y = top + 100  # Menu bar area

    pyautogui.click(abs_x, abs_y)
    time.sleep(0.5)
    return True


def find_image_on_screen(template_path, region=None, confidence=0.8):
    """Find an image on screen and return its center coordinates."""
    try:
        location = pyautogui.locateOnScreen(template_path, region=region, confidence=confidence)
        if location:
            center = pyautogui.center(location)
            return (center.x, center.y)
    except Exception as e:
        print(f"  [DEBUG] Bilderkennung fehlgeschlagen: {e}")
    return None


def find_icon_in_toolbar(icon_name, bounds):
    """Find an icon in the Kindle toolbar by template matching."""
    left, top, width, height = bounds

    # Search only in toolbar area (top 130px)
    toolbar_region = (left, top, width, 130)

    script_dir = Path(__file__).parent if not getattr(sys, 'frozen', False) else Path(sys.executable).parent
    template_path = script_dir / "templates" / f"{icon_name}.png"

    if template_path.exists():
        pos = find_image_on_screen(str(template_path), region=toolbar_region, confidence=0.7)
        if pos:
            return pos

    return None


def navigate_to_title_page():
    """Navigate to title page using 'Gehe zu' -> 'Titelseite' via OCR."""
    print("[INFO] Navigiere zur Titelseite...")

    # Get OCR engine
    engine = get_ocr_engine()
    if not engine:
        print("  [WARNUNG] Windows OCR nicht verfuegbar, verwende Fallback...")
        return navigate_to_title_page_fallback()

    # Get Kindle window with pywinauto
    kindle_window, bounds = get_kindle_window_pywinauto()
    if not kindle_window or not bounds:
        print("  [WARNUNG] pywinauto fehlgeschlagen, verwende Fallback...")
        return navigate_to_title_page_fallback()

    win_left, win_top, win_width, win_height = bounds
    print(f"  Fenster: left={win_left}, top={win_top}, width={win_width}, height={win_height}")

    # Step 1: Find "Gehe zu" in menu bar via OCR
    print("  Suche 'Gehe zu' im Hauptmenue...")
    menu_region = (win_left, win_top - 5, win_left + 300, win_top + 30)
    result = find_text_in_region(engine, menu_region, "Gehe", take_topmost=True)

    if not result:
        print("  [WARNUNG] 'Gehe zu' nicht gefunden, verwende Fallback...")
        return navigate_to_title_page_fallback()

    goto_x, goto_y, goto_info = result
    print(f"  'Gehe zu' gefunden bei ({goto_x}, {goto_y})")

    # Step 2: Click on "Gehe zu"
    click_at(goto_x, goto_y)
    time.sleep(0.8)

    # Step 3: Find "Titelseite" in dropdown via OCR
    print("  Suche 'Titelseite' im Dropdown...")
    dropdown_region = (
        goto_x - 50,
        goto_y + 5,
        goto_x + 150,
        goto_y + 180
    )
    result = find_text_in_region(engine, dropdown_region, "Titelseite", take_topmost=True)

    if not result:
        print("  [WARNUNG] 'Titelseite' nicht gefunden!")
        pyautogui.press('escape')
        return False

    title_x, title_y, title_info = result
    print(f"  'Titelseite' gefunden bei ({title_x}, {title_y})")

    # Step 4: Click on "Titelseite"
    click_at(title_x, title_y)
    time.sleep(1.0)

    print("[OK] Zur Titelseite navigiert")
    return True


def navigate_to_title_page_fallback():
    """Fallback navigation using fixed positions."""
    bounds = activate_and_get_kindle()
    if not bounds:
        return False

    left, top, width, height = bounds

    # Fixed position fallback
    goto_x = left + int(width * 0.31)
    goto_y = top + 100
    print(f"  Verwende Fallback-Position: ({goto_x}, {goto_y})")

    click_at(goto_x, goto_y)
    time.sleep(0.7)

    title_x = goto_x
    title_y = goto_y + 72
    click_at(title_x, title_y)
    time.sleep(1.0)

    print("[OK] Zur Titelseite navigiert (Fallback)")
    return True


def click_fullscreen_button():
    """Click the fullscreen button via Ansicht -> Vollbildmodus menu."""
    print("[INFO] Aktiviere Vollbildmodus...")

    # Get OCR engine
    engine = get_ocr_engine()
    if not engine:
        print("  [WARNUNG] Windows OCR nicht verfuegbar, verwende Fallback...")
        return click_fullscreen_button_fallback()

    # Get Kindle window with pywinauto
    kindle_window, bounds = get_kindle_window_pywinauto()
    if not kindle_window or not bounds:
        print("  [WARNUNG] pywinauto fehlgeschlagen, verwende Fallback...")
        return click_fullscreen_button_fallback()

    win_left, win_top, win_width, win_height = bounds

    # Step 1: Find "Ansicht" in menu bar via OCR
    print("  Suche 'Ansicht' im Hauptmenue...")
    menu_region = (win_left, win_top - 5, win_left + 300, win_top + 30)
    result = find_text_in_region(engine, menu_region, "Ansicht", take_topmost=True)

    if not result:
        print("  [WARNUNG] 'Ansicht' nicht gefunden, verwende Fallback...")
        return click_fullscreen_button_fallback()

    ansicht_x, ansicht_y, ansicht_info = result
    print(f"  'Ansicht' gefunden bei ({ansicht_x}, {ansicht_y})")

    # Step 2: Click on "Ansicht"
    click_at(ansicht_x, ansicht_y)
    time.sleep(0.8)

    # Step 3: Find "Vollbildmodus" in dropdown via OCR
    print("  Suche 'Vollbildmodus' im Dropdown...")
    dropdown_region = (
        ansicht_x - 30,
        ansicht_y + 5,
        ansicht_x + 250,
        ansicht_y + 200
    )
    result = find_text_in_region(engine, dropdown_region, "Vollbild", take_topmost=True)

    if not result:
        print("  [WARNUNG] 'Vollbildmodus' nicht gefunden!")
        pyautogui.press('escape')
        return False

    vollbild_x, vollbild_y, vollbild_info = result
    print(f"  'Vollbildmodus' gefunden bei ({vollbild_x}, {vollbild_y})")

    # Step 4: Click on "Vollbildmodus"
    click_at(vollbild_x, vollbild_y)
    time.sleep(1.0)

    # Switch to fullscreen mode
    set_fullscreen_mode(True)

    print("[OK] Vollbildmodus aktiviert")
    return True


def click_fullscreen_button_fallback():
    """Fallback fullscreen using fixed position."""
    bounds = activate_and_get_kindle()
    if not bounds:
        return False

    left, top, width, height = bounds

    # Fixed position fallback for fullscreen icon
    fs_x = left + int(width * 0.47)
    fs_y = top + 100
    print(f"  Verwende Fallback-Position: ({fs_x}, {fs_y})")

    click_at(fs_x, fs_y)
    time.sleep(0.5)

    set_fullscreen_mode(True)

    print("[OK] Vollbildmodus aktiviert (Fallback)")
    return True


def wait_for_fullscreen_message_to_disappear():
    """Wait until the 'Press F11 to exit' message disappears using OCR."""
    print("[INFO] Warte bis Vollbild-Hinweis verschwindet...")

    engine = get_ocr_engine()
    if not engine:
        # Fallback: just wait a fixed time
        print("  [INFO] OCR nicht verfuegbar, warte 5 Sekunden...")
        time.sleep(5.0)
        return True

    max_wait = 10  # max seconds to wait
    check_interval = 0.5

    screen_width, screen_height = pyautogui.size()

    for i in range(int(max_wait / check_interval)):
        check_stop_and_exit()

        # Check center region where the message appears
        center_region = (
            screen_width // 2 - 200,
            screen_height // 2 - 50,
            screen_width // 2 + 200,
            screen_height // 2 + 50
        )

        result = find_text_in_region(engine, center_region, "Beenden", take_topmost=True, scale_factor=2)

        if result is None:
            print(f"[OK] Vollbild-Hinweis verschwunden nach {(i+1) * check_interval:.1f}s")
            return True

        time.sleep(check_interval)

    print("[WARNUNG] Timeout beim Warten auf Vollbild-Hinweis")
    return True  # Continue anyway


def prepare_kindle_for_capture():
    """Complete preparation sequence: find window, goto title, fullscreen.
    Returns the detected book region or None on failure."""
    print()
    print("[SCHRITT 1/6] Kindle-Fenster finden...")
    if not find_and_activate_kindle():
        return None

    time.sleep(0.5)

    print()
    print("[SCHRITT 2/6] Zur Titelseite navigieren...")
    if not navigate_to_title_page():
        print("[WARNUNG] Navigation zur Titelseite fehlgeschlagen")
        # Continue anyway - user might already be at title

    time.sleep(0.5)

    print()
    print("[SCHRITT 3/6] Vollbildmodus aktivieren...")
    if not click_fullscreen_button():
        print("[WARNUNG] Vollbildmodus-Aktivierung fehlgeschlagen")
        return None

    print()
    print("[SCHRITT 4/6] Warte auf Vollbildmodus...")
    wait_for_fullscreen_message_to_disappear()

    # Extra wait to ensure everything is stable
    time.sleep(1.0)

    print()
    print("[SCHRITT 5/6] Erkenne Buchbereich...")
    screenshot = grab_kindle_screenshot()
    if screenshot is None:
        print("[FEHLER] Konnte Screenshot nicht erstellen!")
        return None

    screen_width, screen_height = screenshot.size

    # In fullscreen: top=0, bottom=screen_height (always full height)
    # Only left/right margins need to be detected from title page
    print("[INFO] Erkenne linken/rechten Rand von Titelseite...")
    title_region = detect_book_region_from_title_page(screenshot)
    if not title_region:
        print("[FEHLER] Konnte Buchbereich nicht erkennen!")
        return None

    t_left, _, t_right, _ = title_region

    # Use full screen height, only crop left/right
    book_region = (t_left, 0, t_right, screen_height)
    print(f"[OK] Buchbereich: ({t_left}, 0) bis ({t_right}, {screen_height})")
    print(f"     Groesse: {t_right - t_left} x {screen_height} Pixel")

    print()
    print("[SCHRITT 6/6] Erkenne Navigationspfeile...")
    detect_arrow_positions(book_region)

    print()
    print("[OK] Kindle bereit fuer Erfassung!")
    return book_region

# Global flag to indicate fullscreen mode
FULLSCREEN_MODE = False
FULLSCREEN_BOUNDS = None

def grab_kindle_screenshot():
    """Take screenshot of Kindle window or fullscreen."""
    global FULLSCREEN_MODE, FULLSCREEN_BOUNDS

    if FULLSCREEN_MODE:
        # In fullscreen, grab the whole screen or cached bounds
        if FULLSCREEN_BOUNDS:
            left, top, width, height = FULLSCREEN_BOUNDS
            return ImageGrab.grab(bbox=(left, top, left + width, top + height))
        else:
            return ImageGrab.grab()

    # Normal window mode
    bounds = activate_and_get_kindle()
    if bounds:
        left, top, width, height = bounds
        return ImageGrab.grab(bbox=(left, top, left + width, top + height))
    return None


def set_fullscreen_mode(enabled, bounds=None):
    """Set fullscreen mode and optionally cache screen bounds."""
    global FULLSCREEN_MODE, FULLSCREEN_BOUNDS
    FULLSCREEN_MODE = enabled
    FULLSCREEN_BOUNDS = bounds

# ============================================================
# Image Analysis
# ============================================================

def detect_background_color(screenshot):
    """Detect if background is dark or light."""
    img_array = np.array(screenshot)
    if len(img_array.shape) == 3:
        gray = np.mean(img_array[:, :, :3], axis=2)
    else:
        gray = img_array

    height, width = gray.shape
    samples = [
        gray[10, 10],
        gray[10, width-10],
        gray[height-10, 10],
        gray[height-10, width-10],
    ]
    avg_corner = np.mean(samples)
    return 'dark' if avg_corner < 128 else 'light'


def find_arrow_button(screenshot, side='right'):
    """Find the gray navigation arrow button on the margin.

    The arrow is a gray ">" or "<" symbol on the black/white margin.
    Returns (x, y) position of the arrow center, or None if not found.

    Args:
        screenshot: PIL Image of the screen
        side: 'right' for next page arrow, 'left' for previous page arrow
    """
    img_array = np.array(screenshot)
    if len(img_array.shape) == 3:
        gray = img_array[:, :, :3].mean(axis=2).astype(np.uint8)
    else:
        gray = img_array

    height, width = gray.shape

    # Determine margin brightness (black ~0 or white ~255)
    if side == 'right':
        margin_strip = gray[height//4:3*height//4, width-30:width]
    else:
        margin_strip = gray[height//4:3*height//4, 0:30]

    margin_brightness = np.mean(margin_strip)
    is_dark_margin = margin_brightness < 128

    # The arrow is a light gray (~200-230) on white background or darker gray on black
    # We need to detect subtle differences

    if side == 'right':
        # Search in the right margin area (last 20% of width, excluding edge)
        search_left = int(width * 0.80)
        search_right = width - 5
    else:
        # Search in the left margin area (first 20% of width, excluding edge)
        search_left = 5
        search_right = int(width * 0.20)

    # Search vertically in the middle 60% of height
    search_top = int(height * 0.20)
    search_bottom = int(height * 0.80)

    # Extract the search region
    search_region = gray[search_top:search_bottom, search_left:search_right].astype(float)

    # Calculate local variance to find the arrow (edges have higher variance)
    # Use a simple gradient approach: find vertical edges
    # The arrow ">" has strong vertical gradients

    # Calculate horizontal gradient (difference between adjacent columns)
    gradient = np.abs(np.diff(search_region, axis=1))

    # The arrow should have gradients in the middle vertically
    # Find columns with significant gradients
    col_gradient_sum = np.sum(gradient, axis=0)

    # Find the column with maximum gradient (this is where the arrow edge is)
    if len(col_gradient_sum) == 0:
        print(f"  [DEBUG] Kein Gradient auf {side} Seite")
        return None

    max_gradient_col = np.argmax(col_gradient_sum)
    max_gradient_value = col_gradient_sum[max_gradient_col]

    # Check if gradient is significant
    mean_gradient = np.mean(col_gradient_sum)
    if max_gradient_value < mean_gradient * 2:
        print(f"  [DEBUG] Kein signifikanter Gradient auf {side} Seite (max={max_gradient_value:.0f}, mean={mean_gradient:.0f})")
        return None

    # Find the vertical center of the gradient in that column
    col_gradients = gradient[:, max_gradient_col]
    gradient_indices = np.where(col_gradients > np.mean(col_gradients) * 2)[0]

    if len(gradient_indices) == 0:
        print(f"  [DEBUG] Keine Gradient-Positionen auf {side} Seite")
        return None

    # Calculate center of arrow
    center_y_local = int(np.mean(gradient_indices))
    center_x_local = max_gradient_col

    # Convert back to full image coordinates
    abs_x = search_left + center_x_local
    abs_y = search_top + center_y_local

    print(f"  Pfeil gefunden bei ({abs_x}, {abs_y}) - Gradient: {max_gradient_value:.0f}")

    return (abs_x, abs_y)

def detect_book_region_from_title_page(screenshot):
    """Detect book region from title page.

    Uses variance-based detection: the book content has pixel variation (text, images),
    while uniform margins (black or white) have near-zero variance.
    """
    img_array = np.array(screenshot)
    if len(img_array.shape) == 3:
        gray = np.mean(img_array[:, :, :3], axis=2)
    else:
        gray = img_array.astype(float)

    height, width = gray.shape

    # Get margin info for debugging
    margin_brightness = np.mean([gray[height//2, 5], gray[height//2, width-5]])
    print(f"  Margin-Helligkeit: {margin_brightness:.0f}")

    # Calculate variance for each column - content columns have higher variance
    # Use a sliding window to smooth out noise
    col_variance = np.zeros(width)
    for col in range(width):
        col_variance[col] = np.var(gray[height//4:3*height//4, max(0,col-2):min(width,col+3)])

    # Find threshold: margins have very low variance, content has higher
    max_variance = np.max(col_variance)
    variance_threshold = max_variance * 0.05  # 5% of max variance indicates content

    print(f"  Max Varianz: {max_variance:.0f}, Schwelle: {variance_threshold:.0f}")

    # Find left edge - first column with significant variance
    left = 0
    for col in range(width // 2):
        if col_variance[col] > variance_threshold:
            left = col
            break

    # Find right edge - last column with significant variance
    right = width
    for col in range(width - 1, width // 2, -1):
        if col_variance[col] > variance_threshold:
            right = col + 1
            break

    # Calculate row variance for top/bottom detection
    row_variance = np.zeros(height)
    for row in range(height):
        row_variance[row] = np.var(gray[max(0,row-2):min(height,row+3), left:right])

    row_variance_threshold = np.max(row_variance) * 0.05

    # Find top edge
    top = 0
    for row in range(height // 2):
        if row_variance[row] > row_variance_threshold:
            top = row
            break

    # Find bottom edge
    bottom = height
    for row in range(height - 1, height // 2, -1):
        if row_variance[row] > row_variance_threshold:
            bottom = row + 1
            break

    book_width = right - left
    book_height = bottom - top

    print(f"  Erkannte Grenzen: left={left}, right={right}, top={top}, bottom={bottom}")
    print(f"  Buchgroesse: {book_width} x {book_height}")

    # Validate: book should be reasonable size
    if book_width > width * 0.2 and book_height > height * 0.3:
        return (left, top, right, bottom)

    print("  [WARNUNG] Varianz-Erkennung fehlgeschlagen, verwende Fallback...")

    # Fallback: assume centered content
    content_width = int(width * 0.55)
    left = (width - content_width) // 2
    right = left + content_width
    top = int(height * 0.05)
    bottom = int(height * 0.85)

    print(f"  Fallback-Grenzen: left={left}, right={right}, top={top}, bottom={bottom}")
    return (left, top, right, bottom)

def margin_and_book_same_color(screenshot):
    """Check if margin and book content have similar colors."""
    img_array = np.array(screenshot)
    if len(img_array.shape) == 3:
        gray = np.mean(img_array[:, :, :3], axis=2)
    else:
        gray = img_array

    height, width = gray.shape

    margin_samples = [
        gray[10, 10],
        gray[10, width-10],
        gray[height-10, 10],
        gray[height-10, width-10],
    ]
    margin_color = np.mean(margin_samples)

    center_y = height // 2
    center_x = width // 2
    center_samples = [
        gray[center_y-50:center_y+50, center_x-50:center_x+50].mean()
    ]
    center_color = np.mean(center_samples)

    threshold = 50
    return abs(margin_color - center_color) < threshold

def find_text_bounds(screenshot):
    """Find the text region by detecting text pixels."""
    img_array = np.array(screenshot)

    if len(img_array.shape) == 3:
        gray = np.mean(img_array[:, :, :3], axis=2)
    else:
        gray = img_array

    height, width = gray.shape

    bg_type = detect_background_color(screenshot)

    if bg_type == 'dark':
        has_text = lambda arr: np.max(arr) > 50
    else:
        has_text = lambda arr: np.min(arr) < 200

    top_section = int(height * 0.1)
    bottom_section = int(height * 0.9)

    left = 0
    for col in range(width):
        if has_text(gray[:, col]):
            left = col
            break

    right = width
    for col in range(width - 1, -1, -1):
        if has_text(gray[:top_section, col]) or has_text(gray[bottom_section:, col]):
            right = col + 1
            break

    top = 0
    for row in range(height):
        if has_text(gray[row, left:right]):
            top = row
            break

    bottom = height
    for row in range(height - 1, -1, -1):
        if has_text(gray[row, left:right]):
            bottom = row + 1
            break

    return (left, top, right, bottom)

def images_are_similar(img1, img2, threshold=0.99):
    """Compare two images and return True if they are very similar."""
    if img1 is None or img2 is None:
        return False

    if img1.size != img2.size:
        return False

    arr1 = np.array(img1)
    arr2 = np.array(img2)

    diff = np.abs(arr1.astype(float) - arr2.astype(float))
    similarity = 1 - (np.mean(diff) / 255)

    return similarity > threshold

# ============================================================
# Book Region Calibration
# ============================================================

def calibrate_book_region(max_pages_to_search=50, consecutive_matches_needed=3, tolerance=20):
    """Detect book region."""
    screenshot = grab_kindle_screenshot()
    if screenshot is None:
        print("[FEHLER] Konnte Kindle-Screenshot nicht erstellen!")
        return None

    screen_width, screen_height = screenshot.size

    if not margin_and_book_same_color(screenshot):
        print("[INFO] Erkenne Buchbereich von Titelseite...")
        region = detect_book_region_from_title_page(screenshot)
        if region:
            left, top, right, bottom = region
            print(f"[OK] Buchbereich erkannt: ({left}, {top}, {right}, {bottom})")
            return region
        print("[INFO] Direkte Erkennung fehlgeschlagen, wechsle zu Seitensuche...")

    print(f"[INFO] Suche nach {consecutive_matches_needed} aufeinanderfolgenden Seiten mit gleichem rechten Rand...")

    min_left = 99999
    min_top = 99999
    max_bottom = 0
    pages_navigated = 0

    recent_right_edges = []
    found_stable_right = False
    stable_right_edge = 0
    full_width_count = 0  # Count pages that extend to full width

    for i in range(max_pages_to_search):
        check_stop_and_exit()

        time.sleep(0.3)
        screenshot = grab_kindle_screenshot()
        if screenshot is None:
            print("[FEHLER] Konnte Kindle-Screenshot nicht erstellen!")
            return None

        bounds = find_text_bounds(screenshot)
        left, top, right, bottom = bounds

        min_left = min(min_left, left)
        min_top = min(min_top, top)
        max_bottom = max(max_bottom, bottom)

        # Track all right edges, including full-width ones
        recent_right_edges.append(right)

        # Check if this is a full-width page (no visible right margin)
        is_full_width = right >= screen_width * 0.9
        if is_full_width:
            full_width_count += 1
            print(f"  Seite +{i+1}: rechter Rand = {right} (Vollbreite)")
        else:
            print(f"  Seite +{i+1}: rechter Rand = {right}")

        # Check for stable right edge (3 consecutive similar values)
        if len(recent_right_edges) >= consecutive_matches_needed:
            last_edges = recent_right_edges[-consecutive_matches_needed:]
            min_edge = min(last_edges)
            max_edge = max(last_edges)

            if max_edge - min_edge <= tolerance:
                stable_right_edge = max(last_edges)
                found_stable_right = True
                print(f"  --> Stabiler Rand gefunden: {stable_right_edge}")
                pages_navigated = i + 1
                break

        click_next_page()
        time.sleep(0.5)
        pages_navigated = i + 1

    if not found_stable_right:
        # If most pages were full-width, use full screen width
        if full_width_count >= pages_navigated * 0.7:
            print("[INFO] Buch nutzt volle Bildschirmbreite (keine sichtbaren Raender)")
            stable_right_edge = screen_width
        elif recent_right_edges:
            # Use the most common right edge
            stable_right_edge = max(recent_right_edges)
            print(f"[INFO] Verwende maximalen rechten Rand: {stable_right_edge}")
        else:
            stable_right_edge = int(screen_width * 0.6)
            print(f"[WARNUNG] Kein Rand gefunden, verwende Schaetzung: {stable_right_edge}")

    padding = 20
    screenshot = grab_kindle_screenshot()
    if screenshot:
        screen_width, screen_height = screenshot.size

        final_left = max(0, min_left - padding)
        final_right = min(screen_width, stable_right_edge + padding)
        final_top = max(0, min_top - padding)
        final_bottom = min(screen_height, max_bottom + padding)

        print(f"[INFO] Buchbereich: ({final_left}, {final_top}, {final_right}, {final_bottom})")

        print(f"[INFO] Navigiere {pages_navigated} Seiten zurueck...")
        for i in range(pages_navigated):
            click_prev_page()
            time.sleep(0.5)

        time.sleep(0.3)
        return (final_left, final_top, final_right, final_bottom)

    return None

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

def capture_pages(output_folder, book_region):
    """Capture all pages from Kindle book."""
    global STOP_FLAG

    page_num = 1
    no_change_count = 0
    last_page_image = None

    try:
        while no_change_count < MAX_NO_CHANGE_COUNT:
            check_stop_and_exit()

            screenshot = grab_kindle_screenshot()
            if screenshot is None:
                print("[FEHLER] Kindle-Fenster verloren!")
                break

            book_page = screenshot.crop(book_region)

            if images_are_similar(book_page, last_page_image):
                no_change_count += 1
                print(f"[INFO] Keine Aenderung erkannt ({no_change_count}/{MAX_NO_CHANGE_COUNT})")
            else:
                no_change_count = 0

                filename = f"page_{page_num:04d}.png"
                filepath = output_folder / filename
                book_page.save(filepath, "PNG")

                while not filepath.exists():
                    check_stop_and_exit()
                    time.sleep(0.05)

                print(f"[OK] Gespeichert: {filename}")

                last_page_image = book_page.copy()
                page_num += 1

            click_next_page()
            time.sleep(WAIT_AFTER_CLICK)

    except KeyboardInterrupt:
        print("\n[INFO] Erfassung vom Benutzer gestoppt.")
    except SystemExit:
        raise

    return page_num - 1

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

    # Prepare Kindle: find window, goto title page, enter fullscreen, detect book region
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
