"""
Create PDF from Book Pages
==========================
Converts captured book page images into a searchable PDF with OCR text.
Uses Windows OCR (built-in, no external dependencies).

Features:
- JPEG compression for smaller file size (optimized for AI processing)
- Optional image scaling for even smaller PDFs
- OCR text layer for searchability

Usage:
1. Change to the book folder containing page_*.png files
2. Run: create_pdf.exe
3. The PDF will be created in the same folder

Requirements:
- Windows 10/11 or Windows Server 2016+
- OCR language pack installed (see README)

Author: Claude
"""

import sys
import asyncio
import io
from pathlib import Path
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader

# ============================================================
# Configuration for AI-optimized output
# ============================================================
JPEG_QUALITY = 75  # JPEG quality (1-100), 75 is good balance
MAX_WIDTH = 1400   # Max image width in pixels (None to disable)
SCALE_FACTOR = 0.75  # PDF scale factor

# Windows OCR imports
try:
    from winsdk.windows.media.ocr import OcrEngine
    from winsdk.windows.globalization import Language
    from winsdk.windows.graphics.imaging import BitmapDecoder
    from winsdk.windows.storage import StorageFile
    WINDOWS_OCR_AVAILABLE = True
except ImportError:
    WINDOWS_OCR_AVAILABLE = False

# ============================================================
# Windows OCR
# ============================================================

def check_ocr_languages():
    """Check available OCR languages and return the best engine."""
    if not WINDOWS_OCR_AVAILABLE:
        return None, "winsdk nicht installiert"

    # Try German first, then English
    for lang_tag in ['de-DE', 'de', 'en-US', 'en']:
        try:
            lang = Language(lang_tag)
            if OcrEngine.is_language_supported(lang):
                engine = OcrEngine.try_create_from_language(lang)
                if engine:
                    return engine, lang_tag
        except Exception:
            continue

    # Try system default
    try:
        engine = OcrEngine.try_create_from_user_profile_languages()
        if engine:
            return engine, "system-default"
    except Exception:
        pass

    return None, "Keine OCR-Sprache installiert"


async def ocr_image_async(engine, img_path):
    """Perform OCR using Windows OCR asynchronously."""
    try:
        # Convert path to absolute Windows path
        abs_path = str(Path(img_path).resolve())

        # Open file via Windows Storage API
        storage_file = await StorageFile.get_file_from_path_async(abs_path)

        # Open stream and decode image
        stream = await storage_file.open_read_async()
        decoder = await BitmapDecoder.create_async(stream)

        # Get SoftwareBitmap
        bitmap = await decoder.get_software_bitmap_async()

        # Perform OCR
        result = await engine.recognize_async(bitmap)

        words = []
        if result and result.lines:
            for line in result.lines:
                for word in line.words:
                    text = word.text.strip()
                    if text:
                        rect = word.bounding_rect
                        x = int(rect.x)
                        y = int(rect.y)
                        w = int(rect.width)
                        h = int(rect.height)
                        words.append((text, x, y, w, h))

        return words

    except Exception as e:
        print(f"[WARNUNG] OCR fehlgeschlagen: {e}", end=" ")
        return []


def ocr_image_windows(engine, img_path):
    """Synchronous wrapper for Windows OCR."""
    try:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            return loop.run_until_complete(ocr_image_async(engine, img_path))
        finally:
            loop.close()
    except Exception as e:
        print(f"[WARNUNG] OCR fehlgeschlagen: {e}", end=" ")
        return []

# ============================================================
# PDF Creation
# ============================================================

def compress_image_to_jpeg(img_path, max_width=MAX_WIDTH, quality=JPEG_QUALITY):
    """Load image, optionally resize, and compress to JPEG in memory.
    Returns (ImageReader, original_width, original_height, new_width, new_height)."""
    img = Image.open(img_path)
    orig_width, orig_height = img.size

    # Convert to RGB if necessary (JPEG doesn't support alpha)
    if img.mode in ('RGBA', 'P'):
        img = img.convert('RGB')

    # Resize if needed
    new_width, new_height = orig_width, orig_height
    if max_width and orig_width > max_width:
        ratio = max_width / orig_width
        new_width = max_width
        new_height = int(orig_height * ratio)
        img = img.resize((new_width, new_height), Image.LANCZOS)

    # Compress to JPEG in memory
    buffer = io.BytesIO()
    img.save(buffer, format='JPEG', quality=quality, optimize=True)
    buffer.seek(0)

    return ImageReader(buffer), orig_width, orig_height, new_width, new_height


def create_pdf(output_folder):
    """Create searchable PDF from captured pages with JPEG compression."""
    folder = Path(output_folder)

    # Look for PNGs in 'pages' subfolder first, then in folder itself
    pages_dir = folder / "pages"
    if pages_dir.exists():
        files = sorted(pages_dir.glob("page_*.png"))
        print(f"[INFO] Lese Seiten aus: {pages_dir}")
    else:
        files = sorted(folder.glob("page_*.png"))

    if not files:
        print("[FEHLER] Keine page_*.png Dateien gefunden!")
        return False

    print(f"[INFO] {len(files)} Seiten gefunden")
    print(f"[INFO] JPEG-Qualitaet: {JPEG_QUALITY}%, Max-Breite: {MAX_WIDTH}px")

    # Initialize Windows OCR
    engine, lang_info = check_ocr_languages()
    if not engine:
        print(f"[FEHLER] Windows OCR nicht verfuegbar: {lang_info}")
        print("  Installiere OCR-Sprachpaket mit PowerShell (als Admin):")
        print('  Add-WindowsCapability -Online -Name "Language.OCR~~~de-DE~0.0.1.0"')
        return False

    print(f"[INFO] Windows OCR Sprache: {lang_info}")
    print()

    output_filename = folder.name + ".pdf"
    output_path = folder / output_filename

    # Get dimensions from first image
    first_img = Image.open(files[0])
    orig_width, orig_height = first_img.size

    # Calculate scaled dimensions
    if MAX_WIDTH and orig_width > MAX_WIDTH:
        ratio = MAX_WIDTH / orig_width
        img_width = MAX_WIDTH
        img_height = int(orig_height * ratio)
    else:
        img_width, img_height = orig_width, orig_height

    page_width = img_width * SCALE_FACTOR
    page_height = img_height * SCALE_FACTOR

    print(f"[INFO] Original: {orig_width}x{orig_height} -> Skaliert: {img_width}x{img_height}")
    print(f"[INFO] PDF-Seite: {int(page_width)}x{int(page_height)}")
    print(f"[INFO] Ausgabe: {output_path}")
    print()

    c = canvas.Canvas(str(output_path), pagesize=(page_width, page_height))

    try:
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))
        font_name = 'Arial'
    except:
        font_name = 'Helvetica'

    pages_processed = 0
    total_original_size = 0
    total_compressed_size = 0

    for i, file in enumerate(files):
        print(f"[{i+1}/{len(files)}] {file.name}...", end=" ", flush=True)

        # Track original file size
        total_original_size += file.stat().st_size

        # OCR on original image (before compression)
        words = ocr_image_windows(engine, file)

        # Compress image and add to PDF
        img_reader, orig_w, orig_h, new_w, new_h = compress_image_to_jpeg(file)

        # Calculate scale factors for OCR coordinates
        ocr_scale_x = (new_w / orig_w) * SCALE_FACTOR if orig_w != new_w else SCALE_FACTOR
        ocr_scale_y = (new_h / orig_h) * SCALE_FACTOR if orig_h != new_h else SCALE_FACTOR

        c.drawImage(img_reader, 0, 0, width=page_width, height=page_height)

        if words:
            text_object = c.beginText()
            text_object.setTextRenderMode(3)  # Invisible
            text_object.setFont(font_name, 10)

            for text, x, y, w, h in words:
                pdf_x = x * ocr_scale_x
                pdf_y = page_height - (y + h) * ocr_scale_y

                font_size = max(h * ocr_scale_y * 0.8, 6)

                text_object.setFont(font_name, font_size)
                text_object.setTextOrigin(pdf_x, pdf_y)
                try:
                    text_object.textOut(text + " ")
                except:
                    pass

            c.drawText(text_object)

        c.showPage()
        pages_processed = i + 1
        print(f"OK ({len(words)} Woerter)")

    c.save()

    # Calculate compression stats
    pdf_size = output_path.stat().st_size
    compression_ratio = (1 - pdf_size / total_original_size) * 100 if total_original_size > 0 else 0

    print()
    print("=" * 60)
    print("  PDF-ERSTELLUNG ABGESCHLOSSEN")
    print("=" * 60)
    print(f"  Seiten: {pages_processed}")
    print(f"  Original PNGs: {total_original_size / (1024*1024):.1f} MB")
    print(f"  PDF-Groesse: {pdf_size / (1024*1024):.1f} MB")
    print(f"  Kompression: {compression_ratio:.0f}%")
    print(f"  Ausgabe: {output_path}")
    print("=" * 60)

    return True

def main():
    print("=" * 60)
    print("  PDF-ERSTELLUNG MIT OCR")
    print("=" * 60)
    print()

    output_folder = Path.cwd()
    success = create_pdf(output_folder)

    if success:
        sys.exit(0)
    else:
        sys.exit(1)

if __name__ == "__main__":
    main()
