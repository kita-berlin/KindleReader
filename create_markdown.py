"""
Create Markdown from Book Pages
================================
Converts captured book page images (PNGs) or PDF files into a structured
Markdown file with extracted text (OCR) and preserved images (charts, diagrams).

Input sources (in priority order):
1. pages/page_*.png  - PNG screenshots from kindle_capture.py
2. *.pdf              - Any PDF in the book folder

Output structure:
  book_folder/
    markdown/        <- output
      <bookname>.md  <- full text with image references
      page_NNNN.jpg  <- extracted chart/diagram images (JPEG)

The script analyzes each page:
- Text-heavy pages: OCR text only (no image saved)
- Image-heavy pages (charts, diagrams): saved as JPEG + referenced in markdown
- Mixed pages: OCR text + image reference

Usage:
1. Change to the book folder
2. Run: python create_markdown.py
3. The markdown/ folder and ZIP will be created

Requirements:
- Windows 10/11 with OCR language pack
- PIL/Pillow, numpy, winsdk
- PyMuPDF (fitz) for PDF input

Author: Claude
"""

import sys
import asyncio
import io
import tempfile
import shutil
from pathlib import Path
from PIL import Image
import numpy as np

# PyMuPDF for PDF page extraction
try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except ImportError:
    PYMUPDF_AVAILABLE = False

# ============================================================
# Configuration
# ============================================================
IMAGE_QUALITY = 80          # JPEG quality for saved images
IMAGE_MAX_WIDTH = 1200      # Max width for saved images
TEXT_VARIANCE_THRESHOLD = 0.15  # Pages with text coverage below this are "image pages"
MIN_TEXT_WORDS = 10         # Minimum words to consider a page as "text page"
PDF_DPI = 200               # DPI for rendering PDF pages as images

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

    for lang_tag in ['de-DE', 'de', 'en-US', 'en']:
        try:
            lang = Language(lang_tag)
            if OcrEngine.is_language_supported(lang):
                engine = OcrEngine.try_create_from_language(lang)
                if engine:
                    return engine, lang_tag
        except Exception:
            continue

    try:
        engine = OcrEngine.try_create_from_user_profile_languages()
        if engine:
            return engine, "system-default"
    except Exception:
        pass

    return None, "Keine OCR-Sprache installiert"


async def ocr_image_async(engine, img_path):
    """Perform OCR and return list of lines with their text."""
    try:
        abs_path = str(Path(img_path).resolve())
        storage_file = await StorageFile.get_file_from_path_async(abs_path)
        stream = await storage_file.open_read_async()
        decoder = await BitmapDecoder.create_async(stream)
        bitmap = await decoder.get_software_bitmap_async()
        result = await engine.recognize_async(bitmap)

        lines = []
        if result and result.lines:
            for line in result.lines:
                line_text = " ".join(w.text.strip() for w in line.words if w.text.strip())
                if line_text:
                    rect = line.words[0].bounding_rect
                    lines.append({
                        'text': line_text,
                        'y': int(rect.y),
                        'height': int(rect.height),
                    })
        return lines

    except Exception as e:
        print(f"  [WARNUNG] OCR fehlgeschlagen: {e}")
        return []


def ocr_image(engine, img_path):
    """Synchronous wrapper for Windows OCR. Returns list of line dicts."""
    try:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            return loop.run_until_complete(ocr_image_async(engine, img_path))
        finally:
            loop.close()
    except Exception as e:
        print(f"  [WARNUNG] OCR fehlgeschlagen: {e}")
        return []


# ============================================================
# PDF Page Extraction
# ============================================================

def extract_pdf_pages(pdf_path, temp_dir, dpi=PDF_DPI):
    """Extract PDF pages as PNG images into temp_dir.
    Returns list of PNG file paths."""
    if not PYMUPDF_AVAILABLE:
        print("[FEHLER] PyMuPDF nicht installiert! pip install PyMuPDF")
        return []

    doc = fitz.open(str(pdf_path))
    files = []
    zoom = dpi / 72.0
    matrix = fitz.Matrix(zoom, zoom)

    print(f"[INFO] Extrahiere {len(doc)} Seiten aus PDF (DPI={dpi})...")

    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=matrix)
        img_path = Path(temp_dir) / f"page_{i+1:04d}.png"
        pix.save(str(img_path))
        files.append(img_path)

    doc.close()
    return files


# ============================================================
# Page Analysis
# ============================================================

def analyze_page(img_path, ocr_lines):
    """Analyze whether a page is text-heavy, image-heavy, or mixed.

    Uses OCR text regions to mask known text areas. Remaining non-text,
    non-empty areas that span at least 2x text line height are graphics.

    Returns:
        'text'  - mostly text, no need to save image
        'image' - mostly image/chart/diagram, save image
        'mixed' - has both significant text and images
    """
    img = Image.open(img_path)
    img_array = np.array(img)

    if len(img_array.shape) == 3:
        gray = np.mean(img_array[:, :, :3], axis=2)
    else:
        gray = img_array.astype(float)

    height, width = gray.shape

    # Count words from OCR
    word_count = sum(len(line['text'].split()) for line in ocr_lines)

    # Very few words -> likely a chart/diagram page
    if word_count < MIN_TEXT_WORDS:
        return 'image'

    # Estimate text line height from OCR
    line_heights = [l['height'] for l in ocr_lines if l.get('height', 0) > 5]
    text_line_h = int(np.median(line_heights)) if line_heights else 30
    min_gfx_height = text_line_h * 2  # graphics must be at least 2x line height

    # Row-level analysis: which rows have content (not empty)?
    row_var = np.var(gray, axis=1)
    row_mean = np.mean(gray, axis=1)
    # A row has content if it has visual variation AND is not pure white
    content_rows = (row_var > 100) & (row_mean < 245)

    # Mask out rows covered by OCR text lines (these are known text, not graphics)
    text_row_mask = np.zeros(height, dtype=bool)
    for line in ocr_lines:
        y_start = line.get('y', 0)
        h_line = line.get('height', 0)
        if h_line < 5:
            continue
        y_end = min(height, y_start + h_line)
        text_row_mask[y_start:y_end] = True

    # Non-text content rows = potential graphics
    gfx_rows = content_rows & ~text_row_mask

    # Find max consecutive non-text content rows
    max_consecutive = 0
    current = 0
    for g in gfx_rows:
        if g:
            current += 1
            max_consecutive = max(max_consecutive, current)
        else:
            current = 0

    has_graphics = max_consecutive > min_gfx_height

    # Decision
    if not has_graphics:
        return 'text'

    text_coverage = sum(l['height'] for l in ocr_lines)
    text_coverage_ratio = text_coverage / height if height > 0 else 0

    if text_coverage_ratio < 0.15:
        return 'image'   # Very little text + graphics = image page
    return 'mixed'       # Both text and graphics


def save_page_image(img_path, output_path, max_width=IMAGE_MAX_WIDTH, quality=IMAGE_QUALITY):
    """Save page as compressed JPEG for markdown reference."""
    img = Image.open(img_path)

    # Convert to RGB
    if img.mode in ('RGBA', 'P'):
        img = img.convert('RGB')

    # Resize if needed
    if max_width and img.width > max_width:
        ratio = max_width / img.width
        new_height = int(img.height * ratio)
        img = img.resize((max_width, new_height), Image.LANCZOS)

    img.save(output_path, format='JPEG', quality=quality, optimize=True)
    return output_path.stat().st_size


def detect_headings(lines):
    """Try to detect headings based on font size (line height) and position."""
    if not lines:
        return lines

    # Calculate average line height
    heights = [l['height'] for l in lines]
    avg_height = np.mean(heights) if heights else 0

    for line in lines:
        # Lines significantly taller than average are likely headings
        if avg_height > 0 and line['height'] > avg_height * 1.4:
            line['is_heading'] = True
        else:
            line['is_heading'] = False

    return lines


# ============================================================
# Markdown Generation
# ============================================================

def find_input_pages(folder):
    """Find page images: PNGs in pages/ subfolder, or extract from PDF.
    Returns (list_of_files, temp_dir_or_None, source_description)."""

    # Priority 1: PNGs in pages/ subfolder (Kindle screenshots)
    pages_dir = folder / "pages"
    if pages_dir.exists():
        files = sorted(pages_dir.glob("page_*.png"))
        if files:
            return files, None, f"pages/ ({len(files)} PNGs)"

    # Priority 2: PNGs in folder root
    files = sorted(folder.glob("page_*.png"))
    if files:
        return files, None, f"root ({len(files)} PNGs)"

    # Priority 3: PDF file
    pdfs = sorted(folder.glob("*.pdf"))
    if pdfs:
        pdf_path = pdfs[0]  # Use first PDF found
        temp_dir = tempfile.mkdtemp(prefix="kindle_md_")
        files = extract_pdf_pages(pdf_path, temp_dir)
        if files:
            return files, temp_dir, f"PDF: {pdf_path.name} ({len(files)} Seiten)"

    return [], None, "keine Quelle gefunden"


def create_markdown(output_folder):
    """Create structured markdown from captured book pages or PDF."""
    folder = Path(output_folder)

    # Find input pages
    files, temp_dir, source_desc = find_input_pages(folder)

    if not files:
        print(f"[FEHLER] Keine Seiten gefunden! ({source_desc})")
        return False

    print(f"[INFO] Quelle: {source_desc}")
    print(f"[INFO] {len(files)} Seiten gefunden")

    # Initialize OCR
    engine, lang_info = check_ocr_languages()
    if not engine:
        print(f"[FEHLER] Windows OCR nicht verfuegbar: {lang_info}")
        if temp_dir:
            shutil.rmtree(temp_dir, ignore_errors=True)
        return False
    print(f"[INFO] OCR Sprache: {lang_info}")

    # Create output directory
    md_dir = folder / "markdown"
    md_dir.mkdir(exist_ok=True)

    book_name = folder.name
    md_path = md_dir / f"{book_name}.md"

    print(f"[INFO] Ausgabe: {md_dir}")
    print()

    # Process pages
    md_lines = []
    md_lines.append(f"# {book_name}\n")
    md_lines.append("")

    stats = {'text': 0, 'image': 0, 'mixed': 0, 'total_words': 0, 'images_saved': 0}

    try:
        for i, file in enumerate(files):
            page_num = i + 1
            print(f"[{page_num}/{len(files)}] {file.name}...", end=" ", flush=True)

            # OCR
            ocr_lines = ocr_image(engine, file)
            ocr_lines = detect_headings(ocr_lines)

            word_count = sum(len(l['text'].split()) for l in ocr_lines)
            stats['total_words'] += word_count

            # Analyze page type via pixel analysis + OCR word count
            page_type = analyze_page(file, ocr_lines)
            stats[page_type] += 1

            # Page separator
            md_lines.append(f"<!-- Seite {page_num} -->")
            md_lines.append("")

            # Save image for image/mixed pages
            if page_type in ('image', 'mixed'):
                img_filename = f"page_{page_num:04d}.jpg"
                img_path = md_dir / img_filename
                img_size = save_page_image(file, img_path)
                stats['images_saved'] += 1

                md_lines.append(f"![Seite {page_num}]({img_filename})")
                md_lines.append("")
                print(f"{page_type} ({word_count} Woerter, Bild: {img_size//1024}KB)", end="")
            else:
                print(f"text ({word_count} Woerter)", end="")

            # Add OCR text
            if ocr_lines:
                for line in ocr_lines:
                    text = line['text'].strip()
                    if not text:
                        continue

                    if line.get('is_heading', False):
                        md_lines.append(f"## {text}")
                    else:
                        md_lines.append(text)
                md_lines.append("")

            print(" OK")

    finally:
        # Clean up temp dir if we extracted from PDF
        if temp_dir:
            shutil.rmtree(temp_dir, ignore_errors=True)

    # Write markdown file
    md_content = "\n".join(md_lines)
    md_path.write_text(md_content, encoding='utf-8')

    # Calculate sizes
    md_size = md_path.stat().st_size
    images_total = sum(f.stat().st_size for f in md_dir.glob("*.jpg"))

    print()
    print("=" * 60)
    print("  MARKDOWN-ERSTELLUNG ABGESCHLOSSEN")
    print("=" * 60)
    print(f"  Seiten: {len(files)}")
    print(f"  Woerter: {stats['total_words']}")
    print(f"  Seitentypen: {stats['text']} Text, {stats['image']} Bild, {stats['mixed']} Gemischt")
    print(f"  Gespeicherte Bilder: {stats['images_saved']}")
    print(f"  Markdown: {md_size / 1024:.0f} KB")
    print(f"  Bilder (JPEG): {images_total / (1024*1024):.1f} MB")
    print(f"  Gesamt: {(md_size + images_total) / (1024*1024):.1f} MB")
    print(f"  Ausgabe: {md_dir}")
    print("=" * 60)

    return True


def main():
    print("=" * 60)
    print("  MARKDOWN-ERSTELLUNG MIT OCR")
    print("=" * 60)
    print()

    output_folder = Path.cwd()
    success = create_markdown(output_folder)

    if success:
        sys.exit(0)
    else:
        sys.exit(1)


if __name__ == "__main__":
    main()
