# MIT License
# Copyright (c) 2025 Quantrosoft
# See LICENSE file for full license text.

"""
Batch PDF -> Markdown Converter
===============================
Converts a whole folder tree of book PDFs into structured Markdown using the
KindleReader method (same output format as create_markdown.py), skipping the
Kindle capture step.

Per PDF the converter decides ONCE how to get at the text:
  TEXT book - the PDF has a real text layer: text is extracted directly per page
              (exact, no OCR). Page classification (text/image/mixed) uses the
              SAME pixel analysis as the existing method
              (create_markdown.analyze_page_array), fed with the text-layer line
              boxes instead of OCR lines - so chart pages (raster AND vector
              drawn) are saved as JPEG and referenced exactly like the existing
              method does. Heading detection uses font size with the same 1.4x
              factor as create_markdown.detect_headings.
  SCAN book - the PDF is a pure image scan (no text layer): pages are rendered
              and run through the EXISTING OCR pipeline imported from
              create_markdown.py (ocr_image subprocess isolation, analyze_page,
              detect_headings, save_page_image) - identical method, identical
              output format. The OCR language is auto-detected per book by
              trial-OCR of sample pages with the German and English engines
              (stopword scoring), passed down via KINDLE_OCR_LANG.

Duplicate files (same MD5) are converted once; every further location gets a
copy of the finished markdown folder.

Output per book (next to the PDF):
  <pdf_dir>/markdown/<pdfname>/<pdfname>.md
  <pdf_dir>/markdown/<pdfname>/page_NNNN.jpg   (chart/figure pages)

Existing non-empty outputs are skipped, so an aborted run can simply be
restarted. Failures are reported loudly per book and summarized at the end;
the exit code is non-zero if any book failed. Source PDFs are never modified.

Usage:
  python batch_pdf2md.py <root_folder> [--ocr-workers N]

Author: Claude
"""

import argparse
import ctypes
import hashlib
import io
import os
import shutil
import sys
import tempfile
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import numpy as np
from PIL import Image

try:
    import fitz  # PyMuPDF
except ImportError:
    print("[FEHLER] PyMuPDF nicht installiert! pip install PyMuPDF")
    sys.exit(1)

# Reuse the existing method: OCR, page analysis, heading detection, image saving.
sys.path.insert(0, str(Path(__file__).resolve().parent))
import create_markdown as cm

# ============================================================
# Configuration
# ============================================================
TEXT_LAYER_MIN_CHARS_PER_PAGE = 200  # Sampled avg below this -> book is a SCAN book
ANALYZE_DPI = 100             # Render DPI for the pixel page analysis of TEXT books
MIXED_IMAGE_COVERAGE = 0.15   # Raster images covering >= this page fraction -> page image saved too
DEFAULT_OCR_WORKERS = 4       # Parallel per-page OCR subprocesses for SCAN books
LANG_SAMPLE_PAGES = 3         # Pages trial-OCRed per language for language detection

# Small stopword sets for OCR language detection (lowercase)
_STOPWORDS = {
    'de-DE': {'der', 'die', 'das', 'und', 'ist', 'nicht', 'mit', 'ein', 'eine',
              'zu', 'den', 'von', 'sich', 'auf', 'werden', 'auch', 'oder'},
    'en-US': {'the', 'and', 'of', 'to', 'in', 'is', 'that', 'for', 'it', 'with',
              'as', 'are', 'this', 'will', 'you', 'not', 'your'},
}


def set_console_title(text):
    """Show the current activity in the console window title."""
    try:
        ctypes.windll.kernel32.SetConsoleTitleW(text)
    except Exception:
        pass


def md5_file(path, chunk=1 << 20):
    h = hashlib.md5()
    with open(path, 'rb') as f:
        while True:
            block = f.read(chunk)
            if not block:
                break
            h.update(block)
    return h.hexdigest()


def out_dir_for(pdf_path):
    return pdf_path.parent / "markdown" / pdf_path.stem


def md_path_for(pdf_path):
    return out_dir_for(pdf_path) / f"{pdf_path.stem}.md"


# ============================================================
# Book classification
# ============================================================

def classify_book(doc):
    """TEXT (has a usable text layer) or SCAN (image-only, needs OCR).
    Samples up to 7 pages spread across the book."""
    n = doc.page_count
    idx = sorted({min(n - 1, max(0, round(f * (n - 1))))
                  for f in (0.05, 0.2, 0.35, 0.5, 0.65, 0.8, 0.95)})
    chars = sum(len(doc[i].get_text()) for i in idx)
    return 'TEXT' if chars / max(1, len(idx)) > TEXT_LAYER_MIN_CHARS_PER_PAGE else 'SCAN'


# ============================================================
# TEXT books: direct text-layer extraction
# ============================================================

def _page_lines(page):
    """Extract text lines with their max font size and line box (PDF points,
    y/height for the pixel analysis), in document order."""
    lines = []
    for block in page.get_text("dict")["blocks"]:
        if block.get("type") != 0:  # 0 = text block
            continue
        for ln in block["lines"]:
            spans = [s for s in ln["spans"] if s["text"].strip()]
            if not spans:
                continue
            x0, y0, x1, y1 = ln["bbox"]
            lines.append({
                'text': " ".join(s["text"].strip() for s in spans),
                'size': max(s["size"] for s in spans),
                'y': y0,
                'height': y1 - y0,
            })
    return lines


def _render_page(page, dpi):
    zoom = dpi / 72.0
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
    return Image.open(io.BytesIO(pix.tobytes("png")))


def _raster_image_coverage(page):
    """Fraction of the page area covered by embedded raster images. Complements
    the pixel analysis: thin light chart lines on white pages fall below its
    row-variance/row-mean thresholds (tuned on dark Kindle screenshots), but an
    embedded raster chart is directly visible in the page's image list."""
    page_area = abs(page.rect)
    if page_area <= 0:
        return 0.0
    covered = 0.0
    try:
        for info in page.get_image_info():
            covered += abs(fitz.Rect(info["bbox"]) & page.rect)
    except Exception:
        return 0.0
    return min(1.0, covered / page_area)


def _save_render_as_jpg(page, out_path):
    """Render the page and save it like create_markdown.save_page_image does."""
    img = _render_page(page, cm.PDF_DPI).convert('RGB')
    if img.width > cm.IMAGE_MAX_WIDTH:
        ratio = cm.IMAGE_MAX_WIDTH / img.width
        img = img.resize((cm.IMAGE_MAX_WIDTH, int(img.height * ratio)), Image.LANCZOS)
    img.save(out_path, format='JPEG', quality=cm.IMAGE_QUALITY, optimize=True)


def convert_text_book(pdf_path, doc, out_dir):
    """Direct text-layer extraction into the standard markdown format.
    Page classification runs the SAME pixel analysis as the existing OCR method
    (analyze_page_array), fed with the text-layer line boxes - this also catches
    vector-drawn charts, which have no embedded raster image to look for."""
    md = [f"# {pdf_path.stem}\n", ""]
    stats = {'text': 0, 'image': 0, 'mixed': 0, 'words': 0, 'images_saved': 0}

    for i, page in enumerate(doc):
        page_num = i + 1
        lines = _page_lines(page)
        md.append(f"<!-- Seite {page_num} -->")
        md.append("")

        # Pixel page analysis on a moderate-DPI render, text lines scaled to it
        zoom = ANALYZE_DPI / 72.0
        arr = np.array(_render_page(page, ANALYZE_DPI))
        scaled = [{'text': l['text'],
                   'y': int(l['y'] * zoom),
                   'height': max(1, int(l['height'] * zoom))} for l in lines]

        gray = arr[:, :, :3].mean(axis=2) if arr.ndim == 3 else arr
        if (gray < 245).mean() < 0.01:
            page_type = 'text'  # blank page: nothing to save
        else:
            page_type = cm.analyze_page_array(arr, scaled)
            # Union with the raster-image signal: embedded raster charts whose
            # thin lines the pixel analysis cannot see still get their page saved.
            if page_type == 'text' and _raster_image_coverage(page) >= MIXED_IMAGE_COVERAGE:
                word_count = sum(len(l['text'].split()) for l in lines)
                page_type = 'mixed' if word_count >= cm.MIN_TEXT_WORDS else 'image'
        stats[page_type] += 1

        if page_type in ('image', 'mixed'):
            img_name = f"page_{page_num:04d}.jpg"
            _save_render_as_jpg(page, out_dir / img_name)
            stats['images_saved'] += 1
            md.append(f"![Seite {page_num}]({img_name})")
            md.append("")

        if lines:
            # Same heading rule as create_markdown.detect_headings: noticeably
            # larger than the page's typical size -> heading.
            median_size = float(np.median([l['size'] for l in lines]))
            for l in lines:
                text = l['text'].strip()
                if not text:
                    continue
                stats['words'] += len(text.split())
                if median_size > 0 and l['size'] > median_size * 1.4:
                    md.append(f"## {text}")
                else:
                    md.append(text)
            md.append("")

        if page_num % 50 == 0 or page_num == doc.page_count:
            print(f"    [{page_num}/{doc.page_count}] Seiten extrahiert", flush=True)

    (out_dir / f"{pdf_path.stem}.md").write_text("\n".join(md), encoding='utf-8')
    return stats


# ============================================================
# SCAN books: existing OCR pipeline from create_markdown.py
# ============================================================

def detect_scan_language(files):
    """Detect the book language by trial-OCRing sample pages with the German and
    English engines and scoring recognized stopwords. Returns a language tag for
    KINDLE_OCR_LANG. Ties (e.g. all-chart samples) keep the German default."""
    n = len(files)
    samples = sorted({files[min(n - 1, round(f * (n - 1)))] for f in (0.25, 0.5, 0.75)})[:LANG_SAMPLE_PAGES]
    scores = {}
    for tag, stopwords in _STOPWORDS.items():
        os.environ['KINDLE_OCR_LANG'] = tag
        words = []
        for f in samples:
            words += [w.strip('.,;:!?()"\'').lower()
                      for l in cm.ocr_image(None, f) for w in l['text'].split()]
        scores[tag] = sum(1 for w in words if w in stopwords)
    os.environ.pop('KINDLE_OCR_LANG', None)
    best = max(scores, key=lambda t: scores[t])
    print(f"    Spracherkennung: {scores} -> {best}", flush=True)
    return best if scores[best] > 0 else 'de-DE'


def convert_scan_book(pdf_path, out_dir, ocr_workers):
    """Render pages and run the EXISTING OCR method (create_markdown.py)."""
    temp_dir = tempfile.mkdtemp(prefix="pdf2md_")
    try:
        files = cm.extract_pdf_pages(pdf_path, temp_dir)
        if not files:
            raise RuntimeError("PDF-Seiten-Extraktion lieferte nichts")

        # Per-book OCR language: detected once, inherited by all OCR subprocesses
        os.environ['KINDLE_OCR_LANG'] = detect_scan_language(files)
        engine, lang_info = cm.check_ocr_languages()
        if not engine:
            raise RuntimeError(f"Windows OCR nicht verfuegbar: {lang_info}")

        # OCR pages in parallel; ocr_image isolates each page in a subprocess,
        # so worker threads only wait on their subprocess (thread-safe).
        print(f"    OCR ({lang_info}, {ocr_workers} parallel) auf {len(files)} Seiten...", flush=True)
        done = 0
        def ocr_one(f):
            nonlocal done
            result = cm.ocr_image(engine, f)
            done += 1
            if done % 10 == 0 or done == len(files):
                print(f"    [OCR {done}/{len(files)}]", flush=True)
            return result
        with ThreadPoolExecutor(max_workers=ocr_workers) as pool:
            ocr_results = list(pool.map(ocr_one, files))

        md = [f"# {pdf_path.stem}\n", ""]
        stats = {'text': 0, 'image': 0, 'mixed': 0, 'words': 0, 'images_saved': 0}

        for i, (file, ocr_lines) in enumerate(zip(files, ocr_results)):
            page_num = i + 1
            ocr_lines = cm.detect_headings(ocr_lines)
            stats['words'] += sum(len(l['text'].split()) for l in ocr_lines)

            page_type = cm.analyze_page(file, ocr_lines)
            stats[page_type] += 1

            md.append(f"<!-- Seite {page_num} -->")
            md.append("")

            if page_type in ('image', 'mixed'):
                img_name = f"page_{page_num:04d}.jpg"
                cm.save_page_image(file, out_dir / img_name)
                stats['images_saved'] += 1
                md.append(f"![Seite {page_num}]({img_name})")
                md.append("")

            for line in ocr_lines:
                text = line['text'].strip()
                if not text:
                    continue
                md.append(f"## {text}" if line.get('is_heading') else text)
            if ocr_lines:
                md.append("")

        (out_dir / f"{pdf_path.stem}.md").write_text("\n".join(md), encoding='utf-8')
        return stats
    finally:
        os.environ.pop('KINDLE_OCR_LANG', None)
        shutil.rmtree(temp_dir, ignore_errors=True)


# ============================================================
# Batch driver
# ============================================================

def copy_book_output(src_dir, dst_pdf):
    """Copy a finished markdown folder to a duplicate PDF's location, renaming
    the .md if the duplicate has a different filename stem."""
    dst_dir = out_dir_for(dst_pdf)
    dst_dir.mkdir(parents=True, exist_ok=True)
    for f in src_dir.iterdir():
        if f.suffix == '.md':
            shutil.copy2(f, dst_dir / f"{dst_pdf.stem}.md")
        else:
            shutil.copy2(f, dst_dir / f.name)


def main():
    ap = argparse.ArgumentParser(description="Batch PDF -> Markdown (KindleReader-Methode)")
    ap.add_argument("root", help="Wurzelordner, wird rekursiv nach *.pdf durchsucht")
    ap.add_argument("--ocr-workers", type=int, default=DEFAULT_OCR_WORKERS,
                    help=f"Parallele OCR-Subprozesse fuer SCAN-Buecher (Default {DEFAULT_OCR_WORKERS})")
    args = ap.parse_args()

    root = Path(args.root)
    if not root.is_dir():
        print(f"[FEHLER] Ordner nicht gefunden: {root}")
        sys.exit(1)

    pdfs = sorted(p for p in root.rglob("*.pdf") if "markdown" not in p.parts)
    print("=" * 60)
    print("  BATCH PDF -> MARKDOWN")
    print("=" * 60)
    print(f"[INFO] Wurzel: {root}")
    print(f"[INFO] {len(pdfs)} PDFs gefunden")
    print()

    seen = {}      # md5 -> Path of first (converted) occurrence
    failures = []  # (Path, error text)
    counts = {'TEXT': 0, 'SCAN': 0, 'COPY': 0, 'SKIP': 0}
    t0 = time.time()

    for k, pdf in enumerate(pdfs, 1):
        rel = pdf.relative_to(root)
        prefix = f"[{k}/{len(pdfs)}]"
        set_console_title(f"pdf2md {prefix} {pdf.stem}")

        try:
            if pdf.stat().st_size == 0:
                raise RuntimeError("Datei ist 0 Bytes (leer/defekt)")

            md_path = md_path_for(pdf)
            if md_path.exists() and md_path.stat().st_size > 1024:
                counts['SKIP'] += 1
                print(f"{prefix} SKIP  {rel} (Markdown existiert)", flush=True)
                continue

            digest = md5_file(pdf)
            if digest in seen:
                copy_book_output(out_dir_for(seen[digest]), pdf)
                counts['COPY'] += 1
                print(f"{prefix} COPY  {rel}  (Duplikat von {seen[digest].relative_to(root)})", flush=True)
                continue

            doc = fitz.open(pdf)
            try:
                kind = classify_book(doc)
                print(f"{prefix} {kind}  {rel}  ({doc.page_count} Seiten)", flush=True)
                out_dir = out_dir_for(pdf)
                out_dir.mkdir(parents=True, exist_ok=True)
                if kind == 'TEXT':
                    stats = convert_text_book(pdf, doc, out_dir)
                else:
                    doc.close()
                    doc = None
                    stats = convert_scan_book(pdf, out_dir, args.ocr_workers)
                counts[kind] += 1
                seen[digest] = pdf
                print(f"    OK: {stats['words']} Woerter, "
                      f"{stats['text']}T/{stats['image']}B/{stats['mixed']}G, "
                      f"{stats['images_saved']} Bilder", flush=True)
            finally:
                if doc is not None:
                    doc.close()

        except Exception as e:
            failures.append((rel, str(e)))
            print(f"{prefix} [FEHLER] {rel}: {e}", flush=True)

    print()
    print("=" * 60)
    print("  BATCH ABGESCHLOSSEN" if not failures else "  BATCH MIT FEHLERN ABGESCHLOSSEN")
    print("=" * 60)
    print(f"  Konvertiert: {counts['TEXT']} Text-Buecher, {counts['SCAN']} Scan-Buecher (OCR)")
    print(f"  Duplikate kopiert: {counts['COPY']}   Uebersprungen: {counts['SKIP']}")
    print(f"  Dauer: {(time.time() - t0) / 60:.1f} min")
    if failures:
        print(f"  FEHLER: {len(failures)}")
        for rel, err in failures:
            print(f"    - {rel}: {err}")
    print("=" * 60)
    sys.exit(1 if failures else 0)


if __name__ == "__main__":
    main()
