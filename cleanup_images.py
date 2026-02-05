"""
Cleanup: Remove text-only images from markdown folders
=======================================================
Scans all JPG images in markdown/ folders and removes those that
contain only text (no charts, diagrams, or visual content).
Also removes the corresponding ![...] references from the .md file.

Two-step approach:
1. Remove unreferenced images (not in .md file)
2. Analyze remaining images via pixel analysis and remove text-only ones

Usage:
  python cleanup_images.py <path>

  <path> can be:
  - A single book folder (containing markdown/)
  - A parent folder (scans all subfolders for markdown/)

Author: Claude
"""

import sys
import re
from pathlib import Path
from PIL import Image
import numpy as np

# ============================================================
# Configuration
# ============================================================
MIN_TEXT_WORDS = 10         # Minimum words to consider "has text"

# ============================================================
# Page Analysis (pixel-based)
# ============================================================

def is_text_only_image(img_path):
    """Analyze if an image is text-only (no charts/diagrams).
    Graphics must span at least 2x text line height (~60px) in
    consecutive rows with actual content (not empty space).
    Without OCR, uses row variance pattern to distinguish text
    (short bursts of ~1 line) from graphics (tall contiguous blocks).
    Returns True if the image should be removed."""
    try:
        img = Image.open(img_path)
        img_array = np.array(img)

        if len(img_array.shape) == 3:
            gray = np.mean(img_array[:, :, :3], axis=2)
        else:
            gray = img_array.astype(float)

        height, width = gray.shape

        # Minimum graphic height: ~2x typical text line height
        min_gfx_height = 60

        # Row-level analysis
        row_var = np.var(gray, axis=1)
        row_mean = np.mean(gray, axis=1)
        # A row has content if it has visual variation AND is not pure white
        content_rows = (row_var > 100) & (row_mean < 245)

        # Find max consecutive content rows
        # Text: short bursts (~20px per line), graphics: tall blocks (>60px)
        max_consecutive = 0
        current = 0
        for c in content_rows:
            if c:
                current += 1
                max_consecutive = max(max_consecutive, current)
            else:
                current = 0

        if max_consecutive > min_gfx_height:
            return False  # Has a graphic region -> keep

        return True  # Text-only -> remove

    except Exception as e:
        print(f"    [WARNUNG] Analyse fehlgeschlagen: {e}")
        return False


def cleanup_markdown_folder(md_dir):
    """Remove text-only images from a markdown folder and update .md file.
    Returns (total_images, removed_count)."""
    jpgs = sorted(md_dir.glob("page_*.jpg"))
    if not jpgs:
        return 0, 0

    # Find the .md file
    md_files = list(md_dir.glob("*.md"))
    if not md_files:
        return 0, 0
    md_file = md_files[0]

    md_content = md_file.read_text(encoding='utf-8')

    # Step 1: Find which images are referenced in the .md file
    referenced = set(re.findall(r'!\[[^\]]*\]\(([^)]+\.jpg)\)', md_content))

    to_remove = []

    for jpg in jpgs:
        filename = jpg.name
        if filename not in referenced:
            # Unreferenced image -> remove silently
            to_remove.append(jpg)
        else:
            # Referenced but text-only -> remove with reference cleanup
            if is_text_only_image(jpg):
                to_remove.append(jpg)

    if not to_remove:
        return len(jpgs), 0

    # Remove images and update .md content
    for jpg in to_remove:
        filename = jpg.name
        # Remove the ![...](...) line referencing this image (if exists)
        pattern = rf'!\[[^\]]*\]\({re.escape(filename)}\)\n?'
        md_content = re.sub(pattern, '', md_content)
        # Delete the file
        jpg.unlink()

    # Clean up double blank lines
    md_content = re.sub(r'\n{3,}', '\n\n', md_content)

    md_file.write_text(md_content, encoding='utf-8')

    return len(jpgs), len(to_remove)


def main():
    if len(sys.argv) < 2:
        print("Usage: python cleanup_images.py <path>")
        print("  <path> = book folder or parent folder")
        sys.exit(1)

    target = Path(sys.argv[1])
    if not target.exists():
        print(f"[FEHLER] Pfad nicht gefunden: {target}")
        sys.exit(1)

    print("=" * 60)
    print("  CLEANUP: TEXT-ONLY BILDER ENTFERNEN")
    print("=" * 60)
    print()

    # Find all markdown/ folders
    md_dirs = []
    if (target / "markdown").exists():
        md_dirs.append(target / "markdown")
    else:
        for d in sorted(target.rglob("markdown")):
            if d.is_dir():
                md_dirs.append(d)

    if not md_dirs:
        print("[INFO] Keine markdown/ Ordner gefunden.")
        sys.exit(0)

    print(f"[INFO] {len(md_dirs)} Buecher gefunden")
    print()

    total_images = 0
    total_removed = 0

    for md_dir in md_dirs:
        book_name = md_dir.parent.name
        images_count, removed_count = cleanup_markdown_folder(md_dir)
        total_images += images_count
        total_removed += removed_count

        if images_count > 0:
            kept = images_count - removed_count
            print(f"  {book_name}: {images_count} Bilder, {removed_count} entfernt, {kept} behalten")

    print()
    print("=" * 60)
    print(f"  Gesamt: {total_images} Bilder, {total_removed} entfernt")
    print(f"  Behalten: {total_images - total_removed}")
    print("=" * 60)


if __name__ == "__main__":
    main()
