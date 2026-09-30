# KindleReader Setup & Workflow

## Quick Start

Sag einfach `kindle` und ich führe dich durch den kompletten Workflow mit interaktiven Fragen:

**Step 1: Check preparation**
I ask: Is Kindle running and the book open in the reader (see "Voraussetzungen")?
- Pick an option from the buttons

**Step 2: Start the scan**
Start `scan.bat` WITHOUT argument in an external console. It reads the open book's
title itself, creates the book folder and continues - no target folder needs to be
asked for.

## Technical execution

IMPORTANT: use exactly these commands!

1. **Start scan.bat** - normally WITHOUT argument:
   ```
   powershell -Command "Start-Process -FilePath '<PATH_TO_THIS_FOLDER>\scan.bat'"
   ```
   - `<PATH_TO_THIS_FOLDER>` = absolute path of this KindleReader folder
   - Without argument (also on double-click) scan.bat runs `kindle_book_title.py`: it
     reads the title of the book currently open in Kindle from **Kindle's own log**
     (`%LOCALAPPDATA%\Packages\AMZNKindle.AmazonKindleReadingApp_*\LocalState\logs\kindle.log`),
     creates `KINDLE_ROOT\<title>` (`KINDLE_ROOT` is set at the top of scan.bat,
     currently `D:\GoogleDriveData\ShareFile\eBooks\Trading\_Kindle`) and changes into
     it. Folder name = title with `:` turned into ` - ` and other characters Windows
     forbids replaced. An existing folder is reused (existing outputs are skipped).
     Log format, checks and failure cases: docstring of `kindle_book_title.py`.
   - With argument `-ArgumentList '<BOOK_FOLDER>'` that **existing** book folder is
     used instead (e.g. to redo PDF/markdown for a book captured earlier).
   - IMPORTANT: ONLY use `powershell Start-Process`! `start` and `cmd /k` do NOT work from Claude Code!

2. **After the start**: tell the user the console is open and the scan is running.

---

## Workflow Details

`scan.bat` runs these steps:

0. **Book folder** (`kindle_book_title.py`, only without argument) - title of the open book from Kindle's log -> `KINDLE_ROOT\<title>`
1. **Kindle Capture** (`kindle_capture.py`) - captures screenshots of the Kindle book into `pages/`
2. **PDF creation** (`create_pdf.py`) - creates a searchable PDF `<folder name>.pdf`
3. **Markdown generation** (`create_markdown.py`) - creates markdown `markdown/<folder name>.md`

Outputs that already exist are skipped automatically.

## Batch: vorhandene PDFs → Markdown (ohne Kindle)

`pdf2md.bat <Wurzelordner>` (Treiber: `batch_pdf2md.py`) konvertiert rekursiv **alle**
`*.pdf` unter dem Wurzelordner in Markdown — gleiche Methode/gleiches Format wie
`create_markdown.py`, nur ohne den Kindle-Erfassungsschritt:

- **TEXT-Bücher** (PDF hat Text-Layer): Text wird pro Seite direkt extrahiert (exakt, kein
  OCR). Seiten-Klassifikation (Text/Bild/Gemischt) über **dieselbe Pixel-Analyse** wie die
  bestehende Methode (`analyze_page_array`, gefüttert mit den Text-Layer-Zeilenboxen)
  **vereinigt** mit einem Raster-Bild-Abdeckungs-Signal (dünne helle Chart-Linien auf weißen
  Seiten liegen unter den Pixel-Schwellen, eingebettete Raster-Charts sieht die Bildliste
  direkt). Headings über Fontgröße (gleicher 1.4×-Faktor wie `detect_headings`).
- **SCAN-Bücher** (reines Bild-PDF): Seiten werden gerendert und durch die **bestehende
  OCR-Strecke** aus `create_markdown.py` geschickt (Subprozess-Isolation, `analyze_page`,
  `detect_headings`, `save_page_image`). OCR-Sprache wird **pro Buch automatisch erkannt**
  (Probeseiten mit de/en-Engine, Stoppwort-Score) und via Env-Variable `KINDLE_OCR_LANG`
  an die OCR-Subprozesse vererbt (`check_ocr_languages` liest sie; Default bleibt de zuerst).
- **Duplikate** (gleiche MD5) werden einmal konvertiert, weitere Fundorte bekommen eine
  Kopie des fertigen markdown-Ordners. Vorhandene nicht-leere .md werden übersprungen
  (abgebrochener Lauf einfach neu starten). Fehler werden pro Buch laut gemeldet und am
  Ende zusammengefasst (Exit-Code ≠ 0). Quell-PDFs werden nie verändert.
- Ausgabe je Buch: `<pdf_ordner>\markdown\<pdfname>\<pdfname>.md` + `page_NNNN.jpg`.
- Log per tee nach `<Wurzelordner>\pdf2md.log`; Konsolen-Titel zeigt laufend `[k/N] Buch`.

---

## Voraussetzungen (WICHTIG — vor dem Start prüfen!)

Vor jedem Lauf müssen diese Punkte stimmen, sonst wird das Ergebnis falsch:

- **Kindle muss LAUFEN.** Das Tool startet Kindle **NICHT** selbst — läuft kein Kindle-Fenster, bricht es mit klarer Fehlermeldung ab. (Neues WinUI-Kindle für PC, ohne Menüleiste; Steuerung sprachunabhängig über Hotkeys + ein Fokus-Klick.)
- **Buch muss geladen sein** — im Reader geöffnet (nicht in der Bibliothek).
- **The book must be open in the CURRENT Kindle session**: the title is taken from the
  log lines of the session of the running Kindle.exe. The open marker is Kindle's
  `[Renderer] Open book ...` line - written both when the user opens a book and when
  Kindle restores the open book after a restart (`openItem` is written only in the
  first case; measured 29.09.2026, see the docstring of `kindle_book_title.py`). If the
  session has no Renderer open line, or the book was closed again after it,
  `kindle_book_title.py` aborts with a message -> open the book from the library, then
  start again.
- **Leseposition auf den ersten paar Seiten oder auf der Titelseite.** Das Tool blättert per PageUp zum Cover zurück — von weit hinten dauert das unnötig lange.
- **Kindle-Einstellungen setzen** (oben rechts **Aa** → *Seiteneinstellungen*):
  - **Layout: „Einzelne Spalte"** — sonst zeigt Kindle im Vollbild **zwei** Buchseiten nebeneinander (= 2 Seiten pro Bild).
  - **Rand: Schieberegler ganz nach RECHTS** (max) — macht die Textspalte schmal/porträt, damit die Seiten dasselbe Format wie die Titelseite haben.
  - **Ausrichtung: „Links".**
  - **Abstand: „Mittel".**
- **Während der Erfassung (~5 Min) Maus/Tastatur NICHT anfassen** — jeder Tastendruck stoppt den Lauf, Mausbewegung kann die Toolbar einblenden (landet sonst mit im Bild).
- Bildschirm darf **nicht sperren**. Das Tool hält die Session per `SetThreadExecutionState` wach (Screensaver/Display-Timeout), eine per GPO/Policy erzwungene Sperre kann es aber nicht verhindern — dann bricht das Blättern ab.
- Python 3.x installiert (Abhängigkeiten installiert scan.bat automatisch).

Hinweis: `create_pdf.py`/`create_markdown.py` nutzen Windows-OCR (`winsdk`) auf den **erfassten Seitenbildern** (Sprache automatisch de/en).

---

## Code-Regeln (UNBEDINGT BEACHTEN!)

1. **KEIN FALLBACK, KEIN WORKAROUND, KEINE ALTERNATIVEN!** Wenn etwas fehlschlägt → sofortiger Abbruch mit `sys.exit(1)` und klarer Fehlermeldung. Niemals mit geschätzten/festen Positionen weiterarbeiten. KEINE "Plan B"-Logik, KEINE Retry-Schleifen, KEINE alternativen Wege zum Ziel. Entweder der direkte Weg funktioniert oder das Script bricht ab.

2. **KEINE hardcodierten Pixel-Positionen!** Navigation läuft über **Hotkeys** (F11 Vollbild, PageUp/PageDown blättern) — nicht über Menü-/Pfeil-Klicks. Maus-Einsatz nur: **ein** Fokus-Klick in den **linken schwarzen Rand** (~15% der Breite — NICHT die Mitte! Die kann bei einer Link-Tabelle einen Hyperlink treffen und öffnet dann den Browser) + Parken des Cursors — alles aus der Fenstergröße berechnet, nichts hardcodiert. Erfassung per `PrintWindow` (fensterbezogen), nicht per absolute Screen-Region.

3. **Alle Python-Abhängigkeiten sind REQUIRED.** Die `scan.bat` installiert automatisch aus `requirements.txt`. Imports wie `winsdk`, `pywinauto` etc. dürfen NICHT optional sein - bei Fehlen → `sys.exit(1)`.

4. **Keine Experimente!** Vor Änderungen am bestehenden Code: Git-Version prüfen. Funktionierende Logik nicht durch ungetestete Alternativen ersetzen.

   **Language (user order 29.09.2026):** the tool's user interface - every console
   message the user sees (`[INFO]`, `[FEHLER]`, banners, prompts) - stays **German**.
   Code, identifiers, comments, docstrings and documentation are **English**. Messages
   use the ASCII transliteration already in place (ae/oe/ue); `.bat` files must stay
   pure ASCII.

5. **Ablauf in kindle_capture.py (Hotkey-basiert, neues WinUI-Kindle):**
   1. Kindle-Fenster finden + aktivieren
   2. **Einmal** in den **linken schwarzen Rand** (~15% der Breite) klicken → gibt dem WinUI-Reader den Tastaturfokus (**nötig für F11 UND die Seitentasten** — F11 ist NICHT App-weit!). **NICHT die Mitte** — die kann bei Link-Tabellen einen Hyperlink treffen und öffnet den Browser. Dann **F11** → Vollbild (setzt auf eine saubere Seite zurück; die Toolbar-Chrome des Klicks wird nicht mit-übernommen)
   3. **Wait until the screen is still** (`wait_until_still`): 4 consecutive grabs PIXEL-IDENTICAL, no threshold; max 20 s, then fail-loud with `_debug_unruhe_1.png`/`_2.png`. Used after F11, on the title page and after the re-focus click. ⚠️ The former mean-difference check ("stable" = mean |diff| ≤ 1.0) let a late change through: measured 29.09.2026, right after "stabil", 4 grabs of the still page differed by `[0, 0, 22073]` px.
   4. **Seitentasten-Beweis (Fokus-Check):** PageDown (oder am Buchende PageUp) muss **sichtbar** blättern. Reagiert der Reader auf **keine** der beiden Tasten → **fail-loud** mit Beweis-Screenshot `_debug_fokus.png` (Ursachen: kein Tastaturfokus oder PrintWindow liefert eingefrorene Frames). ⚠️ Ohne diesen Beweis ist „PageUp ändert nichts" mehrdeutig — im Feld wurde „3× PageUp ohne Wirkung" fälschlich als „Cover erreicht" gemeldet, während Kindle unberührt auf Seite 3 stand.
   5. Zum Cover: **PageUp** bis sich die Seite nicht mehr ändert. Ab hier **nur noch Tasten** + Maus in neutraler **Fenster**-Mitte geparkt (kein weiterer Klick → keine Chrome). **KEIN Ctrl+G** (dessen Dialog stiehlt den Fokus und killt die Seitentasten). **Multi-Monitor:** Vollbild-Erkennung und Maus-Parken rechnen gegen den Monitor, **auf dem das Kindle-Fenster liegt** (`MonitorFromWindow`), nicht gegen den Primärmonitor (`pyautogui.size()` kennt nur den Primären!).
   6. **Determine the page format from the title page** (`detect_page_region_from_cover`). **Definition: the page is what lies BETWEEN the uniform letterbox bars left and right** (black in dark mode, white in light mode). **Only the LEFT and RIGHT edge are searched; vertically the page is the full window height** - in Kindle's fullscreen the text runs from the very top to the very bottom of the screen (user decision 29.09.2026). The page itself may contain the letterbox colour (e.g. a black bar printed on a white title page). If detection fails → **fail-loud** with evidence screenshot `_debug_titelseite.png` in the book folder. NOT via variance, NOT via "longest run with >50% coverage":
      - **Reference rows:** rows whose far-left/far-right edge strips (10 px) are letterbox. A title bar spans the full width, so its rows drop out - chrome cannot make letterbox columns look like content.
      - **Columns:** over the reference rows a letterbox column is ≤2% non-letterbox (`LETTERBOX_MAX`). `_page_span` walks inward from both edges (past a stray 1px edge column - measured: x=1919 white on the fullscreen title page - across the bar) to the first non-letterbox column. Everything between is page, interior letterbox-coloured stretches included.
      - **Plausibility check (never moves the edges):** a letterboxed title page differs from the letterbox colour over >50% of its area; a plain text page does not (then left/right would only be the text block). Measured: title page 97.5%, text pages 7/8 1.0% → fail-loud.
      - ⚠️ **Why not "longest run with >50% coverage"** (the former method): measured 29.09.2026 (Cybernetic Analysis ..., fullscreen) it returned top=157 instead of 0 - rows 141-156 (a black bar on the cover) had only 44% non-black pixels, split the page into two runs, the longer one (157-1078) won. Every text page lost its top lines; page 7 (text only at the top) was captured pure white. Offline on the user's fullscreen title-page screenshot: old (102,157,912,1078), new (102,0,912,1078); on the run's own uncropped title grab: (554,0,1364,1080).
      - ⚠️ **Why not variance:** chrome has variance too - the variance scan returned `1443x834` (landscape, whole window) instead of `516x804` (cover).
      - **Alle** Seiten werden auf dieses Format zugeschnitten (nicht der ganze Bildschirm), damit jede Seite dasselbe Format wie die Titelseite hat. Ändert sich die Fenstergröße während des Laufs, bricht die Erfassung **fail-loud** ab (der feste Zuschnitt wäre sonst falsch).
   7. Jede Seite per **`PrintWindow(PW_RENDERFULLCONTENT)`** erfassen (funktioniert auch im geschützten/exklusiven Vollbild, wo GDI-Screengrab schwarz liefert), **auf das Titelseiten-Format gecroppt**, mit **PageDown** vorwärts bis Buchende.
      **End of book = the content no longer changes when paging forward - nothing else (user order 29.09.2026).** "Changed" is decided by an EXACT pixel comparison with **NO threshold** (`images_are_similar`: identical = no change, any differing pixel = change); the same comparison is used for the focus check and for paging back to the cover. Precondition, established by `wait_until_still` (step 3): grabs of the same still page are pixel-identical - otherwise "no longer changes" could never occur, and `wait_until_still` fails loud after 20 s with the pixel counts and evidence images. ⚠️ The former rule "changed < 0.6% = no turn" ended the capture after 7 of 326 pages: pages 7→8 (short front matter) changed 3095 px = 0.354% of the crop (measured 29.09.2026).
   8. **Evidence files, always written (uncropped window grabs) into the book folder:**
      `_debug_titelseite_voll.png` (the grab the crop region is derived from),
      and at the end-of-book decision `_debug_letzte_seite_voll.png` (last saved page)
      plus `_debug_buchende_voll.png` (window at that moment). The end decision compares
      CROPPED images only, so content outside the crop is invisible to it - a wrong crop
      shows up here first (see the 29.09.2026 case under step 6).

**Warum PrintWindow statt Screenshot:** Kindles Vollbild kann in einen exklusiven/geschützten Modus gehen, in dem `PIL.ImageGrab` (GDI) schwarz/Fehler liefert. `PrintWindow` liest das Eigen-Rendering des Fensters (WinUI + WebView2) und ist davon unabhängig. Braucht `pywin32`.
