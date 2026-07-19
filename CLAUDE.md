# KindleReader Setup & Workflow

## Quick Start

Sag einfach `kindle` und ich führe dich durch den kompletten Workflow mit interaktiven Fragen:

**Schritt 1: Vorbereitung überprüfen**
Ich frage: Ist Kindle offen und das Buch vollständig geladen?
- Wähle eine Option aus den Buttons aus

**Schritt 2: Zielordner angeben**
Ich frage: Wo soll das verarbeitete Buch hin?
- Wähle eine Option aus den Buttons aus oder gib einen eigenen Pfad ein (z.B. `C:\Books\MeinBuchname`)

**Schritt 3: Scan starten**
Zielordner erstellen (falls nötig) und `scan.bat` in externer Console starten.

## Technische Ausführung

WICHTIG: Genau diese Befehle verwenden!

1. **Zielordner erstellen** (falls er nicht existiert):
   ```
   powershell -Command "New-Item -ItemType Directory -Force -Path '<ZIELORDNER>'"
   ```

2. **scan.bat starten** - Buch-Ordner als **Argument** übergeben:
   ```
   powershell -Command "Start-Process -FilePath '<PFAD_ZU_DIESEM_ORDNER>\scan.bat' -ArgumentList '<ZIELORDNER>'"
   ```
   - `<ZIELORDNER>` = der Buch-Ordner (dort landen `pages\`, PDF, `markdown\`)
   - `<PFAD_ZU_DIESEM_ORDNER>` = absoluter Pfad zu diesem KindleReader-Ordner
   - scan.bat wechselt selbst in den übergebenen Ordner (Ordnername = Buchname). **Ohne** Argument (z.B. User-Doppelklick) fragt scan.bat den Buch-Ordner per **Ordner-Dialog** ab (Start: `_Kindle`).
   - WICHTIG: NUR `powershell Start-Process` verwenden! `start` und `cmd /k` funktionieren NICHT aus Claude Code heraus!

3. **Nach dem Start**: Dem User mitteilen, dass die Console offen ist und der Scan läuft.

---

## Workflow Details

Die `scan.bat` führt folgende Schritte aus:

1. **Kindle Capture** (`kindle_capture.py`) - Erfasst Screenshots des Kindle-Buches in `pages/`
2. **PDF-Erstellung** (`create_pdf.py`) - Erstellt durchsuchbares PDF als `[BOOKNAME].pdf`
3. **Markdown-Generierung** (`create_markdown.py`) - Erzeugt Markdown in `markdown/<Ordnername>.md`

Bereits vorhandene Outputs werden automatisch übersprungen.

---

## Voraussetzungen

- Python 3.x installiert
- **Neues WinUI-Kindle für PC** (ohne Menüleiste, Hotkey-Steuerung). Die Navigation läuft über Hotkeys (+ **ein** Fokus-Klick, um dem Reader den Tastaturfokus zu geben) → **sprachunabhängig** (kein deutsches Menü mehr nötig).
- Kindle-App offen mit geladenem Buch (Fenstermodus)
- Bildschirm darf während des Laufs **nicht sperren**. Das Tool hält die Session per `SetThreadExecutionState` wach (Screensaver/Display-Timeout), eine per GPO/Policy erzwungene Sperre kann es aber nicht verhindern — dann bricht das Blättern ab.

Hinweis: `create_pdf.py`/`create_markdown.py` nutzen weiterhin Windows-OCR (`winsdk`) — aber auf den **erfassten Seitenbildern** (Sprache automatisch de/en), nicht auf Menüs.

---

## Code-Regeln (UNBEDINGT BEACHTEN!)

1. **KEIN FALLBACK, KEIN WORKAROUND, KEINE ALTERNATIVEN!** Wenn etwas fehlschlägt → sofortiger Abbruch mit `sys.exit(1)` und klarer Fehlermeldung. Niemals mit geschätzten/festen Positionen weiterarbeiten. KEINE "Plan B"-Logik, KEINE Retry-Schleifen, KEINE alternativen Wege zum Ziel. Entweder der direkte Weg funktioniert oder das Script bricht ab.

2. **KEINE hardcodierten Pixel-Positionen!** Navigation läuft über **Hotkeys** (F11 Vollbild, PageUp/PageDown blättern) — nicht über Menü-/Pfeil-Klicks. Maus-Einsatz nur: **ein** Fokus-Klick in die Fenster-/Bildschirmmitte (gibt dem Reader den Tastaturfokus) + Parken des Cursors — beides aus der Fenster-/Bildschirmgröße berechnet, nichts hardcodiert. Erfassung per `PrintWindow` (fensterbezogen), nicht per absolute Screen-Region.

3. **Alle Python-Abhängigkeiten sind REQUIRED.** Die `scan.bat` installiert automatisch aus `requirements.txt`. Imports wie `winsdk`, `pywinauto` etc. dürfen NICHT optional sein - bei Fehlen → `sys.exit(1)`.

4. **Keine Experimente!** Vor Änderungen am bestehenden Code: Git-Version prüfen. Funktionierende Logik nicht durch ungetestete Alternativen ersetzen.

5. **Ablauf in kindle_capture.py (Hotkey-basiert, neues WinUI-Kindle):**
   1. Kindle-Fenster finden + aktivieren
   2. **Einmal** in die Bildschirmmitte klicken → gibt dem WinUI-Reader den Tastaturfokus (**nötig für F11 UND die Seitentasten** — F11 ist NICHT App-weit!), dann **F11** → Vollbild. Vollbild setzt auf eine saubere Seite zurück; die Toolbar-Chrome des Klicks wird **nicht** mit-übernommen (nur ein kurzer „Drücke F11"-Hinweis, der ausfadet)
   3. Warten bis Bildschirm stabil (Hinweis ausgefadet)
   4. Zum Cover: **PageUp** bis sich die Seite nicht mehr ändert. Ab hier **nur noch Tasten** + Maus in neutraler Mitte geparkt (kein weiterer Klick → keine Chrome). **KEIN Ctrl+G** (dessen Dialog stiehlt den Fokus und killt die Seitentasten)
   5. Jede Vollbild-Seite per **`PrintWindow(PW_RENDERFULLCONTENT)`** erfassen (funktioniert auch im geschützten/exklusiven Vollbild, wo GDI-Screengrab schwarz liefert), mit **PageDown** vorwärts bis Buchende

**Warum PrintWindow statt Screenshot:** Kindles Vollbild kann in einen exklusiven/geschützten Modus gehen, in dem `PIL.ImageGrab` (GDI) schwarz/Fehler liefert. `PrintWindow` liest das Eigen-Rendering des Fensters (WinUI + WebView2) und ist davon unabhängig. Braucht `pywin32`.
