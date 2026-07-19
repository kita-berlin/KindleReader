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

## Voraussetzungen (WICHTIG — vor dem Start prüfen!)

Vor jedem Lauf müssen diese Punkte stimmen, sonst wird das Ergebnis falsch:

- **Kindle muss LAUFEN.** Das Tool startet Kindle **NICHT** selbst — läuft kein Kindle-Fenster, bricht es mit klarer Fehlermeldung ab. (Neues WinUI-Kindle für PC, ohne Menüleiste; Steuerung sprachunabhängig über Hotkeys + ein Fokus-Klick.)
- **Buch muss geladen sein** — im Reader geöffnet (nicht in der Bibliothek).
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

5. **Ablauf in kindle_capture.py (Hotkey-basiert, neues WinUI-Kindle):**
   1. Kindle-Fenster finden + aktivieren
   2. **Einmal** in den **linken schwarzen Rand** (~15% der Breite) klicken → gibt dem WinUI-Reader den Tastaturfokus (**nötig für F11 UND die Seitentasten** — F11 ist NICHT App-weit!). **NICHT die Mitte** — die kann bei Link-Tabellen einen Hyperlink treffen und öffnet den Browser. Dann **F11** → Vollbild (setzt auf eine saubere Seite zurück; die Toolbar-Chrome des Klicks wird nicht mit-übernommen)
   3. Warten bis Bildschirm stabil (Hinweis ausgefadet)
   4. Zum Cover: **PageUp** bis sich die Seite nicht mehr ändert. Ab hier **nur noch Tasten** + Maus in neutraler Mitte geparkt (kein weiterer Klick → keine Chrome). **KEIN Ctrl+G** (dessen Dialog stiehlt den Fokus und killt die Seitentasten)
   5. **Seitenformat von der Titelseite bestimmen:** Die letterboxte Titelseite hebt sich mit komplett schwarzem/weißem Rand vom Hintergrund ab; per Spalten-/Zeilen-**Varianz** wird ihr Rechteck erkannt (`detect_page_region_from_cover`). **Alle** Seiten werden auf dieses Format zugeschnitten (nicht der ganze Bildschirm), damit jede Seite dasselbe Format wie die Titelseite hat.
   6. Jede Seite per **`PrintWindow(PW_RENDERFULLCONTENT)`** erfassen (funktioniert auch im geschützten/exklusiven Vollbild, wo GDI-Screengrab schwarz liefert), **auf das Titelseiten-Format gecroppt**, mit **PageDown** vorwärts bis Buchende

**Warum PrintWindow statt Screenshot:** Kindles Vollbild kann in einen exklusiven/geschützten Modus gehen, in dem `PIL.ImageGrab` (GDI) schwarz/Fehler liefert. `PrintWindow` liest das Eigen-Rendering des Fensters (WinUI + WebView2) und ist davon unabhängig. Braucht `pywin32`.
