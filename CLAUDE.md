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

2. **scan.bat starten** - MUSS in einer externen Console laufen:
   ```
   powershell -Command "Start-Process -FilePath '<PFAD_ZU_DIESEM_ORDNER>\scan.bat' -WorkingDirectory '<ZIELORDNER>'"
   ```
   - `<ZIELORDNER>` = der vom User gewählte Pfad (Schritt 2)
   - `<PFAD_ZU_DIESEM_ORDNER>` = absoluter Pfad zu diesem KindleReader-Ordner
   - Die scan.bat MUSS im Zielordner als Working Directory laufen (nutzt Ordnername als Buchname)
   - WICHTIG: NUR `powershell Start-Process` verwenden! `start` und `cmd /k` funktionieren NICHT aus Claude Code heraus!

3. **Nach dem Start**: Dem User mitteilen, dass die Console offen ist und der Scan läuft.

---

## Workflow Details

Die `scan.bat` führt folgende Schritte aus:

1. **Kindle Capture** (`kindle_capture.py`) - Erfasst Screenshots des Kindle-Buches in `pages/`
2. **PDF-Erstellung** (`create_pdf.py`) - Erstellt durchsuchbares PDF als `[BOOKNAME].pdf`
3. **Markdown-Generierung** (`create_markdown.py`) - Erzeugt Markdown in `markdown/book.md`

Bereits vorhandene Outputs werden automatisch übersprungen.

---

## Voraussetzungen

- Python 3.x installiert
- Kindle-App offen mit geladenem Buch

---

## Code-Regeln (UNBEDINGT BEACHTEN!)

1. **KEIN FALLBACK, KEIN WORKAROUND, KEINE ALTERNATIVEN!** Wenn etwas fehlschlägt → sofortiger Abbruch mit `sys.exit(1)` und klarer Fehlermeldung. Niemals mit geschätzten/festen Positionen weiterarbeiten. KEINE "Plan B"-Logik, KEINE Retry-Schleifen, KEINE alternativen Wege zum Ziel. Entweder der direkte Weg funktioniert oder das Script bricht ab.

2. **KEINE absoluten Koordinaten!** Alle UI-Elemente (Menüs, Buttons, Pfeile) MÜSSEN dynamisch per OCR oder Bilderkennung gefunden werden. Was auf einem Bildschirm bei Pixel X sitzt, sitzt auf einem anderen woanders.

3. **Alle Python-Abhängigkeiten sind REQUIRED.** Die `scan.bat` installiert automatisch aus `requirements.txt`. Imports wie `winsdk`, `pywinauto` etc. dürfen NICHT optional sein - bei Fehlen → `sys.exit(1)`.

4. **Keine Experimente!** Vor Änderungen am bestehenden Code: Git-Version prüfen. Funktionierende Logik nicht durch ungetestete Alternativen ersetzen.

5. **Ablauf in kindle_capture.py:**
   1. Kindle-Fenster finden (KEIN Maximize)
   2. Zur Titelseite navigieren (Gehe zu → Titelseite) — Menüleiste ist noch sichtbar
   3. Vollbildmodus aktivieren (Ansicht → Vollbildmodus)
   4. Warten bis Vollbild-Hinweis verschwindet
   5. Buchbereich erkennen
   6. Navigationspfeile erkennen
