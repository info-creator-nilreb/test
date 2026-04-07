# Word VBA Beispielordner

Lege dein Beispiel-Dokument (`.docx`) in den Unterordner `beispiel-dokument/`.

## Enthalten

- `HomematicFormatierung.bas`: VBA-Modul mit dem Makro `FormatHomematicDokumentationUndSpeichern`.
- `beispiel-dokument/`: Ablageordner für dein echtes `.docx`.

## Verwendung in Word

1. `Alt + F11` drücken, um den VBA-Editor zu öffnen.
2. `Datei -> Datei importieren...` und `HomematicFormatierung.bas` auswählen.
3. Das zu formatierende `.docx` aus `beispiel-dokument/` in Word öffnen.
4. Makro `FormatHomematicDokumentationUndSpeichern` ausführen.
5. Ergebnis: Das Dokument wird formatiert und als neue Datei mit Suffix `_formatiert.docx` im selben Ordner gespeichert.

## Hinweis

Wenn das Dokument noch nie gespeichert wurde (kein Dateipfad), bricht das Makro mit einem Hinweis ab. Speichere das Dokument in dem Fall zuerst als `.docx`.
