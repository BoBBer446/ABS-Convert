# ABS-Convert

PowerShell-Skript zum Konvertieren und Vereinheitlichen von Hörbuch-Ordnern für Audiobookshelf (ABS).

## Funktionen

- Rekursive Suche nach Audio-Dateien (`.mp3`, `.m4b`, `.m4a`, `.aac`, `.flac`, `.ogg`, `.opus`, `.wma`, `.wav`)
- Erkennung von Buch-Ordnern inkl. Disc/CD-Unterordnern (`CD1`, `Disc 2`, `Teil 3`, ...)
- Natürliche Sortierung über Disc- und Tracknummern
- Metadaten-Schreiben via `ffmpeg` (Artist, Album, Title, Genre, Track, Disc)
- Cover-Übernahme (`cover`, `folder`, `front`, ...)
- Skip-Logik für unsichere Track-Reihenfolgen
- Dry-Run-Modus für sichere Vorschau

## Voraussetzungen

- PowerShell 5.1 oder neuer
- `ffmpeg` im PATH oder als expliziter Pfad

## Verwendung

```powershell
# Standardlauf
.\Convert-ToABS.ps1

# Dry Run
.\Convert-ToABS.ps1 -DryRun $true

# Mit festem Autor und eigener Quelle/Ziel
.\Convert-ToABS.ps1 `
  -SourceRoot "E:\JDownloader\Neuer Ordner" `
  -TargetRoot "E:\JDownloader\Neuer Ordner\ABS" `
  -FixedAuthor "Max Mustermann"
```

## Wichtige Parameter

- `-SourceRoot`: Quellpfad mit Hörbuchordnern
- `-TargetRoot`: Zielpfad für strukturierte Ausgabe
- `-FixedAuthor`: Erzwingt einen Autor statt Pfad-Ableitung
- `-FfmpegPath`: ffmpeg-Binary/Pfad
- `-DryRun`: simuliert Operationen
- `-WriteTags`: schreibt Metadaten
- `-CoverOverwrite`: überschreibt bestehende Cover
- `-MaxUnknownTrackRatio`: Schwellwert für fehlende Tracknummern
- `-TrackTitleMode`: `index`, `bookindex`, `source`
- `-AlbumMode`: `folder` oder `title`
- `-Genre`: Genre-Tag (Standard: `Audiobook`)

## Ausgabe

Pro verarbeitetem Buch:
- Zielordner `TargetRoot\Autor\Buchtitel`
- `_ORDER.txt` mit der finalen Sortierreihenfolge
- optional `cover.*`

Globale Ausgabe:
- `_SKIPPED.txt` mit übersprungenen Büchern und Gründen

## Testen

Syntaxcheck:

```powershell
pwsh -NoLogo -NoProfile -Command "[void][System.Management.Automation.Language.Parser]::ParseFile('Convert-ToABS.ps1',[ref]`$null,[ref]`$null)"
```

Trockentest mit Dummy-Struktur:

```powershell
.\Convert-ToABS.ps1 -SourceRoot "<TEST_SRC>" -TargetRoot "<TEST_DST>" -DryRun $true
```
