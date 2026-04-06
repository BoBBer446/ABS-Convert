#requires -version 5.1
[CmdletBinding()]
param(
    [string]$SourceRoot = "E:\JDownloader\Neuer Ordner",
    [string]$TargetRoot = "E:\JDownloader\Neuer Ordner\ABS",
    [string]$FixedAuthor = "",
    [string]$FfmpegPath = "ffmpeg",
    [bool]$DryRun = $false,
    [bool]$WriteTags = $true,
    [bool]$CoverOverwrite = $false,
    [double]$MaxUnknownTrackRatio = 0.70,
    [ValidateSet("index","bookindex","source")]
    [string]$TrackTitleMode = "bookindex",
    [ValidateSet("folder","title")]
    [string]$AlbumMode = "title",
    [string]$Genre = "Audiobook"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# =========================
# INTERN / STATE
# =========================

$script:ProcessedBooks = 0
$script:Skipped = New-Object System.Collections.Generic.List[object]

$AudioExt = @(".mp3", ".m4b", ".m4a", ".aac", ".flac", ".ogg", ".opus", ".wma", ".wav")
$AudioExtSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($e in $AudioExt) {
    [void]$AudioExtSet.Add($e)
}

$IgnorePathParts = @(
    "\booklet\", "\booklets\",
    "\ebook\", "\ebooks\",
    "\das buch\", "\dasbuch\",
    "\files\",
    "\uploaded\",
    "\support\",
    "\unterstützung\",
    "\unterstuetzung\",
    "\_extras\",
    "\extras\"
)

$IgnorePathPartsLower = @($IgnorePathParts | ForEach-Object { $_.ToLowerInvariant() })

$RegexOptions = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor `
                [System.Text.RegularExpressions.RegexOptions]::Compiled

$DiscRegexes = @(
    New-Object System.Text.RegularExpressions.Regex -ArgumentList '[\\\/]\s*cd\s*0*([0-9]{1,3})\s*[\\\/]', $RegexOptions
    New-Object System.Text.RegularExpressions.Regex -ArgumentList '[\\\/]\s*disc\s*0*([0-9]{1,3})\s*[\\\/]', $RegexOptions
    New-Object System.Text.RegularExpressions.Regex -ArgumentList '[\\\/]\s*disk\s*0*([0-9]{1,3})\s*[\\\/]', $RegexOptions
    New-Object System.Text.RegularExpressions.Regex -ArgumentList '[\\\/]\s*teil\s*0*([0-9]{1,3})\s*[\\\/]', $RegexOptions
    New-Object System.Text.RegularExpressions.Regex -ArgumentList '[\\\/]\s*part\s*0*([0-9]{1,3})\s*[\\\/]', $RegexOptions
)

$TrackRegexes = @(
    New-Object System.Text.RegularExpressions.Regex -ArgumentList '^\s*0*([0-9]{1,6})\s*[-_. ]', $RegexOptions
    New-Object System.Text.RegularExpressions.Regex -ArgumentList '^\s*\[\s*0*([0-9]{1,6})\s*\]', $RegexOptions
    New-Object System.Text.RegularExpressions.Regex -ArgumentList '^\s*\(\s*0*([0-9]{1,6})\s*\)', $RegexOptions
    New-Object System.Text.RegularExpressions.Regex -ArgumentList '\btrack\s*(?:no\.?\s*)?0*([0-9]{1,6})\b', $RegexOptions
    New-Object System.Text.RegularExpressions.Regex -ArgumentList '\bkapitel\s*0*([0-9]{1,6})\b', $RegexOptions
    New-Object System.Text.RegularExpressions.Regex -ArgumentList '\bchapter\s*0*([0-9]{1,6})\b', $RegexOptions
    New-Object System.Text.RegularExpressions.Regex -ArgumentList '[-_. ]0*([0-9]{1,6})\s*\.[^.]+$', $RegexOptions
    New-Object System.Text.RegularExpressions.Regex -ArgumentList '^\s*teil\s*0*([0-9]{1,6})\D', $RegexOptions
)

$TrackAlphaRegex = New-Object System.Text.RegularExpressions.Regex -ArgumentList '^\s*0*([0-9]{1,3})([a-z])([0-9]+)?\b', $RegexOptions
$DiscFolderNameRegex = New-Object System.Text.RegularExpressions.Regex -ArgumentList '^(?i)\s*(cd|disc|disk|teil|part)\s*\d{1,3}\s*$', $RegexOptions

# =========================
# FUNKTIONEN
# =========================

function Ensure-Dir {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        if ($DryRun) {
            Write-Host "  [DryRun] mkdir: $Path"
        }
        else {
            New-Item -ItemType Directory -Path $Path -Force | Out-Null
        }
    }
}

function Test-FfmpegAvailable {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Executable
    )

    try {
        & $Executable -version *> $null
        return $true
    }
    catch {
        return $false
    }
}

function Remove-InvalidFileNameChars {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
    $result = $Name

    foreach ($char in $invalidChars) {
        $escaped = [Regex]::Escape([string]$char)
        $result = $result -replace $escaped, "_"
    }

    $result = $result -replace '\s+', ' '
    $result = $result.Trim()
    $result = $result.Trim('.', ' ')

    if ([string]::IsNullOrWhiteSpace($result)) {
        return "Unbekannt"
    }

    return $result
}

function Get-NaturalSortKey {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    return [regex]::Replace($Text, '\d+', {
        param($match)
        $match.Value.PadLeft(20, '0')
    })
}

function Should-IgnorePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FullName
    )

    $lower = $FullName.ToLowerInvariant()

    foreach ($part in $IgnorePathPartsLower) {
        if ($lower.IndexOf($part, [System.StringComparison]::Ordinal) -ge 0) {
            return $true
        }
    }

    return $false
}

function Get-NumberFromRegexes {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,

        [Parameter(Mandatory = $true)]
        [System.Text.RegularExpressions.Regex[]]$Regexes
    )

    foreach ($rx in $Regexes) {
        $m = $rx.Match($Text)
        if ($m.Success) {
            return [int]$m.Groups[1].Value
        }
    }

    return $null
}

function Get-DiscNumberFromPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FullPath
    )

    $n = Get-NumberFromRegexes -Text $FullPath -Regexes $DiscRegexes
    if ($null -ne $n) {
        return $n
    }

    return 0
}

function Get-TrackNumber {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileName
    )

    $n = Get-NumberFromRegexes -Text $FileName -Regexes $TrackRegexes
    if ($null -ne $n) {
        return $n
    }

    return 999999
}

function Get-TrackSortKey {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileName
    )

    $track = Get-TrackNumber -FileName $FileName
    if ($track -eq 999999) {
        return 999999000
    }

    $m = $TrackAlphaRegex.Match($FileName)
    if ($m.Success) {
        $t = [int]$m.Groups[1].Value
        $letter = $m.Groups[2].Value.ToLowerInvariant()
        $letterIndex = ([int][char]$letter) - ([int][char]'a') + 1
        if ($letterIndex -lt 1) {
            $letterIndex = 50
        }

        $sub = 0
        if ($m.Groups[3].Success) {
            $sub = [int]$m.Groups[3].Value
        }

        return ($t * 1000) + ($letterIndex * 10) + $sub
    }

    return ($track * 1000)
}

function Get-BookTitleOnly {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BookFolderName
    )

    $name = $BookFolderName.Trim()

    if ($name -match '^\d{4}\s*-\s*(.+)$') {
        return $Matches[1].Trim()
    }

    return $name
}

function Get-CleanBookTitle {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Author,

        [Parameter(Mandatory = $true)]
        [string]$FolderName
    )

    $bookTitle = $FolderName.Trim()

    if (-not [string]::IsNullOrWhiteSpace($Author)) {
        if ($bookTitle -like "$Author -*") {
            $bookTitle = $bookTitle.Substring($Author.Length).Trim()
            $bookTitle = $bookTitle -replace '^\-\s*', ''
        }
    }

    $bookTitle = Remove-InvalidFileNameChars -Name $bookTitle
    return $bookTitle
}

function Get-RelativePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BasePath,

        [Parameter(Mandatory = $true)]
        [string]$FullPath
    )

    $base = [System.IO.Path]::GetFullPath($BasePath)
    $full = [System.IO.Path]::GetFullPath($FullPath)

    if (-not $base.EndsWith([System.IO.Path]::DirectorySeparatorChar)) {
        $base = $base + [System.IO.Path]::DirectorySeparatorChar
    }

    $baseUri = [System.Uri]$base
    $fullUri = [System.Uri]$full
    $relativeUri = $baseUri.MakeRelativeUri($fullUri)
    $relative = [System.Uri]::UnescapeDataString($relativeUri.ToString())

    return ($relative -replace '/', '\\')
}

function Get-AuthorNameFromBookPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BookPath,

        [Parameter(Mandatory = $true)]
        [string]$SourceBase
    )

    if (-not [string]::IsNullOrWhiteSpace($FixedAuthor)) {
        return (Remove-InvalidFileNameChars -Name $FixedAuthor)
    }

    $relative = Get-RelativePath -BasePath $SourceBase -FullPath $BookPath
    $parts = @($relative -split '[\\\/]') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    if ($parts.Count -ge 2) {
        return (Remove-InvalidFileNameChars -Name $parts[0])
    }

    return "Unbekannt"
}

function Get-BookFolderForAudioFile {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File
    )

    $parent = $File.Directory
    if ($null -eq $parent) {
        return $null
    }

    if ($DiscFolderNameRegex.IsMatch($parent.Name)) {
        if ($null -ne $parent.Parent) {
            return $parent.Parent.FullName
        }
    }

    return $parent.FullName
}

function Test-BookOrderConfidence {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Files
    )

    $filesArr = @($Files)
    $total = $filesArr.Length

    if ($total -eq 0) {
        return [pscustomobject]@{
            Ok = $false
            Reason = "Keine Audio-Dateien"
        }
    }

    if ($total -eq 1) {
        return [pscustomobject]@{
            Ok = $true
            Reason = "OK (single file)"
        }
    }

    $unknown = 0
    foreach ($f in $filesArr) {
        $trk = Get-TrackNumber -FileName $f.Name
        if ($trk -eq 999999) {
            $unknown++
        }
    }

    $unknownRatio = $unknown / $total
    if ($unknownRatio -gt $MaxUnknownTrackRatio) {
        return [pscustomobject]@{
            Ok = $false
            Reason = ("Zu viele Dateien ohne Tracknummer: {0:P0} (Limit {1:P0})" -f $unknownRatio, $MaxUnknownTrackRatio)
        }
    }

    return [pscustomobject]@{
        Ok = $true
        Reason = "OK"
    }
}

function Get-CoverFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BookSourcePath
    )

    $nameCandidates = @("cover","folder","front","artwork","album","poster")
    $extCandidates = @(".jpg",".jpeg",".png",".webp")

    $searchDirs = @($BookSourcePath)

    $coverDir1 = Join-Path $BookSourcePath "cover"
    $coverDir2 = Join-Path $BookSourcePath "covers"

    if (Test-Path -LiteralPath $coverDir1) { $searchDirs += $coverDir1 }
    if (Test-Path -LiteralPath $coverDir2) { $searchDirs += $coverDir2 }

    foreach ($dir in $searchDirs) {
        foreach ($base in $nameCandidates) {
            foreach ($ext in $extCandidates) {
                $candidate = Join-Path $dir ($base + $ext)
                if (Test-Path -LiteralPath $candidate) {
                    return (Get-Item -LiteralPath $candidate)
                }
            }
        }
    }

    $fallback = Get-ChildItem -LiteralPath $BookSourcePath -File -ErrorAction SilentlyContinue | Where-Object {
        @(".jpg",".jpeg",".png",".webp") -contains $_.Extension.ToLowerInvariant()
    } | Select-Object -First 1

    return $fallback
}

function Copy-BookCover {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BookSourcePath,

        [Parameter(Mandatory = $true)]
        [string]$BookTargetPath
    )

    $coverSource = Get-CoverFile -BookSourcePath $BookSourcePath
    if ($null -eq $coverSource) {
        return
    }

    $srcExt = $coverSource.Extension.ToLowerInvariant()
    $destName = "cover$srcExt"
    $coverDest = Join-Path $BookTargetPath $destName

    if ((Test-Path -LiteralPath $coverDest) -and (-not $CoverOverwrite)) {
        return
    }

    if ($DryRun) {
        Write-Host ("  [DryRun] COVER  {0}  ->  {1}" -f $coverSource.FullName, $coverDest)
        return
    }

    Copy-Item -LiteralPath $coverSource.FullName -Destination $coverDest -Force
    Write-Host ("  COVER  {0}  ->  {1}" -f $coverSource.FullName, $coverDest)
}

function Build-FfmpegArguments {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath,

        [Parameter(Mandatory = $true)]
        [string]$AuthorValue,

        [Parameter(Mandatory = $true)]
        [string]$AlbumValue,

        [Parameter(Mandatory = $true)]
        [string]$TitleValue,

        [Parameter(Mandatory = $true)]
        [int]$TrackNo,

        [Parameter(Mandatory = $true)]
        [int]$TrackCount,

        [Parameter(Mandatory = $true)]
        [int]$DiscNo,

        [Parameter(Mandatory = $true)]
        [int]$DiscCount
    )

    $args = @(
        "-hide_banner",
        "-loglevel", "error",
        "-nostdin",
        "-y",
        "-i", $InputPath,
        "-map", "0",
        "-map_metadata", "-1",
        "-c", "copy"
    )

    if ($WriteTags) {
        if (-not [string]::IsNullOrWhiteSpace($AuthorValue)) {
            $args += @("-metadata", "artist=$AuthorValue")
            $args += @("-metadata", "album_artist=$AuthorValue")
        }

        if (-not [string]::IsNullOrWhiteSpace($AlbumValue)) {
            $args += @("-metadata", "album=$AlbumValue")
        }

        if (-not [string]::IsNullOrWhiteSpace($TitleValue)) {
            $args += @("-metadata", "title=$TitleValue")
        }

        if (-not [string]::IsNullOrWhiteSpace($Genre)) {
            $args += @("-metadata", "genre=$Genre")
        }

        if ($TrackNo -gt 0 -and $TrackCount -gt 0) {
            $args += @("-metadata", "track=$TrackNo/$TrackCount")
        }

        if ($DiscNo -gt 0 -and $DiscCount -gt 0) {
            $args += @("-metadata", "disc=$DiscNo/$DiscCount")
        }

        $args += @("-id3v2_version", "3")
    }

    $args += $OutputPath
    return $args
}

function Invoke-FfmpegCopy {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath,

        [Parameter(Mandatory = $true)]
        [string]$AuthorValue,

        [Parameter(Mandatory = $true)]
        [string]$AlbumValue,

        [Parameter(Mandatory = $true)]
        [string]$TitleValue,

        [Parameter(Mandatory = $true)]
        [int]$TrackNo,

        [Parameter(Mandatory = $true)]
        [int]$TrackCount,

        [Parameter(Mandatory = $true)]
        [int]$DiscNo,

        [Parameter(Mandatory = $true)]
        [int]$DiscCount
    )

    $ffmpegArgs = Build-FfmpegArguments `
        -InputPath $InputPath `
        -OutputPath $OutputPath `
        -AuthorValue $AuthorValue `
        -AlbumValue $AlbumValue `
        -TitleValue $TitleValue `
        -TrackNo $TrackNo `
        -TrackCount $TrackCount `
        -DiscNo $DiscNo `
        -DiscCount $DiscCount

    $ffmpegOutput = & $FfmpegPath @ffmpegArgs 2>&1
    $ffmpegExitCode = $LASTEXITCODE

    if ($ffmpegExitCode -ne 0) {
        $message = "ffmpeg Fehler bei Datei: $InputPath"

        if (Test-Path -LiteralPath $OutputPath) {
            try {
                Remove-Item -LiteralPath $OutputPath -Force -ErrorAction SilentlyContinue
            }
            catch {
            }
        }

        if ($ffmpegOutput) {
            $detail = ($ffmpegOutput | ForEach-Object { $_.ToString() }) -join [Environment]::NewLine
            throw "$message`n$detail"
        }

        throw $message
    }
}

function Export-Book {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BookPath,

        [Parameter(Mandatory = $true)]
        [object[]]$FilesForBook
    )

    $bookName = Split-Path -Path $BookPath -Leaf
    Write-Host ""
    Write-Host "=== Export: $bookName ==="

    if ($FilesForBook.Length -eq 0) {
        Write-Host "  SKIP: Keine Audio-Dateien"
        $script:Skipped.Add([pscustomobject]@{
            Book = $bookName
            Reason = "Keine Audio-Dateien"
            Path = $BookPath
        })
        return
    }

    $check = Test-BookOrderConfidence -Files $FilesForBook
    if (-not $check.Ok) {
        Write-Host "  SKIP: $($check.Reason)"
        $script:Skipped.Add([pscustomobject]@{
            Book = $bookName
            Reason = $check.Reason
            Path = $BookPath
        })
        return
    }

    $authorValue = Get-AuthorNameFromBookPath -BookPath $BookPath -SourceBase $SourceRoot
    $bookTitleOnly = Get-BookTitleOnly -BookFolderName $bookName
    $bookTitle = Get-CleanBookTitle -Author $authorValue -FolderName $bookName

    $albumValue = if ($AlbumMode -eq "title") { $bookTitleOnly } else { $bookName }

    $targetAuthorFolder = Join-Path $TargetRoot $authorValue
    Ensure-Dir -Path $targetAuthorFolder

    $bookTargetPath = Join-Path $targetAuthorFolder $bookTitle
    Ensure-Dir -Path $bookTargetPath

    Copy-BookCover -BookSourcePath $BookPath -BookTargetPath $bookTargetPath

    $fileMeta = New-Object System.Collections.Generic.List[object]
    $maxDisc = 0

    foreach ($f in $FilesForBook) {
        $discRaw = Get-DiscNumberFromPath -FullPath $f.FullName
        if ($discRaw -gt $maxDisc) {
            $maxDisc = $discRaw
        }

        $sortKey = Get-TrackSortKey -FileName $f.Name

        [void]$fileMeta.Add([pscustomobject]@{
            File = $f
            DiscRaw = $discRaw
            SortKey = $sortKey
            Name = $f.Name
            FullName = $f.FullName
            Ext = $f.Extension.ToLowerInvariant()
        })
    }

    $discCount = [int]$maxDisc
    if ($discCount -le 0) {
        $discCount = 1
    }

    $sorted = @(
        $fileMeta | ForEach-Object {
            $discNorm = $_.DiscRaw
            if ($maxDisc -gt 0 -and $discNorm -eq 0) {
                $discNorm = 1
            }

            [pscustomobject]@{
                File = $_.File
                Disc = [int]$discNorm
                SortKey = $_.SortKey
                Name = $_.Name
                FullName = $_.FullName
                Ext = $_.Ext
            }
        } | Sort-Object Disc, SortKey, Name
    )

    $orderLog = Join-Path $bookTargetPath "_ORDER.txt"
    if ($DryRun) {
        Write-Host "  [DryRun] würde Order-Log schreiben: $orderLog"
    }
    else {
        $lines = $sorted | ForEach-Object {
            "Disc={0:D3} Key={1:D6} | {2}" -f $_.Disc, $_.SortKey, $_.FullName
        }
        Set-Content -LiteralPath $orderLog -Value $lines -Encoding UTF8
    }

    $trackCount = $sorted.Length
    $pad = [Math]::Max(3, $trackCount.ToString().Length)

    $i = 1
    foreach ($m in $sorted) {
        try {
            $newName = ("{0:D$pad} - {1}{2}" -f $i, $bookTitle, $m.Ext)
            $newName = Remove-InvalidFileNameChars -Name $newName
            $dest = Join-Path $bookTargetPath $newName

            if (Test-Path -LiteralPath $dest) {
                Write-Host "  SKIP: Zielkollision: $newName"
                $script:Skipped.Add([pscustomobject]@{
                    Book = $bookName
                    Reason = "Zielkollision: $newName existiert"
                    Path = $bookTargetPath
                })
                return
            }

            Write-Host ("  COPY  {0}  ->  {1}" -f $m.FullName, $dest)

            if ($TrackTitleMode -eq "index") {
                $titleValue = ("Track {0:D$pad}" -f $i)
            }
            elseif ($TrackTitleMode -eq "source") {
                $titleValue = [System.IO.Path]::GetFileNameWithoutExtension($m.Name)
            }
            else {
                $titleValue = ("{0} - Track {1:D$pad}" -f $bookTitleOnly, $i)
            }

            if ($DryRun) {
                Write-Host ("  [DryRun] ffmpeg tags -> Artist={0} Album={1} Title={2} Track={3}/{4} Disc={5}/{6}" -f `
                    $authorValue, $albumValue, $titleValue, $i, $trackCount, $m.Disc, $discCount)
            }
            else {
                Invoke-FfmpegCopy `
                    -InputPath $m.FullName `
                    -OutputPath $dest `
                    -AuthorValue $authorValue `
                    -AlbumValue $albumValue `
                    -TitleValue $titleValue `
                    -TrackNo $i `
                    -TrackCount $trackCount `
                    -DiscNo $m.Disc `
                    -DiscCount $discCount
            }

            $i++
        }
        catch {
            Write-Host "  ERROR: Ausnahme bei Datei" -ForegroundColor Red
            Write-Host ("         Quelle: {0}" -f $m.FullName) -ForegroundColor Red
            Write-Host ("         Fehler: {0}" -f $_.Exception.Message) -ForegroundColor Red

            if (Test-Path -LiteralPath $dest) {
                try {
                    Remove-Item -LiteralPath $dest -Force -ErrorAction SilentlyContinue
                }
                catch {
                }
            }

            $script:Skipped.Add([pscustomobject]@{
                Book = $bookName
                Reason = "Fehler bei Datei: $($m.FullName)"
                Path = $BookPath
            })

            return
        }
    }

    $script:ProcessedBooks++
    Write-Host "  Fertig: $trackCount Datei(en) -> $bookTargetPath"
}

# =========================
# START
# =========================

if (-not (Test-Path -LiteralPath $SourceRoot)) {
    throw "SourceRoot existiert nicht: $SourceRoot"
}

if (-not (Test-FfmpegAvailable -Executable $FfmpegPath)) {
    throw "ffmpeg wurde nicht gefunden. Prüfe PATH oder setze -FfmpegPath explizit."
}

Ensure-Dir -Path $TargetRoot

Write-Host ("Scan: {0}" -f $SourceRoot)

$allFiles = Get-ChildItem -LiteralPath $SourceRoot -Recurse -File -ErrorAction SilentlyContinue

$audioFilesAll = New-Object System.Collections.Generic.List[object]
$bookFolderSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

foreach ($f in $allFiles) {
    if (-not $AudioExtSet.Contains($f.Extension)) {
        continue
    }

    if (Should-IgnorePath -FullName $f.FullName) {
        continue
    }

    [void]$audioFilesAll.Add($f)

    $bookPath = Get-BookFolderForAudioFile -File $f
    if (-not [string]::IsNullOrWhiteSpace($bookPath)) {
        [void]$bookFolderSet.Add($bookPath)
    }
}

$bookPaths = @($bookFolderSet) | Sort-Object

Write-Host ("Gefundene Buchordner: {0}" -f $bookPaths.Length)
Write-Host ("Gefundene Audiofiles: {0}" -f $audioFilesAll.Count)

$bookToFiles = @{}
foreach ($bp in $bookPaths) {
    $bookToFiles[$bp] = New-Object System.Collections.Generic.List[object]
}

foreach ($f in $audioFilesAll) {
    $bp = Get-BookFolderForAudioFile -File $f
    if ($null -ne $bp -and $bookToFiles.ContainsKey($bp)) {
        [void]$bookToFiles[$bp].Add($f)
    }
}

foreach ($bookPath in $bookPaths) {
    $filesForBook = @($bookToFiles[$bookPath])
    Export-Book -BookPath $bookPath -FilesForBook $filesForBook
}

Write-Host ""
Write-Host "==================== SUMMARY ===================="
Write-Host ("Processed: {0}" -f $script:ProcessedBooks)
Write-Host ("Skipped:   {0}" -f $script:Skipped.Count)

if ($script:Skipped.Count -gt 0) {
    Write-Host ""
    Write-Host "--- Skipped Books ---"

    foreach ($s in $script:Skipped) {
        Write-Host ("- {0} | {1}" -f $s.Book, $s.Reason)
    }

    $skipFile = Join-Path $TargetRoot "_SKIPPED.txt"

    if ($DryRun) {
        Write-Host "  [DryRun] würde Skip-Liste schreiben: $skipFile"
    }
    else {
        $script:Skipped | ForEach-Object {
            "{0} | {1} | {2}" -f $_.Book, $_.Reason, $_.Path
        } | Set-Content -LiteralPath $skipFile -Encoding UTF8
    }
}

Write-Host "DONE."
