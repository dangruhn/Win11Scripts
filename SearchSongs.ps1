param (
    [string]$InputFile = "C:\Users\Dan.Gruhn\SongList.txt",
    [switch]$Debug
)
$musicDirs  = ("C:\Users\Dan.Gruhn\Dropbox\dgruhn-home\Music\MP3",
               "C:\Users\Dan.Gruhn\Dropbox\dgruhn-home\Music\Archives\New")
$logFile   = "C:\Users\Dan.Gruhn\logs\SongMatchLog.txt"
$logMatches = @()
$logNoMatches = @()

# Timestamp header
"--- Song Match Log ---" | Out-File $logFile
"Run Time: $(Get-Date)" | Out-File $logFile -Append
"" | Out-File $logFile -Append

# Verify that the input song list file exists
if (-not (Test-Path $InputFile)) {
    Write-Error "Input file not found: $InputFile"
    exit 1
}

# Load all MP3 files recursively with full paths
$files = Get-ChildItem -Path $musicDirs -Filter *.mp3 -File -Recurse
if ($Debug) { Write-Host "Loaded $($files.Count) MP3 files from $musicDirs" }

# Helper: Build loose regex pattern from song title
function Build-LoosePattern($songName) {
    $clean = $songName -replace '[^\w\s]', ''   # Remove punctuation
    $tokens = $clean -split '\s+'               # Split into words
    return ($tokens -join '.*')                 # Join with wildcards
}

# Parse each line from the input file
Get-Content $InputFile | ForEach-Object {
    if ($Debug) { Write-Host "`nProcessing line: $_" }

    if ($_ -match '"([^"]+)"\t(.+)') {
        $song = $matches[1].Trim()
        $author = $matches[2].Trim()
        if ($Debug) { Write-Host "Extracted Song: '$song' | Author: '$author'" }

        # Build loose pattern
        $pattern = Build-LoosePattern $song
        if ($Debug) { Write-Host "Loose Regex Pattern: $pattern" }

        # Match against filename portion only, but return full path
        $matchesFound = $files | Where-Object {
            # Normalize filename before matching
            $filename = ($_ | Split-Path -Leaf) -replace '[^\w\s]', '' -replace '\s+', ' '
            $filename -match "(?i)$pattern"
        } | Select-Object -ExpandProperty FullName

        # If we found a match
        if ($matchesFound.Count -gt 0) {
            if ($Debug) {
                Write-Host "Matches found:"
                $matchesFound | ForEach-Object { Write-Host "  $_" }
            }
            # With this:
            $logMatches += [PSCustomObject]@{
                Song    = $song
                Author  = $author
                Matches = $matchesFound.Count -gt 0 ? $matchesFound : @("No match")
            }
        } else {
            if ($Debug) { Write-Host "No match found for '$song'" }
            $logNoMatches += [PSCustomObject]@{
                Song    = $song
                Author  = $author
                Matches = @("No match")
            }
        }
    } else {
        if ($Debug) { Write-Host "Line did not match expected format: $_" }
    }
}
# Final summary to console
Write-Host "`n--- Matches Summary ---"
foreach ($entry in $logMatches) {
    Write-Host "`n$($entry.Song) - $($entry.Author)"
    $entry.Matches | ForEach-Object { Write-Host "  $_" }
}
Write-Host "`n--- No Matches Summary ---"
foreach ($entry in $logNoMatches) {
    Write-Host "`n$($entry.Song) - $($entry.Author)"
    $entry.Matches | ForEach-Object { Write-Host "  $_" }
}

# Write matches
$sortedLog = $logMatches | Sort-Object Song
foreach ($entry in $sortedLog) {
    "`n$($entry.Song) - $($entry.Author)" | Out-File $logFile -Append
    $entry.Matches | ForEach-Object { "  $_" | Out-File $logFile -Append }
}

# Write no matches
$sortedLog = $logNoMatches | Sort-Object Song
foreach ($entry in $sortedLog) {
    "`n$($entry.Song) - $($entry.Author)" | Out-File $logFile -Append
    $entry.Matches | ForEach-Object { "  $_" | Out-File $logFile -Append }
}