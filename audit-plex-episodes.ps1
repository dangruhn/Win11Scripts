param (
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [ArgumentCompleter({
        param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
        Get-ChildItem -Path "M:\Video\TV Shows" -Directory | ForEach-Object { $_.FullName }
    })]
    [string]$RootPath = "M:\Video\TV Shows",

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$LogPath = "$env:USERPROFILE\logs\media_audit.log",

    [Parameter(Mandatory = $false)]
    [switch]$DryRun,

    [Parameter(Mandatory = $false)]
    [switch]$VerboseOutput,

    [Parameter(Mandatory = $false)]
    [switch]$UseGui
)

# Load Windows Forms for GUI elements
Add-Type -AssemblyName System.Windows.Forms

# Logging function
function Write-Log {
    param (
        [string]$Message,
        [switch]$Force
    )
    if ($VerboseOutput -or $Force) {
        Write-Host $Message
    }
    Add-Content -Path $LogPath -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Message"
}

# GUI folder picker fallback
function Select-Folder {
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.SelectedPath = $RootPath
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.SelectedPath
    }
    return $null
}

# Detect missing seasons
function Get-MissingSeasons {
    param ([string]$ShowPath)

    $seasonFolders = Get-ChildItem -Path $ShowPath -Directory |
        Where-Object { $_.Name -match '^Season\s+(\d{1,2})$' }

    $existingSeasons = $seasonFolders.Name |
        ForEach-Object {
            if ($_ -match '^Season\s+(\d{1,2})$') {
                [int]$matches[1]
            }
        }

    if ($existingSeasons.Count -eq 0) { return @() }

    $min = ($existingSeasons | Measure-Object -Minimum).Minimum
    $max = ($existingSeasons | Measure-Object -Maximum).Maximum
    $expected = $min..$max
    $missing = $expected | Where-Object { $_ -notin $existingSeasons }

    return $missing
}

# Detect missing episodes and last episode
function Get-EpisodeStats {
    param ([string]$SeasonPath)

    $episodeFiles = Get-ChildItem -Path $SeasonPath -File |
        Where-Object { $_.Name -match '(?i)(?:Episode\s*|Ep\s*|E)(\d{1,3})' }

    $episodeNumbers = $episodeFiles.Name |
        ForEach-Object {
            if ($_ -match '(?i)(?:Episode\s*|Ep\s*|E)(\d{1,3})') {
                [int]$matches[1]
            }
        }

    if ($episodeNumbers.Count -eq 0) { return @{ Missing = @(); Last = $null } }

    $min = ($episodeNumbers | Measure-Object -Minimum).Minimum
    $max = ($episodeNumbers | Measure-Object -Maximum).Maximum
    $expected = $min..$max
    $missing = $expected | Where-Object { $_ -notin $episodeNumbers }

    return @{ Missing = $missing; Last = $max }
}

# Validate RootPath
if ($UseGui -or -not (Test-Path $RootPath)) {
    Write-Log "Launching folder picker..." -Force
    $selected = Select-Folder
    if (-not $selected) {
        Write-Log "No folder selected. Exiting script." -Force
        exit 1
    }
    $RootPath = $selected
    Write-Log "User selected folder: $RootPath" -Force
}

Write-Log "Starting media audit for: $RootPath" -Force
if ($DryRun) {
    Write-Log "Dry-run mode enabled — no changes will be made."
}

$AuditReport = @()

function Audit-Show {
    param ([string]$ShowPath)

    $showName = Split-Path $ShowPath -Leaf
    Write-Log "📺 Auditing show: $showName"

    $missingSeasons = Get-MissingSeasons -ShowPath $ShowPath
    if ($missingSeasons.Count -gt 0) {
        $formattedSeasons = $missingSeasons | ForEach-Object { "Season {0:D2}" -f $_ }
        Write-Log "⚠️ Missing seasons: $($formattedSeasons -join ', ')" -Force
    }

    Get-ChildItem -Path $ShowPath -Directory | Where-Object { $_.Name -match '^Season\s+(\d{1,2})$' } | ForEach-Object {
        $seasonPath = $_.FullName
        $seasonName = $_.Name
        $stats = Get-EpisodeStats -SeasonPath $seasonPath

        $missingEpisodes = @()
        if ($stats.Missing) {
            $missingEpisodes = $stats.Missing | ForEach-Object { "Ep{0:D2}" -f $_ }
        }

        $lastEpisode = if ($stats.Last -is [int] -and $stats.Last -gt 0) {
            "Ep{0:D2}" -f $stats.Last
        } else {
            "None"
        }
        if ($missingEpisodes.Count -gt 0) {
            Write-Log "❌ ${seasonName}: Missing episodes: $($missingEpisodes -join ', ')" -Force
        } else {
            Write-Log "✅ ${seasonName}: All episodes present"
        }

        Write-Log "📌 Last episode in ${seasonName}: $lastEpisode"

        $AuditReport += [PSCustomObject]@{
            Show = $showName
            Season = $seasonName
            MissingEpisodes = $missingEpisodes -join ', '
            LastEpisode = $lastEpisode
        }
    }
}

# Determine if RootPath is a show folder or container
$seasonFolders = Get-ChildItem -Path $RootPath -Directory | Where-Object { $_.Name -match '^Season\s+(\d{1,2})$' }
if ($seasonFolders.Count -gt 0) {
    Audit-Show -ShowPath $RootPath
} else {
    Get-ChildItem -Path $RootPath -Directory | ForEach-Object {
        Audit-Show -ShowPath $_.FullName
    }
}

Write-Log "=== Audit Summary ===" -Force
$AuditReport | Format-Table -AutoSize

Write-Log "Script completed." -Force