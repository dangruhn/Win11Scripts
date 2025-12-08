#****************************************************************************************************
#    handledownload.ps1
#
#    - Logs the start of the script execution for traceability.
#    - Supports multiple download handling modes (e.g., TV, Movies, F1 events) via conditional blocks.
#    - Integrates with FileBot to organize and rename downloaded media files (TV shows, movies, music) 
#      using custom naming formats.
#    - Passes metadata (such as torrent name, category, file paths) to FileBot for automated processing.
#    - Handles both single-file and multi-file (RAR or directory) downloads.
#    - Provides a dedicated mode for processing F1 event downloads with custom logic.
#    - Implements a global mutex (PlexDownloadQueue) to prevent concurrent script executions, ensuring 
#      only one instance processes downloads at a time.
#    - Uses structured logging for monitoring and debugging.
#    - Designed for extensibility and safe concurrent operation in automated media workflows.
#****************************************************************************************************
#   Modification Log:
#   2025-08-09  dgruhn
#   : Start mod log.
#   : Add displayed progress bar for F1 uploading.
#   : Made plans for same for other media files.
#   : Fixed CopyF1File function.
#   : Modified LogOutput to optionally take a log file path.
# 
#****************************************************************************************************
[CmdletBinding()]
param(
    [Parameter()]
    [String]$TorrentName,   # %N: Torrent name
    [String]$Category,	    # %L: Category
    [String]$Tags,	        # %G: Tags (separated by comma)
    [String]$ContentPath,   # %F: Content path (same as root path for multi-file torrent)
    [String]$RootPath,	    # %R: Root path (first torrent subdirectory path)
    [String]$SavePath,	    # %D: Save path
    [int]$NumberOfFiles,    # %C: Number of files
    [long]$TorrentSize,	    # %Z: Torrent size (bytes)
    [String]$TorrentId	    # %K: Torrent ID (either sha-1 or truncated sha-256 info)
)


# Load Windows Forms and Drawing assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# Logging directory
$qBittorrentLogsDir = "$env:USERPROFILE\logs"
$global:logFile = "$qBittorrentLogsDir\handledownload.log"


#******************************************************************************
# Log the input parameters to the log file with a timestamp
#******************************************************************************
function LogOutput {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
        $Args,

        [Parameter(Mandatory = $false)]
        [string]$Color,

        [Parameter(Mandatory = $false)]
        [string]$Prefix,

        [Parameter(Mandatory = $false)]
        [string]$LogFile
    )

    # Only use global log file if -LogFile was NOT passed explicitly
    if (-not $PSBoundParameters.ContainsKey('LogFile') -and $global:logFile)
    {
        $LogFile = $global:logFile
    }

    foreach ($Arg in $Args)
    {
        # Split into lines if it's a string with newlines
        $lines = if ($Arg -is [string]) { $Arg -split "`r?`n" } else { @("$Arg") }

        foreach ($line in $lines)
        {
            $Date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $PrefixText = if ($Prefix) { "$Prefix " } else { "" }
            $LogLine = "$Date $PID $PrefixText$line"

            try {
                Write-Host $LogLine -ForegroundColor $Color
            } catch {
                Write-Host $LogLine  # Fallback if color is invalid
            }

            if ($LogFile)
            {
                try {
                    Add-Content -Path $LogFile -Value $LogLine
                } catch {
                    Write-Warning "Failed to write to log file: $LogFile"
                }
            }
        }
    }
}

# Script wide variables
LogOutput "TorrentName: $TorrentName" `
            "Category: $Category" `
            "ContentPath: $ContentPath" `
            "RootPath: $RootPath" `
            "SavePath: $SavePath" `
            "NumberOfFiles: $NumberOfFiles" `
            "TorrentSize: $TorrentSize" `
            "TorrentId: $TorrentId" `
            "TorrentSize: $TorrentSize"

$global:VideoName = ""

# Development directory
$DevelopmentDir = "C:\Users\Dan.Gruhn\Dropbox\dgruhn-home\Documents\Development"
$F1InfoDir = "$DevelopmentDir\Formula1"

# Formula 1 information database files directory
$F1DBDir = "$F1InfoDir\f1db"

# Formula 1 information database files
$F1CircuitsFilestem = "circuits"
$F1RacesFilestem = "races"

# Formula 1 destination root
$F1DestRoot = "M:\Video\Sports"

# F1 Information
$script:F1Circuits = @()
$script:F1Races = @()


#********************************************************************************
# Circuit info
# This is used along with the info from the F1 info files, correlated by CircuitRef
#********************************************************************************

$formula1Circuits = @(
    @{
       circuitRef = "yas_marina"
       Pattern = "Abu Dhabi/Abu-Dhabi/Yas Marina Circuit/Abu Dhabi Grand Prix"
       CircuitName = "Abu Dhabi"
    },
    @{
       circuitRef = "shanghai"
       Pattern = "Chinese/China/Shanghai International Circuit/Shanghai/China Grand Prix/Chinese Grand Prix"
       CircuitName = "Shanghai"
    },
    @{
       circuitRef = "albert_park"
       Pattern = "Australia/Albert Park Circuit/Australian Grand Prix"
       CircuitName = "Australia"
    },
    @{
       circuitRef = "red_bull_ring"
       Pattern = "Austria/Red Bull Ring/Austrian Grand Prix"
       CircuitName = "Austria"
    },
    @{
       circuitRef = "baku"
       Pattern = "Azerbaijan/Baku City Circuit/Azerbaijan Grand Prix"
       CircuitName = "Azerbaijan"
    },
    @{
       circuitRef = "bahrain"
       Pattern = "Bahrain/Bahrain International Circuit/Bahrain Grand Prix"
       CircuitName = "Bahrain"
    },
    @{
       circuitRef = "spa"
       Pattern = "Belgium/Circuit de Spa-Francorchamps/Belgian Grand Prix"
       CircuitName = "Belgium"
    },
    @{
       circuitRef = "interlagos"
       Pattern = "Brazil/Sao Paulo/Autodromo Jose Carlos Pace/Brazilian Grand Prix/Brazilian"
       CircuitName = "Brazil"
    },
    @{
       circuitRef = "villeneuve"
       Pattern = "Canada/Circuit Gilles Villeneuve/Canadian Grand Prix"
       CircuitName = "Canada"
    },
    @{
       circuitRef = "imola"
       Pattern = "Emilia Romagna/Emilia-Romagna/Autodromo Enzo e Dino Ferrari/Emilia-Romagna Grand Prix/Emilia Romagna Grand Prix"
       CircuitName = "Imola"
    },
    @{
       circuitRef = "catalunya"
       Pattern = "Spain/Circuit de Barcelona-Catalunya/Spanish Grand Prix"
       CircuitName = "Barcelona"
    },
    @{
       circuitRef = "hungaroring"
       Pattern = "Hungary Race/Hungary/Hungaroring/Hungarian Grand Prix"
       CircuitName = "Hungary"
    },
    @{
       circuitRef = "monza"
       Pattern = "Italy/Autodromo Nazionale di Monza/Italian/Italian Grand Prix"
       CircuitName = "Monza"
    },
    @{
       circuitRef = "suzuka"
       Pattern = "Japan/Suzuka International Circuit/Japanese Grand Prix"
       CircuitName = "Suzuka"
    },
    @{
       circuitRef = "vegas"
       Pattern = "Las Vegas/Las Vegas Motor Speedway/Las Vegas Grand Prix"
       CircuitName = "Las Vegas"
    },
    @{
       circuitRef = "silverstone"
       Pattern = "UK/U.K./Great Britain/British/British Grand Prix"
       CircuitName = "Great Britain"
    },
    @{
       circuitRef = "rodriguez"
       Pattern = "Mexico/Autodromo Hermanos Rodr�guez/Mexican Grand Prix/Mexican"
       CircuitName = "Mexico"
    },
    @{
       circuitRef = "miami"
       Pattern = "Miami/Miami International Autodrome/Miami Grand Prix"
       CircuitName = "Miami"
    },
    @{
       circuitRef = "monaco"
       Pattern = "Monaco/Circuit de Monaco/Monaco Grand Prix"
       CircuitName = "Monaco"
    },
   @{
       circuitRef = "zandvoort"
       Pattern = "Netherlands/Circuit Zandvoort/Dutch Grand Prix/Dutch"
       CircuitName = "Netherlands"
    },
    @{
       circuitRef = "losail"
       Pattern = "Qatar/Losail International Circuit/Qatar Grand Prix"
       CircuitName = "Qatar"
    },
    @{
       circuitRef = "jeddah"
       Pattern = "Saudi Arabia/Jeddah Corniche Circuit/Saudi Arabian Grand Prix"
       CircuitName = "Saudi Arabia"
    },
    @{
       circuitRef = "montjuic"
       Pattern = "Spain/Circuit de Barcelona-Catalunya/Spanish Grand Prix"
       CircuitName = "Spain"
    },
    @{
       circuitRef = "marina_bay"
       Pattern = "Singapore/Marina Bay Street Circuit/Singapore Grand Prix"
       CircuitName = "Singapore"
    },
    @{
       circuitRef = "americas"
       Pattern = "United States/USA/COTA/Circuit of the Americas/American Grand Prix/United States Grand Prix/American/austin"
       CircuitName = "United States COTA"
    }
)


#********************************************************************************
# Event info
#********************************************************************************

$eventTypes = @(
    @{
        Pattern = "(?<=\W|^)drivers\W+press\W+conference(?=\W|$)"
        Name = "01 - Drivers Press Conference"
    },
    @{
        Pattern = "(?<=\W|^)team\W+principals\W+press\W+conference(?=\W|$)"
        Name = "02 - Team Principals Press Conference"
    },
    @{
        Pattern = "(?<=\W|^)f1\W+show(?=\W|$)"
        Name = "05 - F1 Show"
    },
    @{
        Pattern = "(?<=\W|^)(?:free\W+)?practice\W+(?:one|1)(?=\W|$)/(?<=\W|^)FP1(?=\W|$)"
        Name = "11 - Free Practice 1"
    },
    @{
        Pattern = "(?<=\W|^)(?:free\W+)?practice\W+(?:two|2)(?=\W|$)/(?<=\W|^)FP2(?=\W|$)"
        Name = "12 - Free Practice 2"
    },
    @{
        Pattern = "(?<=\W|^)(?:free\W+)?practice\W+(?:three|3)(?=\W|$)/(?<=\W|^)FP3(?=\W|$)"
        Name = "13 - Free Practice 3"
    },
    @{
        Pattern = "(?<=\W|^)tech\W+talk\.1(?=\W|$)"
        Name = "21 - Tech Talk 1"
    },
    @{
        Pattern = "(?<=\W|^)tech\W+talk\W+2(?=\W|$)"
        Name = "21 - Tech Talk 2"
    },
    @{
        Pattern = "(?<=\W|^)tech\W+talk\W+3(?=\W|$)"
        Name = "21 - Tech Talk 3"
    },
    @{
        Pattern = "(?<=\W|^)pre[\W-]+qualifying\W+buildup(?=\W|$)"
        Name = "51 - Pre-Qualifying Buildup"
    },
    @{
        Pattern = "(?<=\W|^)post[\W-]+qualifying\W+analysis(?=\W|$)"
        Name = "53 - Post-Qualifying Analysis"
    },
    @{
        Pattern = "(?<=\W|^)ted'?s\W+qualifying\W+notebook(?=\W|$)"
        Name = "54 - Ted's Qualifying Notebook"
    },

   @{
        Pattern = "(?<=\W|^)pre[\W-]+shootou\W+buildup(?<=\W|^)/(?<=\W|^)pre[\W-]+shooutout\W+buildup(?<=\W|^)/(?<=\W|^)pre[\W-]+sprint\W+qualifying\W+buildup(?<=\W|^)"
        Name = "31 - Pre-Sprint Qualifying Buildup"
    },
   @{
        Pattern = "(?<=\W|^)post[\W-]+shootout.analysis(?=\W|$)/(?<=\W|^)post[\W-]+shooutout\W+analysis(?=\W|$)/(?<=\W|^)post[\W-]+sprint\W+qualifying\W+analysis(?=\W|$)"
        Name = "33 - Post-Sprint Qualifying Analysis"
    },
    @{
        Pattern = "(?<=\W|^)sprint.qualifying(?=\W|$)/(?<=\W|^)shootout\W+session(?=\W|$)/(?<=\W|^)sprint\W+shootout(?=\W|$)/(?<=\W|^)shooutout\W+session(?=\W|$)/(?<=\W|^)sprint\W+shooutout(?=\W|$)/(?<=\W|^)sprint\W+qualifying\W+session(?=\W|$)"
        Name = "32 - Sprint Qualifying"
    },
   @{
        Pattern = "(?<=\W|^)pre[\W-]+sprint\W+buildup"
        Name = "41 - Pre-Sprint Buildup"
    },
   @{
        Pattern = "(?<=\W|^)post[\W-]+sprint.analysis(?=\W|$)/(?<=\W|^)post[\W-]+Sprint\W+Qualifying\W+Analysis(?=\W|$)"
        Name = "43 - Post-Sprint Analysis"
    },
   @{
        Pattern = "(?<=\W|^)sprint\W+session(?=\W|$)/(?<=\W|^)sprint\W+race(?=\W|$)"
        Name = "42 - Sprint Race"
    },
   @{
        Pattern = "(?<=\W|^)qualifying(?=\W|$)/(?<=\W|^)qualifying(?=\W|$)/(?<=\W|^)qualifying,notebook(?=\W|$)/(?<=\W|^)quaifying(?=\W|$)/(?<=\W|^)qualy(?=\W|$)"
        Name = "52 - Qualifying"
    },
    @{
        Pattern = "(?<=\W|^)ted'?s\W+sprint\W+notebook(?=\W|$)"
        Name = "44 - Ted's Sprint Notebook"
    },
    @{
        Pattern = "(?<=\W|^)pre[\W-]+race\W+show(?=\W|$)/(?<=\W|^)grand\W+prix\W+sunday(?=\W|$)/(?<=\W|^)pre[\W-]+race\W+buildup(?=\W|$)/(?<=\W|^)race\W+build\W+up(?=\W|$)/(?<=\W|^)on\W+the\W+grid(?=\W|$)/(?<=\W|^)pre[\W-]+race(?=\W|$)"
        Name = "61 - Pre-Race Buildup"
    },
    @{
        Pattern = "(?<=\W|^)race.ssession(?=\W|$)"
        Name = "62 - Grand Prix"
    },
    @{
        Pattern = "(?<=\W|^)chequered.flag(?=\W|$)/(?<=\W|^)post[\W-]+race\W+show(?=\W|$)/(?<=\W|^)post[\W-]+race\W+analysis(?=\W|$)"
        Name = "63 - Post-Race Analysis"
    },
    @{
        Pattern = "(?<=\W|^)ted'?s\W+notebook(?=\W|$)"
        Name = "71 - Ted's Notebook"
    },
    # Note: These are at the end because they might match other parts of the filename wrongly
    #       We need to give all other patterns a chance first
    @{
        Pattern = "(?<=\W|^)sprint(?=\W|$)"
        Name = "42 - Sprint Race"
    },
    @{
        Pattern = "(?<=\W|^)race(?=\W|$)/(?<=\W|^)grand\W+prix(?=\W|$)"
        Name = "62 - Grand Prix"
    }

)


#****************************************************************************************************
# .SYNOPSIS
#     Imports Formula 1 circuit and race information from a remote API and saves it locally.
#
# .DESCRIPTION
#     The Import-F1Information function checks if the local F1 circuits and races data is older than a
#     week. If so, it downloads the latest data from the Ergast API in chunks of 100 entries. The
#     function handles rate limiting by checking for a 429 status code and pauses the execution if
#     necessary. After downloading, it imports the circuit and race information into script-scoped
#     variables.
#
# .PARAMETER F1DBDir
#     The directory where the F1 data files are stored.
#
# .PARAMETER F1CircuitsFilestem
#     The base filename for the F1 circuits data files.
#
# .PARAMETER F1RacesFilestem
#     The base filename for the F1 races data files.
#
# .EXAMPLE
#     Import-F1Information
#     This command will execute the function to import the latest F1 circuit and race information.
#
# .NOTES
#     Author: [Your Name]
#     Date: [Date]
#     This function requires internet access to download data from the Ergast API.
#****************************************************************************************************
Function Import-F1Information
{
    $limit = (Get-Date).AddDays(-7)
    $lastWriteTime = $limit

    # If we already have downloaded the F1 circuits file
    $dstPathname = "$F1DBDir\${F1CircuitsFilestem}0.json"
    if  (Test-Path $dstPathname)
    {
        # Get its file time
        $lastWriteTime = (Get-Item $dstPathname).LastWriteTime
    }
    # If the F1 database files are a week or more old
    $dnlChunkSize = 100
    if ($lastWriteTime -le $limit)
    {
        # Download a new copy
        LogOutput "Downloading a new copy of F1 information"

        # While there are more circuits to download
        $dnlOffset = 0
        do {
            $statusCode = 200
            $dstPathname = "$F1DBDir\$F1CircuitsFilestem$dnlOffset.json"
            try
            {
                Invoke-WebRequest -UseBasicParsing -Uri "https://api.jolpi.ca//ergast/f1/circuits/?limit=100&offset=$dnlOffset" -OutFile "$dstPathname"

                # This will only execute if the Invoke-WebRequest is successful.
                $jsonInfo = Get-Content $dstPathname | ConvertFrom-Json
                $dnlOffset += $dnlChunkSize
            }
            catch
            {
                $statusCode = $_.Exception.Response.StatusCode.value__
                if ($statusCode -eq 429)
                {
                    LogOutput "Error: Too many requests: $response.StatusCode"
                    Start-Sleep -Seconds 1
                }
            }
        } while (([int]$jsonInfo.MRData.offset + $dnlChunkSize) -lt [int]$jsonInfo.MRData.total)

        # While there are more races to download
        $dnlOffset = 0
        do {
            $statusCode = 200
            $dstPathname = "$F1DBDir\$F1RacesFilestem$dnlOffset.json"
            try
            {
                Invoke-WebRequest -Uri "https://api.jolpi.ca///ergast/f1/races/?limit=100&offset=$dnlOffset" -OutFile $dstPathname

                # This will only execute if the Invoke-WebRequest is successful.
                $jsonInfo = Get-Content $dstPathname | ConvertFrom-Json
                $dnlOffset += $dnlChunkSize
            }
            catch
            {
                $statusCode = $_.Exception.Response.StatusCode.value__
                if ($statusCode -eq 429)
                {
                    LogOutput "Error: Too many requests: $response.StatusCode"
                    Start-Sleep -Seconds 1
                }
            }
        } while (([int]$jsonInfo.MRData.offset + $dnlChunkSize) -lt [int]$jsonInfo.MRData.total)
    }
    # Import circuit information
    $dnlOffset = 0
    $script:F1Circuits = @()
    $srcPathname = "$F1DBDir\$F1CircuitsFilestem$dnlOffset.json"
    do {
        $jsonInfo = Get-Content $srcPathname | ConvertFrom-Json
        $script:F1Circuits += $jsonInfo.MRData.CircuitTable.Circuits
        $dnlOffset += $dnlChunkSize
        $srcPathname = "$F1DBDir\$F1CircuitsFilestem$dnlOffset.json"

        # While there are more circuits to load
    } while (Test-Path -Path $srcPathname)

    # Import race information
    $dnlOffset = 0
    $script:F1Races = @()
    $srcPathname = "$F1DBDir\$F1RacesFilestem$dnlOffset.json"
    do {
        $jsonInfo = Get-Content $srcPathname | ConvertFrom-Json
        $script:F1Races += $jsonInfo.MRData.RaceTable.Races
        $dnlOffset += $dnlChunkSize
        $srcPathname = "$F1DBDir\$F1RacesFilestem$dnlOffset.json"
    } while (Test-Path -Path $srcPathname)
}


#********************************************************************************
# .SYNOPSIS
#     Creates a new temporary folder with a unique name based on a random hexadecimal value.
#
# .DESCRIPTION
#     The New-TemporaryFolder function generates a temporary folder in the system's temp directory. 
#     The folder name is constructed using a random number converted to a hexadecimal string, ensuring 
#     that each folder created has a unique name. The folder is created using the New-Item cmdlet.
#
# .PARAMETER None
#     This function does not take any parameters.
#
# .EXAMPLE
#     New-TemporaryFolder
#     This command creates a new temporary folder in the system's temp directory.
#********************************************************************************
Function New-TemporaryFolder
{
    # Make a new folder based upon a TempFileName
    $TemporaryFolder="$($Env:temp)\tmp$([convert]::tostring((get-random 65535),16).padleft(4,'0')).tmp"
    $null = New-Item -ItemType Directory -Path $TemporaryFolder
   return $TemporaryFolder
}


#********************************************************************************
# Get successively longer leaves in the given path
#********************************************************************************
Function Get-SuccessivePathLeaves
{
    $pathname = $args[0]

    $parent = Split-Path -Path "$pathname" -Parent
    $leaf = Split-Path -Path "$pathname" -Leaf
    $result = @($leaf)
    while ($parent.Length -gt 0)
    {
        $parentLeaf = Split-Path -Path $parent -Leaf
        $parent = Split-Path -Path $parent -Parent
        $leaf = $parentLeaf, $leaf -join '\'
        $result += $leaf
    }
    return $result

}
#********************************************************************************
# Convert a byte count into a human-readable string with appropriate units.
#********************************************************************************
function Format-ByteSize {
    param (
        [Parameter(Mandatory = $true)] [double] $Bytes,
        [Int16]                                 $DecimalPlaces = 2
    )

    $units = @(
        @{ Label = "TB"; Factor = 1TB },
        @{ Label = "GB"; Factor = 1GB },
        @{ Label = "MB"; Factor = 1MB },
        @{ Label = "KB"; Factor = 1KB },
        @{ Label = "bytes"; Factor = 1 }
    )

    foreach ($unit in $units) {
        if ($Bytes -ge $unit.Factor) {
            $rounded = [math]::Round($Bytes / $unit.Factor, $DecimalPlaces)
            $formatted = "{0:F$DecimalPlaces}" -f $rounded
            return "$formatted $($unit.Label)"
        }
    }

    return "$Bytes bytes"
}

#********************************************************************************
# Convert seconds to HH:MM:SS format
# Only shows hours if there are hours, only shows minutes if there are minutes
#********************************************************************************
function Convert-SecondsToHHMMSS {
    param (
        [Parameter(Mandatory = $true)]
        [int]$TotalSeconds
    )

    $hours   = [int][math]::Floor($TotalSeconds / 3600)
    $minutes = [int][math]::Floor(($TotalSeconds % 3600) / 60)
    $seconds = $TotalSeconds % 60

    if ($hours -gt 0) {
        return "{0:D}:{1:D2}:{2:D2}" -f $hours, $minutes, $seconds
    } elseif ($minutes -gt 0) {
        return "{0:D}:{1:D2}" -f $minutes, $seconds
    } else {
        return "{0:D} secs" -f $seconds
    }
}
#********************************************************************************
# Copy a file with a progress bar
#********************************************************************************
function Copy-WithProgress {
    param (
        [string]$SourcePath,
        [string]$DestinationPath
    )

    $window = New-Object Windows.Window
    $window.Title = "Path Confirmation"
    $window.Width = 960
    $window.Height = 300 
    $window.WindowStartupLocation = 'CenterScreen'
    $window.Topmost = $true


    $stackPanel = New-Object Windows.Controls.StackPanel
    $stackPanel.Margin = '10'

    # Source Path TextBlock
    $srcText = New-Object Windows.Controls.TextBlock
    $srcText.Text = "Source: $SourcePath"
    $srcText.Width = 920
    $srcText.Height = 30
    $srcText.Padding = '6,2,6,2'
    $srcText.TextWrapping = 'NoWrap'
    $srcText.Margin = '0,0,0,8'
    $srcText.ToolTip = $SourcePath
    $srcText.FontFamily = 'Consolas'

    # Destination Path TextBlock
    $destText = New-Object Windows.Controls.TextBlock
    $destText.Text = "Destination: $DestinationPath"
    $destText.Width = 920
    $destText.Height = 30
    $destText.Padding = '6,2,6,2'
    $destText.TextWrapping = 'NoWrap'
    $destText.Margin = '0,0,0,8'
    $destText.ToolTip = $DestinationPath
    $destText.FontFamily = 'Consolas'

    # File Size Label
    $sizeLabel = New-Object Windows.Controls.TextBlock
    $sizeLabel.Text = "File Size: Calculating..."
    $sizeLabel.Margin = '0,0,0,6'

    # Progress Bar
    $progressBar = New-Object Windows.Controls.ProgressBar
    $progressBar.Minimum = 0
    $progressBar.Height = 20
    $progressBar.Width = 480
    $progressBar.Margin = '0,0,0,6'

    # Percent Complete Label
    $percentLabel = New-Object Windows.Controls.TextBlock
    $percentLabel.Text = "Progress: 0%"
    $percentLabel.Margin = '0,0,0,6'

    # Time Remaining Label
    $timeLabel = New-Object Windows.Controls.TextBlock
    $timeLabel.Text = "Time Remaining: Calculating..."
    $timeLabel.Margin = '0,0,0,10'

    # Cancel Button
    $cancelButton = New-Object Windows.Controls.Button
    $cancelButton.Content = "Cancel"
    $cancelButton.Width = 80
    $cancelButton.Margin = '0,0,0,0'
    $cancelled = $false
    $cancelButton.Add_Click({ $cancelled = $true })

    # Add controls to panel
    [void]$stackPanel.Children.Add($srcText)
    [void]$stackPanel.Children.Add($destText)
    [void]$stackPanel.Children.Add($sizeLabel)
    [void]$stackPanel.Children.Add($progressBar)
    [void]$stackPanel.Children.Add($percentLabel)
    [void]$stackPanel.Children.Add($timeLabel)
    [void]$stackPanel.Children.Add($cancelButton)

    $window.Content = $stackPanel
    $window.Show()

    # Ensure focus and topmost behavior
    $null = $window.Dispatcher.BeginInvoke([Action]{
        $window.Activate()
        $window.Focus()
    })

    # Start copy
    $sourceStream = [System.IO.File]::OpenRead($SourcePath)
    $destStream = [System.IO.File]::Create($DestinationPath)

    $buffer = New-Object byte[] (2MB)
    $totalBytes = $sourceStream.Length
    $progressBar.Maximum = $totalBytes
    $bytesCopied = 0
    $startTime = Get-Date

    # Format file size
    $sizeText = Format-ByteSize $totalBytes 1
    $sizeLabel.Text = "File Size: $sizeText"

    while (($read = $sourceStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
        if ($cancelled) {
            $sourceStream.Close()
            $destStream.Close()
            Remove-Item $DestinationPath -ErrorAction SilentlyContinue
            $window.Close()
            Write-Host "Copy cancelled."
            return
        }

        $destStream.Write($buffer, 0, $read)
        $bytesCopied += $read
        $progressBar.Value = $bytesCopied

        # Percent complete
        $percent = [math]::Round(($bytesCopied / $totalBytes) * 100, 0)
        $percentLabel.Text = "Progress: $percent%"

        # Estimate time remaining
        $elapsed = (Get-Date) - $startTime
        if ($bytesCopied -gt 0 -and $elapsed.TotalSeconds -gt 0) {
            $rate = $bytesCopied / $elapsed.TotalSeconds
            if ($rate -gt 0) {
                $rateValue = Format-ByteSize $rate 1
                $rateLabel = $rateValue + "/sec"
                $remainingBytes = Format-ByteSize ($totalBytes - $bytesCopied) 1
                $remainingTime = Convert-SecondsToHHMMSS ([math]::Ceiling(($totalBytes - $bytesCopied) / $rate))
                $timeLabel.Text = "Remaining: $remainingTime, $remainingBytes ($rateLabel)"
            }
        }
        # Allow UI to update
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
    }

    $sourceStream.Close()
    $destStream.Close()
    $window.Close()
    Write-Host "Copy completed successfully."
}

function Copy-WithPathCheck {
    param(
        [Parameter(Mandatory=$true)]
        [string]$SourcePath,

        [Parameter(Mandatory=$true)]
        [string]$DestinationPath
    )

    try {
        # Ensure the source exists
        if (-not (Test-Path $SourcePath)) {
            throw "Source path does not exist: $SourcePath"
        }

        # Get the parent directory of the destination
        $destDir = Split-Path $DestinationPath -Parent

        # Ensure the destination directory exists, or try to create it
        if (-not (Test-Path $destDir)) {
            try {
                New-Item -ItemType Directory -Path $destDir -Force | Out-Null
            }
            catch {
                throw "Failed to create destination directory: $destDir. Error: $_"
            }
        }

        # Perform the copy
        # Copy-Item -Path $SourcePath -Destination $DestinationPath -Force
        Copy-WithProgress -SourcePath $SourcePath -DestinationPath $DestinationPath

        Write-Host "Copied $SourcePath → $DestinationPath successfully."
    }
    catch {
        Write-Host "Copy failed. Error: $_"
    }
}

function Extract-RarWith7Zip {
    param(
        [Parameter(Mandatory=$true)]
        [string]$RarFile,

        [Parameter(Mandatory=$true)]
        [string]$OutputDir
    )

    # Candidate paths for 7z.exe
    $paths = @(
        (Join-Path $env:ProgramFiles "7-Zip\7z.exe"),
        (Join-Path ${env:ProgramFiles(x86)} "7-Zip\7z.exe")
    )

    # Find the first valid path
    $sevenZipPath = $paths | Where-Object { Test-Path $_ } | Select-Object -First 1

    if (-not $sevenZipPath)
    {
        throw "7z.exe not found in Program Files or Program Files (x86). Please verify 7-Zip is installed."
    }

    # Ensure output directory exists
    if (-not (Test-Path $OutputDir))
    {
        New-Item -ItemType Directory -Path $OutputDir | Out-Null
    }

    # Run 7-Zip extraction
    $errorOut = & $sevenZipPath x $RarFile "-o$OutputDir" -y 1> $null 2>&1
    $exitCode = $LASTEXITCODE
    
    # Add any error output to the log
    LogOutput "$errorOut"

    return $exitCode
}


#********************************************************************************
# Copy and F1 file to the proper location based on 
#********************************************************************************

Function CopyF1File
{
    $SrcPath = $args[0]
    $EventInfo = $args[1]
    $Suffix = $args[2]

    $DestDir = "$($EventInfo.RaceDate) Formula1 $($EventInfo.CircuitName)"
    $baseName = ($DestDir, $EventInfo.EventName, $EventInfo.ResolutionBits -join " - ")
    if (![string]::IsNullOrEmpty($Suffix)) {
        $dstName = "$baseName.$Suffix"
    } else {
        $dstName = $baseName
    }


    $DestDirPath = "$F1DestRoot", $DestDir -join "\"
    $DestPath = "$DestDirPath", "$dstName" -join "\"

    # Get the video name for the final completion dialog box
    $global:VideoName = Split-Path $DestPath -Leaf

    # Make sure the destination directory exists
    if (!(Test-Path -Path $DestDirPath -PathType Container))
    {
        LogOutput "mkdir $DestDirPath"
        New-Item -Path "$DestDirPath" -ItemType Directory -Force  | Out-Null
    }
    # If the destination file doesn't exist or isn't the correct size
    if ((!(Test-Path -Path $DestPath)) -or ((Get-Item -Path $SrcPath).Length -ne (Get-Item -Path $DestPath).Length))
    {
        Copy-WithProgress -SourcePath $SrcPath -DestinationPath $DestPath
        #Start-BitsTransfer -Source  -Destination $DestPath -Priority High -DisplayName $displayName -Description "$SrcPath to $DestPath"

        LogOutput "Remove read-only from $DestPath"
        attrib -r "$DestPath"
    }
}

#********************************************************************************
# Process a file
#********************************************************************************

Function Invoke-FileProcessing
{
    $srcFolder = $args[0]
    $SrcPathname = $args[1]

    # If the file is a directory
    if (Test-Path $SrcPathname -PathType Container)
    {
        # Get the relevant files in the directory
        $fileList = Get-ChildItem -Path "$SrcPathname\*" -Name -Include *.mkv,*.rar,*.nfo,*.mp4

        # Process each subfile
        foreach ($subfile in $fileList)
        {
            Invoke-FileProcessing $SrcPathname "$SrcPathname\$subfile"
        }
    }
    else
    {
        switch -Wildcard ($SrcPathname)
        {
            "*.mkv" {
                $eventInfo = Get-EventInfoF1 "$srcFolder" "$SrcPathname"

                # If we cannot determine this to be a valid race
                if ($eventInfo.RaceDate -eq "xxxx-xx-xx")
                {
                    LogOutput "Error: Unrecognized input: $srcFolder $SrcPathname"
                }
                else
                {
                    CopyF1File "$SrcPathname" $eventInfo "mkv"
                }
            }
            "*.rar" {
                $TempFolder = New-TemporaryFolder

                LogOutput "Extract $SrcPathname to $TempFolder"


                & "C:\Program Files\WinRAR\Rar.exe" -y -idq e "$SrcPathname" $TempFolder

                # Process the expanded files in the temp folder
                Invoke-FileProcessing $srcFolder "$TempFolder"

                # Remove the temp folder and expanded files
                Remove-Item -path "$TempFolder" -recurse -force

            }
            "*.nfo" {
                $eventInfo = Get-EventInfoF1 "$srcFolder" "$SrcPathname"

                # If we cannot determine this to be a valid race
                if ($eventInfo.RaceDate -eq "xxxx-xx-xx")
                {
                    LogOutput "Error: Unrecognized input: $srcFolder $SrcPathname"
                }
                else
                {
                    CopyF1File "$SrcPathname" $eventInfo "nfo"
                }
            }
            "*.mp4" {
                $eventInfo = Get-EventInfoF1 "$srcFolder" "$SrcPathname"

                # If we cannot determine this to be a valid race
                if ($eventInfo.RaceDate -eq "xxxx-xx-xx")
                {
                    LogOutput "Error: Unrecognized input: $srcFolder $SrcPathname"
                }
                else
                {
                    CopyF1File "$SrcPathname" $eventInfo "mp4"
                }
            }
        }
    }
}


#***************************************************************************************************
<#
.SYNOPSIS
    Extracts Formula 1 event information from a video file path and folder structure.

.DESCRIPTION
    This function analyzes file and folder names to identify Formula 1 race details including year,
    circuit, event type, and video specifications. It uses cached data for performance and 
    cross-references circuit information with F1 race data to determine race dates and rounds.

.PARAMETER args[0]
    The source folder path containing the video file.

.PARAMETER args[1]
    The full pathname of the source video file.

.OUTPUTS
    PSCustomObject
    Returns an object with the following properties:
    - RaceDate: The race date in yyyy-MM-dd format
    - CircuitName: The name of the F1 circuit
    - CircuitId: The unique identifier for the circuit
    - EventName: The type of F1 event (practice, qualifying, race, etc.)
    - ResolutionBits: Video resolution and bit depth (e.g., "1080p - 8b")
    - Round: The round number of the race in the season

.NOTES
    - Uses a static cache ($script:ResolutionBitsCache) to store video resolution/bitdepth results
    - Requires external dependencies: filebot for media info, $formula1Circuits, $script:F1Circuits,
      $script:F1Races, $eventTypes
    - Falls back to current year if no year is found in the path
    - Falls back to "1080p - 8b" if video specs cannot be determined
    - Returns placeholder values if circuit/race cannot be identified

#>
#***************************************************************************************************
Function Get-EventInfoF1
{
    $SrcFolder = $args[0]
    $SrcPathname = $args[1]

    # Static hashtable for caching resolution/bitdepth results
    if (-not $script:ResolutionBitsCache)
    {
        $script:ResolutionBitsCache = @{}
    }

    # Get successively longer leaves of the path
    $PathLeaves = Get-SuccessivePathLeaves $SrcPathname
    $PathLeaves += Get-SuccessivePathLeaves $SrcFolder

    $year = ""
    foreach ($srcName in $PathLeaves)
    {
        # Get the year
        if ($srcName -match '.*(20[0-9][0-9]).*')
        {
            $year = $Matches[1]
            break
        }

    }
    # If no year
    if ($year.Length -eq 0)
    {
        # Use current year
        $year = [datetime]::Now.Year
    }

    # Get resolution and bit depth from the video file, using cache if available
    if ($script:ResolutionBitsCache.ContainsKey($SrcPathname))
    {
        $resolutionBits = $script:ResolutionBitsCache[$SrcPathname]
    }
    else
    {
        $resolutionBits = &filebot -mediainfo -r "$SrcPathname" --format "{height}p - {bitdepth}b" 2>$null
        $script:ResolutionBitsCache[$SrcPathname] = $resolutionBits
    }

    # If no resolution and bits
    if ($resolutionBits.Length -eq 0)
    {
        # 1080p and 8 bits
        $resolutionBits = "1080p - 8b"
    }

    # Decode the circuit name
    $circuitName = ""
    :doneCircuit
    foreach ($srcName in $PathLeaves)
    {
        foreach ($circuit in $formula1Circuits)
        {
            # Check our local match patterns
            foreach ($pattern in @($circuit.Pattern -Split "/" ))
            {
                $pattern = $pattern -replace ' ', '\.'
                if ($srcName -imatch $pattern)
                {
                    $circuitName = $circuit.CircuitName
                    $circuitRef = $circuit.circuitRef

                    # Get the circuit ID
                    $circuitObj = $script:F1Circuits | Where-Object { $_.circuitId -eq $circuitRef }
                    if ($null -ne $circuitObj) {
                        $circuitId = $circuitObj.circuitId
                    } else {
                        $circuitId = $null
                    }
                    break doneCircuit
                }
            }
        }
        # Check the downloaded file match patterns
        foreach ($circuit in $script:F1Circuits)
        {
            foreach ($pattern in @($circuit.circuitName, $circuit.Location.locality, $circuit.Location.country))
            {
                $pattern = $pattern -replace ' ', '\.'
                if ($srcName -imatch $pattern)
                {
                    $circuitName = $circuit.CircuitName
                    $circuitId = $circuit.circuitId
                    break doneCircuit
                }
            }
        }
    }

    # Get all races at the given circuit
    $FileRaceList = $script:F1Races | Where-Object { $_.Circuit.circuitId -eq $circuitId }

    # Get the race date in the year
    $FileRace = ($FileRaceList | Where-Object { $_.season -eq $year } | Select-Object -First 1)
    $raceDate = $FileRace.date
    $raceRound = $FileRace.round

    # If we can determine race/date
    if ($raceDate)
    {
        # Get the date in the proper format
        [DateTime]$dateTime = $raceDate
        $raceDate = $dateTime.ToString("yyyy-MM-dd")


        # Decode the event of the race
        $eventName = ""

        :doneEvent
        foreach ($srcName in $PathLeaves)
        {
            foreach ($eventType in $eventTypes)
            {
                foreach ($pattern in @($eventType.Pattern -Split "/" ))
                {
                    $pattern = $pattern -replace ' ', '[\s\.]+'
                    if ($srcName -imatch $pattern)
                    {
                        $eventName = $eventType.Name
                        break doneEvent
                    }
                }
            }
        }

    }
    else
    {
        $raceDate = "xxxx-xx-xx"
        $circuitName = "UnknownCircuit"
        $eventName = "NoEvent"
    }

    # Return the decoded pieces
    return [pscustomobject]@{
            RaceDate = $raceDate
            CircuitName = $circuitName
            CircuitId = $circuitId
            EventName = $eventName
            ResolutionBits = $resolutionBits
            Round = $raceRound
            }
}


#******************************************************************************
# Determine if the given string is an F1 race event designation
#******************************************************************************

Function IsF1()
{
    $testStr = $args[0];
    return(
            ($testStr -imatch "\.formula1\."  ) -or
            ($testStr -imatch "\.formula\.1\.") -or
            ($testStr -imatch "^formula1\."   ) -or
            ($testStr -imatch "^formula.1\."   ) -or
            ($testStr -imatch "\.f1\."        ) -or
            ($testStr -imatch "^f1\."         )
          );
}
function Invoke-MovieTvFileProcessing {
    [CmdletBinding()]
    param (
        [string]$TorrentName,
        [ValidateSet("TV", "Movie")] [string]$Category,
        [string]$ContentPath
    )

    $ExePath = "C:\Program Files\FileBot\filebot.exe"

    # Format strings
    $SeriesFormat = @"
TV Shows/{n}/{episode.special ? 'Specials' : 'Season '+s.pad(2)}/{n} - {episode.special ? 'S00E'+special.pad(2) : s00e00} - {t} - {vf} - {bitdepth}b
"@

    $MovieFormat = @"
MoviesTmp/{n} ({y})/{n} ({y}) - {vf} - {bitdepth}b
"@
    $OutputRoot =  "M:/Video"
    #$OutputRoot =  "C:/Users/Dan.Gruhn/tmp"

    # Run FileBot AMC script in test mode and capture output so we know how to handle files
    # NOTE: We add -tmp to the formats to allow actual renaming when we get the artwork
    LogOutput "Getting files to process from '$ContentPath'"
    $output = `
        & "$ExePath" -script fn:amc `
            -non-strict `
            -rename `
            "$ContentPath" `
            --output "$OutputRoot" `
            --action test `
            --def seriesFormat="${SeriesFormat}-tmp" `
            --def movieFormat="$MovieFormat-tmp" 2>&1

    # LogOutput "$output"

    # Filter archive extraction lines
    $extractions = $output | Where-Object { $_ -match '^Read archive' }

    # Filter copy/rename lines
    $copies = $output | Where-Object { $_ -match '^\[TEST\] from' }

    # Process extractions first
    foreach ($line in $extractions)
    {
        # Match lines like: Read archive [example.rar] and extract to [C:\Downloads\Example Movie]
        if ($line -match 'Read archive \[(?<archive>[^\]]+)\] and extract to \[(?<dest>[^\]]+)\]')
        {
            $RarFile = "$ContentPath\$($matches['archive'])"
            $OutputDir = "$($matches['dest'])"
            LogOutput "Extracting archive: $RarFile to $OutputDir"
            Extract-RarWith7Zip -RarFile $RarFile -OutputDir $OutputDir >$null 2>&1
        }
    }

    # Process files copying (with renaming) and artwork fetching
    foreach ($line in $copies)
    {
        # Match lines like: [TEST] from [C:\Downloads\Example Movie\example.mkv] to [M:/Video/MoviesTmp/Example Movie (2023)/example.mkv-tmp]
        if ($line -match 'from \[(?<src>[^\]]+)\] to \[(?<dst>[^\]]+)\]')
        {
            # Copy the file
            $SourcePath = "$($matches['src'])"
            $DestinationPath = "$($matches['dst'])"

            # Remove the -tmp suffix for to get the final path
            $finalFilepath = $DestinationPath -replace '-tmp(?=\.\w+$)', ''

            # Get the video name for the final completion dialog box
            $global:VideoName = Split-Path $finalFilepath -Leaf

            # If the final file already exists
            if (Test-Path $finalFilepath)
            {
                $finalSize = (Get-Item $finalFilepath).Length
                $sourceSize = (Get-Item $SourcePath).Length

                # If the file is the correct size
                if ($finalSize -eq $sourceSize)
                {
                    # If the file already exists and is the same size, skip copying and just rename
                    LogOutput "Renaming file: $finalFilepath to $DestinationPath"
                    Rename-Item -Path $finalFilepath -NewName $DestinationPath
                }
                else
                {
                    LogOutput "$finalFilepath not the correct size, removing it"
                    Remove-Item $finalFilepath -Force
                }
            }
            else
            {
                LogOutput "Copying file: $SourcePath to temporary file $DestinationPath"
                Copy-WithPathCheck -SourcePath $SourcePath -DestinationPath $DestinationPath
            }

            # Add artwork and move to final name
            LogOutput "Rename $DestinationPath to $finalFilepath and add artwork"
            & "$ExePath" -script fn:amc `
                -non-strict `
                -rename `
                "$DestinationPath" `
                --output "$OutputRoot" `
                --action move `
                --def seriesFormat="$SeriesFormat" `
                --def movieFormat="$MovieFormat" `
                --def artwork=y >$null 2>&1
        }
    }
}

function New-ControlPoint {
    param (
        [int]$x,
        [int]$y
    )
    return New-Object System.Drawing.Point($x, $y)
}

function New-ControlSize {
    param (
        [int]$width,
        [int]$height
    )
    return New-Object System.Drawing.Size($width, $height)
}


#******************************************************************************
# Begin Execution
#******************************************************************************

LogOutput "****************************************************************************************"
LogOutput "handledownload.ps1 -TorrentName '$TorrentName' -Category '$Category' -Tags '$Tags' -ContentPath '$ContentPath' -RootPath '$RootPath' -SavePath '$SavePath' -NumberOfFiles $NumberOfFiles -TorrentSize $TorrentSize -TorrentId $TorrentId"
LogOutput "****************************************************************************************"


# Set to $True for manual testing
if ($False)
{
    # Name of file or containing directory (if RAR or multiple files)
    $TorrentName = "star.trek.strange.new.worlds.s03e05.hdr.2160p.web.h265-successfulcrab.mkv"
    # Movie or TV
    $Category = "TV"

    # Path to file or containing directory (if RAR or multiple files)
    $ContentPath = "E:\Downloads\TOR\Done\Star.Trek.Strange.New.Worlds.S03E05.HDR.2160p.WEB.H265-SuccessfulCrab\star.trek.strange.new.worlds.s03e05.hdr.2160p.web.h265-successfulcrab.mkv"

    # Path to torrent done directory (usually doesn't change)
    $SavePath = "F:\Downloads\TOR\Done"
    LogOutput "C:\Program Files\FileBot\filebot.launcher.exe" -script "fn:amc" --output "M:/Video" --log-file "$logFile" --action copy --conflict auto -non-strict --def "seriesFormat=TV Shows/{n}/{episode.special ? 'Specials' : 'Season '+s.pad(2)}/{n} -  {episode.special ? 'S00E'+special.pad(2) : s00e00} - {t} - {vf} - {bitdepth}b" --def "movieFormat=MoviesTmp/{n} ({y})/{n} ({y}) - {vf} - {bitdepth}b" --def "reportError=y" --def "myepisodes=dangruhn:Mess#1024iah" --def "music=y" "artwork=y" "ut_label=" "ut_state=5" "ut_title=$TorrentName" "ut_kind=$Category" "ut_file=$ContentPath" "ut_dir=$SavePath" --def "clean=y" --def "minFileSize=0" --def "minLengthMS=0"

    & "C:\Program Files\FileBot\filebot.launcher.exe" -script "fn:amc" --output "M:/Video" --log-file "$logFile" --action copy --conflict auto -non-strict --def "seriesFormat=TV Shows/{n}/{episode.special ? 'Specials' : 'Season '+s.pad(2)}/{n} -  {episode.special ? 'S00E'+special.pad(2) : s00e00} - {t} - {vf} - {bitdepth}b" --def "movieFormat=MoviesTmp/{n} ({y})/{n} ({y}) - {vf} - {bitdepth}b" --def "reportError=y" --def "myepisodes=dangruhn:Mess#1024iah" --def "music=y" "artwork=y" "ut_label=" "ut_state=5" "ut_title=$TorrentName" "ut_kind=$Category" "ut_file=$ContentPath" "ut_dir=$SavePath" --def "clean=y" --def "minFileSize=0" --def "minLengthMS=0"

}
# Set to $True for F1 testing
elseif ($False)
{
   Import-F1Information
   $SavePath = "E:\Downloads\TOR\Done\10.F1.2025.R18.Singapore.Grand.Prix.Race.Sky.Sports.F1.UHD.2160p.mkv"
   $ContentPath = $SavePath
   Invoke-FileProcessing  $SavePath $ContentPath
   exit 0
}
else
{

    Add-Type -TypeDefinition @"
using System.Threading;
public class MutexLocker {
    public static Mutex GetMutex(string name) {
        return new Mutex(false, name);
    }
}
"@
    $mutexName = "Global\PlexDownloadQueue"
    $mutex = [MutexLocker]::GetMutex($mutexName)

    LogOutput "Waiting for mutex: $mutexName"

    try
    {
        # Wait until the mutex is acquired
        $mutex.WaitOne()

        LogOutput "Mutex acquired. Processing download..."

        # If this is a Formula 1 file
        if (($Tags -eq "F1") -or ($Category -eq "F1") -or (IsF1 $SavePath) -or (IsF1 $ContentPath))
        {

            Import-F1Information
            LogOutput "Processing F1 files in $SavePath with content $ContentPath"
            Invoke-FileProcessing $SavePath $ContentPath
        }
        else
        {
            LogOutput "******************************************************************************"
            LogOutput "${Category}:        $TorrentName"
            LogOutput "******************************************************************************"

            Invoke-MovieTvFileProcessing `
                -TorrentName "$TorrentName" `
                -Category "$Category" `
                -ContentPath "$ContentPath"

            # Optional: Trigger Plex library refresh via API or webhook

        }
    }
    finally
    {
        $mutex.ReleaseMutex()
        LogOutput "Mutex released."
    }
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # Layout constants
    $formWidth = 1800
    $hPadding = 10
    $vPadding = 10

    # Create form
    $form = New-Object Windows.Forms.Form
    $normalized = (Resolve-Path $ContentPath).Path
    $finalComponent = Split-Path $normalized -Leaf
    $form.Text = "Download Complete, Exiting"
    $form.Size = New-Object System.Drawing.Size($formWidth, $formHeight)
    $form.StartPosition = "Manual"
    $form.Location = New-Object System.Drawing.Point (
        (([System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Width  - $formWidth) / 2),
        (([System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Height - $formHeight) / 2))
    $form.Topmost = $true

    # Track current Y position
    $currentY = $vPadding

    # Title label
    $titleLabel = New-Object Windows.Forms.Label
    $titleLabel.Text = ("$($global:VideoName)" -ne "") ? "$global:VideoName" : "$finalComponent"
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $titleLabel.AutoSize = $true
    $titleLabel.Location = New-Object System.Drawing.Point($hPadding, $currentY)
    $form.Controls.Add($titleLabel)

    # Move down after label
    $currentY += $titleLabel.Height + $vPadding

    # Progress bar
    $progressBar = New-Object Windows.Forms.ProgressBar
    $progressBar.Size = New-Object System.Drawing.Size(($formWidth - (3 * $hPadding)), 20)
    $progressBar.Location = New-Object System.Drawing.Point($hPadding, $currentY)
    $progressBar.Value = 0
    $form.Controls.Add($progressBar)

    # Move down after progress bar
    $currentY += $progressBar.Height + $vPadding

    # Exit button
    $exitButton = New-Object Windows.Forms.Button
    $exitButton.Text = "Exit"
    $exitButton.Size = New-Object System.Drawing.Size(80, 30)
    $exitButton.Location = New-Object System.Drawing.Point((($formWidth - $exitButton.Width) / 2), $currentY)
    $exitButton.Add_Click({ $form.Close() })
    $form.Controls.Add($exitButton)

    # Pause button (placed next to Exit)
    $pauseButton = New-Object Windows.Forms.Button
    $pauseButton.Text = "Pause"
    $pauseButton.Size = New-Object System.Drawing.Size(80, 30)
    $pauseButton.Location = New-Object System.Drawing.Point(($exitButton.Location.X + $exitButton.Width + $hPadding), $currentY)
    $form.Controls.Add($pauseButton)

    # Track bottom-most control
    $bottomControl = $form.Controls | Sort-Object { $_.Bottom } -Descending | Select-Object -First 1

    # Add padding at bottom
    $form.Height = $bottomControl.Bottom + $bottomControl.Height + (3 * $vPadding)

    # Set up global state for timer and add click and tick handlers
    $global:timerCancelled = $false
    $global:timerRunning = $true

    # Exit logic
    $global:exited = $false
    $exitButton.Add_Click({
        $global:exited = $true
        $global:timerCancelled = $true
        $form.Close()
    })

    # Timer logic
    $state = [pscustomobject]@{ Counter = 0 }
    $timer = New-Object Windows.Forms.Timer
    $timer.Interval = 200  # 200ms = 20s total
    $timer.Add_Tick({
        if ($global:timerCancelled)
        {
            $timer.Stop()
            return
        }
        if ($global:timerRunning)
        {
            $state.Counter++
            $progressBar.Value = [Math]::Min($state.Counter, $progressBar.Maximum)
            if ($state.Counter -ge $progressBar.Maximum)
            {
                $timer.Stop()
                $form.Close()
            }
        }
        [System.Windows.Forms.Application]::DoEvents()
    })

    # Toggle pause/start state
    $pauseButton.Add_Click({
        # If timer is running, pause it
        if ($global:timerRunning)
        {
            $pauseButton.Text = "Start"
            $global:timerRunning = $false
            $timer.Stop()
        }
        else # Timer is paused, start it
        {
            $pauseButton.Text = "Pause"
            $global:timerRunning = $true
            $timer.Start()
        }
    })

    # Start timer after form is shown
    $form.Add_Shown({
            $form.Activate()
            $form.Focus()
            $form.BringToFront()
            Start-Sleep -Milliseconds 500
            $timer.Start()
        })

    # Run the form with message loop
    [Windows.Forms.Application]::Run($form)
}