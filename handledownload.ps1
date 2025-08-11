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
# Script wide variables

# Logging directory
$qBittorrentLogsDir = "$env:USERPROFILE\logs"
$global:logFile = "$qBittorrentLogsDir\handledownload.log"

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
# Copy file
#********************************************************************************

Function Copy-F1File
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

    # Make sure the destination directory exists
    if (!(Test-Path -Path $DestDirPath -PathType Container))
    {
        LogOutput "mkdir $DestDirPath"
    }
    # If the destination file doesn't exist or isn't the correct size
    if ((!(Test-Path -Path $DestPath)) -or ((Get-Item -Path $SrcPath).Length -ne (Get-Item -Path $DestPath).Length))
    {
        $displayName = ("Uploading", $EventInfo.RaceDate, $EventInfo.CircuitName, $EventInfo.EventName) -join " "
        Start-BitsTransfer -Source $SrcPath -Destination $DestPath -Priority High -DisplayName $displayName -Description "$SrcPath to  $DestPath"

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
                    Copy-F1File "$SrcPathname" $eventInfo "mkv"
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
                    Copy-F1File "$SrcPathname" $eventInfo "mp4"
                }
            }
        }
    }
}


#********************************************************************************
# Get the information about an F1 event
#********************************************************************************

Function Get-EventInfoF1
{
    $SrcFolder = $args[0]
    $SrcPathname = $args[1]

    # Static hashtable for caching resolution/bitdepth results
    if (-not $script:ResolutionBitsCache) {
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
    if ($script:ResolutionBitsCache.ContainsKey($SrcPathname)) {
        $resolutionBits = $script:ResolutionBitsCache[$SrcPathname]
    } else {
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
# Log the input parameters to the log file with a timestamp
#******************************************************************************

function LogOutput {
    param (
        [string]$LogFile,
        [string]$Color = "White",              # Optional Write-Host color
        [string]$Prefix = "",                  # Optional line prefix
        [Parameter(ValueFromRemainingArguments = $true)]
        $Args
    )

    # Fallback to global log file if no log file given
    if (-not $LogFile -and $global:logFile) {
        $LogFile = $global:logFile
    }

    foreach ($Arg in $Args) {
        # Split into lines if it's a string with newlines
        $lines = if ($Arg -is [string]) { $Arg -split "`r?`n" } else { @("$Arg") }

        foreach ($line in $lines) {
            $Date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $PrefixText = if ($Prefix) { "$Prefix " } else { "" }
            $LogLine = "$Date $PID $PrefixText$line"

            try {
                Write-Host $LogLine -ForegroundColor $Color
            } catch {
                Write-Host $LogLine  # Fallback if color is invalid
            }

            if ($LogFile) {
                try {
                    Add-Content -Path $LogFile -Value $LogLine
                } catch {
                    Write-Warning "Failed to write to log file: $LogFile"
                }
            }
        }
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


#******************************************************************************
# Invoke the filebot process
#******************************************************************************

function Invoke-FileBot {
    [CmdletBinding()]
    param (
        [string]$TorrentName,
        [ValidateSet("TV", "Movie")] [string]$Category,
        [string]$ContentPath,
        [string]$RootPath,
        [string]$SavePath,
        [int]$NumberOfFiles,
        [long]$TorrentSize,
        [string]$TorrentId,
        [switch]$DryRun
    )
    $ExePath = '"C:\Program Files\FileBot\filebot.launcher.exe"'  # quoted
    $FilebotLogPath = 'C:\Users\Dan.Gruhn\logs\filebot.log'

    # Precompute conditional values
    $dbSource = if ($Category -eq 'TV') { 'TheTVDB' } else { 'TheMovieDB' }
    $actionType = if ($DryRun) { 'test' } else { 'copy' }

    # Ensure the log directory exists
    $LogDir = Split-Path -Path $FilebotLogPath -Parent
    if (-not (Test-Path -Path $LogDir)) {
        New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
    }

    # Format strings
    $SeriesFormat = @"
TV Shows/{n}/{episode.special ? 'Specials' : 'Season '+s.pad(2)}/{n} - {episode.special ? 'S00E'+special.pad(2) : s00e00} - {t} - {vf} - {bitdepth}b
"@

    $MovieFormat = @"
MoviesTmp/{n} ({y})/{n} ({y}) - {vf} - {bitdepth}b
"@


    $argList = @(
        "-script fn:amc",
        "--output M:/Video",
        "--conflict auto",
        "-non-strict",
        "--action $actionType",
        "--def seriesFormat=`"$SeriesFormat`"",
        "--def movieFormat=`"$MovieFormat`"",
        "--def reportError=y",
        "--def myepisodes=dangruhn:Mess#1024iah",
        "--def music=y",
        "--def artwork=y",
        "--def ut_label=",
        "--def ut_state=5",
        "--def ut_title=`"$TorrentName`"",
        "--def ut_kind=`"$Category`"",
        "--def ut_file=`"$ContentPath`"",
        "--def ut_dir=`"$SavePath`"",
        "--def clean=y",
        "--def minFileSize=0",
        "--def minLengthMS=0"
    )
    $wrappedCmd = "$ExePath $argList >> $FilebotLogPath 2>&1"

    if ($DryRun) {
        LogOutput "Dry run: $wrappedCmd"
        return
    }
    
    LogOutput "Filebot command: $wrappedCmd"
    # Timestamped log entry
    LogOutput -LogFile $FilebotLogPath ""
    LogOutput -LogFile $FilebotLogPath "Starting FileBot..."
    LogOutput -LogFile $FilebotLogPath "$wrappedCmd"

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $ExePath
    $psi.Arguments = $argList -join ' '
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.CreateNoWindow = $true

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    $proc.Start() | Out-Null

    # Capture output
    $stdout = $proc.StandardOutput.ReadToEnd()
    $stderr = $proc.StandardError.ReadToEnd()
    # $proc.WaitForExit()

    # Access exit code
    $exitCode = $proc.ExitCode

    # Log or display
    LogOutput -LogFile $FilebotLogPath -Color "Green" -Prefix "[Output] " $stdout
    if ($stderr) {
        LogOutput -LogFile $FilebotLogPath -Color "Red" -Prefix "[Error] " $stderr
    }
    LogOutput -LogFile $FilebotLogPath "Exit Code: $exitCode`nFileBot completed."

    if ($proc.ExitCode -ne 0) {
        LogOutput  "FileBot exited with code $($proc.ExitCode). Check $FilebotLogPath for details."
    }
    else {
        LogOutput  "FileBot completed successfully."
    }

}

#******************************************************************************
# Begin Execution
#******************************************************************************

LogOutput "******************************************************************************"
LogOutput "Starting handledownload.ps1 script"
LogOutput "******************************************************************************"

if ($False)
{
    # Name of file or containing directory (if RAR or multiple files)
    $TorrentName = "The Rookie S07E17 Mutiny and the Bounty 1080p AMZN WEB-DL DDP5 1 HEVC-YELLO"

    # Movie or TV
    $Category = "TV"

    # Path to file or containing directory (if RAR or multiple files)
    $ContentPath = "F:\Downloads\TOR\Done\The Rookie S07E17 Mutiny and the Bounty 1080p AMZN WEB-DL DDP5 1 HEVC-YELLO\The Rookie S07E17 Mutiny and the Bounty 1080p AMZN WEB-DL DDP5 1 HEVC-YELLO.mkv"

    # Path to torrent done directory (usually doesn't change)
    $SavePath = "F:\Downloads\TOR\Done"
    LogOutput "C:\Program Files\FileBot\filebot.launcher.exe" -script "fn:amc" --output "P:/DLNA/Video" --log-file "$logFile" --action copy --conflict auto -non-strict --def "seriesFormat=TV Shows/{n}/{episode.special ? 'Specials' : 'Season '+s.pad(2)}/{n} -  {episode.special ? 'S00E'+special.pad(2) : s00e00} - {t} - {vf} - {bitdepth}b" --def "movieFormat=MoviesTmp/{n} ({y})/{n} ({y}) - {vf} - {bitdepth}b" --def "reportError=y" --def "myepisodes=dangruhn:Mess#1024iah" --def "music=y" "artwork=y" "ut_label=" "ut_state=5" "ut_title=$TorrentName" "ut_kind=$Category" "ut_file=$ContentPath" "ut_dir=$SavePath" --def "clean=y" --def "minFileSize=0" --def "minLengthMS=0"

    & "C:\Program Files\FileBot\filebot.launcher.exe" -script "fn:amc" --output "P:/DLNA/Video" --log-file "$logFile" --action copy --conflict auto -non-strict --def "seriesFormat=TV Shows/{n}/{episode.special ? 'Specials' : 'Season '+s.pad(2)}/{n} -  {episode.special ? 'S00E'+special.pad(2) : s00e00} - {t} - {vf} - {bitdepth}b" --def "movieFormat=MoviesTmp/{n} ({y})/{n} ({y}) - {vf} - {bitdepth}b" --def "reportError=y" --def "myepisodes=dangruhn:Mess#1024iah" --def "music=y" "artwork=y" "ut_label=" "ut_state=5" "ut_title=$TorrentName" "ut_kind=$Category" "ut_file=$ContentPath" "ut_dir=$SavePath" --def "clean=y" --def "minFileSize=0" --def "minLengthMS=0"

}
# Set to $True for F1 download
elseif ($False)
{
   Import-F1Information
   $SavePath = "F:\Downloads\TOR\Done\Formula1.2025.Round05.Saudi.Arabia.FP1.F1TV.WEB-DL.2160p.HLG.H265.Multi-MWR"
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

    try {
        # Wait until the mutex is acquired
        $mutex.WaitOne()  # Blocks until available

        LogOutput "Mutex acquired. Processing download..."

        # If this is a Formula 1 file
        if (($Tags -eq "F1") -or
            ($Category -eq "F1") -or
            (IsF1 $SavePath) -or
            (IsF1 $ContentPath)) {
            # Handle it ourselves
            Import-F1Information
            LogOutput "Processing F1 files in $SavePath with content $ContentPath"
            Invoke-FileProcessing $SavePath $ContentPath
        }
        else {
            Invoke-FileBot `
                -TorrentName "$TorrentName" `
                -Category "$Category" `
                -ContentPath "$ContentPath" `
                -RootPath "$RootPath" `
                -SavePath "$SavePath" `
                -NumberOfFiles $NumberOfFiles `
                -TorrentSize $TorrentSize`
                -TorrentId "$TorrentId"
            # Optional: Trigger Plex library refresh via API or webhook

        }
    }
    finally {
        $mutex.ReleaseMutex()
        LogOutput "Mutex released."
    }
} 
