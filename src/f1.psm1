<#
    Module: f1
    Extracted Formula 1 helper functions from handledownload.ps1
#>

function Import-F1InformationModule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$F1DBDir,
        [Parameter(Mandatory=$true)][string]$F1CircuitsFilestem,
        [Parameter(Mandatory=$true)][string]$F1RacesFilestem
    )

    $limit = (Get-Date).AddDays(-7)
    $lastWriteTime = $limit

    $dnlChunkSize = 100

    # If we already have downloaded the F1 circuits file
    $dstPathname = "$F1DBDir\${F1CircuitsFilestem}0.json"
    if (Test-Path $dstPathname) { $lastWriteTime = (Get-Item $dstPathname).LastWriteTime }

    if ($lastWriteTime -le $limit) {
        LogOutput "Downloading a new copy of F1 information"
        $dnlOffset = 0
        do {
            $statusCode = 200
            $dstPathname = "$F1DBDir\$F1CircuitsFilestem$dnlOffset.json"
            try {
                Invoke-WebRequest -UseBasicParsing -Uri "https://api.jolpi.ca//ergast/f1/circuits/?limit=100&offset=$dnlOffset" -OutFile "$dstPathname"
                $jsonInfo = Get-Content $dstPathname | ConvertFrom-Json
                $dnlOffset += $dnlChunkSize
            } catch {
                $statusCode = $null
                if ($_.Exception -and $_.Exception.Response) { try { $statusCode = $_.Exception.Response.StatusCode.value__ } catch { $statusCode = $null } }
                if ($statusCode -eq 429) { LogOutput "Error: Too many requests: $statusCode"; Start-Sleep -Seconds 1 } else { LogOutput "Error downloading circuits: $($_.Exception.Message)" }
            }
        } while (([int]$jsonInfo.MRData.offset + $dnlChunkSize) -lt [int]$jsonInfo.MRData.total)

        $dnlOffset = 0
        do {
            $statusCode = 200
            $dstPathname = "$F1DBDir\$F1RacesFilestem$dnlOffset.json"
            try {
                Invoke-WebRequest -Uri "https://api.jolpi.ca///ergast/f1/races/?limit=100&offset=$dnlOffset" -OutFile $dstPathname
                $jsonInfo = Get-Content $dstPathname | ConvertFrom-Json
                $dnlOffset += $dnlChunkSize
            } catch {
                $statusCode = $null
                if ($_.Exception -and $_.Exception.Response) { try { $statusCode = $_.Exception.Response.StatusCode.value__ } catch { $statusCode = $null } }
                if ($statusCode -eq 429) { LogOutput "Error: Too many requests: $statusCode"; Start-Sleep -Seconds 1 } else { LogOutput "Error downloading races: $($_.Exception.Message)" }
            }
        } while (([int]$jsonInfo.MRData.offset + $dnlChunkSize) -lt [int]$jsonInfo.MRData.total)
    }

    # Import circuit information
    $dnlOffset = 0
    $F1Circuits = @()
    $srcPathname = "$F1DBDir\$F1CircuitsFilestem$dnlOffset.json"
    do {
        $jsonInfo = Get-Content $srcPathname | ConvertFrom-Json
        $F1Circuits += $jsonInfo.MRData.CircuitTable.Circuits
        $dnlOffset += $dnlChunkSize
        $srcPathname = "$F1DBDir\$F1CircuitsFilestem$dnlOffset.json"
    } while (Test-Path -Path $srcPathname)

    # Import race information
    $dnlOffset = 0
    $F1Races = @()
    $srcPathname = "$F1DBDir\$F1RacesFilestem$dnlOffset.json"
    do {
        $jsonInfo = Get-Content $srcPathname | ConvertFrom-Json
        $F1Races += $jsonInfo.MRData.RaceTable.Races
        $dnlOffset += $dnlChunkSize
        $srcPathname = "$F1DBDir\$F1RacesFilestem$dnlOffset.json"
    } while (Test-Path -Path $srcPathname)

    return @{ Circuits = $F1Circuits; Races = $F1Races }
}

function Get-EventInfoF1Module {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$SrcFolder,
        [Parameter(Mandatory=$true)][string]$SrcPathname,
        [Parameter(Mandatory=$true)][array]$Formula1Circuits,
        [Parameter(Mandatory=$true)][array]$F1Circuits,
        [Parameter(Mandatory=$true)][array]$F1Races,
        [Parameter(Mandatory=$true)][array]$EventTypes,
        [string]$FileBotExe = 'filebot'
    )

    if (-not $script:ResolutionBitsCache) { $script:ResolutionBitsCache = @{} }

    $PathLeaves = @()
    $PathLeaves += (Get-SuccessivePathLeaves $SrcPathname)
    $PathLeaves += (Get-SuccessivePathLeaves $SrcFolder)

    $year = ""
    foreach ($srcName in $PathLeaves) { if ($srcName -match '.*(20[0-9][0-9]).*') { $year = $Matches[1]; break } }
    if ($year.Length -eq 0) { $year = [datetime]::Now.Year }

    if ($script:ResolutionBitsCache.ContainsKey($SrcPathname)) { $resolutionBits = $script:ResolutionBitsCache[$SrcPathname] } else { $resolutionBits = & $FileBotExe -mediainfo -r "$SrcPathname" --format "{height}p - {bitdepth}b" 2>$null; $script:ResolutionBitsCache[$SrcPathname] = $resolutionBits }
    if ($resolutionBits.Length -eq 0) { $resolutionBits = "1080p - 8b" }

    $circuitName = ""; $circuitId = $null
    :doneCircuit
    foreach ($srcName in $PathLeaves) {
        foreach ($circuit in $Formula1Circuits) {
            $pattern = $circuit.Pattern -replace ' ', '[\\s\\._-]+'
            if ($srcName -imatch $pattern) { $circuitName = $circuit.CircuitName; $circuitRef = $circuit.circuitRef; $circuitObj = $F1Circuits | Where-Object { $_.circuitId -eq $circuitRef } | Select-Object -First 1; if ($null -ne $circuitObj) { $circuitId = $circuitObj.circuitId } else { $circuitId = $null }; break doneCircuit }
        }
        foreach ($circuit in $F1Circuits) {
            foreach ($pattern in @($circuit.circuitName, $circuit.Location.locality, $circuit.Location.country)) {
                $pattern = $pattern -replace ' ', '[\\s\\._-]+'
                if ($srcName -imatch $pattern) { $circuitName = $circuit.circuitName; $circuitId = $circuit.circuitId; break doneCircuit }
            }
        }
    }

    $FileRaceList = $F1Races | Where-Object { $_.Circuit.circuitId -eq $circuitId }
    $FileRace = ($FileRaceList | Where-Object { $_.season -eq [int]$year } | Select-Object -First 1)
    $raceDate = $FileRace.date
    $raceRound = $FileRace.round

    if ($null -ne $raceDate -and $raceDate -ne '') {
        [DateTime]$dateTime = $raceDate; $raceDate = $dateTime.ToString('yyyy-MM-dd')
        $eventName = ""
        :doneEvent
        foreach ($srcName in $PathLeaves) {
            foreach ($eventType in $EventTypes) {
                $pattern = $eventType.Pattern -replace ' ', '[\\s\\._-]+'
                $m = [regex]::Match($srcName, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                if ($m.Success) { if ($eventType.Name -is [scriptblock]) { $eventName = & $eventType.Name $m } else { $eventName = $eventType.Name }; break doneEvent }
            }
        }
    } else {
        $raceDate = 'xxxx-xx-xx'; $circuitName = 'UnknownCircuit'; $eventName = 'NoEvent'
    }

    return [pscustomobject]@{ RaceDate = $raceDate; CircuitName = $circuitName; CircuitId = $circuitId; EventName = $eventName; ResolutionBits = $resolutionBits; Round = $raceRound }
}

function CopyF1FileModule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$SrcPath,
        [Parameter(Mandatory=$true)][psobject]$EventInfo,
        [string]$Suffix,
        [Parameter(Mandatory=$true)][string]$F1DestRoot
    )

    $DestDir = "$($EventInfo.RaceDate) Formula1 $($EventInfo.CircuitName)"
    $baseName = ($DestDir, $EventInfo.EventName, $EventInfo.ResolutionBits -join ' - ')
    if (![string]::IsNullOrEmpty($Suffix)) { $dstName = "$baseName.$Suffix" } else { $dstName = $baseName }
    $DestDirPath = "$F1DestRoot\$DestDir"
    $DestPath = Join-Path $DestDirPath $dstName
    $global:VideoName = Split-Path $DestPath -Leaf
    if (!(Test-Path -Path $DestDirPath -PathType Container)) { LogOutput "mkdir $DestDirPath"; New-Item -Path $DestDirPath -ItemType Directory -Force | Out-Null }
    if ((!(Test-Path -Path $DestPath)) -or ((Get-Item -Path $SrcPath).Length -ne (Get-Item -Path $DestPath).Length)) { Copy-WithProgress -SourcePath $SrcPath -DestinationPath $DestPath; attrib -r "$DestPath" }
}

function Invoke-F1FileProcessingModule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$SrcFolder,
        [Parameter(Mandatory=$true)][string]$SrcPathname,
        [Parameter(Mandatory=$true)][string]$F1DestRoot,
        [Parameter(Mandatory=$true)][array]$Formula1Circuits,
        [Parameter(Mandatory=$true)][array]$F1Circuits,
        [Parameter(Mandatory=$true)][array]$F1Races,
        [Parameter(Mandatory=$true)][array]$EventTypes
    )

    if (Test-Path $SrcPathname -PathType Container) {
        $fileList = Get-ChildItem -Path "$SrcPathname\*" -Name -Include *.mkv,*.rar,*.nfo,*.mp4
        foreach ($subfile in $fileList) { Invoke-F1FileProcessingModule -SrcFolder $SrcFolder -SrcPathname "$SrcPathname\$subfile" -F1DestRoot $F1DestRoot -Formula1Circuits $Formula1Circuits -F1Circuits $F1Circuits -F1Races $F1Races -EventTypes $EventTypes }
    } else {
        switch -Wildcard ($SrcPathname) {
            '*.mkv' { $eventInfo = Get-EventInfoF1Module -SrcFolder $SrcFolder -SrcPathname $SrcPathname -Formula1Circuits $Formula1Circuits -F1Circuits $F1Circuits -F1Races $F1Races -EventTypes $EventTypes; if ($eventInfo.RaceDate -eq 'xxxx-xx-xx') { LogOutput "Error: Unrecognized input: $SrcFolder $SrcPathname" } else { CopyF1FileModule -SrcPath $SrcPathname -EventInfo $eventInfo -Suffix 'mkv' -F1DestRoot $F1DestRoot } }
            '*.rar' {
                $TempFolder = New-TemporaryFolder
                LogOutput "Extract $SrcPathname to $TempFolder"
                $winRarPath = 'C:\Program Files\WinRAR\Rar.exe'
                if (-not (Validate-Executable -Path $winRarPath -Name 'WinRAR Rar.exe')) { LogOutput "Cannot extract RAR: WinRAR not found. Skipping $SrcPathname" } else { & "$winRarPath" -y -idq e "$SrcPathname" $TempFolder }
                Invoke-F1FileProcessingModule -SrcFolder $SrcFolder -SrcPathname $TempFolder -F1DestRoot $F1DestRoot -Formula1Circuits $Formula1Circuits -F1Circuits $F1Circuits -F1Races $F1Races -EventTypes $EventTypes
                Remove-Item -Path $TempFolder -Recurse -Force
            }
            '*.nfo' { $eventInfo = Get-EventInfoF1Module -SrcFolder $SrcFolder -SrcPathname $SrcPathname -Formula1Circuits $Formula1Circuits -F1Circuits $F1Circuits -F1Races $F1Races -EventTypes $EventTypes; if ($eventInfo.RaceDate -eq 'xxxx-xx-xx') { LogOutput "Error: Unrecognized input: $SrcFolder $SrcPathname" } else { CopyF1FileModule -SrcPath $SrcPathname -EventInfo $eventInfo -Suffix 'nfo' -F1DestRoot $F1DestRoot } }
            '*.mp4' { $eventInfo = Get-EventInfoF1Module -SrcFolder $SrcFolder -SrcPathname $SrcPathname -Formula1Circuits $Formula1Circuits -F1Circuits $F1Circuits -F1Races $F1Races -EventTypes $EventTypes; if ($eventInfo.RaceDate -eq 'xxxx-xx-xx') { LogOutput "Error: Unrecognized input: $SrcFolder $SrcPathname" } else { CopyF1FileModule -SrcPath $SrcPathname -EventInfo $eventInfo -Suffix 'mp4' -F1DestRoot $F1DestRoot } }
        }
    }
}

function IsF1Module {
    param([Parameter(Mandatory=$true)][string]$TestStr)
    return (($TestStr -imatch "\.formula1\.") -or ($TestStr -imatch "\.formula\.1\.") -or ($TestStr -imatch "^formula1\.") -or ($TestStr -imatch "^formula.1\.") -or ($TestStr -imatch "\.f1\.") -or ($TestStr -imatch "^f1\."))
}

Export-ModuleMember -Function Import-F1InformationModule, Get-EventInfoF1Module, CopyF1FileModule, Invoke-F1FileProcessingModule, IsF1Module
