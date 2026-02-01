<#
    Module: handledownload
    Extracted utility functions for handledownload.ps1
#>

function New-TemporaryFolder {
    [CmdletBinding()]
    param()
    $TemporaryFolder = "$($Env:temp)\tmp$([convert]::tostring((get-random 65535),16).padleft(4,'0')).tmp"
    $null = New-Item -ItemType Directory -Path $TemporaryFolder -Force
    return $TemporaryFolder
}

function Get-SuccessivePathLeaves {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$Pathname
    )
    $parent = Split-Path -Path "$Pathname" -Parent
    $leaf = Split-Path -Path "$Pathname" -Leaf
    $result = @($leaf)
    while ($parent.Length -gt 0) {
        $parentLeaf = Split-Path -Path $parent -Leaf
        $parent = Split-Path -Path $parent -Parent
        $leaf = $parentLeaf, $leaf -join '\'
        $result += $leaf
    }
    return $result
}

function Format-ByteSize {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] [double] $Bytes,
        [Int16] $DecimalPlaces = 2
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

function Convert-SecondsToHHMMSS {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] [int]$TotalSeconds
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

function Copy-WithProgress {
    param (
        [string]$SourcePath,
        [string]$DestinationPath
    )

    $window = New-Object Windows.Window
    $window.Title = "Path Confirmation"
    $window.WindowStartupLocation = 'CenterScreen'
    $window.Topmost = $true
    $window.SizeToContent = 'WidthAndHeight'

    $stackPanel = New-Object Windows.Controls.StackPanel
    $stackPanel.Margin = '10'

    $srcText = New-Object Windows.Controls.TextBlock
    $srcText.Text = "Source: $SourcePath"
    $srcText.MaxWidth = 1250
    $srcText.Padding = '6,2,6,2'
    $srcText.TextWrapping = 'Wrap'
    $srcText.Margin = '0,0,0,8'
    $srcText.ToolTip = $SourcePath
    $srcText.FontFamily = 'Segoe UI'

    $destText = New-Object Windows.Controls.TextBlock
    $destText.Text = "Destination: $DestinationPath"
    $destText.MaxWidth = 1250
    $destText.Padding = '6,2,6,2'
    $destText.TextWrapping = 'Wrap'
    $destText.Margin = '0,0,0,8'
    $destText.ToolTip = $DestinationPath
    $destText.FontFamily = 'Segoe UI'

    $sizeLabel = New-Object Windows.Controls.TextBlock
    $sizeLabel.Text = "File Size: Calculating..."
    $sizeLabel.Margin = '0,0,0,6'

    $progressBar = New-Object Windows.Controls.ProgressBar
    $progressBar.Minimum = 0
    $progressBar.Height = 20
    $progressBar.Margin = '0,0,0,6'
    $progressBar.HorizontalAlignment = 'Stretch'

    $percentLabel = New-Object Windows.Controls.TextBlock
    $percentLabel.Text = "Progress: 0%"
    $percentLabel.Margin = '0,0,0,6'

    $timeLabel = New-Object Windows.Controls.TextBlock
    $timeLabel.Text = "Time Remaining: Calculating..."
    $timeLabel.Margin = '0,0,0,10'

    $cancelButton = New-Object Windows.Controls.Button
    $cancelButton.Content = "Cancel"
    $cancelButton.Width = 80
    $cancelButton.Margin = '0,0,0,0'
    $cancelled = $false
    $cancelButton.Add_Click({ $cancelled = $true })

    [void]$stackPanel.Children.Add($srcText)
    [void]$stackPanel.Children.Add($destText)
    [void]$stackPanel.Children.Add($sizeLabel)
    [void]$stackPanel.Children.Add($progressBar)
    [void]$stackPanel.Children.Add($percentLabel)
    [void]$stackPanel.Children.Add($timeLabel)
    [void]$stackPanel.Children.Add($cancelButton)

    $window.Content = $stackPanel
    $window.Show()

    $null = $window.Dispatcher.BeginInvoke([Action]{
        $window.Activate()
        $window.Focus()
    })

    $sourceStream = [System.IO.File]::OpenRead($SourcePath)
    $destStream = [System.IO.File]::Create($DestinationPath)

    $buffer = New-Object byte[] (2MB)
    $totalBytes = $sourceStream.Length
    [void]($progressBar.Maximum = $totalBytes)
    $bytesCopied = 0
    $startTime = Get-Date

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
        [void]($progressBar.Value = $bytesCopied)

        $percent = [math]::Round(($bytesCopied / $totalBytes) * 100, 0)
        $percentLabel.Text = "Progress: $percent%"

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
        if (-not (Test-Path $SourcePath)) {
            throw "Source path does not exist: $SourcePath"
        }

        $destDir = Split-Path $DestinationPath -Parent

        if (-not (Test-Path $destDir)) {
            try {
                New-Item -ItemType Directory -Path $destDir -Force | Out-Null
            }
            catch {
                throw "Failed to create destination directory: $destDir. Error: $_"
            }
        }

        Copy-WithProgress -SourcePath $SourcePath -DestinationPath $DestinationPath

        Write-Host "Copied $SourcePath → $DestinationPath successfully."
    }
    catch {
        Write-Host "Copy failed. Error: $_"
    }
}

function Extract-RarWith7Zip {
    param(
        [Parameter(Mandatory=$true)][string]$RarFile,
        [Parameter(Mandatory=$true)][string]$OutputDir
    )

    $paths = @(
        (Join-Path $env:ProgramFiles "7-Zip\7z.exe"),
        (Join-Path ${env:ProgramFiles(x86)} "7-Zip\7z.exe")
    )

    $sevenZipPath = $paths | Where-Object { Test-Path $_ } | Select-Object -First 1

    if (-not $sevenZipPath) {
        Write-Host "7z.exe not found in Program Files or Program Files (x86). Please verify 7-Zip is installed."
        return 1
    }

    if (-not (Test-Path $OutputDir)) {
        New-Item -ItemType Directory -Path $OutputDir | Out-Null
    }

    $proc = Start-Process -FilePath $sevenZipPath -ArgumentList @('x', $RarFile, "-o$OutputDir", '-y') -NoNewWindow -Wait -PassThru
    $exitCode = $proc.ExitCode

    if ($exitCode -ne 0) { Write-Host "7z exit code: $exitCode for file $RarFile" }
    return $exitCode
}

function Invoke-MovieTvFileProcessing {
    [CmdletBinding()]
    param (
        [string]$TorrentName,
        [ValidateSet("TV", "Movie")] [string]$Category,
        [string]$ContentPath
    )

    $ExePath = "C:\Program Files\FileBot\filebot.exe"

    if (-not (Test-Path $ExePath)) {
        Write-Host "FileBot not found at $ExePath. Cannot process TV/Movie files."
        return
    }

    $SeriesFormat = @"
TV Shows/{n}/{episode.special ? 'Specials' : 'Season '+s.pad(2)}/{n} - {episode.special ? 'S00E'+special.pad(2) : s00e00} - {t} - {vf} - {bitdepth}b
"@

    $MovieFormat = @"
MoviesTmp/{n} ({y})/{n} ({y}) - {vf} - {bitdepth}b
"@
    $OutputRoot =  "M:/Video"

    Write-Host "Getting files to process from '$ContentPath'"
    $output = `
        & "$ExePath" -script fn:amc `
            -non-strict `
            -rename `
            "$ContentPath" `
            --output "$OutputRoot" `
            --action test `
            --def seriesFormat="${SeriesFormat}-tmp" `
            --def movieFormat="$MovieFormat-tmp" 2>&1

    $extractions = $output | Where-Object { $_ -match '^Read archive' }
    $copies = $output | Where-Object { $_ -match '^\[TEST\] from' }

    foreach ($line in $extractions) {
        if ($line -match 'Read archive \[(?<archive>[^\]]+)\] and extract to \[(?<dest>[^\]]+)\]') {
            $RarFile = "$ContentPath\$($matches['archive'])"
            $OutputDir = "$($matches['dest'])"
            Write-Host "Extracting archive: $RarFile to $OutputDir"
            $null = Extract-RarWith7Zip -RarFile $RarFile -OutputDir $OutputDir >$null 2>&1
        }
    }

    foreach ($line in $copies) {
        if ($line -match 'from \[(?<src>[^\]]+)\] to \[(?<dst>[^\]]+)\]') {
            $SourcePath = "$($matches['src'])"
            $DestinationPath = "$($matches['dst'])"
            $finalFilepath = $DestinationPath -replace '-tmp(?=\.\w+$)', ''
            $global:VideoName = Split-Path $finalFilepath -Leaf

            if (Test-Path $finalFilepath) {
                $finalSize = (Get-Item $finalFilepath).Length
                $sourceSize = (Get-Item $SourcePath).Length
                if ($finalSize -ne $sourceSize) {
                    Write-Host "$finalFilepath not the correct size, removing it"
                    Remove-Item $finalFilepath -Force
                    Write-Host "Copying file: $SourcePath to temporary file $DestinationPath"
                    Copy-WithPathCheck -SourcePath $SourcePath -DestinationPath $DestinationPath
                }
            }
            else {
                Write-Host "Copying file: $SourcePath to temporary file $DestinationPath"
                Copy-WithPathCheck -SourcePath $SourcePath -DestinationPath $DestinationPath
            }

            Write-Host (("Rename {0} to {1} and add artwork" -f (Split-Path $finalFilepath -Leaf), (Split-Path $SourcePath -Leaf)))
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

Export-ModuleMember -Function New-TemporaryFolder,Get-SuccessivePathLeaves,Format-ByteSize,Convert-SecondsToHHMMSS,Copy-WithProgress,Copy-WithPathCheck,Extract-RarWith7Zip,Invoke-MovieTvFileProcessing
