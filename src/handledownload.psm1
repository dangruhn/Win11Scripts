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

Export-ModuleMember -Function New-TemporaryFolder,Get-SuccessivePathLeaves,Format-ByteSize,Convert-SecondsToHHMMSS
