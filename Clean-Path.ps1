param (
    [switch]$Apply,
    [switch]$DryRun,
    [switch]$Help
)

function Show-Help {
@"
Clean-Path.ps1 — Audits and cleans your PATH environment variable

USAGE:
    .\Clean-Path.ps1 [-Apply] [-DryRun] [-Help]

PARAMETERS:
    -Apply     Apply cleaned PATH to User scope
    -DryRun    Preview changes without applying
    -Help      Show this help message

FEATURES:
    ✅ Combines System and User PATH
    🧹 Removes duplicates
    ❌ Filters out missing folders
    🔄 Optionally updates User PATH

EXAMPLES:
    .\Clean-Path.ps1 -DryRun
        Preview cleaned PATH entries

    .\Clean-Path.ps1 -Apply
        Apply cleaned PATH to User scope

    .\Clean-Path.ps1 -Help
        Show this help message
"@ | Write-Host
}

function Log { param($msg); Write-Host "[+] $msg" }

function Get-PathEntries {
    param ($scope)
    [Environment]::GetEnvironmentVariable("Path", $scope) -split ';' |
        Where-Object { $_ -ne "" } |
        ForEach-Object { $_.Trim('"') }
}

function Clean-Path {
    $system = Get-PathEntries "Machine"
    $user = Get-PathEntries "User"

    $combined = $system + $user
    $unique = [System.Collections.Generic.HashSet[string]]::new()
    $valid = @()
    $invalid = @()

    foreach ($entry in $combined) {
        $clean = $entry.Trim()
        if (-not $unique.Contains($clean)) {
            $null = $unique.Add($clean)  # Suppress 'True' output
            if (Test-Path $clean) {
                $valid += $clean
            }
            else {
                $invalid += $clean
            }
        }
    }

    Log "✅ Valid entries: $($valid.Count)"
    Log "❌ Missing folders removed: $($invalid.Count.Count)"
    Log "⚠️ Duplicates removed: $($combined.Count - $unique.Count)"

    $newPath = ($valid -join ";")

    if ($DryRun) {
        Log "Dry-run: would set User PATH to:"
        Write-Host $newPath
    }
    elseif ($Apply) {
        [Environment]::SetEnvironmentVariable("Path", $newPath, "User")
        Log "✅ Cleaned PATH written to User scope"
    }

    return @{
        Valid   = $valid
        Invalid = $invalid
    }
}

# Entry point
if ($Help) {
    Show-Help
} else {
    Clean-Path
}