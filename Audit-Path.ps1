param (
    [string]$Executable = "git.exe"
)

function Log { param($msg); Write-Host "[+] $msg" }

function Audit-Path {
    $pathEntries = $env:Path -split ';' | Where-Object { $_ -ne "" }
    $seen = @{}
    $valid = @()
    $missing = @()
    $duplicates = @()
    $shadowed = @()

    foreach ($entry in $pathEntries) {
        $clean = $entry.Trim('"')
        if ($seen.ContainsKey($clean)) {
            $duplicates += $clean
            continue
        }
        $seen[$clean] = $true

        if (Test-Path $clean) {
            $valid += $clean
            $exePath = Join-Path $clean $Executable
            if (Test-Path $exePath) {
                $shadowed += $exePath
            }
        } else {
            $missing += $clean
        }
    }

    Log "✅ Valid PATH entries: $($valid.Count)"
    Log "❌ Missing folders: $($missing.Count)"
    Log "⚠️ Duplicates: $($duplicates.Count)"
    Log "🔄 '$Executable' found in: $($shadowed.Count) locations"

    if ($shadowed.Count -gt 0) {
        Log "🔍 First '$Executable' resolved by PATH:"
        Write-Host "    $($shadowed[0])"
    }

    return @{
        Valid = $valid
        Missing = $missing
        Duplicates = $duplicates
        Shadowed = $shadowed
    }
}

# Run audit
Audit-Path