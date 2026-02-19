# ============================================================
#  Merge System + User PATH, Deduplicate, Apply to User PATH
# ============================================================

# 1. Backup existing PATHs
$systemPath = [Environment]::GetEnvironmentVariable("PATH", "Machine")
$userPath   = [Environment]::GetEnvironmentVariable("PATH", "User")

$timestamp = (Get-Date).ToString("yyyyMMdd-HHmmss")
$backupDir = "$env:USERPROFILE\PATH-merge-backups"
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null

Set-Content -Path "$backupDir\SystemPATH-$timestamp.txt" -Value $systemPath
Set-Content -Path "$backupDir\UserPATH-$timestamp.txt"   -Value $userPath

Write-Host "Backed up System and User PATH to $backupDir"

# 2. Merge System + User PATH
$merged = @()
$merged += ($systemPath -split ';')
$merged += ($userPath   -split ';')

# 3. Normalize entries (trim, remove empties)
$normalized = $merged |
    ForEach-Object { $_.Trim() } |
    Where-Object { $_ -ne "" }

# 4. Deduplicate while preserving order
$seen = @{}
$deduped = foreach ($entry in $normalized) {
    $key = $entry.ToLower()
    if (-not $seen.ContainsKey($key)) {
        $seen[$key] = $true
        $entry
    }
}

# 5. Reassemble cleaned PATH
$cleanPath = ($deduped -join ';')

# 6. Apply cleaned PATH to User environment only (safe)
[Environment]::SetEnvironmentVariable("PATH", $cleanPath, "User")
Write-Host "Updated User PATH with merged + deduped PATH."

# 7. Update current session PATH
$env:PATH = $cleanPath
Write-Host "Updated current session PATH."
