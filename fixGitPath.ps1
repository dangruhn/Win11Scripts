param (
    [switch]$DryRun
)

function Log {
    param ($msg)
    Write-Host "[+] $msg"
}

function Find-GitExe {
    $searchRoots = @(
        "$env:ProgramFiles\Git",
        "$env:ProgramFiles(x86)\Git",
        "$env:LocalAppData\Programs\Git",
        "C:\"
    )
    foreach ($root in $searchRoots) {
        try {
            $result = Get-ChildItem -Path $root -Recurse -Filter "git.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($result) { return $result.DirectoryName }
        } catch {}
    }
    return $null
}

function Is-InPath {
    param ($folder)
    $env:Path -split ';' | ForEach-Object { $_.Trim('"') } | Where-Object { $_ -eq $folder }
}

function Add-ToPath {
    param ($folder)
    $currentPath = [Environment]::GetEnvironmentVariable("Path", "User")
    $newPath = "$currentPath;$folder"
    if (-not $DryRun) {
        [Environment]::SetEnvironmentVariable("Path", $newPath, "User")
    }
    Log "Added '$folder' to user PATH"
}

# Main
Log "Searching for git.exe..."
$gitFolder = Find-GitExe

if (-not $gitFolder) {
    Log "❌ git.exe not found on system"
    exit 1
}

Log "✅ Found git.exe in: $gitFolder"

if (Is-InPath $gitFolder) {
    Log "✔️ Already in PATH"
} else {
    Log "🔧 Not in PATH"
    Add-ToPath $gitFolder
    if ($DryRun) {
        Log "Dry-run mode: no changes made"
    } else {
        Log "✅ PATH updated. Restart terminal to apply."
    }
}