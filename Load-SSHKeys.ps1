
param(
    [string]$BasePath = "C:\Users\dgruhn\Dropbox\dgruhn-home\Documents\Akamai\Security\Keys"
)
$logFile = "$HOME\ssh-agent-load.log"

# Helper to write log with timestamp
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    "$timestamp $Message" | Out-File -FilePath $logFile -Append
}

function Add-Key {
    param([string]$keyPath)

    # Check if key is already loaded
    $loadedKeys = & ssh-add -l 2>&1
    if ($loadedKeys -match [regex]::Escape($keyPath)) {
        Write-Log "Key already loaded: $keyPath"
        return
    }

    # Try to add the key
    try {
        ssh-add $keyPath 2>&1 | ForEach-Object { Write-Log $_ }
        Write-Log "Loaded key: $keyPath"
    }
    catch {
        Write-Log "Failed to load key: $keyPath - $_"
    }
}


# Check if ssh-agent is running
try {
    $null = & ssh-add -l 2>&1
} catch {
    Write-Log "ssh-agent not running. Attempting to start ssh-agent."
    $agentOutput = & ssh-agent | Out-String
    Write-Log $agentOutput
    # Re-check after starting
    $null = & ssh-add -l 2>&1
}

# Find latest key for each type
$keyTypes = @('internal', 'deployed', 'github')

foreach ($type in $keyTypes) {
    $pattern = "dgruhn-*-${type}-linux"
    $files = Get-ChildItem -Path $BasePath -Filter $pattern | Sort-Object LastWriteTime -Descending
    if ($files.Count -gt 0) {
        $latestKey = $files[0].FullName
        # Copy latest key to dgruhn-latest-${type} if needed
        $destPath = Join-Path -Path $BasePath -ChildPath "dgruhn-latest-${type}"
        $copyNeeded = $true
        if (Test-Path $destPath) {
            $srcInfo = Get-Item $latestKey
            $destInfo = Get-Item $destPath
            if ($srcInfo.Length -eq $destInfo.Length -and $srcInfo.LastWriteTime -le $destInfo.LastWriteTime) {
                $copyNeeded = $false
            }
        }
        if ($copyNeeded) {
            try {
                if (Test-Path $destPath) {
                    Remove-Item $destPath -Force
                }
                Copy-Item -Path $latestKey -Destination $destPath -Force
                Write-Log "Copied $latestKey to $destPath"
            } catch {
                Write-Log "Failed to copy $latestKey to ${destPath}: $_"
            }
        } else {
            Write-Log "No copy needed for ${type}: $destPath is up to date."
        }
        Add-Key $latestKey
    } else {
        Write-Log "No keys found for type: ${type}"
    }
}

