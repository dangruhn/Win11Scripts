param (
    [ValidateSet("VSCode", "Neovim", "Python", "All")]
    [string]$Target = "VSCode",
    [switch]$DryRun,
    [switch]$Verbose,
    [switch]$Confirm,
    [switch]$Help
)
$summary = @{
    Deleted = @()
    Skipped = @()
    NotFound = @()
}

$logFile = "$env:USERPROFILE\Desktop\dev_cleanup_log.txt"

function Show-Help {
    Write-Host @"
Dev Environment Cleanup Script - Usage:

    .\dev_cleanup.ps1 [-Target <VSCode|Neovim|Python|All>] [-DryRun] [-Verbose] [-Confirm] [-Help]

Switches:
    -Target     Specify which tool to clean (default: VSCode)
    -DryRun     Preview actions without making changes
    -Verbose    Show detailed output for each step
    -Confirm    Prompt before deleting files or registry keys
    -Help       Show this help message and exit
"@ -ForegroundColor Cyan
    exit
}

if ($Help) { Show-Help }

function Log {
    param (
        [string]$message,
        [string]$color = "White",
        [switch]$always = $false
    )
    if ($Verbose -or $always) {
        Write-Host $message -ForegroundColor $color
    }
    Add-Content -Path $logFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $message"
}

Log "=== Dev Cleanup Script Started ===" "Cyan" -always
Log "Target: $Target" "Yellow" -always
Log "Dry-run mode: $($DryRun.IsPresent)" "Yellow" -always
Log "Verbose mode: $($Verbose.IsPresent)" "Yellow" -always
Log "Confirm mode: $($Confirm.IsPresent)" "Yellow" -always

# Generic deletion function
function Delete-ItemSafe {
    param (
        [string]$Path,
        [string]$Type = "Item"
    )
    if (Test-Path $Path) {
        Log "Found ${Type}: $Path" "Green"
        if (-not $DryRun) {
            if ($Confirm) {
                $response = Read-Host "Delete $Type $Path? (Y/N)"
                if ($response -notin @("Y", "y")) {
                    Log "Skipped deletion of $Path by user choice." "DarkGray"
                    $summary.Skipped += "${Type}: $Path"
                    return
                }
            }
            Log "Deleting ${Type}: $Path" "Red"
            Remove-Item -Recurse -Force -Path $Path
            $summary.Deleted += "${Type}: $Path"
        } else {
            $summary.Skipped += "$Type (dry-run): $Path"
        }
    } else {
        Log "$Type not found: $Path" "DarkGray"
        $summary.NotFound += "${Type}: $Path"
    }
}

function Uninstall-VSCode {
    $uninstallKeys = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    foreach ($keyPath in $uninstallKeys) {
        $keys = Get-ChildItem $keyPath | Where-Object {
            ($_ | Get-ItemProperty).DisplayName -like "*Visual Studio Code*"
        }
        foreach ($key in $keys) {
            $props = $key | Get-ItemProperty
            Log "Found uninstall string: $($props.UninstallString)" "Yellow"
            if (-not $DryRun) {
                if ($Confirm) {
                    $response = Read-Host "Uninstall VS Code? (Y/N)"
                    if ($response -notin @("Y", "y")) {
                        Log "Skipped uninstall by user choice." "DarkGray"
                        continue
                    }
                }
                Log "Uninstalling VS Code..." "Red"
                Start-Process -FilePath "cmd.exe" -ArgumentList "/c", $props.UninstallString -Wait
            }
        }
    }
}


# VS Code Cleanup
function Cleanup-VSCode {
    Log "--- Cleaning VS Code ---" "Cyan"
    Uninstall-VSCode

    $vsPaths = @(
        "$env:APPDATA\Code",
        "$env:LOCALAPPDATA\Programs\Microsoft VS Code",
        "$env:LOCALAPPDATA\Code",
        "$env:USERPROFILE\.vscode"
    )
    foreach ($path in $vsPaths) {
        Delete-ItemSafe -Path $path -Type "VSCode Folder"
    }

    $vsRegKeys = @(
        "HKCU:\Software\Microsoft\VSCommon",
        "HKCU:\Software\Microsoft\VisualStudio",
        "HKCU:\Software\Microsoft\VSCode"
    )
    foreach ($reg in $vsRegKeys) {
        Delete-ItemSafe -Path $reg -Type "VSCode Registry Key"
    }
}

# Neovim Cleanup
function Cleanup-Neovim {
    Log "--- Cleaning Neovim ---" "Cyan"
    $nvimPaths = @(
        "$env:LOCALAPPDATA\nvim",
        "$env:APPDATA\nvim",
        "$env:USERPROFILE\AppData\Local\nvim-data",
        "$env:USERPROFILE\AppData\Roaming\nvim"
    )
    foreach ($path in $nvimPaths) {
        Delete-ItemSafe -Path $path -Type "Neovim Config"
    }
}

# Python Cleanup
function Cleanup-Python {
    Log "--- Cleaning Python ---" "Cyan"
    $pyPaths = @(
        "$env:LOCALAPPDATA\Programs\Python",
        "$env:USERPROFILE\AppData\Local\Programs\Python",
        "$env:APPDATA\Python",
        "$env:USERPROFILE\.python"
    )
    foreach ($path in $pyPaths) {
        Delete-ItemSafe -Path $path -Type "Python Folder"
    }

    $pyRegKeys = @(
        "HKCU:\Software\Python",
        "HKLM:\Software\Python"
    )
    foreach ($reg in $pyRegKeys) {
        Delete-ItemSafe -Path $reg -Type "Python Registry Key"
    }
}

# Dispatch based on target
switch ($Target) {
    "VSCode" { Cleanup-VSCode }
    "Neovim" { Cleanup-Neovim }
    "Python" { Cleanup-Python }
    "All"    {
        Cleanup-VSCode
        Cleanup-Neovim
        Cleanup-Python
    }
}
Log "`n=== Summary Report ===" "Cyan" -always

if ($summary.Deleted.Count -gt 0) {
    Log "`n✅ Deleted:" "Green" -always
    $summary.Deleted | ForEach-Object { Log $_ "Green" -always }
}

if ($summary.Skipped.Count -gt 0) {
    Log "`n❌ Skipped:" "Yellow" -always
    $summary.Skipped | ForEach-Object { Log $_ "Yellow" -always }
}

if ($summary.NotFound.Count -gt 0) {
    Log "`n🚫 Not Found:" "DarkGray" -always
    $summary.NotFound | ForEach-Object { Log $_ "DarkGray" -always }
}

Log "`n=== Cleanup Complete ===" "Cyan" -always
