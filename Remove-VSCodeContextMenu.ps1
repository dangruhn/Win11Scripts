# Remove-VSCodeContextMenu.ps1
param (
    [switch]$DryRun,
    [switch]$Verbose
)

function Remove-ContextMenu {
    $keyPath = "HKCR\batfile\shell\Open with VS Code"

    if ($DryRun) {
        Write-Host "Dry run: Would remove registry key $keyPath"
    } else {
        if (Test-Path "Registry::$keyPath") {
            Remove-Item -Path "Registry::$keyPath" -Recurse -Force
            Write-Host "🧼 Removed context menu: Open with VS Code"
        } else {
            Write-Warning "Context menu not found at $keyPath"
        }
    }

    if ($Verbose) {
        Write-Host "Operation complete."
    }
}

Remove-ContextMenu