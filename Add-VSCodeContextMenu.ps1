# Add-VSCodeContextMenu.ps1
param (
    [switch]$DryRun,
    [switch]$Verbose
)

function Get-VSCodePath {
    $paths = @(
        "$env:LOCALAPPDATA\Programs\Microsoft VS Code\Code.exe",
        "C:\Program Files\Microsoft VS Code\Code.exe"
    )

    foreach ($path in $paths) {
        if (Test-Path $path) {
            if ($Verbose) { Write-Host "Found VS Code at: $path" }
            return $path
        }
    }

    # Fallback to PATH
    $cliPath = & where code 2>$null
    if ($cliPath) {
        $resolved = (Get-Command $cliPath).Source
        if ($Verbose) { Write-Host "Found VS Code via PATH: $resolved" }
        return $resolved
    }

    throw "VS Code not found. Please install it or add to PATH."
}

function Add-ContextMenu {
    param (
        [string]$VSCodePath
    )

    $keyBase = "Registry::HKCR\batfile\shell\Open with VS Code"
    $keyCommand = "$keyBase\command"
    $command = "`"$VSCodePath`" `"%1`""

    if ($DryRun) {
        Write-Host "Dry run: Would set registry key $keyCommand with command:"
        Write-Host $command
    } else {
        New-Item -Path $keyBase -Force | Out-Null
        New-Item -Path $keyCommand -Force | Out-Null
        Set-ItemProperty -Path $keyCommand -Name "(default)" -Value $command
        Write-Host "✅ Context menu added: Open with VS Code"
    }
}

try {
    $vsCodePath = Get-VSCodePath
    Add-ContextMenu -VSCodePath $vsCodePath
} catch {
    Write-Error $_
}