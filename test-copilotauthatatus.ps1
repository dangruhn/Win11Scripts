function Test-CopilotAuthStatus {
    param (
        [switch]$PromptIfMissing,
        [switch]$DryRun
    )

    $codeCmd = "$env:USERPROFILE\AppData\Local\Programs\Microsoft VS Code\bin\code.cmd"
    $copilotExt = "github.copilot"

    # Check if Copilot extension is installed
    Write-Host "🔍 Checking GitHub Copilot extension installed." -ForegroundColor Cyan

    $installed = & $codeCmd --list-extensions
    if ($installed -notcontains $copilotExt) {
        Write-Host "❌ GitHub Copilot extension not installed." -ForegroundColor Red
        return $false
    }

    # Check Copilot status via VS Code command
    $statusCmd = "--command github.copilot.status"
    Write-Host "🔍 Checking Copilot authentication status..." -ForegroundColor Cyan
    if (-not $DryRun) {
        Start-Process -FilePath $codeCmd -ArgumentList $statusCmd
    }

    # Optional: Prompt for re-authentication
    if ($PromptIfMissing) {
        Write-Host "🔁 Prompting GitHub login via Copilot..." -ForegroundColor Yellow
        if (-not $DryRun) {
            Start-Process -FilePath $codeCmd -ArgumentList "--folder-uri vscode://github-authentication", "--command workbench.view.account"
        }
    }

    return $true
}
Test-CopilotAuthStatus -PromptIfMissing