param (
    [switch]$VerifyCopilot,
    [switch]$DryRun,
    [switch]$Verbose
)

# Logging setup
$logFile = "$env:USERPROFILE\Desktop\vscode_install_log.txt"
$summary = @()

function Log {
    param (
        [string]$message,
        [string]$color = "White",
        [switch]$always = $false,
        [switch]$verboseOnly = $false
    )
    if ($verboseOnly -and -not $Verbose) { return }
    if ($DryRun -or $always -or $Verbose) { Write-Host $message -ForegroundColor $color }
    Add-Content -Path $logFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $message"
}

function Add-Summary {
    param ([string]$entry)
    $summary += $entry
}

Log "=== VS Code Install Script Started ===" "Cyan" -always
Log "Dry-run mode: $DryRun" "Yellow" -always
Log "Verbose mode: $Verbose" "Yellow" -always

$vsCodeUrl = "https://update.code.visualstudio.com/latest/win32-x64-user/stable"
$installerPath = "$env:TEMP\VSCodeSetup.exe"
$extensions = @(
    "asvetliakov.vscode-neovim"
    "eamodio.gitlens"
    "esbenp.prettier-vscode"
    "github.copilot"
    "github.copilot-chat"
    "github.vscode-pull-request-github"
    "ms-python.debugpy"
    "ms-python.python"
    "ms-python.vscode-pylance"
    "ms-python.vscode-python-envs"
    "ms-vscode-remote.remote-ssh"
    "ms-vscode-remote.remote-ssh-edit"
    "ms-vscode-remote.vscode-remote-extensionpack"
    "ms-vscode.cmake-tools"
    "ms-vscode.cpptools"
    "ms-vscode.cpptools-extension-pack"
    "ms-vscode.cpptools-themes"
    "ms-vscode.powershell"
    "ms-vscode.remote-explorer"
    "ms-vscode.remote-server"
    "pkief.material-icon-theme"
    "sdras.night-owl"
    "vscode-icons-team.vscode-icons"
)

$extensionAffinityConfig = @{
    "asvetliakov.vscode-neovim" = 1
    "ms-vscode.vscode-typescript-next" = 2
}

function Set-VSCodeIconTheme {
    param ([string]$theme = "vscode-icons")
    $settingsPath = "$env:APPDATA\Code\User\settings.json"
    $backupPath = "$settingsPath.bak"

    if (-not (Test-Path $settingsPath)) {
        New-Item -ItemType File -Path $settingsPath -Force | Out-Null
        Set-Content $settingsPath "{}" -Encoding UTF8
    }

    if (-not (Test-Path $backupPath)) {
        Copy-Item $settingsPath $backupPath -Force
    }

    $raw = Get-Content $settingsPath -Raw
    $settings = @{}
    if ($raw.Trim()) {
        try { $settings = $raw | ConvertFrom-Json } catch {
            Log "Invalid JSON in settings.json. Skipping theme update." "Red"
            return
        }
    }

    $settingsHash = @{}
    foreach ($prop in $settings.PSObject.Properties) {
        $settingsHash[$prop.Name] = $prop.Value
    }

    if ($settingsHash["workbench.iconTheme"] -ne $theme) {
        $settingsHash["workbench.iconTheme"] = $theme
        if (-not $DryRun) {
            $settingsHash | ConvertTo-Json -Depth 10 | Set-Content $settingsPath -Encoding UTF8
        }
        Log "✅ Icon theme set to '$theme'" "Green"
        Add-Summary "Icon theme set to '$theme'"
    } else {
        Log "🔄 Icon theme already set to '$theme'" "Yellow"
        Add-Summary "Icon theme already set to '$theme'"
    }
}

function Set-VSCodeExtensionAffinities {
    param ([hashtable]$config)
    $settingsPath = "$env:APPDATA\Code\User\settings.json"
    $backupPath = "$settingsPath.bak"

    if (-not (Test-Path $settingsPath)) {
        New-Item -ItemType File -Path $settingsPath -Force | Out-Null
        Set-Content $settingsPath "{}" -Encoding UTF8
    }

    if (-not (Test-Path $backupPath)) {
        Copy-Item $settingsPath $backupPath -Force
    }

    $raw = Get-Content $settingsPath -Raw
    $settings = @{}
    if ($raw.Trim()) {
        try { $settings = $raw | ConvertFrom-Json } catch {
            Log "Invalid JSON in settings.json. Skipping affinity update." "Red"
            return
        }
    }

    $settingsHash = @{}
    foreach ($prop in $settings.PSObject.Properties) {
        $settingsHash[$prop.Name] = $prop.Value
    }

    if (-not $settingsHash["extensions.experimental.affinity"]) {
        $settingsHash["extensions.experimental.affinity"] = @{}
    }

    $affinity = $settingsHash["extensions.experimental.affinity"]
    if ($affinity -isnot [hashtable]) {
        $affinityHash = @{}
        foreach ($prop in $affinity.PSObject.Properties) {
            $affinityHash[$prop.Name] = $prop.Value
        }
        $affinity = $affinityHash
    }

    $changed = $false
    foreach ($ext in $config.Keys) {
        $desiredAffinity = $config[$ext]
        if ($affinity[$ext] -ne $desiredAffinity) {
            $affinity[$ext] = $desiredAffinity
            Log "✅ Affinity set for '$ext' to $desiredAffinity" "Green"
            Add-Summary "Affinity set for '$ext' to $desiredAffinity"
            $changed = $true
        } else {
            Log "🔄 Affinity already set for '$ext'" "Yellow"
            Add-Summary "Affinity already set for '$ext'"
        }
    }

    if ($changed -and -not $DryRun) {
        $settingsHash["extensions.experimental.affinity"] = $affinity
        $settingsHash | ConvertTo-Json -Depth 10 | Set-Content $settingsPath -Encoding UTF8
    }
}

function Set-VSCodeWindowBehavior {
    $settingsPath = "$env:APPDATA\Code\User\settings.json"
    $backupPath = "$settingsPath.bak"

    if (-not (Test-Path $settingsPath)) {
        New-Item -ItemType File -Path $settingsPath -Force | Out-Null
        Set-Content $settingsPath "{}" -Encoding UTF8
    }

    if (-not (Test-Path $backupPath)) {
        Copy-Item $settingsPath $backupPath -Force
    }

    $raw = Get-Content $settingsPath -Raw
    $settings = @{}
    if ($raw.Trim()) {
        try { $settings = $raw | ConvertFrom-Json } catch {
            Log "Invalid JSON in settings.json. Skipping window behavior update." "Red"
            return
        }
    }

    $settingsHash = @{}
    foreach ($prop in $settings.PSObject.Properties) {
        $settingsHash[$prop.Name] = $prop.Value
    }

    if ($settingsHash["window.newWindowDimensions"] -ne "inherit") {
        $settingsHash["window.newWindowDimensions"] = "inherit"
        if (-not $DryRun) {
            $settingsHash | ConvertTo-Json -Depth 10 | Set-Content $settingsPath -Encoding UTF8
        }
        Log "✅ Set 'window.newWindowDimensions' to 'inherit'" "Green"
        Add-Summary "'window.newWindowDimensions' set to 'inherit'"
    } else {
        Log "🔄 'window.newWindowDimensions' already set to 'inherit'" "Yellow"
        Add-Summary "'window.newWindowDimensions' already set"
    }
}


function Set-VSCodeColorTheme {
    param ([string]$theme = "Visual Studio Light")
    $settingsPath = "$env:APPDATA\Code\User\settings.json"
    $backupPath = "$settingsPath.bak"

    if (-not (Test-Path $settingsPath)) {
        New-Item -ItemType File -Path $settingsPath -Force | Out-Null
        Set-Content $settingsPath "{}" -Encoding UTF8
    }

    if (-not (Test-Path $backupPath)) {
        Copy-Item $settingsPath $backupPath -Force
    }

    $raw = Get-Content $settingsPath -Raw
    $settings = @{}
    if ($raw.Trim()) {
        try { $settings = $raw | ConvertFrom-Json } catch {
            Log "Invalid JSON in settings.json. Skipping color theme update." "Red"
            return
        }
    }

    $settingsHash = @{}
    foreach ($prop in $settings.PSObject.Properties) {
        $settingsHash[$prop.Name] = $prop.Value
    }

    if ($settingsHash["workbench.colorTheme"] -ne $theme) {
        $settingsHash["workbench.colorTheme"] = $theme
        if (-not $DryRun) {
            $settingsHash | ConvertTo-Json -Depth 10 | Set-Content $settingsPath -Encoding UTF8
        }
        Log "✅ Color theme set to '$theme'" "Green"
        Add-Summary "Color theme set to '$theme'"
    } else {
        Log "🔄 Color theme already set to '$theme'" "Yellow"
        Add-Summary "Color theme already set to '$theme'"
    }
}

# Download and install VS Code
Log "Downloading VS Code installer..." "Cyan"
Add-Summary "VS Code installer download initiated"
if (-not $DryRun) {
    Invoke-WebRequest -Uri $vsCodeUrl -OutFile $installerPath
}

Log "Installing VS Code silently..." "Cyan"
Add-Summary "VS Code silent install initiated"
if (-not $DryRun) {
    Start-Process -FilePath $installerPath -ArgumentList "/VERYSILENT /MERGETASKS=!runcode" -Wait
}

# Add 'code' to PATH
$codeCmd = "$env:USERPROFILE\AppData\Local\Programs\Microsoft VS Code\bin"
$currentPath = [Environment]::GetEnvironmentVariable("Path", "User")

if ($currentPath -notlike "*$codeCmd*") {
    if (-not $DryRun) {
        [Environment]::SetEnvironmentVariable("Path", "$currentPath;$codeCmd", "User")
    }
    Log "✅ Added VS Code bin to PATH" "Green"
    Add-Summary "VS Code bin added to PATH"
} else {
    Log "🔄 VS Code bin already in PATH" "Yellow"
    Add-Summary "VS Code bin already in PATH"
}

# Install extensions
foreach ($ext in $extensions) {
    Log "Installing extension: $ext" "Magenta"
    Add-Summary "Extension processed: $ext"
    if (-not $DryRun) {
        & "$codeCmd\code.cmd" --install-extension $ext
    }
}

# Apply icon theme and affinity
Set-VSCodeIconTheme -theme "vscode-icons"
Set-VSCodeExtensionAffinities -config $extensionAffinityConfig
Set-VSCodeWindowBehavior
Set-VSCodeColorTheme -theme "Visual Studio Light"

# Launch VS Code for GitHub login
Log "Launching VS Code for GitHub authentication..." "Cyan"
Add-Summary "VS Code launched for GitHub login"
if (-not $DryRun) {
    Start-Process -FilePath "$codeCmd\code.cmd" -ArgumentList "--folder-uri vscode://github-authentication", "--command workbench.view.account"
}

# Optional: Open Copilot subscription page
if ($VerifyCopilot) {
    Log "Opening GitHub Copilot subscription page..." "Cyan"
    Add-Summary "GitHub Copilot subscription page opened"
    if (-not $DryRun) {
        Start-Process "https://github.com/settings/copilot"
    }
}

# Final summary
Log "=== Summary Report ===" "Cyan" -always
foreach ($entry in $summary) {
    Log "• $entry" "Gray" -always
}
Log "=== VS Code installation and GitHub setup complete ===" "Cyan" -always