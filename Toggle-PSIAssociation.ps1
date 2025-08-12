param (
    [switch]$Silent
)

Add-Type -AssemblyName PresentationFramework

# 🔍 Detect PowerShell 7 path
function Get-PowerShell7Path {
    $installRoot = "C:\Program Files\PowerShell"
    $candidates = @()

    if (Test-Path $installRoot) {
        $candidates = Get-ChildItem -Path $installRoot -Directory |
            Where-Object { Test-Path (Join-Path $_.FullName "pwsh.exe") } |
            Sort-Object Name -Descending
    }

    foreach ($folder in $candidates) {
        if ($folder.Name -notmatch "preview") {
            return Join-Path $folder.FullName "pwsh.exe"
        }
    }

    # Fallback to preview if no stable found
    if ($candidates.Count -gt 0) {
        return Join-Path $candidates[0].FullName "pwsh.exe"
    }

    # Fallback to PATH
    $cmd = Get-Command pwsh -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }

    return $null
}

# Paths
$ps7Path = Get-PowerShell7Path
if (-not $ps7Path) {
    Write-Warning "⚠️ PowerShell 7 executable not found. Registry association may fail."
    $ps7Path = "pwsh.exe"  # fallback to generic
}

Add-Type -AssemblyName PresentationFramework

# Paths
$vsCodePath = "C:\Users\$env:USERNAME\AppData\Local\Programs\Microsoft VS Code\Code.exe"
$extension = ".ps1"
$ps7ProgId = "PowerShellScript.7"
$vsCodeProgId = "VSCode.ps1"
$extKey = "Registry::HKCU\Software\Classes\$extension"

function Get-CurrentAssociation {
    if (Test-Path $extKey) {
        $current = (Get-ItemProperty -Path $extKey -Name "(default)" -ErrorAction SilentlyContinue)."(default)"
        return $current
    }
    return $null
}

function Set-Association-PowerShell7 {
    $progIdKey = "Registry::HKCU\Software\Classes\$ps7ProgId"

    New-Item -Path $extKey -Force | Out-Null
    Set-ItemProperty -Path $extKey -Name "(default)" -Value $ps7ProgId

    New-Item -Path $progIdKey -Force | Out-Null
    Set-ItemProperty -Path $progIdKey -Name "(default)" -Value "PowerShell Script (pwsh.exe)"
    New-Item -Path "$progIdKey\DefaultIcon" -Force | Out-Null
    Set-ItemProperty -Path "$progIdKey\DefaultIcon" -Name "(default)" -Value "`"$ps7Path`",0"
    New-Item -Path "$progIdKey\shell\open\command" -Force | Out-Null
    Set-ItemProperty -Path "$progIdKey\shell\open\command" -Name "(default)" -Value "`"$ps7Path`" -NoLogo -ExecutionPolicy Bypass -File `"%1`""

    Write-Host "✅ .ps1 files now run with PowerShell 7."
}

function Set-Association-VSCode {
    $progIdKey = "Registry::HKCU\Software\Classes\$vsCodeProgId"

    New-Item -Path $extKey -Force | Out-Null
    Set-ItemProperty -Path $extKey -Name "(default)" -Value $vsCodeProgId

    New-Item -Path $progIdKey -Force | Out-Null
    Set-ItemProperty -Path $progIdKey -Name "(default)" -Value "PowerShell Script (VS Code)"
    New-Item -Path "$progIdKey\DefaultIcon" -Force | Out-Null
    Set-ItemProperty -Path "$progIdKey\DefaultIcon" -Name "(default)" -Value "`"$vsCodePath`",0"
    New-Item -Path "$progIdKey\shell\open\command" -Force | Out-Null
    Set-ItemProperty -Path "$progIdKey\shell\open\command" -Name "(default)" -Value "`"$vsCodePath`" `"%1`""

    Write-Host "✅ .ps1 files now open in VS Code."
}

function Show-GUISelector {
    $current = Get-CurrentAssociation
    $isPS7 = $current -eq $ps7ProgId
    $isVSCode = $current -eq $vsCodeProgId

    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Toggle .ps1 Association" Height="200" Width="360" WindowStartupLocation="CenterScreen">
    <StackPanel Margin="20">
        <TextBlock FontSize="14" FontWeight="Bold" Margin="0,0,0,10">Choose default action for .ps1 files:</TextBlock>
        <RadioButton Name="PS7Radio" Content="Run with PowerShell 7" Margin="0,5"/>
        <RadioButton Name="VSCodeRadio" Content="Open in VS Code" Margin="0,5"/>
        <Button Name="ApplyButton" Content="Apply Selection" Margin="0,20,0,0" Padding="5"/>
    </StackPanel>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)

    $ps7Radio = $window.FindName("PS7Radio")
    $vsCodeRadio = $window.FindName("VSCodeRadio")
    $applyButton = $window.FindName("ApplyButton")

    if ($isPS7) { $ps7Radio.IsChecked = $true }
    elseif ($isVSCode) { $vsCodeRadio.IsChecked = $true }

    $applyButton.Add_Click({
        if ($ps7Radio.IsChecked) { Set-Association-PowerShell7 }
        elseif ($vsCodeRadio.IsChecked) { Set-Association-VSCode }
        $window.Close()
    })

    $window.ShowDialog() | Out-Null
}

# 🔄 Silent mode logic
if ($Silent) {
    $current = Get-CurrentAssociation
    if ($current -eq $ps7ProgId) {
        Set-Association-VSCode
    } else {
        Set-Association-PowerShell7
    }
    Write-Host "`n🔁 Silent toggle complete. Current association: $(Get-CurrentAssociation)"
} else {
    Show-GUISelector
    Write-Host "`n💡 Tip: Restart Explorer or log out/in to apply changes."
}