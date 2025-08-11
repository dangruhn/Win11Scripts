param (
    [switch]$Cleanup,
    [switch]$Associate,
    [switch]$AddContextFile,
    [switch]$AddContextFolder,
    [switch]$All,
    [switch]$Help
)


if ($Help) {
    $helpText = @"
Setup-ConvertToOFX.ps1 — Configure .qfx file behavior and context menus

Available switches:
  -Cleanup           Remove existing .qfx associations and ProgIDs
  -Associate         Associate .qfx with ConvertToOFX.exe and msmoney.exe icon
  -AddContextFile    Add 'Convert to OFX Tool' to .qfx file context menu
  -AddContextFolder  Add 'Convert All .qfx in Folder' to folder context menu
  -All               Run all setup steps in sequence
  -Help              Show this help message

Examples:
  .\Setup-ConvertToOFX.ps1 -All
  .\Setup-ConvertToOFX.ps1 -Cleanup -AddContextFolder
"@

    # Console output with color
    Write-Host "`n=== Setup-ConvertToOFX.ps1 Help ===" -ForegroundColor Cyan
    Write-Host $helpText -ForegroundColor Gray

    # GUI fallback if launched non-interactively
    if (-not ($Host.Name -match "Console")) {
        Add-Type -AssemblyName PresentationFramework
        [System.Windows.MessageBox]::Show($helpText, "Setup-ConvertToOFX Help", 'OK', 'Information')
    }

    exit
}


$extension = ".qfx"
$progId = "ConvertToOFX.QFX"
$openExe = "C:\Users\Dan.Gruhn\bin\ConvertToOFX.exe"
$iconExe = "C:\Program Files (x86)\Microsoft Money Plus\MNYCoreFiles\msmoney.exe"
$menuLabelSingle = "Convert to OFX Tool"
$menuLabelFolder = "Convert All .qfx in Folder"

function Remove-QFXAssociations {
    $extPath = "Registry::HKCU\Software\Classes\$extension"
    if (Test-Path $extPath) {
        Remove-Item -Path $extPath -Recurse -Force
        Write-Host "🧹 Removed HKCU association for $extension"
    }

    $knownProgIds = @("ConvertToOFX.QFX", "Money.QFX")
    foreach ($progId in $knownProgIds) {
        $progIdPath = "Registry::HKCU\Software\Classes\$progId"
        if (Test-Path $progIdPath) {
            Remove-Item -Path $progIdPath -Recurse -Force
            Write-Host "🧹 Removed ProgID: $progId"
        }
    }
}

function Create-QFXAssociation {
    $progIdPath = "Registry::HKCU\Software\Classes\$progId"
    New-Item -Path $progIdPath -Force | Out-Null
    Set-ItemProperty -Path $progIdPath -Name "(default)" -Value "QFX File (ConvertToOFX)"
    New-Item -Path "$progIdPath\DefaultIcon" -Force | Out-Null
    Set-ItemProperty -Path "$progIdPath\DefaultIcon" -Name "(default)" -Value "`"$iconExe`",0"
    New-Item -Path "$progIdPath\shell\open\command" -Force | Out-Null
    Set-ItemProperty -Path "$progIdPath\shell\open\command" -Name "(default)" -Value "`"$openExe`" `"%1`""

    $extPath = "Registry::HKCU\Software\Classes\$extension"
    New-Item -Path $extPath -Force | Out-Null
    Set-ItemProperty -Path $extPath -Name "(default)" -Value $progId

    Write-Host "🔗 .qfx files now open with ConvertToOFX.exe and show msmoney.exe icon."
}

function Add-ContextMenu-SingleFile {
    $keyPath = "Registry::HKCU\Software\Classes\SystemFileAssociations\$extension\shell\ConvertToOFX"
    New-Item -Path $keyPath -Force | Out-Null
    Set-ItemProperty -Path $keyPath -Name "(default)" -Value $menuLabelSingle
    Set-ItemProperty -Path $keyPath -Name "Icon" -Value "`"$iconExe`",0"

    $cmdPath = "$keyPath\command"
    New-Item -Path $cmdPath -Force | Out-Null
    Set-ItemProperty -Path $cmdPath -Name "(default)" -Value "`"$openExe`" `"%1`""

    Write-Host "🖱️ Context menu 'Convert to OFX Tool' added for all .qfx files."
}

function Add-ContextMenu-FolderBulk {
    $folderKey = "Registry::HKCU\Software\Classes\Directory\shell\ConvertAllQFX"
    New-Item -Path $folderKey -Force | Out-Null
    Set-ItemProperty -Path $folderKey -Name "(default)" -Value $menuLabelFolder
    Set-ItemProperty -Path $folderKey -Name "Icon" -Value "`"$iconExe`",0"

    $folderCmd = "$folderKey\command"
    New-Item -Path $folderCmd -Force | Out-Null
    $bulkCommand = "`"$openExe`" `"%1\*.qfx`""
    Set-ItemProperty -Path $folderCmd -Name "(default)" -Value $bulkCommand

    Write-Host "📁 Context menu 'Convert All .qfx in Folder' added to folders."
}


function Show-GUISelector {
    Add-Type -AssemblyName PresentationFramework

    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Setup ConvertToOFX" Height="300" Width="400" WindowStartupLocation="CenterScreen">
    <StackPanel Margin="20">
        <TextBlock FontSize="16" FontWeight="Bold" Margin="0,0,0,10">Select Setup Actions:</TextBlock>
        <CheckBox Name="CleanupBox" Content="🧹 Remove existing .qfx associations" Margin="0,5"/>
        <CheckBox Name="AssociateBox" Content="🔗 Associate .qfx with ConvertToOFX.exe" Margin="0,5"/>
        <CheckBox Name="ContextFileBox" Content="🖱️ Add context menu for .qfx files" Margin="0,5"/>
        <CheckBox Name="ContextFolderBox" Content="📁 Add context menu for folders" Margin="0,5"/>
        <Button Name="RunButton" Content="Run Selected Actions" Margin="0,20,0,0" Padding="5"/>
    </StackPanel>
</Window>
"@

    try {
        $reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
        $window = [Windows.Markup.XamlReader]::Load($reader)

        $CleanupBox = $window.FindName("CleanupBox")
        $AssociateBox = $window.FindName("AssociateBox")
        $ContextFileBox = $window.FindName("ContextFileBox")
        $ContextFolderBox = $window.FindName("ContextFolderBox")
        $RunButton = $window.FindName("RunButton")

        $RunButton.Add_Click({
            if ($CleanupBox.IsChecked) { Remove-QFXAssociations }
            if ($AssociateBox.IsChecked) { Create-QFXAssociation }
            if ($ContextFileBox.IsChecked) { Add-ContextMenu-SingleFile }
            if ($ContextFolderBox.IsChecked) { Add-ContextMenu-FolderBulk }
            $window.Close()
        })

        $window.ShowDialog() | Out-Null
    } catch {
        Write-Warning "⚠️ Failed to load GUI selector. Falling back to CLI."
    }
}


if (-not ($PSBoundParameters.Count) -and -not $Help) {
    Show-GUISelector
    exit
}

# 🔄 Execute based on parameters
if ($All) {
    Remove-QFXAssociations
    Create-QFXAssociation
    Add-ContextMenu-SingleFile
    Add-ContextMenu-FolderBulk
    Write-Host "`n✅ Full setup complete."
} else {
    if ($Cleanup) { Remove-QFXAssociations }
    if ($Associate) { Create-QFXAssociation }
    if ($AddContextFile) { Add-ContextMenu-SingleFile }
    if ($AddContextFolder) { Add-ContextMenu-FolderBulk }
}

Write-Host "`n💡 Tip: Restart Explorer to see icon and context menu changes."