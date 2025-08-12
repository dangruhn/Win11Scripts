Add-Type -AssemblyName System.Windows.Forms

function Show-FolderBrowserWithOptions {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Empty Directory Cleanup"
    $form.Width = 420
    $form.Height = 220
    $form.StartPosition = "CenterScreen"

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Select the root directory:"
    $label.AutoSize = $true
    $label.Top = 20
    $label.Left = 20
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Width = 250
    $textBox.Top = 45
    $textBox.Left = 20
    $form.Controls.Add($textBox)

    $browseButton = New-Object System.Windows.Forms.Button
    $browseButton.Text = "Browse..."
    $browseButton.Top = 43
    $browseButton.Left = 280
    $browseButton.Add_Click({
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = "Choose root directory"
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $textBox.Text = $dialog.SelectedPath
        }
    })
    $form.Controls.Add($browseButton)

    $dryRunCheckbox = New-Object System.Windows.Forms.CheckBox
    $dryRunCheckbox.Text = "Dry Run (preview only)"
    $dryRunCheckbox.Top = 75
    $dryRunCheckbox.Left = 20
    $dryRunCheckbox.Width = 200
    $dryRunCheckbox.Height = 25
    $dryRunCheckbox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.Controls.Add($dryRunCheckbox)

    $includeHiddenCheckbox = New-Object System.Windows.Forms.CheckBox
    $includeHiddenCheckbox.Text = "Include hidden/system folders"
    $includeHiddenCheckbox.Top = 100
    $includeHiddenCheckbox.Left = 20
    $includeHiddenCheckbox.Width = 250
    $includeHiddenCheckbox.Height = 25
    $includeHiddenCheckbox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.Controls.Add($includeHiddenCheckbox)

    $ageLabel = New-Object System.Windows.Forms.Label
    $ageLabel.Text = "Minimum folder age (days):"
    $ageLabel.Top = 130
    $ageLabel.Left = 20
    $ageLabel.Width = 180
    $form.Controls.Add($ageLabel)

    $ageBox = New-Object System.Windows.Forms.NumericUpDown
    $ageBox.Top = 128
    $ageBox.Left = 200
    $ageBox.Width = 60
    $ageBox.Minimum = 0
    $ageBox.Maximum = 3650
    $ageBox.Value = 0
    $form.Controls.Add($ageBox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Top = 160
    $okButton.Left = 220
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Top = 160
    $cancelButton.Left = 300
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)

    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK -and $textBox.Text) {
        return @{
            Path          = $textBox.Text
            DryRun        = $dryRunCheckbox.Checked
            IncludeHidden = $includeHiddenCheckbox.Checked
            MinAgeDays    = [int]$ageBox.Value
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show(
            "No directory selected. Operation cancelled.",
            "Cancelled",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return $null
    }
}

function Show-ProgressWindow {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Cleaning Up..."
    $form.Width = 400
    $form.Height = 150
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.ControlBox = $false

    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Width = 350
    $progressBar.Height = 20
    $progressBar.Top = 20
    $progressBar.Left = 20
    $progressBar.Style = 'Continuous'
    $form.Controls.Add($progressBar)

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Width = 350
    $statusLabel.Top = 50
    $statusLabel.Left = 20
    $statusLabel.Text = "Starting..."
    $form.Controls.Add($statusLabel)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Top = 80
    $cancelButton.Left = 150
    $cancelButton.Width = 100
    $form.Controls.Add($cancelButton)

    $cancelled = $false
    $cancelButton.Add_Click({ $global:cancelled = $true })

    $form.Show()
    return @{
        Form      = $form
        Bar       = $progressBar
        Label     = $statusLabel
        Cancelled = { $global:cancelled }
    }
}

function Write-Log {
    param (
        [string]$Message,
        [string]$LogPath
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File -FilePath $LogPath -Append -Encoding UTF8
}

function Remove-EmptyDirectories {
    param (
        [string] $RootPath,
        [string] $LogPath,
        [bool]   $DryRun,
        [bool]   $IncludeHidden,
        [int]    $MinAgeDays
    )

    $removedCount = 0
    $failedCount  = 0
    $cutoffDate   = (Get-Date).AddDays(-$MinAgeDays)

    if (-not (Test-Path $RootPath)) {
        Write-Log "Invalid path: $RootPath" $LogPath
        [System.Windows.Forms.MessageBox]::Show(
            "Invalid path: $RootPath",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    $dirs = Get-ChildItem -Path $RootPath -Recurse -Directory -Force |
        Where-Object {
            ($_ | Get-ChildItem -Recurse -Force -File).Count -eq 0 -and
            ($IncludeHidden -or -not ($_.Attributes -match "Hidden|System")) -and
            ($_.LastWriteTime -lt $cutoffDate)
        }

    $progress = Show-ProgressWindow
    $progress.Bar.Maximum = $dirs.Count

    $i = 0
    foreach ($dir in $dirs) {
        if ($progress.Cancelled.Invoke()) {
            Write-Log "Cleanup cancelled by user." $LogPath
            $progress.Form.Close()
            [System.Windows.Forms.MessageBox]::Show(
                "Cleanup cancelled by user.",
                "Cancelled",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }

        $progress.Label.Text = "Processing: $($dir.Name)"
        $progress.Bar.Value   = ++$i
        $progress.Form.Refresh()

        if ($DryRun) {
            Write-Host "Dry Run: Would remove $($dir.FullName)" -ForegroundColor Cyan
            Write-Log "Dry Run: Would remove $($dir.FullName)" $LogPath
        } else {
            try {
                Remove-Item -Path $dir.FullName -Force -Recurse -ErrorAction Stop
                Write-Host "Removed: $($dir.FullName)" -ForegroundColor Green
                Write-Log "Removed: $($dir.FullName)" $LogPath
                $removedCount++
            } catch {
                Write-Host "Failed to remove: $($dir.FullName) — $_" -ForegroundColor Yellow
                Write-Log "Failed to remove: $($dir.FullName) — $_" $LogPath
                $failedCount++
            }
        }
    }

    $progress.Form.Close()

    $summary = if ($DryRun) {
        "Dry Run complete.`nTotal empty folders found: $($dirs.Count)"
    } else {
        "Cleanup complete.`nRemoved: $removedCount`nFailed: $failedCount"
    }

    Write-Log $summary $LogPath

    $view = [System.Windows.Forms.MessageBox]::Show(
        "$summary`n`nWould you like to view the log file?",
        "Summary",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )

    if ($view -eq [System.Windows.Forms.DialogResult]::Yes) {
        Start-Process -FilePath $LogPath
    }
}

# Main execution
$result = Show-FolderBrowserWithOptions
if ($result) {
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $logFile   = Join-Path $scriptDir "EmptyDirCleanup.log"

    Remove-EmptyDirectories `
        -RootPath      $result.Path `
        -LogPath       $logFile `
        -DryRun        $result.DryRun `
        -IncludeHidden $result.IncludeHidden `
        -MinAgeDays    $result.MinAgeDays
}