Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Shared state
$paused = [ref]$false
$filterRegex = [ref]$null
$tailJob = $null

# GUI setup
$form = New-Object System.Windows.Forms.Form
$form.Text = "Tail Viewer"
$form.Size = New-Object System.Drawing.Size(800, 680)
$form.StartPosition = "CenterScreen"

$filePathBox = New-Object System.Windows.Forms.TextBox
$filePathBox.Location = New-Object System.Drawing.Point(10, 10)
$filePathBox.Size = New-Object System.Drawing.Size(600, 20)
$form.Controls.Add($filePathBox)

$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "Browse"
$browseButton.Location = New-Object System.Drawing.Point(620, 10)
$browseButton.Size = New-Object System.Drawing.Size(75, 20)
$form.Controls.Add($browseButton)

$pauseButton = New-Object System.Windows.Forms.Button
$pauseButton.Text = "Pause"
$pauseButton.Location = New-Object System.Drawing.Point(700, 10)
$pauseButton.Size = New-Object System.Drawing.Size(75, 20)
$form.Controls.Add($pauseButton)

$regexBox = New-Object System.Windows.Forms.TextBox
$regexBox.Location = New-Object System.Drawing.Point(10, 40)
$regexBox.Size = New-Object System.Drawing.Size(600, 20)
$form.Controls.Add($regexBox)

$applyButton = New-Object System.Windows.Forms.Button
$applyButton.Text = "Apply"
$applyButton.Location = New-Object System.Drawing.Point(620, 40)
$applyButton.Size = New-Object System.Drawing.Size(75, 20)
$applyButton.Add_Click({
    $filterRegex.Value = $regexBox.Text
    $statusBar.Text = "Status: Mode = $($modeDropdown.SelectedItem), Filter = $($filterRegex.Value)"
})
$form.Controls.Add($applyButton)

$clearButton = New-Object System.Windows.Forms.Button
$clearButton.Text = "Clear"
$clearButton.Location = New-Object System.Drawing.Point(700, 40)
$clearButton.Size = New-Object System.Drawing.Size(75, 20)
$clearButton.Add_Click({
    $filterRegex.Value = $null
    $regexBox.Text = ""
    $statusBar.Text = "Status: Mode = $($modeDropdown.SelectedItem), Filter = None"
})
$form.Controls.Add($clearButton)

$modeDropdown = New-Object System.Windows.Forms.ComboBox
$modeDropdown.Location = New-Object System.Drawing.Point(10, 600)
$modeDropdown.Size = New-Object System.Drawing.Size(120, 20)
$modeDropdown.Anchor = "Bottom, Left"
$modeDropdown.Items.AddRange(@("Full file", "Last N lines"))
$modeDropdown.SelectedIndex = 0
$form.Controls.Add($modeDropdown)

$nLinesBox = New-Object System.Windows.Forms.NumericUpDown
$nLinesBox.Location = New-Object System.Drawing.Point(140, 600)
$nLinesBox.Anchor = "Bottom, Left"
$nLinesBox.Size = New-Object System.Drawing.Size(60, 20)
$nLinesBox.Minimum = 1
$nLinesBox.Maximum = 10000
$nLinesBox.Value = 100
$form.Controls.Add($nLinesBox)

$restartButton = New-Object System.Windows.Forms.Button
$restartButton.Text = "Restart View"
$restartButton.Location = New-Object System.Drawing.Point(210, 600)
$restartButton.Anchor = "Bottom, Left"
$restartButton.Size = New-Object System.Drawing.Size(100, 20)
$form.Controls.Add($restartButton)

$restartButton.Add_Click({
    $outputBox.Clear()

    if ($tailJob -and $tailJob.State -eq 'Running') {
        Stop-Job $tailJob.Id
        Remove-Job $tailJob.Id
        $tailJob = $null
    }

    $filePath = $filePathBox.Text
    if (-not (Test-Path $filePath)) {
        $outputBox.AppendText("No valid file selected.`r`n")
        return
    }

    $mode = $modeDropdown.SelectedItem
    $nLines = [int]$nLinesBox.Value

    if ($mode -eq "Last N lines") {
        $lines = Get-Content -Tail $nLines -Path $filePath
    } else {
        $lines = Get-Content -Path $filePath
    }

    foreach ($line in $lines) {
        if (-not $filterRegex.Value -or $line -match $filterRegex.Value) {
            $outputBox.AppendText("$line`r`n")
        }
    }

    $statusBar.Text = "Status: Mode = $mode, Filter = " + ($filterRegex.Value ?? "None")

    $tailJob = Start-ThreadJob -ScriptBlock {
        param($filePath, $paused, $filterRegex, $outputBox)

        try {
            $fs = [System.IO.File]::Open($filePath, 'Open', 'Read', 'ReadWrite')
            $reader = New-Object System.IO.StreamReader($fs)
            $fs.Seek(0, [System.IO.SeekOrigin]::End) | Out-Null

            while ($true) {
                if (-not $paused.Value) {
                    $line = $reader.ReadLine()
                    if ($line) {
                        if (-not $filterRegex.Value -or $line -match $filterRegex.Value) {
                            $outputBox.Invoke([Action]{ $outputBox.AppendText("$line`r`n") })
                        }
                    } else {
                        Start-Sleep -Milliseconds 200
                    }
                } else {
                    Start-Sleep -Milliseconds 500
                }
            }
        } catch {
            $outputBox.Invoke([Action]{ $outputBox.AppendText("Error: $_`r`n") })
        }
    } -ArgumentList $filePath, $paused, $filterRegex, $outputBox
})

$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Multiline = $true
$outputBox.ScrollBars = "Vertical"
$outputBox.Location = New-Object System.Drawing.Point(10, 70)
$outputBox.Size = New-Object System.Drawing.Size(766, 500)
$outputBox.ReadOnly = $true
$outputBox.Anchor = "Top, Left, Right"
$form.Controls.Add($outputBox)

$statusBar = New-Object System.Windows.Forms.Label
$statusBar.Location = New-Object System.Drawing.Point(10, 580)
$statusBar.Anchor = "Bottom, Left"
$statusBar.Size = New-Object System.Drawing.Size(765, 20)
$statusBar.Text = "Status: Ready"
$form.Controls.Add($statusBar)

# Browse button logic
$browseButton.Add_Click({
    if ($tailJob -and $tailJob.State -eq 'Running') {
        Stop-Job $tailJob.Id
        Remove-Job $tailJob.Id
        $tailJob = $null
    }

    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "Log files (*.log;*.txt)|*.log;*.txt|All files (*.*)|*.*"
    if ($dialog.ShowDialog() -eq "OK") {
        $filePathBox.Text = $dialog.FileName

        # Start tailing thread
        Start-ThreadJob -ScriptBlock {
            param($filePath, $paused, $filterRegex, $outputBox)

            try {
                $fs = [System.IO.File]::Open($filePath, 'Open', 'Read', 'ReadWrite')
                $reader = New-Object System.IO.StreamReader($fs)
                $fs.Seek(0, [System.IO.SeekOrigin]::End) | Out-Null

                while ($true) {
                    if (-not $paused.Value) {
                        $line = $reader.ReadLine()
                        if ($line) {
                            if (-not $filterRegex.Value -or $line -match $filterRegex.Value) {
                                $outputBox.Invoke([Action]{ $outputBox.AppendText("$line`r`n") })
                            }
                        } else {
                            Start-Sleep -Milliseconds 200
                        }
                    } else {
                        Start-Sleep -Milliseconds 500
                    }
                }
            } catch {
                $outputBox.Invoke([Action]{ $outputBox.AppendText("Error: $_`r`n") })
            }
        } -ArgumentList $filePathBox.Text, $paused, $filterRegex, $outputBox
    }
})

# Pause button logic
$pauseButton.Add_Click({
    $paused.Value = -not $paused.Value
    $pauseButton.Text = if ($paused.Value) { "Resume" } else { "Pause" }
})

# Apply regex filter
$applyButton.Add_Click({
    $filterRegex.Value = $regexBox.Text
})

# Clear regex filter
$clearButton.Add_Click({
    $filterRegex.Value = $null
    $regexBox.Text = ""
})

# Run the form
[void]$form.ShowDialog()