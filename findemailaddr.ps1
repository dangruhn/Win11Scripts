Add-Type -AssemblyName PresentationFramework

# Function to filter and display matching recipients
function Show-MatchingRecipients {
    param (
        [string]$Pattern
    )
    $resultsBox.Items.Clear()
    try {
        $matches = $data | Where-Object {
            $_.Recipient -match $Pattern
        }
        foreach ($entry in $matches) {
            $resultsBox.Items.Add($entry.Recipient)
        }
        if ($resultsBox.Items.Count -eq 0) {
            $resultsBox.Items.Add("No matches found.")
        }
    } catch {
        $resultsBox.Items.Add("Invalid regex pattern.")
    }
}


# Get the most recent duocircle CSV file based on timestamp
$folder = "C:\Users\Dan.Gruhn\email"
$latestFile = Get-ChildItem -Path $folder -Filter "duocircle_fwdtable_*.csv" |
    Where-Object { $_.Name -match "_\d{14}\.csv$" } |
    Sort-Object {
        [datetime]::ParseExact(
            ($_.Name -replace '^.*_(\d{14})\.csv$', '$1'),
            'yyyyMMddHHmmss',
            [System.Globalization.CultureInfo]::InvariantCulture
        )
    } -Descending | Select-Object -First 1

if (-not $latestFile) {
    [System.Windows.MessageBox]::Show("No matching CSV file found.")
    return
}

# Load CSV content
$data = Import-Csv $latestFile.FullName

# Create WPF window
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    Title="Find Email Address" Height="700" Width="800">
    <StackPanel Margin="10">
        <TextBlock Text="Enter email regex:" Margin="0,0,0,5"/>
        <TextBox Name="RegexBox" Height="25"/>
       <ListBox Name="ResultsBox" Height="580" Margin="0,10,0,0"/>
    </StackPanel>
</Window>
"@

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Find regex input box
$regexBox = $window.FindName("RegexBox")
$regexBox.Add_KeyDown({
    if ($_.Key -eq 'C' -and $_.KeyboardDevice.Modifiers -eq 'Control') {
        # Let TextBox handle Ctrl+C natively
        $_.Handled = $false
    }
})


# Set focus to the regex input box on load
$window.Add_Loaded({
    $regexBox.Focus()
})
# Find other controls
$resultsBox = $window.FindName("ResultsBox")

$regexBox.Add_KeyDown({
    if ($_.Key -eq 'Enter') {
        $pattern = $regexBox.Text
        Show-MatchingRecipients -Pattern $pattern

        if ($_.KeyboardDevice.Modifiers -eq 'Control') {
            $results = $resultsBox.Items | Where-Object { $_ -ne "No matches found." -and $_ -ne "Invalid regex pattern." }
            if ($results.Count -gt 0) {
                [System.Windows.Clipboard]::SetText($results -join "`r`n")
                [System.Windows.MessageBox]::Show("Copied to clipboard.")
            }
        }
    }
})

# Ctrl+C: Copy results to clipboard without re-submitting
$window.Add_KeyDown({
    if ($_.Key -eq 'C' -and $_.KeyboardDevice.Modifiers -eq 'Control') {
        if ($regexBox.IsFocused) {
            # Let TextBox handle Ctrl+C natively
            $_.Handled = $false
        } elseif ($resultsBox.IsKeyboardFocusWithin -and $resultsBox.SelectedItem) {
            # Copy only the selected item
            [System.Windows.Clipboard]::SetText($resultsBox.SelectedItem.ToString())
            [System.Windows.MessageBox]::Show("Copied selected result.")
        } else {
            # Copy all valid results
            $results = $resultsBox.Items | Where-Object {
                $_ -ne "No matches found." -and $_ -ne "Invalid regex pattern."
            }
            if ($results.Count -gt 0) {
                [System.Windows.Clipboard]::SetText($results -join "`r`n")
                [System.Windows.MessageBox]::Show("Copied all results.")
            } else {
                [System.Windows.MessageBox]::Show("No valid results to copy.")
            }
        }
    }
})


# Ctrl+S: Save results to a file
$window.Add_KeyDown({
    if ($_.Key -eq 'S' -and $_.KeyboardDevice.Modifiers -eq 'Control') {
        $results = $resultsBox.Items | Where-Object { $_ -ne "No matches found." -and $_ -ne "Invalid regex pattern." }
        if ($results.Count -gt 0) {
            Add-Type -AssemblyName PresentationFramework
            $dialog = New-Object Microsoft.Win32.SaveFileDialog
            $dialog.Title = "Save Email Addresses"
            $dialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
            $dialog.FileName = "email_filter_results.txt"
            if ($dialog.ShowDialog() -eq $true) {
                $results | Out-File -FilePath $dialog.FileName -Encoding UTF8
                [System.Windows.MessageBox]::Show("Saved to $($dialog.FileName)")
            }
        } else {
            [System.Windows.MessageBox]::Show("No valid results to save.")
        }
    }
})

# Ctrl+R: Reset the form
$window.Add_KeyDown({
    if ($_.Key -eq 'R' -and $_.KeyboardDevice.Modifiers -eq 'Control') {
        $regexBox.Clear()
        $resultsBox.Items.Clear()
        $regexBox.Focus()
    }
})

# F1: Show regex help tooltip
$regexBox.Add_KeyDown({
    if ($_.Key -eq 'F1') {
        $help = @"
Regex Quick Guide:
^      → Start of string
$      → End of string
.      → Any character
.*     → Zero or more of any character
\w+    → One or more word characters
[a-z]  → Character class
@      → Literal '@' symbol

Hotkeys:
Enter         → Submit regex
Ctrl+Enter    → Submit and copy results to clipboard
Ctrl+C        → Copy results to clipboard
Ctrl+S        → Save results to a file
Ctrl+R        → Reset form
F1            → Show this help dialog
"@
        [System.Windows.MessageBox]::Show($help, "Regex & Hotkey Help")
    }
})

$window.ShowDialog() | Out-Null