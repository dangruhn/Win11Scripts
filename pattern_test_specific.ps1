# Recreated test harness with robust parsing of the $eventTypes block
$handledownload = 'c:\Users\Dan.Gruhn\bin\handledownload.ps1'
$found = 'c:\Users\Dan.Gruhn\bin\found_formula_f1.txt'
$out = 'c:\Users\Dan.Gruhn\bin\pattern_test_specific_report.csv'

if (-not (Test-Path $handledownload)) { Write-Error "handledownload.ps1 not found: $handledownload"; exit 1 }
$lines = Get-Content $handledownload

# Extract circuits using a focused regex for their blocks
$circuits = @()
$cpattern = '(?ms)@\{\s*[^}]*?circuitRef\s*=\s*"(?<ref>.*?)"[^}]*?Pattern\s*=\s*"(?<pat>.*?)"[^}]*?CircuitName\s*=\s*"(?<name>.*?)"[^}]*?\}'
[regex]::Matches(($lines -join "`n"), $cpattern) | ForEach-Object {
    $circuits += [pscustomobject]@{
        circuitRef  = $_.Groups['ref'].Value
        Pattern     = $_.Groups['pat'].Value
        CircuitName = $_.Groups['name'].Value
    }
}

# Extract the $eventTypes array block by locating the start and the matching closing ')' on its own line
 $startIdxMatch = $lines | Select-String -Pattern '^[ \t]*\$eventTypes\s*=\s*@\(' | Select-Object -First 1
if (-not $startIdxMatch) { Write-Error "Couldn't find `'$eventTypes`' block"; exit 1 }
$startIdx = $startIdxMatch.LineNumber
if (-not $startIdx) { Write-Error "Couldn't find \$eventTypes block"; exit 1 }
$blockLines = @()
for ($i = $startIdx; $i -le $lines.Count; $i++) {
    $blockLines += $lines[$i-1]
    if ($lines[$i-1] -match '^[ \t]*\)[ \t]*$') { break }
}
$blockText = $blockLines -join "`n"
 

# Find each @ { ... } event block inside the eventTypes text by scanning lines (handles inner scriptblock braces)
$events = @()
$linesInBlock = $blockText -split "`n"
$i = 0
while ($i -lt $linesInBlock.Count) {
    $line = $linesInBlock[$i]
    if ($line -match '^[ \t]*@\{') {
        $bLines = @()
        do {
            $bLines += $linesInBlock[$i]
            $i++
        } while ($i -lt $linesInBlock.Count -and ($linesInBlock[$i] -notmatch '^[ \t]*\},?[ \t]*$'))
        if ($i -lt $linesInBlock.Count) { $bLines += $linesInBlock[$i]; $i++ }
        $b = ($bLines -join "`n")
        $p = [regex]::Match($b, 'Pattern\s*=\s"(?<pat>.*?)"', [System.Text.RegularExpressions.RegexOptions]::Singleline)
        $n = [regex]::Match($b, 'Name\s*=\s*"(?<name>.*?)"', [System.Text.RegularExpressions.RegexOptions]::Singleline)
        $nsb = [regex]::Match($b, 'Name\s*=\s*\{(?<sb>.*?)\}', [System.Text.RegularExpressions.RegexOptions]::Singleline)
        $events += [pscustomobject]@{
            Pattern         = if ($p.Success) { $p.Groups['pat'].Value } else { '' }
            Name            = if ($n.Success) { $n.Groups['name'].Value } else { $null }
            ScriptBlockText = if ($nsb.Success) { $nsb.Groups['sb'].Value } else { $null }
        }
    } else { $i++ }
}

# Load inputs
if (Test-Path $found) { $inputs = Get-Content $found | Where-Object { $_ -and $_.Trim() } } else { Write-Warning "$found not found; using sample"; $inputs = @("Sample.F1.2024.R01.Test.Race.mkv") }

$results = @()
 
foreach ($input in $inputs) {
    $cMatch = ''
    foreach ($c in $circuits) {
        $pat = ($c.Pattern) -replace ' ', '[\\s\\._-]+'
        try { if ($input -imatch $pat) { $cMatch = $c.CircuitName; break } } catch { }
    }

    $eMatch = ''
    foreach ($e in $events) {
        if (-not $e.Pattern) { continue }
        $epat = ($e.Pattern) -replace ' ', '[\\s\\._-]+'
        try {
            $m = [regex]::Match($input, $epat, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            if ($m.Success) {
                    if ($e.Name) {
                    $eMatch = $e.Name
                } elseif ($e.ScriptBlockText) {
                    $sbText = $e.ScriptBlockText.Trim()
                    try {
                        $sb = [scriptblock]::Create($sbText)
                        $res = & $sb $m
                        
                        if ($res -is [System.Array]) { $eMatch = ($res -join '') } else { $eMatch = $res }
                        if ($eMatch) { $eMatch = $eMatch.ToString().Trim() } else { $eMatch = $m.Value.Trim() }
                    } catch {
                        $eMatch = $m.Value.Trim()
                    }
                } else {
                    $eMatch = $m.Value.Trim()
                }
                break
            }
        } catch { }
    }

    # If still not matched, leave Event empty
    if (-not $eMatch) { $eMatch = '' }

    $results += [pscustomobject]@{ Input = $input; Circuit = $cMatch; Event = $eMatch }
}

$results | Export-Csv -Path $out -NoTypeInformation -Encoding UTF8
Write-Host "Wrote report to $out"
$results | Format-Table Input, Circuit, Event -AutoSize
