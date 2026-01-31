$txt = Get-Content 'c:\Users\Dan.Gruhn\bin\handledownload.ps1' -Raw
$blockMatch = [regex]::Match($txt, '(?ms)\$eventTypes\s*=\s*@\((.*?)\)')
if (-not $blockMatch.Success) { Write-Host 'eventTypes block not found'; exit }
$block = $blockMatch.Groups[1].Value
$matches = [regex]::Matches($block, '(?ms)@\{.*?\}')
Write-Host "Found event blocks: $($matches.Count)"
for ($i=0; $i -lt $matches.Count; $i++){
    $b = $matches[$i].Value
    $p = [regex]::Match($b, 'Pattern\s*=\s"(?<pat>.*?)"', [System.Text.RegularExpressions.RegexOptions]::Singleline)
    $s = [regex]::Match($b, 'Name\s*=\s*\{(?<sb>.*?)\}', [System.Text.RegularExpressions.RegexOptions]::Singleline)
    $n = [regex]::Match($b, 'Name\s*=\s*"(?<name>.*?)"', [System.Text.RegularExpressions.RegexOptions]::Singleline)
    if ($p.Success -and ($p.Groups['pat'].Value -match 'Shakedown' -or $s.Success)){
        Write-Host "--- Event #$i ---"
        Write-Host "Pattern: $($p.Groups['pat'].Value)"
        Write-Host "HasNameLiteral: $($n.Success)"
        if ($s.Success) { Write-Host "ScriptBlockText: $($s.Groups['sb'].Value)" } else { Write-Host "No ScriptBlock captured" }
    }
}
Write-Host 'Done'