$file = 'c:\Users\Dan.Gruhn\bin\handledownload.ps1'
$lines = Get-Content $file
$patterns = @()
for ($i=0; $i -lt $lines.Count; $i++) {
    if ($lines[$i] -match 'Pattern\s*=\s*"(.*)"') {
        $pat = $matches[1]
        $name = ''
        for ($j=1; $j -le 6; $j++) {
            if ($i+$j -lt $lines.Count -and $lines[$i+$j] -match 'Name\s*=\s*"(.*)"') { $name = $matches[1]; break }
        }
        $patterns += [PSCustomObject]@{ Pattern = $pat; Name = $name }
    }
}

$sampleInputs = @(
    'Abu Dhabi Grand Prix',
    'Yas Marina Circuit',
    'FP1',
    'Free Practice 1',
    'FP2_Singapore',
    'Free-Practice-3',
    'Shakedown Day2',
    'Day 3 Shakedown',
    'shootout session',
    'sprint qualifying session',
    'pre race buildup',
    'race session',
    'grand prix',
    'chequered.flag',
    "Ted's Notebook",
    'qualifying notebook',
    'COTA',
    'Miami Grand Prix',
    'Las Vegas Motor Speedway',
    'Autodromo Jose Carlos Pace',

    '2024 Abu Dhabi Grand Prix FP1 2160p.mkv',
    '2023 Singapore Grand Prix - Full Race.mkv',
    '2022 Monza - Qualifying.notebook',
    'F1.Sprint.Qualifying.Singapore.2023.mkv',
    'Shakedown.Day1.Yas.Marina.mkv',
    'team.principals.press.conference.mp4',
    'drivers.press.conference.2024.mp4',
    'Ted.s.Qualifying.Notebook.1080p.nfo',
    'sprint.race.2021.mp4',
    'race session 2023 1080p.mkv',
    'chequered.flag.post.race.show.mkv',
    'COTA_Race_2020.mp4',
    'Miami_Grand_Prix_2024_Highlights.mp4',
    'Las.Vegas.Motor.Speedway.Race.mkv',
    'Autodromo_Jose_Carlos_Pace_2021.mkv',
    'pre-race-buildup.mp4',
    'post-sprint-analysis.mkv',
    'qualifying.notebook.2022.nfo'
)

$report = @()
foreach ($s in $sampleInputs) {
    $found = $false
    foreach ($p in $patterns) {
        try {
            $m = [regex]::Match($s, $p.Pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        } catch {
            continue
        }
        if ($m.Success) { 
            $report += [PSCustomObject]@{ Input = $s; MatchedName = $p.Name; MatchedValue = $m.Value; Pattern = $p.Pattern }
            $found = $true; break
        }
    }
    if (-not $found) { $report += [PSCustomObject]@{ Input = $s; MatchedName = ''; MatchedValue = ''; Pattern = '' } }
}

$csvPath = 'c:\Users\Dan.Gruhn\bin\pattern_report.csv'
$report | Export-Csv -NoTypeInformation -Path $csvPath -Encoding UTF8
Write-Host "Wrote report to $csvPath"
$report | Format-Table -AutoSize
