Get-Content 'c:\Users\Dan.Gruhn\bin\handledownload.ps1' | Select-String 'Shakedown' -AllMatches | ForEach-Object { Write-Host "Line $($_.LineNumber): $($_.Line.Trim())" }
