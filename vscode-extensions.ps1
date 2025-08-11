$installedExtensions = code --list-extensions | Sort-Object
Write-Host "Quoted strings for use in your script"
$installedExtensions | ForEach-Object { "`"$($_)`"" }
