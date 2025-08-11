function Get-PowerShellInstallations {
    $installRoot = "C:\Program Files\PowerShell"
    $results = @()

    if (Test-Path $installRoot) {
        Get-ChildItem -Path $installRoot -Directory | ForEach-Object {
            $versionFolder = $_.Name
            $pwshPath = Join-Path $_.FullName "pwsh.exe"

            if (Test-Path $pwshPath) {
                $isPreview = $versionFolder -match "preview"
                $results += [PSCustomObject]@{
                    Version     = $versionFolder
                    Path        = $pwshPath
                    Type        = if ($isPreview) { "Preview" } else { "Stable" }
                    Installed   = $true
                }
            }
        }
    }

    # Also check if pwsh.exe is in PATH
    $pathPwsh = Get-Command pwsh -ErrorAction SilentlyContinue
    if ($pathPwsh) {
        $results += [PSCustomObject]@{
            Version     = "Unknown"
            Path        = $pathPwsh.Source
            Type        = "PATH Reference"
            Installed   = $true
        }
    }

    if ($results.Count -eq 0) {
        Write-Warning "No PowerShell 7 installations found."
    }

    return $results | Sort-Object Version
}

# Run it
Get-PowerShellInstallations | Format-Table -AutoSize