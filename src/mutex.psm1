<#
    Module: mutex
    Export helpers to acquire/release a named global mutex used by handledownload.ps1
#>

function Get-NamedMutex {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$Name
    )

    if (-not ("MutexLocker" -as [type])) {
        Add-Type -TypeDefinition @"
using System.Threading;
public class MutexLocker {
    public static Mutex GetMutex(string name) {
        return new Mutex(false, name);
    }
}
"@ -ErrorAction SilentlyContinue
    }

    return [MutexLocker]::GetMutex($Name)
}

function Acquire-NamedMutex {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$Name
    )

    $mutex = Get-NamedMutex -Name $Name
    if ($null -ne $mutex) {
        $null = $mutex.WaitOne()
    }
    return $mutex
}

function Release-NamedMutex {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][System.Threading.Mutex]$Mutex
    )

    if ($null -ne $Mutex) {
        try { $Mutex.ReleaseMutex() } catch { }
    }
}

Export-ModuleMember -Function Get-NamedMutex, Acquire-NamedMutex, Release-NamedMutex
