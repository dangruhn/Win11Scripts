Import-Module (Join-Path $PSScriptRoot '..\..\src\mutex.psm1') -Force -Scope Local

Describe 'Mutex module (src/mutex.psm1)' {
    It 'Get-NamedMutex returns a Mutex instance' {
        $name = "TestMutex_$([guid]::NewGuid().ToString())"
        $mutex = Get-NamedMutex -Name $name
        ($mutex -ne $null) | Should Be $true
        ($mutex.GetType().FullName -match 'Mutex') | Should Be $true
        Release-NamedMutex -Mutex $mutex
    }

    It 'Acquire-NamedMutex and Release-NamedMutex work without throwing' {
        $name = "TestMutex_$([guid]::NewGuid().ToString())"
        $mutex = Acquire-NamedMutex -Name $name
        ($mutex -ne $null) | Should Be $true
        { Release-NamedMutex -Mutex $mutex } | Should Not Throw
    }
}
