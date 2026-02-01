Import-Module (Join-Path $PSScriptRoot '..\..\src\handledownload.psm1') -Force -Scope Local

# Provide global stubs for dependencies the f1 module expects
function global:Get-SuccessivePathLeaves { param([string]$Pathname) return @((Split-Path -Path $Pathname -Leaf), (Split-Path -Path $Pathname -Parent)) }
function global:LogOutput { param($Args) }
function global:Copy-WithProgress { param([string]$SourcePath,[string]$DestinationPath) Copy-Item -Path $SourcePath -Destination $DestinationPath -Force }

Import-Module (Join-Path $PSScriptRoot '..\..\src\f1.psm1') -Force -Scope Local

Describe 'F1 module (src/f1.psm1)' {

    It 'IsF1Module recognizes F1-like names' {
        (IsF1Module -TestStr 'some.f1.file') | Should Be $true
        (IsF1Module -TestStr 'formula1.something') | Should Be $true
        (IsF1Module -TestStr 'notanf1file') | Should Be $false
    }

    It 'Get-EventInfoF1Module decodes a simple circuit and race' {
        $tempFolder = Join-Path $env:TEMP "TestCircuit_f1test_$([guid]::NewGuid())"
        New-Item -Path $tempFolder -ItemType Directory | Out-Null
        $testFile = Join-Path $tempFolder 'racefile.mkv'
        New-Item -Path $testFile -ItemType File | Out-Null

        # Provide a cached resolution to avoid calling filebot
        $script:ResolutionBitsCache = @{}
        $script:ResolutionBitsCache[$testFile] = '2160p - 10b'

        $formula1Circuits = @(@{ circuitRef='test_ref'; Pattern='TestCircuit' ; CircuitName='TestCircuit' })
        $F1Circuits = @(@{ circuitId='test_ref'; circuitName='TestCircuit'; Location=@{ locality='TestLocal'; country='TestCountry' } })
        $F1Races = @(@{ season = [int](Get-Date).Year; date = (Get-Date).ToString('yyyy-MM-dd'); round = 1; Circuit = @{ circuitId='test_ref' } })
        $eventTypes = @(@{ Pattern = 'Race'; Name = 'Race' })

        $mockFilebot = (Join-Path $PSScriptRoot '..\helpers\mock-filebot.ps1')
        $result = Get-EventInfoF1Module -SrcFolder $tempFolder -SrcPathname $testFile -Formula1Circuits $formula1Circuits -F1Circuits $F1Circuits -F1Races $F1Races -EventTypes $eventTypes -FileBotExe $mockFilebot

        $result.RaceDate | Should Not Be 'xxxx-xx-xx'
        $result.CircuitName | Should Be 'TestCircuit'
        $result.ResolutionBits | Should Match '2160p'

        Remove-Item -Path $tempFolder -Recurse -Force
    }

    It 'CopyF1FileModule copies source to destination (stub Copy-WithProgress)' {
        $src = Join-Path $env:TEMP "f1src_$([guid]::NewGuid()).mkv"
        $dstRoot = Join-Path $env:TEMP "f1dest_$([guid]::NewGuid())"
        New-Item -Path $src -ItemType File | Out-Null

        # stub UI copy function to a simple Copy-Item
        function Copy-WithProgress { param($s,$d) Copy-Item -Path $s -Destination $d -Force }

        $eventInfo = [pscustomobject]@{ RaceDate = (Get-Date).ToString('yyyy-MM-dd'); CircuitName = 'UnitCircuit'; EventName = 'Race'; ResolutionBits = '1080p - 8b' }

        CopyF1FileModule -SrcPath $src -EventInfo $eventInfo -Suffix 'mkv' -F1DestRoot $dstRoot

        $expectedDir = Join-Path $dstRoot ("$($eventInfo.RaceDate) Formula1 $($eventInfo.CircuitName)")
        $files = Get-ChildItem -Path $expectedDir -Recurse -ErrorAction SilentlyContinue
        ($files.Count -gt 0) | Should Be $true

        Remove-Item -Path $dstRoot -Recurse -Force
        Remove-Item -Path $src -Force
    }

}
