# Run in a development checkout, never in a deployed server directory.
# Builds the real server with a test-only entry point; no game startup or database is used.
param([string]$Vb6 = 'C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE')
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$log = Join-Path ([IO.Path]::GetTempPath()) ('ao-config-build-' + [Guid]::NewGuid().ToString('N') + '.log')
$results = Join-Path $root 'managed_config_test_results.txt'
$exe = Join-Path $root 'server.exe'
$started = Get-Date
$flags = 'UsarQueSocket=1:ConUpTime=1:AntiExternos=0:Lac=1:DEBUGGING=0:PYMMO=1:UNLOCK_CPU=1:DIRECT_PLAY=0:UNIT_TEST=1'
$compiler = Start-Process -FilePath $Vb6 -WorkingDirectory $root -ArgumentList "/make Server.VBP /d $flags /out `"$log`"" -Wait -PassThru
if (Test-Path $log) { Get-Content $log }
if ($compiler.ExitCode -ne 0 -or -not (Test-Path $exe) -or (Get-Item $exe).LastWriteTime -lt $started) {
    throw "Native VB6 build failed; see $log"
}
if (Test-Path $results) { Remove-Item $results }
$run = Start-Process -FilePath $exe -WorkingDirectory $root -ArgumentList '--test-managed-config' -PassThru
if (-not $run.WaitForExit(30000)) { $run.Kill(); throw 'Managed configuration tests timed out' }
if ($run.ExitCode -ne 0 -or -not (Test-Path $results)) { throw 'Managed configuration test entry point failed' }
$text = [IO.File]::ReadAllText($results)
Write-Output $text
if ($text -notmatch 'Checks=[1-9][0-9]* Failed=0' -or $text -match '(?m)^(FAIL|ERROR) ') { throw 'Managed configuration regression failed' }

# Exercise the real Main preflight and fatal exit, before any normal startup operation.
$missing = Join-Path ([IO.Path]::GetTempPath()) ('AO absent config ' + [Guid]::NewGuid().ToString('N'))
$run = Start-Process -FilePath $exe -WorkingDirectory $root -ArgumentList "--other=TEST /config-dir=`"$missing`"" -PassThru
if (-not $run.WaitForExit(10000)) { $run.Kill(); throw 'Invalid configuration startup did not fail promptly' }
if ($run.ExitCode -ne 1) { throw "Expected fatal startup exit 1, got $($run.ExitCode)" }
Write-Output 'PASS Main rejects a missing external directory with exit code 1'
Write-Output "Build log: $log"
Write-Output "Fixture results: $results"
