param(
    [string]$Compiler = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\VB98\vb6.exe"
)

$ErrorActionPreference = 'Stop'
$repo = (Resolve-Path (Join-Path $PSScriptRoot '../../..')).Path
$output = Join-Path $repo 'build/trigger-command-tests'
[void](New-Item -ItemType Directory -Force -Path $output)
$encoding = [Text.Encoding]::GetEncoding(1252)

function Read-Source([string]$relative) {
    [IO.File]::ReadAllText((Join-Path $repo $relative), $encoding)
}

function Extract-Block([string]$source, [string]$pattern) {
    $matches = [regex]::Matches($source, $pattern, 'Multiline,Singleline')
    if ($matches.Count -ne 1) { throw "Expected exactly one source block: $pattern" }
    $matches[0].Value
}

# Compile the actual handlers and enum declarations with isolated side effects.
# No server startup, network connection, database, or live map files are used.
$commands = Read-Source 'Codigo/Protocol_GmCommands.bas'
$declares = Read-Source 'Codigo/Declares.bas'
$blocks = @(
    'Attribute VB_Name = "TriggerUnderTest"',
    'Option Explicit',
    (Extract-Block $declares '^Public Enum e_PlayerType\r?\n.*?^End Enum'),
    (Extract-Block $declares '^Public Enum e_Trigger\r?\n.*?^End Enum'),
    (Extract-Block $commands '^Public Sub HandleSetTrigger\(.*?^End Sub'),
    (Extract-Block $commands '^Public Sub HandleAskTrigger\(.*?^End Sub')
)
[IO.File]::WriteAllText((Join-Path $output 'trigger_under_test.bas'), ($blocks -join "`r`n`r`n"), $encoding)
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'harness.bas') -Destination $output
$networkReference = (Read-Source 'Server.vbp') -split '\r?\n' |
    Where-Object { $_ -match '^Reference=.*Aurora\.Network\.dll#' }
if (@($networkReference).Count -ne 1) { throw 'Expected one Aurora.Network reference in Server.vbp' }
$networkReference = $networkReference.Replace('#Aurora.Network.dll#', '#..\..\Aurora.Network.dll#')
$project = @(
    'Type=Exe',
    $networkReference,
    'Module=TriggerUnderTest; trigger_under_test.bas',
    'Module=TriggerHarness; harness.bas',
    'Startup="Sub Main"',
    'Name="TriggerCommandTests"',
    'ExeName32="trigger_command_tests.exe"',
    'CompilationType=0'
)
[IO.File]::WriteAllText((Join-Path $output 'tests.vbp'), ($project -join "`r`n"), $encoding)
$previousCompatibility = $env:__COMPAT_LAYER
Push-Location $output
try {
    # VB6's installer-era elevation heuristic is unnecessary for compilation.
    # Run at the caller's existing privilege instead of requesting administrator access.
    $env:__COMPAT_LAYER = 'RunAsInvoker'
    [IO.File]::WriteAllText((Join-Path $output 'compile.log'), '', $encoding)
    $build = Start-Process -FilePath $Compiler -WorkingDirectory $output -WindowStyle Hidden -PassThru -Wait `
        -ArgumentList @('/make', 'tests.vbp', '/out', 'compile.log')
    $buildLog = [IO.File]::ReadAllText((Join-Path $output 'compile.log'), $encoding)
    Write-Output $buildLog
    if ($build.ExitCode -ne 0 -or $buildLog -notmatch 'succeeded') { throw 'VB6 trigger test compilation failed' }
    $test = Start-Process -FilePath (Join-Path $output 'trigger_command_tests.exe') `
        -WorkingDirectory $output -WindowStyle Hidden -PassThru -Wait
    Get-Content -LiteralPath (Join-Path $output 'results.txt')
    if ($test.ExitCode -ne 0) { throw 'Trigger command regression tests failed' }
} finally {
    $env:__COMPAT_LAYER = $previousCompatibility
    Pop-Location
}
