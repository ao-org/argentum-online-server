# Copyright (C) 2026 Noland Studios LTD
# Licensed under the GNU Affero General Public License, version 3 or later.
param(
    [Parameter(Mandatory = $true)][string]$Fixtures,
    [Parameter(Mandatory = $true)][string]$Maps,
    [string]$Compiler = 'C:\Program Files (x86)\Microsoft Visual Studio\VB98\vb6.exe'
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$outputDirectory = Join-Path $repoRoot 'build\map-schema-tests'
$fixtureDirectory = (Resolve-Path -LiteralPath $Fixtures).Path
$mapDirectory = (Resolve-Path -LiteralPath $Maps).Path
$previousCompatibility = $env:__COMPAT_LAYER
try {
    $env:__COMPAT_LAYER = 'RunAsInvoker'
    & python (Join-Path $PSScriptRoot 'test_map_schema.py') --output $outputDirectory
    if ($LASTEXITCODE -ne 0) { throw 'Could not generate the native test project.' }
    $compileLog = Join-Path $outputDirectory 'compile.log'
    # Use a fresh log: VB6 appends to an existing /out file.
    [System.IO.File]::WriteAllText($compileLog, '')
    $projectPath = Join-Path $outputDirectory 'tests.vbp'
    $compilerArguments = @('/make', ('"' + $projectPath + '"'), '/out', ('"' + $compileLog + '"'), '/outdir', ('"' + $outputDirectory + '"'))
    $buildProcess = Start-Process -FilePath $Compiler -ArgumentList $compilerArguments -WindowStyle Hidden -PassThru -Wait
    $buildResult = Get-Content -LiteralPath $compileLog -Raw
    if ($buildProcess.ExitCode -ne 0 -or $buildResult -notmatch "Build of 'map_schema_tests.exe' succeeded\.") {
        throw "Native test compilation failed: $buildResult"
    }
    $testArguments = '"' + $fixtureDirectory + '|' + $mapDirectory + '"'
    $testProcess = Start-Process -FilePath (Join-Path $outputDirectory 'map_schema_tests.exe') -ArgumentList $testArguments -WindowStyle Hidden -PassThru -Wait
    $results = Get-Content -LiteralPath (Join-Path $outputDirectory 'results.txt') -Raw
    if ($testProcess.ExitCode -ne 0 -or $results -notmatch '\d+ checks, 0 failures, 773 maps' -or $results -match 'FAIL:|FATAL:') {
        throw "Native map schema tests failed: $results"
    }
    ($results -split '\r?\n' | Where-Object { $_ -match 'checks, ' })
}
finally {
    $env:__COMPAT_LAYER = $previousCompatibility
}
