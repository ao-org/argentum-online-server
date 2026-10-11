# Run the existing pure taming regressions with the real VB6 compiler, without
# starting the server, opening sockets, loading resources or touching a database.
param(
    [string]$Vb6 = 'C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE'
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$encoding = [Text.Encoding]::GetEncoding(1252)
$output = Join-Path ([IO.Path]::GetTempPath()) ('ao-taming-tests-' + [Guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($output) | Out-Null

function Read-Source([string]$relative) {
    [IO.File]::ReadAllText((Join-Path $root $relative), $encoding)
}
function Extract-Function([string]$source, [string]$name) {
    $pattern = '(?ms)^(?:Public |Private )?Function ' + [regex]::Escape($name) + '\(.*?^End Function\r?$'
    $matches = [regex]::Matches($source, $pattern)
    if ($matches.Count -ne 1) { throw "Expected one function: $name" }
    $matches[0].Value
}

# Extract verbatim production and regression bodies, not a second implementation.
$work = Read-Source 'Codigo\Trabajo.bas'
$tests = Read-Source 'Codigo\UnitTesting.bas'
$declares = Read-Source 'Codigo\Declares.bas'
$messages = Read-Source 'Codigo\modMessageIDs.bas'
$enum = [regex]::Match($declares, '(?ms)^Public Enum e_Class\r?$.*?^End Enum\r?$')
if (-not $enum.Success) { throw 'e_Class enum missing' }
$parts = @(
    'Attribute VB_Name = "TamingHarness"',
    'Option Explicit',
    'Private Total As Integer',
    'Private Failed As Integer',
    'Private Results As String',
    $enum.Value
)
foreach ($name in @('MSG_NO_PUEDES_DOMAR_ESA_CRIATURA', 'MSG_CRIATURA_HA_ACEPTADO_COMO_SU_AMO', 'MSG_NO_TAME_FAILED')) {
    $line = [regex]::Match($messages, '(?m)^Public Const ' + $name + '\s+As Integer = \d+\r?$')
    if (-not $line.Success) { throw "Message constant missing: $name" }
    $parts += $line.Value
}
foreach ($name in @('GetTamingPower', 'GetTamingAttemptMessage')) {
    $parts += Extract-Function $work $name
}
foreach ($name in @('test_suite_taming_pure', 'test_taming_power', 'test_taming_power_range', 'test_taming_results', 'test_taming_overflow')) {
    $parts += Extract-Function $tests $name
}
$parts += @'
Private Sub RunTest(ByVal TestName As String, ByVal Passed As Boolean)
    Total = Total + 1
    If Not Passed Then Failed = Failed + 1
    Results = Results & TestName & ": " & CStr(Passed) & vbCrLf
End Sub

Public Sub Main()
    On Error GoTo TestError
    Call test_suite_taming_pure
    GoTo Report
TestError:
    Failed = Failed + 1
    Results = Results & "Runtime error: " & Err.Number & " " & Err.Description & vbCrLf
Report:
    Dim File As Integer
    File = FreeFile
    Open App.Path & "\results.txt" For Output As #File
    Print #File, Results
    Print #File, "Total=" & Total & " Failed=" & Failed
    Close #File
End Sub
'@
[IO.File]::WriteAllText((Join-Path $output 'TamingHarness.bas'), (($parts -join "`r`n`r`n") -replace '\r?\n', "`r`n"), $encoding)
$project = @'
Type=Exe
Module=TamingHarness; TamingHarness.bas
Startup="Sub Main"
Name="TamingRegression"
ExeName32="TamingHarness.exe"
CompilationType=0
OptimizationType=0
'@
[IO.File]::WriteAllText((Join-Path $output 'TamingHarness.vbp'), ($project -replace '\r?\n', "`r`n"), $encoding)
Write-Output "Native VB6 test artifacts: $output"
$compiler = Start-Process -FilePath $Vb6 -WorkingDirectory $output -ArgumentList '/make TamingHarness.vbp /out build.log' -Wait -PassThru
$exe = Join-Path $output 'TamingHarness.exe'
if (Test-Path (Join-Path $output 'build.log')) { Get-Content (Join-Path $output 'build.log') }
if ($compiler.ExitCode -ne 0 -or -not (Test-Path $exe)) { throw 'VB6 test build failed' }
$run = Start-Process -FilePath $exe -WorkingDirectory $output -PassThru
if (-not $run.WaitForExit(30000)) { $run.Kill(); throw 'Taming test timed out' }
$result = [IO.File]::ReadAllText((Join-Path $output 'results.txt'), $encoding)
Write-Output $result
if ($result -notmatch 'Total=4 Failed=0') { throw 'Taming regression tests failed' }
