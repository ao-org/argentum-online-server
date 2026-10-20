# Isolated regression tests extracted from production VB6 and its registered suite.
# Default: native VB6 compilation. -ScriptHost: run the pure selection logic under
# Windows VBScript after removing VB6 type declarations/error handlers. This is
# useful without VB6 installed, but is NOT a server build or native VB6 validation.
param(
    [string]$Vb6 = 'C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE',
    [switch]$ScriptHost
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$output = Join-Path $root 'build\world-click-regression'
$encoding = [Text.Encoding]::GetEncoding(1252)
if (-not $ScriptHost -and -not (Test-Path $Vb6)) { throw 'VB6 compiler unavailable. Use -Vb6 with its installed path, or -ScriptHost for logic-only checks.' }
[IO.Directory]::CreateDirectory($output) | Out-Null
function Read-Source([string]$relative) { [IO.File]::ReadAllText((Join-Path $root $relative), $encoding) }
function Extract([string]$source, [string]$kind, [string]$name) {
    $pattern = '(?ms)^(?:Public |Private )?' + $kind + ' ' + [regex]::Escape($name) + '\b.*?^End ' + $kind + '\r?$'
    $matches = [regex]::Matches($source, $pattern)
    if ($matches.Count -ne 1) { throw "Expected one $kind $name" }
    $matches[0].Value
}
function Write-Source([string]$name, [string]$text) {
    [IO.File]::WriteAllText((Join-Path $output $name), ($text -replace '\r?\n', "`r`n"), $encoding)
}
$production = Read-Source 'Codigo\WorldActions.bas'
$tests = Read-Source 'Codigo\Tests\Unit_WorldClick.bas'
$enum = Extract $production 'Enum' 'e_WorldClickCharacter'
$functions = @(
    (Extract $production 'Function' 'ResolveWorldClickCharacter'),
    (Extract $tests 'Function' 'test_suite_world_click'),
    (Extract $tests 'Function' 'CheckWorldClick')
) -join "`r`n`r`n"
$functions = $functions.Replace('UnitTesting.RunTest', 'RunTest')
$runner = @'
Private Total As Integer
Private Failed As Integer
Private Results As String
Public Sub RunTest(ByVal Name As String, ByVal Passed As Boolean)
    Total = Total + 1
    If Not Passed Then Failed = Failed + 1
    Results = Results & Name & ": " & CStr(Passed) & vbCrLf
End Sub
'@
if ($ScriptHost) {
    $constants = [regex]::Matches($enum, '(?m)^\s*(\w+)\s*=\s*(\d+)\s*\r?$') | ForEach-Object {
        'Const ' + $_.Groups[1].Value + ' = ' + $_.Groups[2].Value
    }
    $script = "Option Explicit`r`n" + ($constants -join "`r`n") + "`r`n" + $runner + "`r`n" + $functions
    $script = $script -replace '(?m)^\s*On Error GoTo \w+\s*\r?$', ''
    $script = $script -replace '(?ms)^\w+_Err:.*?^End Function', 'End Function'
    $script = $script -replace ' As (Integer|Boolean|String|e_WorldClickCharacter)\b', ''
    $script = $script -replace '(?m)^(Public|Private) (Function|Sub) ', '$2 '
    $script = $script -replace '(?m)^Private ', 'Dim '
    $script += "`r`nTotal = 0`r`nFailed = 0`r`nResults = vbNullString`r`nCall test_suite_world_click()`r`nWScript.Echo Results`r`nWScript.Echo `"Logic-only VBScript: Total=`" & Total & `" Failed=`" & Failed`r`nIf Total <> 10 Or Failed <> 0 Then WScript.Quit 1`r`n"
    Write-Source 'WorldClickTests.vbs' $script
    & "$env:SystemRoot\System32\cscript.exe" //nologo (Join-Path $output 'WorldClickTests.vbs')
    if ($LASTEXITCODE -ne 0) { throw 'World click logic checks failed' }
    exit 0
}
$main = @'
Public Sub Main()
    Call test_suite_world_click()
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\results.txt" For Output As #f
    Print #f, Results
    Print #f, "Total=" & Total & " Failed=" & Failed
    Close #f
End Sub
Public Sub TraceError(ByVal Number As Long, ByVal Description As String, ByVal Source As String, ByVal Line As Long)
    Err.Raise Number, Source, Description
End Sub
'@
Write-Source 'WorldClickTests.bas' ("Attribute VB_Name = `"WorldClickTests`"`r`nOption Explicit`r`n" + $enum + "`r`n" + $runner + "`r`n" + $functions + "`r`n" + $main)
Write-Source 'WorldClickTests.vbp' @'
Type=Exe
Module=WorldClickTests; WorldClickTests.bas
Startup="Sub Main"
Name="WorldClickTests"
ExeName32="WorldClickTests.exe"
CompilationType=0
'@
Remove-Item (Join-Path $output 'WorldClickTests.exe'), (Join-Path $output 'results.txt'), (Join-Path $output 'build.log') -ErrorAction SilentlyContinue
$compiler = Start-Process $Vb6 -WorkingDirectory $output -ArgumentList '/make WorldClickTests.vbp /out build.log' -PassThru
if (-not $compiler.WaitForExit(60000)) { throw "Compiler still running: $($compiler.Id); inspect its dialog" }
if (Test-Path "$output\build.log") { Get-Content "$output\build.log" }
if (-not (Test-Path "$output\WorldClickTests.exe")) { throw 'VB6 test build failed' }
$run = Start-Process "$output\WorldClickTests.exe" -WorkingDirectory $output -PassThru
if (-not $run.WaitForExit(30000)) { throw "Test still running: $($run.Id)" }
$result = [IO.File]::ReadAllText("$output\results.txt", $encoding)
Write-Output $result
if ($result -notmatch 'Total=10 Failed=0') { throw 'World click regressions failed' }
