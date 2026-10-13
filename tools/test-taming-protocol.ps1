# Native VB6 tests using verbatim production routines, without server startup,
# sockets, game resources or a database. Requires registered Aurora.Network.
param([string]$Vb6 = 'C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE')
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$encoding = [Text.Encoding]::GetEncoding(1252)
$output = Join-Path ([IO.Path]::GetTempPath()) ('ao-taming-protocol-' + [Guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($output) | Out-Null
function Read-Source([string]$relative) {
    [IO.File]::ReadAllText((Join-Path $root $relative), $encoding)
}
function Extract([string]$source, [string]$kind, [string]$name) {
    $pattern = '(?ms)^(?:Public |Private )?' + $kind + ' ' + [regex]::Escape($name) + '\b.*?^End ' + $kind + '\r?$'
    $matches = [regex]::Matches($source, $pattern)
    if ($matches.Count -ne 1) { throw "Expected one $kind $name" }
    $matches[0].Value
}
function Write-Source([string]$name, [string]$text) {
    [IO.File]::WriteAllText((Join-Path $output $name), ($text -replace '\r?\n', "`r`n"), $encoding)
}
$work = Read-Source 'Codigo\Trabajo.bas'
$tests = Read-Source 'Codigo\UnitTesting.bas'
$declares = Read-Source 'Codigo\Declares.bas'
$messages = Read-Source 'Codigo\modMessageIDs.bas'
$parts = @('Attribute VB_Name = "TamingHarness"', 'Option Explicit',
    'Public Writer As Network.Writer', 'Public Captured() As Byte',
    'Public Sends As Integer', 'Public Const ToIndex As Integer = 1',
    'Private Total As Integer', 'Private Failed As Integer', 'Private Results As String',
    (Extract $declares 'Enum' 'e_Class'), (Extract $declares 'Enum' 'e_HooTamingResult'))
foreach ($name in @('MSG_NO_PUEDES_DOMAR_ESA_CRIATURA', 'MSG_CRIATURA_HA_ACEPTADO_COMO_SU_AMO', 'MSG_NO_TAME_FAILED')) {
    $line = [regex]::Match($messages, '(?m)^Public Const ' + $name + '\s+As Integer = \d+\r?$')
    if (-not $line.Success) { throw "Missing $name" }
    $parts += $line.Value
}
foreach ($name in @('GetTamingPower', 'GetTamingAttemptMessage', 'TamingPresentationResultForMessage')) {
    $parts += Extract $work 'Function' $name
}
foreach ($name in @('test_suite_taming_pure', 'test_taming_power', 'test_taming_power_range', 'test_taming_results', 'test_taming_overflow')) {
    $parts += Extract $tests 'Function' $name
}
$parts += Extract (Read-Source 'Codigo\Protocol_Writes.bas') 'Sub' 'WriteHooTamingResult'
$parts += @'
Public Sub TraceError(ByVal Number As Long, ByVal Description As String, ByVal Source As String, ByVal Line As Long)
    Err.Raise Number, Source, Description
End Sub
Private Sub RunTest(ByVal Name As String, ByVal Passed As Boolean)
    Total = Total + 1
    If Not Passed Then Failed = Failed + 1
    Results = Results & Name & ": " & CStr(Passed) & vbCrLf
End Sub
Private Function TestPacket() As Boolean
    Dim outcome As Integer, i As Integer, expected As Variant
    expected = Array(207, 0, 4, 3, 2, 1, 34, 0, 10, 20, 52, 18, 0)
    For outcome = 0 To 2
        Set Writer = New Network.Writer
        Call WriteHooTamingResult(1, &H1020304, 34, 10, 20, &H1234, outcome)
        If UBound(Captured) - LBound(Captured) + 1 <> 13 Then Exit Function
        For i = 0 To 11
            If Captured(LBound(Captured) + i) <> expected(i) Then Exit Function
        Next i
        If Captured(LBound(Captured) + 12) <> outcome Then Exit Function
    Next outcome
    If Sends <> 3 Then Exit Function
    Call WriteHooTamingResult(1, 9, 34, 10, 20, 0, eHooTamingResult_Success)
    If Captured(LBound(Captured) + 10) <> 0 Or Captured(LBound(Captured) + 11) <> 0 Then Exit Function
    Call WriteHooTamingResult(1, 9, 34, 10, 20, 0, eHooTamingResult_Unresolved)
    If Sends <> 4 Then Exit Function
    TestPacket = True
End Function
Public Sub Main()
    On Error GoTo FailedTest
    Call test_suite_taming_pure
    Call RunTest("Mandatory packet 207 bytes, outcomes, missing NPC and unresolved suppression", TestPacket())
    GoTo Report
FailedTest:
    Failed = Failed + 1
    Results = Results & "Error: " & Err.Number & " " & Err.Description & vbCrLf
Report:
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\results.txt" For Output As #f
    Print #f, Results
    Print #f, "Total=" & Total & " Failed=" & Failed
    Close #f
End Sub
'@
Write-Source 'TamingHarness.bas' ($parts -join "`r`n`r`n")
Write-Source 'PacketId.bas' (Read-Source 'Codigo\PacketId.bas')
Write-Source 'modSendData.bas' @'
Attribute VB_Name = "modSendData"
Option Explicit
Public Sub SendData(ByVal Route As Integer, ByVal UserIndex As Integer)
    If Route <> ToIndex Or UserIndex <> 1 Then Err.Raise 5
    Call Writer.GetData(Captured)
    Call Writer.Clear
    Sends = Sends + 1
End Sub
'@
$dll = Join-Path $root 'Aurora.Network.dll'
Write-Source 'TamingHarness.vbp' @"
Type=Exe
Reference=*\G{88D89F7C-BB59-4F13-9850-93CBD874039D}#1.5#0#$dll#Aurora.Network
Module=TamingHarness; TamingHarness.bas
Module=modSendData; modSendData.bas
Module=PacketId; PacketId.bas
Startup="Sub Main"
Name="TamingProtocolRegression"
ExeName32="TamingHarness.exe"
CondComp="PYMMO = 1"
CompilationType=0
OptimizationType=0
"@
Write-Output "Native VB6 artifacts: $output"
$compiler = Start-Process $Vb6 -WorkingDirectory $output -ArgumentList '/make TamingHarness.vbp /out build.log' -PassThru
if (-not $compiler.WaitForExit(120000)) { throw "Compiler still running: $($compiler.Id); inspect its dialog" }
if (Test-Path "$output\build.log") { Get-Content "$output\build.log" }
if (-not (Test-Path "$output\TamingHarness.exe")) { throw 'VB6 test build failed' }
$run = Start-Process "$output\TamingHarness.exe" -WorkingDirectory $output -PassThru
if (-not $run.WaitForExit(30000)) { throw "Test still running: $($run.Id)" }
$result = [IO.File]::ReadAllText("$output\results.txt", $encoding)
Write-Output $result
if ($result -notmatch 'Total=5 Failed=0') { throw 'Taming protocol regressions failed' }
