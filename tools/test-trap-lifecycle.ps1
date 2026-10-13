# Compile the real trap class and verbatim production routines with deterministic
# game/transport fixtures. No server startup, resource loading, sockets or DB.
param([string]$Vb6 = 'C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE')
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$encoding = [Text.Encoding]::GetEncoding(1252)
$output = Join-Path ([IO.Path]::GetTempPath()) ('ao-trap-lifecycle-' + [Guid]::NewGuid().ToString('N'))
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
$declares = Read-Source 'Codigo\Declares.bas'
$definitions = @('Attribute VB_Name = "TrapDefinitions"', 'Option Explicit')
foreach ($name in @('e_ReferenceType','e_EffectOverTimeType','e_EffectType','e_TrapEffect','e_TrapEvent','e_OBJType','e_UssableOnTarget','e_TipoUsuario')) {
    $definitions += Extract $declares 'Enum' $name
}
foreach ($name in @('t_WorldPos','t_AnyReference','t_BaseDotInfo','t_EffectOverTimeList')) {
    $definitions += Extract $declares 'Type' $name
}
$definitions += Extract (Read-Source 'Codigo\CustomScenarios.bas') 'Enum' 'e_DamageSourceType'
foreach ($name in @('e_TextChannel','e_FontTypeNames')) {
    $definitions += Extract (Read-Source 'Codigo\Protocol.bas') 'Enum' $name
}
$definitions += Extract $declares 'Sub' 'SetBaseDot'
Write-Source 'TrapDefinitions.bas' ($definitions -join "`r`n`r`n")
$routines = @('Attribute VB_Name = "TrapProduction"', 'Option Explicit')
$inventory = Read-Source 'Codigo\InvUsuario.bas'
$routines += Extract $inventory 'Function' 'ValidTargetableItem'
foreach ($name in @('UserTargetableItem','PlaceTrap')) { $routines += Extract $inventory 'Sub' $name }
$protocol = Read-Source 'Codigo\Protocol_Writes.bas'
$routines += Extract $protocol 'Function' 'PrepareTrapUpdate'
$routines += Extract $protocol 'Sub' 'WriteUpdateTrapState'
$routines += Extract (Read-Source 'Codigo\ModMap.bas') 'Function' 'CanAddTrapAt'
$routines += Extract (Read-Source 'Codigo\ModMap.bas') 'Sub' 'ActivateTrap'
Write-Source 'TrapProduction.bas' ($routines -join "`r`n`r`n")
$effects = Read-Source 'Codigo\EffectsOverTime.bas'
$parts = @(Read-Source 'Codigo\Tests\TrapLifecycle\EffectsOverTime.bas')
foreach ($name in @('GetNextId','IsTrapEffectDefinition','CreateTrap','UpdateEffect')) { $parts += Extract $effects 'Function' $name }
foreach ($name in @('AddEffectToUpdate','AddEffect','RemoveEffect','RemoveEffectAtPos','ClearEffectList')) { $parts += Extract $effects 'Sub' $name }
Write-Source 'EffectsOverTime.bas' ($parts -join "`r`n`r`n")
foreach ($file in @('Fixtures.bas','Unit_TrapLifecycle.bas','modSendData.bas','NPCs.bas','UserMod.bas','TrapConfig.cls')) {
    Write-Source $file (Read-Source "Codigo\Tests\TrapLifecycle\$file")
}
foreach ($file in @('clsTrap.cls','IBaseEffectOverTime.cls','PacketId.bas')) {
    Write-Source $file (Read-Source "Codigo\$file")
}
# Guard the call sites which are not linked into this isolated harness.
$area = Read-Source 'Codigo\ModAreas.bas'
if ($area -notmatch 'MapData\(x, y, Map\)\.Trap\.SendSnapshot\(UserIndex\)') { throw 'Area snapshot must use the existing trap instance' }
$dispatch = Read-Source 'Codigo\Protocol.bas'
$target = [regex]::Match($dispatch, '(?s)Case e_Skill.TargetableItem(.*?)Call UserTargetableItem\(UserIndex, x, y\)').Value
if ($target -notmatch 'If Not ValidTargetableItem\(UserIndex\)' -or $target -match 'TargetObjInvSlot') { throw 'Unsafe delayed item slot dispatch' }
if (($inventory -split '\.flags.UsingItemObjIndex = ObjIndex').Count -ne 3) { throw 'Both item target requests must capture the item definition' }
$dll = Join-Path $root 'Aurora.Network.dll'
$project = @"
Type=Exe
Reference=*\G{88D89F7C-BB59-4F13-9850-93CBD874039D}#1.5#0#$dll#Aurora.Network
Module=TrapDefinitions; TrapDefinitions.bas
Module=TrapProduction; TrapProduction.bas
Module=EffectsOverTime; EffectsOverTime.bas
Module=Fixtures; Fixtures.bas
Module=Unit_TrapLifecycle; Unit_TrapLifecycle.bas
Module=modSendData; modSendData.bas
Module=NPCs; NPCs.bas
Module=UserMod; UserMod.bas
Module=PacketId; PacketId.bas
Class=clsTrap; clsTrap.cls
Class=IBaseEffectOverTime; IBaseEffectOverTime.cls
Class=TrapConfig; TrapConfig.cls
Startup="Sub Main"
Name="TrapTests"
ExeName32="TrapLifecycle.exe"
CompilationType=0
OptimizationType=0
"@
Write-Source 'TrapLifecycle.vbp' $project
Write-Output "Native VB6 artifacts: $output"
foreach ($pymmo in 0,1) {
    $exe = Join-Path $output 'TrapLifecycle.exe'
    if (Test-Path $exe) { Remove-Item $exe }
    $compiler = Start-Process $Vb6 -WorkingDirectory $output -ArgumentList "/make TrapLifecycle.vbp /d PYMMO=$pymmo /out build-$pymmo.log" -PassThru
    if (-not $compiler.WaitForExit(120000)) { throw "Compiler still running: $($compiler.Id); inspect its dialog" }
    if (Test-Path "$output\build-$pymmo.log") { Get-Content "$output\build-$pymmo.log" }
    if (-not (Test-Path $exe)) { throw "VB6 test build failed (PYMMO=$pymmo)" }
    $results = Join-Path $output 'results.txt'
    if (Test-Path $results) { Remove-Item $results }
    $run = Start-Process $exe -WorkingDirectory $output -PassThru
    if (-not $run.WaitForExit(30000)) { throw "Test still running: $($run.Id)" }
    if ($run.ExitCode -ne 0 -or -not (Test-Path $results)) { throw "Test did not finish (PYMMO=$pymmo)" }
    $result = [IO.File]::ReadAllText($results, $encoding)
    Write-Output "PYMMO=$pymmo`n$result"
    if ($result -notmatch 'Total=14 Failed=0') { throw 'Trap lifecycle regression failed' }
}
Write-Output 'Three production integration source guards passed.'
