# Copyright (C) 2026 Noland Studios LTD
# Licensed under the GNU Affero General Public License, version 3 or later.
param([string]$Python = 'python', [string]$Harness = '')
$ErrorActionPreference = 'Stop'
if ([Environment]::Is64BitProcess) {
    throw 'Run this test with the 32-bit Windows PowerShell in Windows/SysWOW64 to match the VB6 ODBC driver.'
}
$repository = Split-Path -Parent $PSScriptRoot
$outputDirectory = Join-Path $repository 'build/castle-sql-tests'
New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null
$migrationPath = Join-Path $repository 'ScriptsDB/20261007-01-migrate castle identities.sql'
if (-not $Harness) { $Harness = Join-Path $repository 'build/map-schema-tests/map_schema_tests.exe' }
if (-not (Test-Path -LiteralPath $Harness)) { throw 'Build the native harness with tools/test_map_schema.ps1 first.' }
$statementPrefix = Join-Path $outputDirectory ([guid]::NewGuid().ToString('N'))
$exportArguments = '"split-sql|' + $migrationPath + '|' + $statementPrefix + '"'
$exportProcess = Start-Process -FilePath $Harness -ArgumentList $exportArguments -WindowStyle Hidden -PassThru -Wait
if ($exportProcess.ExitCode -ne 0 -or (Test-Path -LiteralPath ($statementPrefix + '.error'))) {
    throw 'The production VB6 SQL splitter rejected the migration.'
}
$statementCount = [int][IO.File]::ReadAllText($statementPrefix + '.count')
$statements = @(for ($index = 1; $index -le $statementCount; $index++) {
    [IO.File]::ReadAllText($statementPrefix + '.' + $index + '.sql', [Text.Encoding]::GetEncoding(1252))
})

function Assert-Equal($Actual, $Expected, [string]$Message) {
    if ($Actual -ne $Expected) { throw "$Message`: expected $Expected, got $Actual" }
}
function Scalar($Connection, [string]$Query) {
    $rows = $Connection.Execute($Query)
    try { return $rows.Fields.Item(0).Value }
    finally { if ($rows.State -ne 0) { $rows.Close() } }
}
function Invoke-CastleSql($Connection, [string[]]$Statements) {
    foreach ($statement in $Statements) {
        # Match Query's prepared synchronous ADO command path.
        $command = New-Object -ComObject ADODB.Command
        try {
            $command.ActiveConnection = $Connection
            $command.CommandType = 1
            $command.Prepared = $true
            $command.CommandText = $statement
            $result = $command.Execute()
            if ($null -eq $result) { throw 'Query must return a recordset object for RunScriptInFile to recognize success.' }
            if ($result.State -ne 0) { $result.Close() }
        } finally { [Runtime.InteropServices.Marshal]::ReleaseComObject($command) | Out-Null }
    }
}

foreach ($invalid in @($false, $true)) {
    $databasePath = Join-Path $outputDirectory (([guid]::NewGuid().ToString('N')) + '.db')
    $fixtureCode = 'import sqlite3,sys; sys.path.insert(0,sys.argv[1]); from test_castle_migration import fixture; c=sqlite3.connect(sys.argv[2]); fixture(c); c.close()'
    & $Python -c $fixtureCode $PSScriptRoot $databasePath
    if ($LASTEXITCODE -ne 0) { throw 'Cannot create isolated castle fixture' }
    $connection = New-Object -ComObject ADODB.Connection
    try {
        $connection.Open("DRIVER={SQLite3 ODBC Driver};DATABASE=$databasePath")
        $connection.Execute('PRAGMA foreign_keys=ON') | Out-Null
        $version = Scalar $connection 'SELECT sqlite_version()'
        if ([version]$version -lt [version]'3.35.0') { throw "SQLite $version does not support this migration" }
        if ($invalid) {
            $connection.Execute('UPDATE castle_coordinates SET outside_x=10 WHERE castle_id=2') | Out-Null
            $failed = $false
            try { Invoke-CastleSql $connection $statements }
            catch { $failed = $true; $connection.Execute('ROLLBACK') | Out-Null }
            Assert-Equal $failed $true 'Partial outside tuple must reject the migration'
            Assert-Equal (Scalar $connection 'SELECT trigger FROM castle WHERE id=1') 21 'Rollback must restore legacy column'
            Assert-Equal (Scalar $connection "SELECT COUNT(*) FROM sqlite_master WHERE name='castle_legacy_trigger_map'") 0 'Rollback must remove mapping table'
        } else {
            Invoke-CastleSql $connection $statements
            Assert-Equal (Scalar $connection 'SELECT COUNT(*) FROM castle') 20 'Castle count'
            Assert-Equal (Scalar $connection 'SELECT COUNT(*) FROM castle_coordinates') 15 'Coordinate count'
            Assert-Equal (Scalar $connection 'SELECT COUNT(*) FROM castle_whitelist') 3 'Whitelist count'
            Assert-Equal (Scalar $connection 'SELECT owner_account_id FROM castle WHERE id=2') 9 'Unplaced castle owner'
            Assert-Equal (Scalar $connection "SELECT CAST(foundation_date AS TEXT) FROM castle WHERE id=2") '2026-08-11 18:46:53.434' 'Fractional foundation date'
            Assert-Equal (Scalar $connection 'SELECT COUNT(*) FROM castle_coordinates WHERE castle_id=2 AND outside_map IS NULL AND outside_x IS NULL AND outside_y IS NULL') 1 'Unplaced outside tuple'
            Assert-Equal (Scalar $connection 'SELECT castle_id FROM castle_legacy_trigger_map WHERE legacy_trigger=23') 3 'Stable castle mapping'
            Assert-Equal (Scalar $connection "SELECT seq FROM sqlite_sequence WHERE name='castle'") 50 'Castle sequence'
            $violations = $connection.Execute('PRAGMA foreign_key_check')
            try { Assert-Equal $violations.EOF $true 'Foreign-key validation' }
            finally { if ($violations.State -ne 0) { $violations.Close() } }
        }
        Write-Output "32-bit ADO SQLite $version`: castle SQL test passed (invalid fixture=$invalid)."
    } finally {
        if ($connection.State -ne 0) { $connection.Close() }
        [Runtime.InteropServices.Marshal]::ReleaseComObject($connection) | Out-Null
    }
}
