# Bug Condition Exploration Test - Property 1: Graceful Shutdown After Tests Complete
# **Validates: Requirements 1.1, 1.2, 1.3, 1.4, 2.1, 2.2, 2.3, 2.4**
#
# This test inspects the source code of General.bas to verify that the
# #If UNIT_TEST = 1 block contains graceful shutdown logic after WriteResultsToFile().
#
# Expected: On UNFIXED code, this test FAILS (shutdown lines are missing).
# After fix: This test PASSES (shutdown lines are present).

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$generalBas = Join-Path $scriptDir "..\General.bas"

if (-not (Test-Path $generalBas)) {
    Write-Host "FAIL: General.bas not found at $generalBas"
    exit 1
}

$content = Get-Content $generalBas -Raw
$lines = Get-Content $generalBas

# Find the #If UNIT_TEST = 1 block that contains WriteResultsToFile
# We need to extract the block between #If UNIT_TEST = 1 and the matching #End If
# that contains the test suite calls (Init, test_suite, WriteResultsToFile)

$inUnitTestBlock = $false
$blockLines = @()
$blockStartLine = -1
$blockEndLine = -1
$foundWriteResults = $false

for ($i = 0; $i -lt $lines.Count; $i++) {
    $line = $lines[$i].Trim()
    
    if ($line -match '^\#If\s+UNIT_TEST\s*=\s*1\s+Then$' -and -not $inUnitTestBlock) {
        # Check if this block contains WriteResultsToFile (the main init block, not the loop one)
        $inUnitTestBlock = $true
        $blockStartLine = $i
        $blockLines = @()
        continue
    }
    
    if ($inUnitTestBlock) {
        $blockLines += $lines[$i]
        
        if ($line -match 'WriteResultsToFile') {
            $foundWriteResults = $true
        }
        
        if ($line -match '^\#End\s+If$') {
            $blockEndLine = $i
            if ($foundWriteResults) {
                break  # Found the right block
            }
            # Reset - this wasn't the right block
            $inUnitTestBlock = $false
            $blockLines = @()
            $foundWriteResults = $false
        }
    }
}

if (-not $foundWriteResults) {
    Write-Host "FAIL: Could not find #If UNIT_TEST = 1 block containing WriteResultsToFile"
    exit 1
}

Write-Host "Found UNIT_TEST block at lines $($blockStartLine+1)-$($blockEndLine+1)"
Write-Host "Block contents:"
$blockLines | ForEach-Object { Write-Host "  $_" }
Write-Host ""

$allPassed = $true

# Test 1: Check for GuardarYCerrar = True
$hasGuardarYCerrar = ($blockLines | Where-Object { $_ -match 'GuardarYCerrar\s*=\s*True' }).Count -gt 0
if ($hasGuardarYCerrar) {
    Write-Host "PASS: Found 'GuardarYCerrar = True' in UNIT_TEST block"
} else {
    Write-Host "FAIL: Missing 'GuardarYCerrar = True' in UNIT_TEST block"
    Write-Host "  Counterexample: The #If UNIT_TEST = 1 block ends after WriteResultsToFile with no shutdown trigger"
    $allPassed = $false
}

# Test 2: Check for Unload frmMain
$hasUnloadFrmMain = ($blockLines | Where-Object { $_ -match 'Unload\s+frmMain' }).Count -gt 0
if ($hasUnloadFrmMain) {
    Write-Host "PASS: Found 'Unload frmMain' in UNIT_TEST block"
} else {
    Write-Host "FAIL: Missing 'Unload frmMain' in UNIT_TEST block - Form_Unload/CerrarServidor() will never be called"
    Write-Host "  Counterexample: No Unload frmMain means CerrarServidor() cleanup chain is never triggered"
    $allPassed = $false
}

# Test 3: Check for Exit Sub
$hasExitSub = ($blockLines | Where-Object { $_ -match 'Exit\s+Sub' }).Count -gt 0
if ($hasExitSub) {
    Write-Host "PASS: Found 'Exit Sub' in UNIT_TEST block"
} else {
    Write-Host "FAIL: Missing 'Exit Sub' in UNIT_TEST block - execution falls through to While(True) loop"
    Write-Host "  Counterexample: Without Exit Sub, Main() continues past #End If into the infinite While(True) game loop"
    $allPassed = $false
}

# Test 4: Verify ordering - GuardarYCerrar before Unload before Exit Sub, all after WriteResultsToFile
if ($hasGuardarYCerrar -and $hasUnloadFrmMain -and $hasExitSub) {
    $writeIdx = -1
    $guardarIdx = -1
    $unloadIdx = -1
    $exitIdx = -1
    for ($i = 0; $i -lt $blockLines.Count; $i++) {
        if ($blockLines[$i] -match 'WriteResultsToFile' -and $writeIdx -eq -1) { $writeIdx = $i }
        if ($blockLines[$i] -match 'GuardarYCerrar\s*=\s*True' -and $guardarIdx -eq -1) { $guardarIdx = $i }
        if ($blockLines[$i] -match 'Unload\s+frmMain' -and $unloadIdx -eq -1) { $unloadIdx = $i }
        if ($blockLines[$i] -match 'Exit\s+Sub' -and $exitIdx -eq -1) { $exitIdx = $i }
    }
    
    if ($writeIdx -lt $guardarIdx -and $guardarIdx -lt $unloadIdx -and $unloadIdx -lt $exitIdx) {
        Write-Host "PASS: Correct ordering - WriteResultsToFile < GuardarYCerrar < Unload frmMain < Exit Sub"
    } else {
        Write-Host "FAIL: Incorrect ordering of shutdown statements"
        $allPassed = $false
    }
}

Write-Host ""
if ($allPassed) {
    Write-Host "=== ALL TESTS PASSED ==="
    Write-Host "The UNIT_TEST block contains proper graceful shutdown logic."
    exit 0
} else {
    Write-Host "=== TESTS FAILED ==="
    Write-Host "Bug confirmed: The #If UNIT_TEST = 1 block lacks graceful shutdown."
    Write-Host "After tests complete, execution falls through to While(True) loop."
    exit 1
}
