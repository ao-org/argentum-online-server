# Preservation Property Tests - Property 2: Production Mode Behavior Unaffected
# **Validates: Requirements 3.1, 3.2, 3.3, 3.4**
#
# These tests inspect source code to verify that production-mode behaviors are preserved:
# 1. While(True) loop remains present and unconditional outside #If UNIT_TEST block
# 2. CerrarServidor() is called inside Form_Unload (cleanup chain preserved)
# 3. Form_QueryUnload checks GuardarYCerrar and shows popup when False
# 4. Command4_Click sets GuardarYCerrar = True and calls Unload frmMain
# 5. The fix (if applied) is entirely inside #If UNIT_TEST = 1 / #End If

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$generalBas = Join-Path $scriptDir "..\General.bas"
$frmMainFrm = Join-Path $scriptDir "..\frmMain.frm"

if (-not (Test-Path $generalBas)) {
    Write-Host "FAIL: General.bas not found at $generalBas"
    exit 1
}
if (-not (Test-Path $frmMainFrm)) {
    Write-Host "FAIL: frmMain.frm not found at $frmMainFrm"
    exit 1
}

$generalLines = Get-Content $generalBas
$frmMainLines = Get-Content $frmMainFrm

$allPassed = $true

# ============================================================
# Property 2a: While(True) loop remains present and unconditional
# outside the #If UNIT_TEST block
# ============================================================
Write-Host "--- Property 2a: While(True) loop preservation ---"

# Find the #End If that closes the UNIT_TEST block containing WriteResultsToFile
$endIfLine = -1
$inBlock = $false
$foundWrite = $false
for ($i = 0; $i -lt $generalLines.Count; $i++) {
    $line = $generalLines[$i].Trim()
    if ($line -match '^\#If\s+UNIT_TEST\s*=\s*1\s+Then$') {
        $inBlock = $true
        continue
    }
    if ($inBlock) {
        if ($line -match 'WriteResultsToFile') { $foundWrite = $true }
        if ($line -match '^\#End\s+If$') {
            if ($foundWrite) { $endIfLine = $i; break }
            $inBlock = $false
            $foundWrite = $false
        }
    }
}

if ($endIfLine -eq -1) {
    Write-Host "FAIL: Could not find #End If for UNIT_TEST block with WriteResultsToFile"
    $allPassed = $false
} else {
    # Check that While (True) exists after the #End If, outside any conditional block
    $foundWhileTrue = $false
    for ($j = $endIfLine + 1; $j -lt [Math]::Min($endIfLine + 5, $generalLines.Count); $j++) {
        if ($generalLines[$j].Trim() -match '^While\s*\(True\)') {
            $foundWhileTrue = $true
            break
        }
    }
    if ($foundWhileTrue) {
        Write-Host "PASS: While(True) loop exists unconditionally after #End If (line $($j+1))"
    } else {
        Write-Host "FAIL: While(True) loop not found after UNIT_TEST #End If"
        $allPassed = $false
    }
}

# ============================================================
# Property 2b: CerrarServidor() called inside Form_Unload
# ============================================================
Write-Host ""
Write-Host "--- Property 2b: CerrarServidor() in Form_Unload ---"

$inFormUnload = $false
$foundCerrar = $false
for ($i = 0; $i -lt $frmMainLines.Count; $i++) {
    $line = $frmMainLines[$i].Trim()
    if ($line -match 'Sub\s+Form_Unload') {
        $inFormUnload = $true
        continue
    }
    if ($inFormUnload) {
        if ($line -match 'CerrarServidor') {
            $foundCerrar = $true
            break
        }
        if ($line -match '^End\s+Sub') { break }
    }
}

if ($foundCerrar) {
    Write-Host "PASS: Form_Unload calls CerrarServidor() for proper cleanup"
} else {
    Write-Host "FAIL: Form_Unload does not call CerrarServidor()"
    $allPassed = $false
}

# ============================================================
# Property 2c: Form_QueryUnload checks GuardarYCerrar
# ============================================================
Write-Host ""
Write-Host "--- Property 2c: Form_QueryUnload popup preservation ---"

$inQueryUnload = $false
$foundGuardarCheck = $false
$foundForceClosePopup = $false
for ($i = 0; $i -lt $frmMainLines.Count; $i++) {
    $line = $frmMainLines[$i].Trim()
    if ($line -match 'Sub\s+Form_QueryUnload') {
        $inQueryUnload = $true
        continue
    }
    if ($inQueryUnload) {
        if ($line -match 'If\s+GuardarYCerrar') { $foundGuardarCheck = $true }
        if ($line -match 'FORZAR.*CIERRE|forzar.*cierre') { $foundForceClosePopup = $true }
        if ($line -match '^End\s+Sub') { break }
    }
}

if ($foundGuardarCheck) {
    Write-Host "PASS: Form_QueryUnload checks GuardarYCerrar flag"
} else {
    Write-Host "FAIL: Form_QueryUnload does not check GuardarYCerrar"
    $allPassed = $false
}

if ($foundForceClosePopup) {
    Write-Host "PASS: Form_QueryUnload shows force-close confirmation popup"
} else {
    Write-Host "FAIL: Form_QueryUnload missing force-close popup"
    $allPassed = $false
}

# ============================================================
# Property 2d: Command4_Click sets GuardarYCerrar and Unloads
# ============================================================
Write-Host ""
Write-Host "--- Property 2d: Command4_Click save-and-close path ---"

$inCommand4 = $false
$foundGuardarSet = $false
$foundUnload = $false
$foundConfirmPopup = $false
for ($i = 0; $i -lt $frmMainLines.Count; $i++) {
    $line = $frmMainLines[$i].Trim()
    if ($line -match 'Sub\s+Command4_Click') {
        $inCommand4 = $true
        continue
    }
    if ($inCommand4) {
        if ($line -match 'guardar\s+y\s+cerrar' -or $line -match 'seguro.*guardar') { $foundConfirmPopup = $true }
        if ($line -match 'GuardarYCerrar\s*=\s*True') { $foundGuardarSet = $true }
        if ($line -match 'Unload\s+frmMain') { $foundUnload = $true }
        if ($line -match '^End\s+Sub' -or $line -match '_Err:') { break }
    }
}

if ($foundConfirmPopup) {
    Write-Host "PASS: Command4_Click shows save-and-close confirmation popup"
} else {
    Write-Host "FAIL: Command4_Click missing confirmation popup"
    $allPassed = $false
}

if ($foundGuardarSet) {
    Write-Host "PASS: Command4_Click sets GuardarYCerrar = True"
} else {
    Write-Host "FAIL: Command4_Click does not set GuardarYCerrar = True"
    $allPassed = $false
}

if ($foundUnload) {
    Write-Host "PASS: Command4_Click calls Unload frmMain"
} else {
    Write-Host "FAIL: Command4_Click does not call Unload frmMain"
    $allPassed = $false
}

# ============================================================
# Property 2e: Any shutdown-related changes are inside #If UNIT_TEST = 1
# ============================================================
Write-Host ""
Write-Host "--- Property 2e: Fix scoped to #If UNIT_TEST = 1 block ---"

# Verify that GuardarYCerrar, Unload frmMain outside of frmMain.frm
# only appear inside #If UNIT_TEST blocks in General.bas
$outsideUnitTestBlock = $true
$violationsFound = $false
$inUnitTest = $false
$nestLevel = 0

for ($i = 0; $i -lt $generalLines.Count; $i++) {
    $line = $generalLines[$i].Trim()
    
    if ($line -match '^\#If\s+UNIT_TEST\s*=\s*1\s+Then$') {
        $inUnitTest = $true
        $nestLevel++
        continue
    }
    if ($line -match '^\#End\s+If$' -and $inUnitTest) {
        $nestLevel--
        if ($nestLevel -le 0) { $inUnitTest = $false; $nestLevel = 0 }
        continue
    }
    
    # Outside UNIT_TEST blocks, there should be no GuardarYCerrar or Unload frmMain in General.bas
    if (-not $inUnitTest) {
        if ($line -match 'GuardarYCerrar\s*=\s*True' -or $line -match 'Unload\s+frmMain') {
            Write-Host "FAIL: Found shutdown code outside #If UNIT_TEST block at line $($i+1): $line"
            $violationsFound = $true
        }
    }
}

if (-not $violationsFound) {
    Write-Host "PASS: No shutdown code (GuardarYCerrar/Unload frmMain) outside #If UNIT_TEST blocks in General.bas"
} else {
    $allPassed = $false
}

Write-Host ""
if ($allPassed) {
    Write-Host "=== ALL PRESERVATION TESTS PASSED ==="
    exit 0
} else {
    Write-Host "=== SOME PRESERVATION TESTS FAILED ==="
    exit 1
}
