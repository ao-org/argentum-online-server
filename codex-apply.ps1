param(
    [string]$PatchFile = "codex.patch"
)

# 1) Leer del portapapeles y guardar
$clip = Get-Clipboard
if (-not $clip) {
    Write-Host "No hay nada en el portapapeles. Pegá el parche a mano en $PatchFile"
    New-Item -ItemType File -Path $PatchFile -Force | Out-Null
} else {
    $clip | Out-File -FilePath $PatchFile -Encoding utf8
    Write-Host "Patch guardado en $PatchFile"
}

Write-Host "Aplicando patch..."

# ejecutamos git apply y capturamos el código de salida
git apply $PatchFile -v
$exitCode = $LASTEXITCODE

if ($exitCode -ne 0) {
    Write-Host "❌ git apply falló con código $exitCode"
    Write-Host "Revisá las secciones del patch que NO se aplicaron (Protocol.bas y frmShop.frm)."
} else {
    Write-Host "✅ Patch aplicado correctamente. Podés ver los cambios con: git status"
}

Read-Host "Presioná Enter para salir..."
