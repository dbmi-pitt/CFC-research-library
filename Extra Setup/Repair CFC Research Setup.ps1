$ErrorActionPreference = "Stop"

Set-Location -LiteralPath $PSScriptRoot
New-Item -ItemType Directory -Force -Path "logs" | Out-Null
$LogFile = "logs\repair_setup.log"
"CFC Research Library repair started at $(Get-Date)" | Set-Content -LiteralPath $LogFile
"Working folder: $(Get-Location)" | Add-Content -LiteralPath $LogFile

Write-Host "Repairing CFC Research Library setup..."
Write-Host ""

if (Test-Path -LiteralPath ".venv") {
    $backupName = ".venv-replaced-$(Get-Random)"
    Write-Host "Moving old Python environment to $backupName..."
    "Moving old Python environment to $backupName..." | Add-Content -LiteralPath $LogFile
    Rename-Item -LiteralPath ".venv" -NewName $backupName
}

$bundledPy = Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

Write-Host "Creating a fresh Python environment..."
if (Test-Path -LiteralPath $bundledPy) {
    & $bundledPy -m venv .venv *>> $LogFile
}
elseif (Get-Command py -ErrorAction SilentlyContinue) {
    py -3 -m venv .venv *>> $LogFile
}
else {
    python -m venv .venv *>> $LogFile
}

Write-Host "Installing all required libraries..."
& ".venv\Scripts\python.exe" -m pip install --upgrade pip *>> $LogFile
& ".venv\Scripts\python.exe" -m pip install -r requirements.txt *>> $LogFile

Write-Host ""
Write-Host "Repair complete."
Write-Host "Now double-click:"
Write-Host "Run CFC Research Update.bat"
Write-Host ""
Write-Host "Full repair log saved to:"
Write-Host $LogFile
Write-Host ""
Read-Host "Press Enter to close"
