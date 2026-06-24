$ErrorActionPreference = "Stop"

Set-Location -LiteralPath $PSScriptRoot
New-Item -ItemType Directory -Force -Path "logs" | Out-Null
New-Item -ItemType Directory -Force -Path ".uv-cache" | Out-Null
New-Item -ItemType Directory -Force -Path ".uv-python" | Out-Null
$env:UV_CACHE_DIR = Join-Path (Get-Location) ".uv-cache"
$env:UV_PYTHON_INSTALL_DIR = Join-Path (Get-Location) ".uv-python"
$LogFile = "logs\last_run.log"
"CFC Research Library update started at $(Get-Date)" | Set-Content -LiteralPath $LogFile
"Working folder: $(Get-Location)" | Add-Content -LiteralPath $LogFile
"UV cache: $env:UV_CACHE_DIR" | Add-Content -LiteralPath $LogFile
"UV Python install dir: $env:UV_PYTHON_INSTALL_DIR" | Add-Content -LiteralPath $LogFile

Write-Host "Starting CFC Research Library update..."
Write-Host ""

if (-not (Test-Path -LiteralPath ".env")) {
    Write-Host "Missing .env file."
    Write-Host ""
    Write-Host "Create a file named .env in this folder using .env.example as the template."
    Write-Host "Add ENTREZ_EMAIL, ZOTERO_GROUP_ID, ZOTERO_API_KEY, and OPENAI_API_KEY."
    "Missing .env file." | Add-Content -LiteralPath $LogFile
    Write-Host ""
    Read-Host "Press Enter to close"
    exit 1
}

if (Test-Path -LiteralPath ".venv\Scripts\python.exe") {
    try {
        & ".venv\Scripts\python.exe" -c "print('venv ok')" *>> $LogFile
    }
    catch {
        Write-Host "Existing Python environment is broken. Rebuilding it..."
        "Existing Python environment is broken. Rebuilding it..." | Add-Content -LiteralPath $LogFile
        Rename-Item -LiteralPath ".venv" -NewName ".venv-broken-$(Get-Random)"
    }
}

if (-not (Test-Path -LiteralPath ".venv\Scripts\python.exe")) {
    Write-Host "Creating local Python environment..."
    $bundledPy = Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
    if (Test-Path -LiteralPath $bundledPy) {
        & $bundledPy -m venv .venv *>> $LogFile
    }
    elseif (Get-Command py -ErrorAction SilentlyContinue) {
        py -3 -m venv .venv *>> $LogFile
    }
    elseif (Get-Command python -ErrorAction SilentlyContinue) {
        python -m venv .venv *>> $LogFile
    }
    elseif (Get-Command uv -ErrorAction SilentlyContinue) {
        uv venv *>> $LogFile
    }
    else {
        Write-Host "No Python installation was found."
        Write-Host "Install Python from https://www.python.org/downloads/ and run this again."
        "No Python installation was found." | Add-Content -LiteralPath $LogFile
        Read-Host "Press Enter to close"
        exit 1
    }
}

Write-Host "Installing or updating required libraries..."
& ".venv\Scripts\python.exe" -m pip install --upgrade pip *>> $LogFile
& ".venv\Scripts\python.exe" -m pip install -r requirements.txt *>> $LogFile

Write-Host ""
Write-Host "Running all-category research update..."
& ".venv\Scripts\python.exe" cfc_research_library.py --all-categories --output "reports\CFC_All_Categories_Master_Review_Report.xlsx" *>> $LogFile

Write-Host ""
Write-Host "Done."
Write-Host "Workbook created at:"
Write-Host "reports\CFC_All_Categories_Master_Review_Report.xlsx"
Write-Host "Full log saved to:"
Write-Host $LogFile
Write-Host ""
Read-Host "Press Enter to close"
