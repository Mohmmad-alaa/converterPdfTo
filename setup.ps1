param(
    [switch]$CheckOnly
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

function Step($message) {
    Write-Host ""
    Write-Host "== $message ==" -ForegroundColor Cyan
}

function Fail($message) {
    Write-Host ""
    Write-Host "ERROR: $message" -ForegroundColor Red
    if (-not $CheckOnly) { Read-Host "Press Enter to close" }
    exit 1
}

Write-Host "Muhawwil setup - Windows OCR edition" -ForegroundColor Green
Write-Host "Project: $root"

if ($env:OS -ne "Windows_NT") {
    Fail "This application requires Windows 10 or Windows 11."
}

$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $CheckOnly -and -not $isAdmin) {
    Write-Host ""
    Write-Host "Administrator permission is required once to install Arabic OCR." -ForegroundColor Yellow
    Write-Host "A new elevated window will open..."
    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    Start-Process powershell.exe -Verb RunAs -ArgumentList $arguments
    exit 0
}

# --- Python ---------------------------------------------------------------
Step "Checking Python"
$pythonExe = "python"
$pythonArgs = @()
if (Get-Command python -ErrorAction SilentlyContinue) {
    $pythonExe = "python"
    $pythonArgs = @()
} elseif (Get-Command py -ErrorAction SilentlyContinue) {
    $pythonExe = "py"
    $pythonArgs = @("-3")
} else {
    Fail "Python was not found. Install Python 3.10 or newer from https://www.python.org/downloads/ and enable 'Add Python to PATH'."
}

try {
    $version = & $pythonExe @pythonArgs -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')"
} catch {
    Fail "Python could not be started. Reinstall Python and enable 'Add Python to PATH'."
}
if ([version]$version -lt [version]"3.10") {
    Fail "Python $version is too old. Install Python 3.10 or newer."
}
if ([version]$version -gt [version]"3.13") {
    Write-Host "Warning: Python $version is newer than the tested versions. If package installation fails, install Python 3.12 or 3.13." -ForegroundColor Yellow
}
Write-Host "Python $version OK" -ForegroundColor Green

$venvPython = Join-Path $root ".venv\Scripts\python.exe"
$venvReady = Test-Path $venvPython

# --- Check-only mode ------------------------------------------------------
if ($CheckOnly) {
    Step "Check summary"
    Write-Host "Python: $version"
    Write-Host "Virtual environment: $(if ($venvReady) { 'found' } else { 'not created yet' })"
    try {
        if ($venvReady) {
            $ocr = & $venvPython -c "import ocr_engine; print('AVAILABLE' if ocr_engine.available() else 'MISSING'); print(','.join(ocr_engine.languages()))"
        } elseif ($pythonExe -eq "py") {
            $ocr = & py -3 -c "import ocr_engine; print('AVAILABLE' if ocr_engine.available() else 'MISSING'); print(','.join(ocr_engine.languages()))"
        } else {
            $ocr = & python -c "import ocr_engine; print('AVAILABLE' if ocr_engine.available() else 'MISSING'); print(','.join(ocr_engine.languages()))"
        }
        Write-Host "OCR: $($ocr -join ' | ')"
    } catch {
        Write-Host "OCR: Python packages are not installed yet." -ForegroundColor Yellow
    }
    try {
        $cap = Get-WindowsCapability -Online | Where-Object { $_.Name -like "Language.OCR*ar-SA*" } | Select-Object -First 1
        if ($cap) {
            Write-Host "Arabic OCR capability: $($cap.State)"
        } else {
            Write-Host "Arabic OCR capability: not found" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "Arabic OCR capability: run setup as administrator to check/install it." -ForegroundColor Yellow
    }
    exit 0
}

# --- Virtual environment and packages ------------------------------------
Step "Preparing Python environment"
if (-not $venvReady) {
    Write-Host "Creating .venv ..."
    if ($pythonExe -eq "py") {
        & py -3 -m venv .venv
    } else {
        & python -m venv .venv
    }
    if ($LASTEXITCODE -ne 0) { Fail "Could not create the virtual environment." }
}
if (-not (Test-Path $venvPython)) {
    Fail "The virtual environment is incomplete. Delete the .venv folder and run setup again."
}

Write-Host "Installing required packages (this may take a few minutes)..."
& $venvPython -m pip install --upgrade pip
if ($LASTEXITCODE -ne 0) { Fail "Could not update pip." }
& $venvPython -m pip install -r requirements.txt
if ($LASTEXITCODE -ne 0) { Fail "Could not install requirements.txt." }

# --- Arabic OCR capability ------------------------------------------------
Step "Checking Arabic OCR"
$cap = Get-WindowsCapability -Online | Where-Object { $_.Name -like "Language.OCR*ar-SA*" } | Select-Object -First 1
if (-not $cap) {
    Write-Host "Arabic OCR capability was not found on this Windows build." -ForegroundColor Yellow
    Write-Host "Install it manually: Settings > Time & Language > Language & region > Arabic > Language options > Optical character recognition."
} elseif ($cap.State -ne "Installed") {
    Write-Host "Installing Arabic OCR capability..."
    try {
        Add-WindowsCapability -Online -Name $cap.Name | Out-Null
    } catch {
        Write-Host "Automatic installation failed. Install it manually from Windows Settings." -ForegroundColor Yellow
    }
} else {
    Write-Host "Arabic OCR is already installed." -ForegroundColor Green
}

try {
    $ocr = & $venvPython -c "import ocr_engine; print('AVAILABLE' if ocr_engine.available() else 'MISSING'); print(','.join(ocr_engine.languages()))"
    Write-Host "OCR status: $($ocr -join ' | ')" -ForegroundColor Green
} catch {
    Write-Host "OCR check failed. The application can still run, but scanned Arabic files may not be recognized." -ForegroundColor Yellow
}

# --- Start ---------------------------------------------------------------
Step "Starting Muhawwil"
$port = 8000
$listening = Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue
if ($listening) {
    Write-Host "The application is already running on port $port." -ForegroundColor Green
    Start-Process "http://127.0.0.1:$port"
    Read-Host "Press Enter to close"
    exit 0
}

Start-Process "http://127.0.0.1:$port"
& $venvPython -m uvicorn main:app --host 127.0.0.1 --port $port

if ($LASTEXITCODE -ne 0) {
    Fail "The application stopped with an error."
}
