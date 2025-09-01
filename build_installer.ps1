$ErrorActionPreference = 'Stop'

# Move to repo root (this script's directory)
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $here

Write-Host "Building SPS30 Tray Logger..." -ForegroundColor Cyan

if (-not (Get-Command py -ErrorAction SilentlyContinue)) {
    Write-Error "Python launcher 'py' not found. Install Python 3 for Windows from https://www.python.org/downloads/windows/ and ensure 'py' is on PATH."
}

# Close any running instance to avoid file lock
Write-Host "Ensuring no running instance is locking dist\\sps30-tray-logger.exe..." -ForegroundColor Yellow
try {
    Get-Process -Name "sps30-tray-logger" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
} catch {}

# Clean previous build artifacts
Remove-Item -Recurse -Force -ErrorAction SilentlyContinue .\build
Remove-Item -Recurse -Force -ErrorAction SilentlyContinue .\dist

Write-Host "Installing/upgrading build dependencies..." -ForegroundColor Yellow
py -m pip install --upgrade pip | Out-Host
# Core deps for building and UI
py -m pip install pyinstaller pystray pillow matplotlib | Out-Host
# Optional sensor drivers. If PyPI fails for sensirion-shdlc-sps, fall back to GitHub.
try {
    py -m pip install sensirion-shdlc-driver | Out-Host

    # Prefer local archive if present
    $localArchive = "C:\\Users\\Sebastian\\Downloads\\shdlc-sps30-0.1.tar.gz"
    if (Test-Path $localArchive) {
        Write-Host "Installing sensirion-shdlc-sps from local archive: $localArchive" -ForegroundColor Yellow
        & py -m pip install "$localArchive" | Out-Host
        $rcLocal = $LASTEXITCODE
        if ($rcLocal -eq 0) { throw 'OK_LOCAL' }
        Write-Warning "Local archive install failed (exit $rcLocal). Falling back to online sources..."
    }

    # Try PyPI first
    & py -m pip install sensirion-shdlc-sps | Out-Host
    $rc = $LASTEXITCODE
    if ($rc -ne 0) {
        Write-Warning "PyPI sensirion-shdlc-sps failed (exit $rc). Falling back to GitHub ZIP..."
        # Fallback to public ZIP (if available)
        & py -m pip install https://github.com/Sensirion/python-shdlc-sps/archive/refs/heads/master.zip | Out-Host
    }
} catch {
    if ($_.Exception.Message -eq 'OK_LOCAL') {
        Write-Host "sensirion-shdlc-sps installed from local archive." -ForegroundColor Green
    } else {
        Write-Warning "Could not install sensirion-shdlc-sps from any source. The EXE will build; device access may not work."
    }
}

Write-Host "Running PyInstaller (single-file, no console)..." -ForegroundColor Yellow
py -m PyInstaller --onefile --noconsole --name sps30-tray-logger .\sps30_tray_logger_win.py | Out-Host

$exe = Join-Path $here 'dist\sps30-tray-logger.exe'
if (-not (Test-Path $exe)) {
    Write-Error "Build failed: $exe not found. See PyInstaller output above."
}

Write-Host "Executable created: $exe" -ForegroundColor Green

# Try to build an installer with Inno Setup if available
$iss = Join-Path $here 'installer\SPS30TrayLogger.iss'
if (Test-Path $iss) {
    $iscc = Get-Command iscc.exe -ErrorAction SilentlyContinue
    if ($iscc) {
        Write-Host "Inno Setup detected. Building installer..." -ForegroundColor Yellow
        & $iscc.Path $iss | Out-Host
        $setupPath = Join-Path $here 'installer\Output\SPS30TrayLoggerSetup.exe'
        if (Test-Path $setupPath) {
            Write-Host "Installer created: $setupPath" -ForegroundColor Green
        } else {
            Write-Warning "ISCC ran, but no installer found at $setupPath. Check the Inno output above."
        }
    } else {
        Write-Warning "Inno Setup (ISCC.exe) not found on PATH. Install from https://jrsoftware.org/isinfo.php to build the .exe installer."
        Write-Host "You can still share the portable EXE: $exe" -ForegroundColor Cyan
    }
} else {
    Write-Warning "Installer script not found: $iss"
}

Write-Host "Done." -ForegroundColor Cyan


