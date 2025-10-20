# Badge Generator Setup Script
# This script checks for Python and installs all required dependencies

$ErrorActionPreference = 'Stop'

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Badge Generator - Setup Script" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check if Python is installed
Write-Host "[1/4] Checking for Python..." -ForegroundColor Yellow
try {
    $pythonVersion = python --version 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "  ✓ Found: $pythonVersion" -ForegroundColor Green
    } else {
        throw "Python not found"
    }
} catch {
    Write-Host "  ✗ Python is not installed or not in PATH" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please install Python first:" -ForegroundColor Yellow
    Write-Host "  1. Download from: https://www.python.org/downloads/" -ForegroundColor White
    Write-Host "  2. During installation, check 'Add Python to PATH'" -ForegroundColor White
    Write-Host "  3. Restart your terminal and run this script again" -ForegroundColor White
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

# Check Python version (should be 3.8+)
Write-Host "[2/4] Verifying Python version..." -ForegroundColor Yellow
$versionString = python -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')" 2>&1
$version = [version]$versionString
if ($version -lt [version]"3.8") {
    Write-Host "  ✗ Python 3.8 or higher is required (found $versionString)" -ForegroundColor Red
    exit 1
}
Write-Host "  ✓ Python version is compatible ($versionString)" -ForegroundColor Green

# Upgrade pip
Write-Host "[3/4] Upgrading pip..." -ForegroundColor Yellow
python -m pip install --upgrade pip --quiet
if ($LASTEXITCODE -eq 0) {
    Write-Host "  ✓ pip upgraded successfully" -ForegroundColor Green
} else {
    Write-Host "  ⚠ Warning: Could not upgrade pip" -ForegroundColor Yellow
}

# Install dependencies from requirements.txt
Write-Host "[4/4] Installing Python packages..." -ForegroundColor Yellow
if (Test-Path "requirements.txt") {
    python -m pip install -r requirements.txt
    if ($LASTEXITCODE -eq 0) {
        Write-Host "  ✓ All packages installed successfully" -ForegroundColor Green
    } else {
        Write-Host "  ✗ Failed to install some packages" -ForegroundColor Red
        exit 1
    }
} else {
    Write-Host "  ✗ requirements.txt not found" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Setup Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Additional requirements:" -ForegroundColor Yellow
Write-Host "  • Microsoft Word (for docx2pdf conversion)" -ForegroundColor White
Write-Host "  • Faustina font family (for proper text rendering)" -ForegroundColor White
Write-Host "    Download: https://github.com/google/fonts/tree/main/ofl/faustina" -ForegroundColor White
Write-Host ""
Write-Host "To generate badges, run:" -ForegroundColor Yellow
Write-Host "  python badge_generator.py" -ForegroundColor White
Write-Host ""
