# PowerShell script to rebuild virtual environment from requirements.txt
# This removes the existing .venv and creates a fresh one with exact package versions

$ProjectRoot = Split-Path -Parent $PSScriptRoot
$VenvPath = Join-Path $ProjectRoot ".venv"
$RequirementsFile = Join-Path $PSScriptRoot "requirements.txt"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Virtual Environment Rebuild Script" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check if requirements.txt exists
if (!(Test-Path $RequirementsFile)) {
    Write-Host "ERROR: requirements.txt not found at:" -ForegroundColor Red
    Write-Host "  $RequirementsFile" -ForegroundColor Red
    exit 1
}

Write-Host "Project Root: $ProjectRoot" -ForegroundColor Gray
Write-Host "Virtual Environment: $VenvPath" -ForegroundColor Gray
Write-Host "Requirements File: $RequirementsFile" -ForegroundColor Gray
Write-Host ""

# Step 1: Remove existing virtual environment
if (Test-Path $VenvPath) {
    Write-Host "Step 1: Removing existing virtual environment..." -ForegroundColor Yellow
    Write-Host "  This may take a moment..." -ForegroundColor Gray
    
    try {
        Remove-Item -Recurse -Force $VenvPath -ErrorAction Stop
        Write-Host "  ✓ Removed existing .venv" -ForegroundColor Green
    } catch {
        Write-Host "  ✗ Failed to remove .venv: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "  Please close any programs using the virtual environment and try again." -ForegroundColor Yellow
        exit 1
    }
} else {
    Write-Host "Step 1: No existing virtual environment found (skipping removal)" -ForegroundColor Yellow
}

Write-Host ""

# Step 2: Create new virtual environment
Write-Host "Step 2: Creating new virtual environment..." -ForegroundColor Yellow

try {
    python -m venv $VenvPath
    
    if (Test-Path $VenvPath) {
        Write-Host "  ✓ Virtual environment created successfully" -ForegroundColor Green
    } else {
        Write-Host "  ✗ Failed to create virtual environment" -ForegroundColor Red
        exit 1
    }
} catch {
    Write-Host "  ✗ Error creating virtual environment: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""

# Step 3: Activate virtual environment and upgrade pip
Write-Host "Step 3: Activating virtual environment and upgrading pip..." -ForegroundColor Yellow

$ActivateScript = Join-Path $VenvPath "Scripts\Activate.ps1"

if (!(Test-Path $ActivateScript)) {
    Write-Host "  ✗ Activation script not found" -ForegroundColor Red
    exit 1
}

try {
    & $ActivateScript
    Write-Host "  ✓ Virtual environment activated" -ForegroundColor Green
    
    # Upgrade pip to latest version
    python -m pip install --upgrade pip --quiet
    Write-Host "  ✓ pip upgraded to latest version" -ForegroundColor Green
} catch {
    Write-Host "  ✗ Error activating or upgrading: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""

# Step 4: Install packages from requirements.txt
Write-Host "Step 4: Installing packages from requirements.txt..." -ForegroundColor Yellow
Write-Host "  This may take several minutes..." -ForegroundColor Gray

try {
    pip install -r $RequirementsFile
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host ""
        Write-Host "  ✓ All packages installed successfully" -ForegroundColor Green
    } else {
        Write-Host ""
        Write-Host "  ⚠ Some packages may have failed to install" -ForegroundColor Yellow
        Write-Host "  Exit code: $LASTEXITCODE" -ForegroundColor Yellow
    }
} catch {
    Write-Host "  ✗ Error installing packages: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "Virtual Environment Rebuild Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Yellow
Write-Host "  1. Close and reopen VS Code (or reload window)" -ForegroundColor White
Write-Host "  2. Select the new Python interpreter:" -ForegroundColor White
Write-Host "     $VenvPath\Scripts\python.exe" -ForegroundColor Gray
Write-Host "  3. Run your application to verify everything works" -ForegroundColor White
Write-Host ""
