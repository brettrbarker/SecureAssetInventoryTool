# Build script for Asset Management System
# This script packages the application into a standalone Windows executable

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Asset Management System - Build Script" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check if virtual environment is activated
if (-not $env:VIRTUAL_ENV) {
    Write-Host "WARNING: Virtual environment not detected!" -ForegroundColor Yellow
    Write-Host "Attempting to activate .venv..." -ForegroundColor Yellow
    
    if (Test-Path ".\.venv\Scripts\Activate.ps1") {
        & ".\.venv\Scripts\Activate.ps1"
        Write-Host "Virtual environment activated." -ForegroundColor Green
    } else {
        Write-Host "ERROR: Could not find .venv\Scripts\Activate.ps1" -ForegroundColor Red
        Write-Host "Please activate your virtual environment manually and try again." -ForegroundColor Red
        exit 1
    }
}

# Check if PyInstaller is installed
Write-Host "Checking PyInstaller installation..." -ForegroundColor Cyan
$pyinstallerCheck = python -m pip show pyinstaller 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "PyInstaller not found. Installing..." -ForegroundColor Yellow
    python -m pip install pyinstaller pyinstaller-hooks-contrib
}

# Clean previous build artifacts
Write-Host ""
Write-Host "Cleaning previous build artifacts..." -ForegroundColor Cyan
if (Test-Path ".\dist") {
    Remove-Item -Recurse -Force ".\dist"
    Write-Host "  Removed dist folder" -ForegroundColor Gray
}
if (Test-Path ".\build") {
    Remove-Item -Recurse -Force ".\build"
    Write-Host "  Removed build folder" -ForegroundColor Gray
}
if (Test-Path ".\SecureAssetInventoryTool.spec") {
    Remove-Item -Force ".\SecureAssetInventoryTool.spec"
    Write-Host "  Removed old spec file" -ForegroundColor Gray
}

# Check for icon file
$iconPath = "assets\fonts\secure_asset_inventory_tool_final.ico"
$iconParam = ""
if (Test-Path $iconPath) {
    Write-Host "Icon file found: $iconPath" -ForegroundColor Green
    $iconParam = "--icon `"$iconPath`""
} else {
    Write-Host "Icon file not found (optional): $iconPath" -ForegroundColor Yellow
}

# Run PyInstaller
Write-Host ""
Write-Host "Building executable with PyInstaller..." -ForegroundColor Cyan
Write-Host "This may take several minutes..." -ForegroundColor Yellow
Write-Host ""

# Build the command with optimizations for smaller size and better performance
# Key optimizations:
# --onefile: Single executable (smaller distribution, slightly slower startup)
# --strip: Remove debug symbols (smaller size)
# --exclude-module: Remove unused large modules
# --noupx: Don't use UPX compression (better compatibility, slightly larger)
# Only include necessary assets folders
$command = @"
pyinstaller --noconfirm ``
    --onefile ``
    --name "SecureAssetInventoryTool" ``
    $iconParam ``
    --strip ``
    --exclude-module matplotlib.tests ``
    --exclude-module pandas.tests ``
    --exclude-module numpy.tests ``
    --exclude-module PIL.tests ``
    --exclude-module pytest ``
    --exclude-module unittest ``
    --add-data "assets/fonts/CustomTkinter_shapes_font.otf;assets/fonts" ``
    --add-data "assets/fonts/Roboto-Medium.ttf;assets/fonts" ``
    --add-data "assets/fonts/Roboto-Regular.ttf;assets/fonts" ``
    --add-data "assets/fonts/Roboto;assets/fonts/Roboto" ``
    --add-data "assets/fonts/secure_asset_inventory_tool_final.ico;assets/fonts" ``
    --add-data "assets/templates/default_template.csv;assets/templates" ``
    --add-data "assets/config.json;assets" ``
    --hidden-import "PIL._tkinter_finder" ``
    --hidden-import "customtkinter" ``
    --hidden-import "tkinter" ``
    --collect-all "customtkinter" ``
    --noupx ``
    main.py
"@

# Execute the command
Invoke-Expression $command

if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "Build completed successfully!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Executable location: .\dist\SecureAssetInventoryTool.exe" -ForegroundColor Cyan
    Write-Host ""
    
    # Show file size
    if (Test-Path ".\dist\SecureAssetInventoryTool.exe") {
        $fileSize = (Get-Item ".\dist\SecureAssetInventoryTool.exe").Length
        $fileSizeMB = [math]::Round($fileSize / 1MB, 2)
        Write-Host "Executable size: $fileSizeMB MB" -ForegroundColor Cyan
    }
    
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Yellow
    Write-Host "  1. Test the executable: .\dist\SecureAssetInventoryTool.exe" -ForegroundColor White
    Write-Host "  2. The executable is a single portable file" -ForegroundColor White
    Write-Host "  3. Note: First run may be slower (unpacking)" -ForegroundColor White
    Write-Host "  4. The app will create/use 'assets' folder in the same directory" -ForegroundColor White
    Write-Host ""
    
    # Prompt to create release
    Write-Host "========================================" -ForegroundColor Cyan
    $createRelease = Read-Host "Would you like to save this build as a release? (Y/N)"
    
    if ($createRelease -eq "Y" -or $createRelease -eq "y") {
        # Extract version from main.py
        $mainPyContent = Get-Content ".\main.py" -Raw
        if ($mainPyContent -match 'VERSION\s*=\s*"([^"]+)"') {
            $version = $matches[1]
            Write-Host "Detected version: $version" -ForegroundColor Green
            
            # Create releases folder if it doesn't exist
            if (-not (Test-Path ".\releases")) {
                New-Item -ItemType Directory -Path ".\releases" | Out-Null
                Write-Host "Created releases folder" -ForegroundColor Green
            }
            
            # Copy exe with versioned filename
            $releaseFileName = "SecureAssetInventoryTool_v$version.exe"
            $releasePath = ".\releases\$releaseFileName"
            
            Copy-Item ".\dist\SecureAssetInventoryTool.exe" -Destination $releasePath
            Write-Host ""
            Write-Host "Release saved: $releasePath" -ForegroundColor Green
            
            # Show release file size
            $releaseSize = (Get-Item $releasePath).Length
            $releaseSizeMB = [math]::Round($releaseSize / 1MB, 2)
            Write-Host "Release size: $releaseSizeMB MB" -ForegroundColor Cyan
            Write-Host ""
        } else {
            Write-Host "ERROR: Could not extract version from main.py" -ForegroundColor Red
            Write-Host "Please ensure VERSION constant is defined in main.py" -ForegroundColor Yellow
        }
    } else {
        Write-Host "Release not created." -ForegroundColor Yellow
    }
    Write-Host ""
} else {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "Build failed!" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "Check the error messages above." -ForegroundColor Yellow
    exit 1
}
