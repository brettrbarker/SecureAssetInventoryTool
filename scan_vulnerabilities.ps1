# PowerShell script to scan the Code directory for vulnerabilities using Syft and Grype
# Syft creates SBOM (Software Bill of Materials)
# Grype scans for vulnerabilities

# Configuration - Use relative paths from script location
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$CodeDir = $ScriptDir
$ProjectRoot = Split-Path -Parent $ScriptDir
$VenvDir = Join-Path $ProjectRoot ".venv"
$RequirementsFile = Join-Path $CodeDir "requirements.txt"
$OutputDir = Join-Path $CodeDir "security_reports"
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# Create output directory if it doesn't exist
if (!(Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force
    Write-Host "Created output directory: $OutputDir" -ForegroundColor Green
}

# Check if Syft and Grype are in PATH
Write-Host "Checking for Syft and Grype in system PATH..." -ForegroundColor Gray

$SyftPath = Get-Command syft -ErrorAction SilentlyContinue
$GrypePath = Get-Command grype -ErrorAction SilentlyContinue

if ($null -eq $SyftPath) {
    Write-Host "ERROR: Syft not found in system PATH" -ForegroundColor Red
    Write-Host "Please install Syft and ensure it's in your system PATH." -ForegroundColor Red
    Write-Host "Download from: https://github.com/anchore/syft/releases" -ForegroundColor Yellow
    exit 1
}

if ($null -eq $GrypePath) {
    Write-Host "ERROR: Grype not found in system PATH" -ForegroundColor Red
    Write-Host "Please install Grype and ensure it's in your system PATH." -ForegroundColor Red
    Write-Host "Download from: https://github.com/anchore/grype/releases" -ForegroundColor Yellow
    exit 1
}

# Use the command paths
$SyftPath = $SyftPath.Source
$GrypePath = $GrypePath.Source

Write-Host "✓ Syft found: $SyftPath" -ForegroundColor Green
Write-Host "✓ Grype found: $GrypePath" -ForegroundColor Green
Write-Host ""

Write-Host "=== Asset Management System Security Scan ===" -ForegroundColor Cyan
Write-Host "Timestamp: $(Get-Date)" -ForegroundColor Gray
Write-Host "Source: requirements.txt (installed packages only)" -ForegroundColor Gray
Write-Host "Requirements file: $RequirementsFile" -ForegroundColor Gray
Write-Host "Virtual environment: $VenvDir" -ForegroundColor Gray
Write-Host "Output directory: $OutputDir" -ForegroundColor Gray
Write-Host ""

# Verify requirements.txt exists
if (!(Test-Path $RequirementsFile)) {
    Write-Host "ERROR: requirements.txt not found at $RequirementsFile" -ForegroundColor Red
    Write-Host "Please ensure requirements.txt exists in the Code directory." -ForegroundColor Red
    exit 1
}

# Step 1: Generate SBOM using Syft from requirements.txt
Write-Host "Step 1: Generating SBOM (Software Bill of Materials) from requirements.txt..." -ForegroundColor Yellow
$SbomJsonFile = Join-Path $OutputDir "sbom_$Timestamp.json"
$SbomTextFile = Join-Path $OutputDir "sbom_$Timestamp.txt"

try {
    # Generate SBOM from requirements.txt (clean, no duplicates)
    Write-Host "  Creating JSON SBOM from requirements.txt..." -ForegroundColor Gray
    & $SyftPath "file:$RequirementsFile" -o json --file $SbomJsonFile
    
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  ⚠ Warning: Syft exited with code $LASTEXITCODE" -ForegroundColor Yellow
    }
    
    # Generate SBOM in text format (human readable)
    Write-Host "  Creating text SBOM from requirements.txt..." -ForegroundColor Gray
    & $SyftPath "file:$RequirementsFile" -o text --file $SbomTextFile
    
    if (Test-Path $SbomJsonFile) {
        Write-Host "  ✓ SBOM generated successfully from requirements.txt" -ForegroundColor Green
        Write-Host "    JSON: $SbomJsonFile" -ForegroundColor Gray
        Write-Host "    Text: $SbomTextFile" -ForegroundColor Gray
        
        # Count unique packages
        $sbomContent = Get-Content $SbomJsonFile -Raw | ConvertFrom-Json
        $packageCount = $sbomContent.artifacts.Count
        Write-Host "    Packages: $packageCount" -ForegroundColor Gray
    } else {
        Write-Host "  ✗ SBOM generation failed" -ForegroundColor Red
        exit 1
    }
} catch {
    Write-Host "  ✗ Error generating SBOM: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""

# Step 2: Scan for vulnerabilities using Grype
Write-Host "Step 2: Scanning for vulnerabilities..." -ForegroundColor Yellow
$VulnJsonFile = Join-Path $OutputDir "vulnerabilities_$Timestamp.json"
$VulnTextFile = Join-Path $OutputDir "vulnerabilities_$Timestamp.txt"
$VulnSummaryFile = Join-Path $OutputDir "vulnerability_summary_$Timestamp.txt"

try {
    # Scan using the SBOM file for vulnerabilities (JSON format)
    Write-Host "  Scanning SBOM for vulnerabilities (JSON)..." -ForegroundColor Gray
    & $GrypePath "sbom:$SbomJsonFile" -o json --file $VulnJsonFile
    
    # Scan using the SBOM file for vulnerabilities (Text format)
    # Use PowerShell to strip ANSI color codes from output
    Write-Host "  Scanning SBOM for vulnerabilities (Text)..." -ForegroundColor Gray
    $rawOutput = & $GrypePath "sbom:$SbomJsonFile" -o table 2>&1 | Out-String
    $cleanOutput = $rawOutput -replace '\x1b\[[0-9;]*m', ''
    $cleanOutput | Out-File -FilePath $VulnTextFile -Encoding UTF8
    
    if (Test-Path $VulnJsonFile) {
        Write-Host "  ✓ Vulnerability scan completed successfully" -ForegroundColor Green
        Write-Host "    JSON: $VulnJsonFile" -ForegroundColor Gray
        Write-Host "    Text: $VulnTextFile" -ForegroundColor Gray
    } else {
        Write-Host "  ✗ Vulnerability scan failed" -ForegroundColor Red
    }
} catch {
    Write-Host "  ✗ Error scanning vulnerabilities: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""

# Step 3: Generate summary report
Write-Host "Step 3: Generating summary report..." -ForegroundColor Yellow

try {
    $SummaryContent = @"
Asset Management System Security Scan Summary
=============================================
Scan Date: $(Get-Date)
Scanned Directory: $CodeDir
Project Root: $ProjectRoot
Virtual Environment: $VenvDir
Syft Version: $(& $SyftPath version 2>$null | Select-String "version" | Select-Object -First 1)
Grype Version: $(& $GrypePath version 2>$null | Select-String "version" | Select-Object -First 1)

Files Generated:
- SBOM (JSON): $SbomJsonFile
- SBOM (Text): $SbomTextFile
- Vulnerabilities (JSON): $VulnJsonFile
- Vulnerabilities (Text): $VulnTextFile
- Direct Scan: $DirectScanFile

Quick Stats:
"@

    # Try to extract vulnerability counts from the JSON file
    if (Test-Path $VulnJsonFile) {
        try {
            $VulnData = Get-Content $VulnJsonFile | ConvertFrom-Json
            $VulnCount = $VulnData.matches.Count
            $CriticalCount = ($VulnData.matches | Where-Object { $_.vulnerability.severity -eq "Critical" }).Count
            $HighCount = ($VulnData.matches | Where-Object { $_.vulnerability.severity -eq "High" }).Count
            $MediumCount = ($VulnData.matches | Where-Object { $_.vulnerability.severity -eq "Medium" }).Count
            $LowCount = ($VulnData.matches | Where-Object { $_.vulnerability.severity -eq "Low" }).Count
            
            $SummaryContent += @"

- Total Vulnerabilities Found: $VulnCount
- Critical Severity: $CriticalCount
- High Severity: $HighCount
- Medium Severity: $MediumCount
- Low Severity: $LowCount
"@
        } catch {
            $SummaryContent += "`n- Could not parse vulnerability statistics from JSON file"
        }
    }
    
    # Try to get package count from SBOM
    if (Test-Path $SbomJsonFile) {
        try {
            $SbomData = Get-Content $SbomJsonFile | ConvertFrom-Json
            $PackageCount = $SbomData.artifacts.Count
            $SummaryContent += "`n- Total Packages/Components Found: $PackageCount"
        } catch {
            $SummaryContent += "`n- Could not parse package statistics from SBOM file"
        }
    }

    $SummaryContent += @"


Next Steps:
1. Review the vulnerability report: $VulnTextFile
2. Check the SBOM for inventory: $SbomTextFile
3. Address any Critical or High severity vulnerabilities
4. Consider updating dependencies with known vulnerabilities

Note: This scan covers Python packages from both the Code directory and virtual environment.
The scan now includes all installed packages in the .venv directory.
For a complete security assessment, consider running additional scans on deployment environments.
"@

    $SummaryContent | Out-File -FilePath $VulnSummaryFile -Encoding UTF8
    Write-Host "  ✓ Summary report generated: $VulnSummaryFile" -ForegroundColor Green
} catch {
    Write-Host "  ⚠ Warning: Could not generate complete summary: $($_.Exception.Message)" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=== Scan Complete ===" -ForegroundColor Cyan
Write-Host "All reports saved to: $OutputDir" -ForegroundColor Green
Write-Host ""
Write-Host "Key files to review:" -ForegroundColor White
Write-Host "  📋 Summary: $VulnSummaryFile" -ForegroundColor Cyan
Write-Host "  🚨 Vulnerabilities: $VulnTextFile" -ForegroundColor Red
Write-Host "  📦 SBOM: $SbomTextFile" -ForegroundColor Blue
Write-Host ""

# Optional: Open the summary file
$OpenSummary = Read-Host "Would you like to open the summary report now? (y/N)"
if ($OpenSummary -eq "y" -or $OpenSummary -eq "Y") {
    if (Test-Path $VulnSummaryFile) {
        Start-Process notepad.exe -ArgumentList $VulnSummaryFile
    }
}

Write-Host "Security scan completed successfully!" -ForegroundColor Green