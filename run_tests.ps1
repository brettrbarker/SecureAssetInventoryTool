# run_tests.ps1 — Run all unit tests for the Secure Asset Inventory Tool.
#
# Usage:
#   .\run_tests.ps1              # Run all tests
#   .\run_tests.ps1 -Verbose     # Run with verbose output
#   .\run_tests.ps1 -File tests/test_asset_database.py  # Run a single file
#
# Exit codes:
#   0  All tests passed
#   1  One or more tests failed

param(
    [switch]$Verbose,
    [string]$File = ""
)

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " Secure Asset Inventory Tool — Unit Tests" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

# ── Locate Python ─────────────────────────────────────────────────────────────
$pythonExe = $null

if ($env:VIRTUAL_ENV) {
    $pythonExe = Join-Path $env:VIRTUAL_ENV "Scripts\python.exe"
} elseif (Test-Path ".\.venv\Scripts\python.exe") {
    $pythonExe = ".\.venv\Scripts\python.exe"
} else {
    $pythonExe = "python"
}

Write-Host "Python: $pythonExe" -ForegroundColor Gray

# ── Determine test target ──────────────────────────────────────────────────────
if ($File -ne "") {
    $testTarget = $File
    Write-Host "Running: $testTarget" -ForegroundColor Yellow
} else {
    $testTarget = "tests"
    Write-Host "Running: all tests in .\tests\" -ForegroundColor Yellow
}

Write-Host ""

# ── Build argument list ────────────────────────────────────────────────────────
$args = @("-m", "unittest", "discover")

if ($File -ne "") {
    # Run a specific file rather than discovery
    $args = @("-m", "unittest", $File.Replace("\", "/").Replace(".py", "").Replace("/", "."))
} else {
    $args = @("-m", "unittest", "discover", "-s", "tests", "-p", "test_*.py")
}

if ($Verbose) {
    $args += "-v"
}

# ── Execute ────────────────────────────────────────────────────────────────────
& $pythonExe @args

$exitCode = $LASTEXITCODE

Write-Host ""
if ($exitCode -eq 0) {
    Write-Host "==========================================" -ForegroundColor Green
    Write-Host " All tests PASSED" -ForegroundColor Green
    Write-Host "==========================================" -ForegroundColor Green
} else {
    Write-Host "==========================================" -ForegroundColor Red
    Write-Host " Tests FAILED (exit code $exitCode)" -ForegroundColor Red
    Write-Host "==========================================" -ForegroundColor Red
}

exit $exitCode
