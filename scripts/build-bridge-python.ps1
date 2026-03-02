Write-Host "================================================" -ForegroundColor Cyan
Write-Host "       SF10 PYTHON BRIDGE BUILDER               " -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan

# 1. Install Dependencies
Write-Host "Installing/Updating dependencies..." -ForegroundColor Yellow
python -m pip install -r scripts/requirements.txt

# 2. Build EXE using PyInstaller
Write-Host "Building standalone executable..." -ForegroundColor Yellow
python -m PyInstaller --onefile --distpath release scripts/excel-bridge-server.py

if ($LASTEXITCODE -eq 0) {
    Write-Host "`nSUCCESS: release/excel-bridge-server.exe updated!" -ForegroundColor Green
}
else {
    Write-Host "`nERROR: Build failed." -ForegroundColor Red
}

Write-Host "`nPress any key to exit..."
cmd /c pause
