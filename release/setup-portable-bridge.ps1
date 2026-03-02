param(
    [string]$AuthToken = "3AKzmN3su2n50gfv8qmMf8ZqpHj_275QXanCMyU8sVv46EYVp",
    [string]$BridgeKey = "sf10-bridge-2024"
)

$ErrorActionPreference = "Stop"
$scriptDir = $PSScriptRoot
$ngrokPath = Join-Path $scriptDir "ngrok.exe"
$configPath = Join-Path $scriptDir "ngrok.yml"

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "       PORTABLE SF10 BRIDGE SETUP               " -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan

# 0. Kill Zombie Processes
Write-Host "Checking for existing processes..." -ForegroundColor Yellow
Get-Process -Name "excel-bridge-server" -ErrorAction SilentlyContinue | Stop-Process -Force
Get-Process -Name "ngrok" -ErrorAction SilentlyContinue | Stop-Process -Force

# Also find and kill anything occupying port 8787 specifically
$portProcess = Get-NetTCPConnection -LocalPort 8787 -ErrorAction SilentlyContinue
if ($portProcess) {
    Write-Host "Clearing port 8787 (PID: $($portProcess[0].OwningProcess))..." -ForegroundColor Cyan
    Stop-Process -Id $portProcess[0].OwningProcess -Force -ErrorAction SilentlyContinue
}

# 1. Check if ngrok exists
if (-not (Test-Path $ngrokPath)) {
    Write-Host "ERROR: ngrok.exe not found in this folder!" -ForegroundColor Red
    Write-Host "Please download it from https://ngrok.com/download" -ForegroundColor Cyan
    Write-Host "and place the 'ngrok.exe' file here." -ForegroundColor Cyan
    Write-Host "------------------------------------------------"
    cmd /c pause
    exit 1
}

# 2. Add Auth Token if provided
if (-not [string]::IsNullOrWhiteSpace($AuthToken)) {
    Write-Host "Configuring ngrok auth token..." -ForegroundColor Cyan
    try {
        & $ngrokPath config add-authtoken $AuthToken --config $configPath
        Write-Host "Auth token added successfully!" -ForegroundColor Green
    }
    catch {
        Write-Host "Warning: Failed to add auth token. It might already be set." -ForegroundColor Yellow
    }
}

$bridgeProcess = $null
$ngrokProcess = $null

try {
    # 3. Start the Bridge and Ngrok
    Write-Host "Starting Excel Bridge Server..." -ForegroundColor Green
    $bridgePath = Join-Path $scriptDir "excel-bridge-server.exe"
    if (-not (Test-Path $bridgePath)) {
        throw "excel-bridge-server.exe not found in this folder."
    }
    
    # Set the Key in the environment so the bridge picks it up
    $env:BRIDGE_API_KEY = $BridgeKey
    
    # Start Bridge
    $bridgeProcess = Start-Process -FilePath $bridgePath -WorkingDirectory $scriptDir -PassThru
    
    # Start Ngrok
    Write-Host "Starting Ngrok Tunnel on port 8787..." -ForegroundColor Green
    $ngrokArgs = @("http", "8787", "--config", $configPath)
    $ngrokProcess = Start-Process -FilePath $ngrokPath -ArgumentList $ngrokArgs -WorkingDirectory $scriptDir -PassThru

    Write-Host "------------------------------------------------" -ForegroundColor Cyan
    Write-Host "Processes started in separate windows!" -ForegroundColor Green
    Write-Host "1. Bridge Server Window: Handles Excel logic."
    Write-Host "2. Ngrok Window: Provides your Public URL."
    Write-Host ""
    Write-Host "BRIDGE API KEY: $BridgeKey" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Monitoring... (Close THIS window to stop both)" -ForegroundColor White
    Write-Host "------------------------------------------------"

    # Monitor
    while ($true) {
        if ($bridgeProcess -and $bridgeProcess.HasExited) { 
            Write-Host "`n[!ERROR] Excel Bridge Server has stopped!" -ForegroundColor Red
            Write-Host "Exit Code: $($bridgeProcess.ExitCode)" -ForegroundColor Yellow
            throw "Bridge binary exited."
        }
        if ($ngrokProcess -and $ngrokProcess.HasExited) { 
            Write-Host "`n[!ERROR] Ngrok has stopped!" -ForegroundColor Red
            Write-Host "Exit Code: $($ngrokProcess.ExitCode)" -ForegroundColor Yellow
            throw "Ngrok exited."
        }
        Start-Sleep -Seconds 2
    }
}
catch {
    if ($_.Exception.Message -notlike "*stopped*") {
        Write-Host "`n[ERROR] $($_.Exception.Message)" -ForegroundColor Red
    }
}
finally {
    Write-Host "`nCleaning up processes..." -ForegroundColor Yellow
    if ($bridgeProcess -and -not $bridgeProcess.HasExited) { 
        Stop-Process -Id $bridgeProcess.Id -Force -ErrorAction SilentlyContinue 
    }
    if ($ngrokProcess -and -not $ngrokProcess.HasExited) { 
        Stop-Process -Id $ngrokProcess.Id -Force -ErrorAction SilentlyContinue 
    }
    Write-Host "`nCleanup complete." -ForegroundColor Gray
    Write-Host "Press any key to exit this window..." -ForegroundColor White
    $null = [Console]::ReadKey()
}
