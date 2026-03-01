param(
    [int]$BridgePort = 8787,
    [string]$VercelOrigin = "",
    [string]$BridgeKey = "",
    [string]$NgrokDomain = "",
    [string]$NgrokPath = "",
    [string]$NgrokConfigPath = ""
)

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot

if ([string]::IsNullOrWhiteSpace($NgrokConfigPath)) {
    $NgrokConfigPath = Join-Path $PSScriptRoot "ngrok.yml"
}

function Resolve-NgrokPath {
    param([string]$ExplicitPath)

    if (-not [string]::IsNullOrWhiteSpace($ExplicitPath)) {
        if (Test-Path $ExplicitPath) { return $ExplicitPath }
        throw "Provided NgrokPath does not exist: $ExplicitPath"
    }

    try {
        $cmd = Get-Command ngrok -ErrorAction Stop
        if ($cmd -and $cmd.Source) { return $cmd.Source }
    } catch {
        # Continue to fallback paths.
    }

    $fallbacks = @(
        "$projectRoot\\tools\\ngrok.exe",
        "$projectRoot\\..\\ngrok-bin\\ngrok.exe",
        "C:\\form137application\\ngrok-bin\\ngrok.exe",
        "$env:USERPROFILE\\ngrok.exe",
        "$env:LOCALAPPDATA\\ngrok\\ngrok.exe",
        "C:\\ngrok\\ngrok.exe"
    )
    foreach ($candidate in $fallbacks) {
        if (Test-Path $candidate) { return $candidate }
    }

    return $null
}

try {
    $resolvedNgrokPath = Resolve-NgrokPath -ExplicitPath $NgrokPath
} catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}

if (-not $resolvedNgrokPath) {
    Write-Host "ngrok is not installed or not in PATH." -ForegroundColor Red
    Write-Host "Install: https://ngrok.com/download" -ForegroundColor Yellow
    Write-Host "Then run: ngrok config add-authtoken <YOUR_TOKEN>" -ForegroundColor Yellow
    Write-Host "Or pass -NgrokPath to this script." -ForegroundColor Yellow
    exit 1
}

if ([string]::IsNullOrWhiteSpace($BridgeKey)) {
    $BridgeKey = [Guid]::NewGuid().ToString("N")
}

$env:BRIDGE_PORT = "$BridgePort"
$env:BRIDGE_API_KEY = $BridgeKey
if (-not [string]::IsNullOrWhiteSpace($VercelOrigin)) {
    $env:BRIDGE_ALLOW_ORIGIN = $VercelOrigin
}

Write-Host "Starting Excel bridge on port $BridgePort..." -ForegroundColor Cyan
$bridgeArgs = @(
    "-NoExit",
    "-Command",
    "cd '$projectRoot'; npm run bridge:excel"
)
$bridgeProcess = Start-Process -FilePath "powershell.exe" -ArgumentList $bridgeArgs -PassThru

Start-Sleep -Seconds 2

Write-Host "Starting ngrok tunnel..." -ForegroundColor Cyan
$ngrokArgs = @("http", "$BridgePort", "--config", "$NgrokConfigPath")
if (-not [string]::IsNullOrWhiteSpace($NgrokDomain)) {
    $ngrokArgs += @("--domain", $NgrokDomain)
}

# Clear broken proxy env vars for ngrok connectivity, then restore.
$proxyVars = @("HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY", "http_proxy", "https_proxy", "all_proxy")
$savedProxy = @{}
foreach ($name in $proxyVars) {
    $value = [Environment]::GetEnvironmentVariable($name, "Process")
    if ($null -ne $value) {
        $savedProxy[$name] = $value
        [Environment]::SetEnvironmentVariable($name, $null, "Process")
    }
}

$ngrokProcess = Start-Process -FilePath $resolvedNgrokPath -ArgumentList $ngrokArgs -PassThru

foreach ($entry in $savedProxy.GetEnumerator()) {
    [Environment]::SetEnvironmentVariable($entry.Key, $entry.Value, "Process")
}

Write-Host "Waiting for ngrok public URL..." -ForegroundColor Cyan
$publicUrl = $null
for ($i = 0; $i -lt 30; $i++) {
    Start-Sleep -Seconds 1
    try {
        $resp = Invoke-RestMethod -Uri "http://127.0.0.1:4040/api/tunnels" -Method Get
        $httpsTunnel = $resp.tunnels | Where-Object { $_.proto -eq "https" } | Select-Object -First 1
        if ($httpsTunnel) {
            $publicUrl = $httpsTunnel.public_url
            break
        }
    } catch {
        # keep waiting
    }
}

Write-Host ""
Write-Host "Bridge process ID: $($bridgeProcess.Id)" -ForegroundColor Green
Write-Host "ngrok process ID:  $($ngrokProcess.Id)" -ForegroundColor Green
Write-Host ""

if ($publicUrl) {
    Write-Host "ngrok public URL: $publicUrl" -ForegroundColor Green
    Write-Host ""
    Write-Host "Set these in Vercel Environment Variables:" -ForegroundColor Yellow
    Write-Host "VITE_EXCEL_BRIDGE_URL=$publicUrl" -ForegroundColor White
    Write-Host "VITE_EXCEL_BRIDGE_KEY=$BridgeKey" -ForegroundColor White
    if (-not [string]::IsNullOrWhiteSpace($VercelOrigin)) {
        Write-Host ""
        Write-Host "Bridge CORS locked to: $VercelOrigin" -ForegroundColor Green
    } else {
        Write-Host ""
        Write-Host "Warning: BRIDGE_ALLOW_ORIGIN not set. CORS currently not restricted." -ForegroundColor Yellow
    }
} else {
    Write-Host "Could not read ngrok URL from local API." -ForegroundColor Yellow
    Write-Host "Open http://127.0.0.1:4040 to copy the tunnel URL." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "To stop:" -ForegroundColor Cyan
Write-Host "Stop-Process -Id $($bridgeProcess.Id),$($ngrokProcess.Id)" -ForegroundColor White
