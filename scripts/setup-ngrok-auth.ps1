param(
    [Parameter(Mandatory = $true)]
    [string]$AuthToken,
    [string]$NgrokPath = "C:\form137application\ngrok-bin\ngrok.exe",
    [string]$NgrokConfigPath = ""
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $NgrokPath)) {
    Write-Host "ngrok executable not found: $NgrokPath" -ForegroundColor Red
    exit 1
}

if ([string]::IsNullOrWhiteSpace($NgrokConfigPath)) {
    $NgrokConfigPath = Join-Path $PSScriptRoot "ngrok.yml"
}

$configDir = Split-Path -Parent $NgrokConfigPath
if (-not (Test-Path $configDir)) {
    New-Item -ItemType Directory -Force $configDir | Out-Null
}

$proxyVars = @("HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY", "http_proxy", "https_proxy", "all_proxy")
$savedProxy = @{}
foreach ($name in $proxyVars) {
    $value = [Environment]::GetEnvironmentVariable($name, "Process")
    if ($null -ne $value) {
        $savedProxy[$name] = $value
        [Environment]::SetEnvironmentVariable($name, $null, "Process")
    }
}

try {
    & $NgrokPath config add-authtoken $AuthToken --config $NgrokConfigPath
    if ($LASTEXITCODE -ne 0) {
        throw "ngrok authtoken setup failed."
    }
    Write-Host "ngrok authtoken configured successfully." -ForegroundColor Green
} finally {
    foreach ($entry in $savedProxy.GetEnumerator()) {
        [Environment]::SetEnvironmentVariable($entry.Key, $entry.Value, "Process")
    }
}
