# Form 137 Excel Bridge - Portable Installer
# Designed to be run from Downloads or any folder containing the .exe and .xlsx

# --- 0. Request Admin Privileges ---
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Requesting Administrator privileges..." -ForegroundColor Yellow
    $arguments = "-ExecutionPolicy Bypass -File `"$PSCommandPath`""
    Start-Process powershell.exe -ArgumentList $arguments -Verb RunAs
    exit
}

$InstallDir = "C:\Form137Bridge"
$ExeName = "excel-bridge-server.exe" 
$NgrokExe = "ngrok.exe"
$TemplateName = "Form137_Template.xlsx"
$ProtocolName = "sf10-bridge"

Write-Host "--- Form 137 Python Bridge & Ngrok Setup ---" -ForegroundColor Cyan

try {
    # 1. Verify files exist in the same folder
    $RequiredFiles = @($ExeName, $TemplateName)
    foreach ($f in $RequiredFiles) {
        if (-not (Test-Path "$PSScriptRoot\$f")) {
            throw "Could not find $f in the current folder ($PSScriptRoot). Please ensure all files are extracted together."
        }
    }

    # 2. Create Installation Directory
    if (-not (Test-Path $InstallDir)) {
        Write-Host "Creating $InstallDir..."
        New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
    }

    # 3. Copy files to C:\Form137Bridge
    Write-Host "Copying bridge files to $InstallDir..."
    Copy-Item -Path "$PSScriptRoot\$ExeName" -Destination "$InstallDir\$ExeName" -Force
    Copy-Item -Path "$PSScriptRoot\$TemplateName" -Destination "$InstallDir\$TemplateName" -Force
    
    if (Test-Path "$PSScriptRoot\$NgrokExe") {
        Write-Host "Copying Ngrok..."
        Copy-Item -Path "$PSScriptRoot\$NgrokExe" -Destination "$InstallDir\$NgrokExe" -Force
    }

    # 4. Create a Startup Batch File that handles both Bridge and Ngrok (if present)
    $StartBatch = "$InstallDir\run-bridge.bat"
    $BatchContent = @"
@echo off
title Form 137 Excel Bridge
cd /d "%~dp0"
echo Starting Excel Bridge Server...
start "" "excel-bridge-server.exe"
if exist "ngrok.exe" (
    echo Starting Ngrok Tunnel...
    start "" "ngrok.exe" http 8787 --log=stdout
)
"@
    Set-Content -Path $StartBatch -Value $BatchContent

    # 5. Register Custom Protocol (pointing to batch file)
    Write-Host "Registering protocol ${ProtocolName}..."
    $RegPath = "HKCU:\Software\Classes\${ProtocolName}"
    if (Test-Path $RegPath) { Remove-Item $RegPath -Recurse -Force }
    New-Item -Path $RegPath -Value "URL:Form 137 Bridge Protocol" -Force | Out-Null
    New-ItemProperty -Path $RegPath -Name "URL Protocol" -Value "" -Force | Out-Null
    New-Item -Path "$RegPath\shell\open\command" -Force | Out-Null
    Set-Item -Path "$RegPath\shell\open\command" -Value "cmd.exe /c `"$StartBatch`" `"%1`""

    # 6. Add to Windows Startup
    Write-Host "Adding to Windows Startup..."
    $StartupReg = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
    New-ItemProperty -Path $StartupReg -Name "Form137Bridge" -Value "`"$StartBatch`"" -PropertyType String -Force | Out-Null

    # 7. Open Firewall Port
    Write-Host "Adding Firewall Rule (Port 8787)..."
    netsh advfirewall firewall add rule name="Form 137 Excel Bridge" dir=in action=allow protocol=TCP localport=8787 | Out-Null

    Write-Host "`nInstallation Complete!" -ForegroundColor Green
    Write-Host "Location: ${InstallDir}"
    Write-Host "The bridge and tunnel will now start automatically.`n"

    # Start it now
    Start-Process -FilePath $StartBatch -WorkingDirectory $InstallDir
}
catch {
    $errorMessage = $_.Exception.Message
    Write-Host "`nFATAL ERROR: $errorMessage" -ForegroundColor Red
    Write-Host "Press any key to exit..."
    $null = [Console]::ReadKey($true)
}
