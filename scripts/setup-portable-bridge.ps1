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
$SourceExe = "$PSScriptRoot\$ExeName"
$TemplateName = "Form 137-SHS-BLANK.xlsx"
$SourceTemplate = "$PSScriptRoot\$TemplateName"
$ProtocolName = "sf10-bridge"

Write-Host "--- Form 137 Bridge Portable Setup ---" -ForegroundColor Cyan

try {
    # 1. Verify files exist in the same folder
    if (-not (Test-Path $SourceExe)) {
        throw "Could not find $ExeName in the current folder ($PSScriptRoot). Please ensure the .exe is next to this script."
    }

    # 2. Create Installation Directory
    if (-not (Test-Path $InstallDir)) {
        Write-Host "Creating $InstallDir..."
        New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
    }

    # 3. Copy files to C:\Form137Bridge
    Write-Host "Copying files to $InstallDir..."
    Copy-Item -Path $SourceExe -Destination "$InstallDir\$ExeName" -Force
    
    if (Test-Path $SourceTemplate) {
        Copy-Item -Path $SourceTemplate -Destination "$InstallDir\$TemplateName" -Force
    }
    else {
        Write-Host "Warning: $TemplateName not found in current folder. Skipping template copy." -ForegroundColor Yellow
    }

    # 4. Register Custom Protocol (sf10-bridge://)
    Write-Host "Registering protocol ${ProtocolName}..."
    $RegPath = "HKCU:\Software\Classes\${ProtocolName}"
    if (Test-Path $RegPath) { Remove-Item $RegPath -Recurse -Force }
    New-Item -Path $RegPath -Value "URL:Form 137 Bridge Protocol" -Force | Out-Null
    New-ItemProperty -Path $RegPath -Name "URL Protocol" -Value "" -Force | Out-Null
    New-Item -Path "$RegPath\shell\open\command" -Force | Out-Null
    Set-Item -Path "$RegPath\shell\open\command" -Value "`"$InstallDir\$ExeName`" `"%1`""

    # 5. Add to Windows Startup
    Write-Host "Adding to Windows Startup..."
    $StartupReg = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
    New-ItemProperty -Path $StartupReg -Name "Form137Bridge" -Value "`"$InstallDir\$ExeName`"" -PropertyType String -Force | Out-Null

    # 6. Open Firewall Port
    Write-Host "Adding Firewall Rule (Port 8787)..."
    netsh advfirewall firewall add rule name="Form 137 Excel Bridge" dir=in action=allow protocol=TCP localport=8787 | Out-Null

    Write-Host "`nInstallation Complete!" -ForegroundColor Green
    Write-Host "Location: ${InstallDir}"
    Write-Host "Protocol: ${ProtocolName}://"
    Write-Host "The bridge will now start automatically with Windows.`n"

    # Start the bridge now
    Write-Host "Starting bridge..."
    Start-Process -FilePath "$InstallDir\$ExeName" -WorkingDirectory $InstallDir

}
catch {
    $errorMessage = $_.Exception.Message
    Write-Host "`nFATAL ERROR: $errorMessage" -ForegroundColor Red
    Write-Host "Please ensure you have permission to write to C:\ and modify the registry."
    Write-Host "Press any key to exit..."
    $null = [Console]::ReadKey($true)
}
