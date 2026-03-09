# Form 137 Excel Bridge Installer
# Run this from the form137-react root directory

$InstallDir = "C:\Form137Bridge"
$ExeName = "excel-bridge-server.exe"
$SourceExe = "$PSScriptRoot\..\release\$ExeName"
$TemplateSource = "$PSScriptRoot\..\public\Form 137-SHS-BLANK.xlsx"
$ProtocolName = "sf10-bridge"

Write-Host "--- Form 137 Bridge Installer ---" -ForegroundColor Cyan

# 1. Create Installation Directory
if (-not (Test-Path $InstallDir)) {
    Write-Host "Creating $InstallDir..."
    New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
}

# 2. Build the Bridge Executable (if not already built or forced)
Write-Host "Building bridge executable..."
npm run build:bridge

if (-not (Test-Path $SourceExe)) {
    Write-Error "Failed to build bridge executable at $SourceExe"
    exit 1
}

# 3. Copy files to C:\Form137Bridge
Write-Host "Copying files to $InstallDir..."
Copy-Item -Path $SourceExe -Destination "$InstallDir\$ExeName" -Force
if (Test-Path $TemplateSource) {
    Copy-Item -Path $TemplateSource -Destination "$InstallDir\Form 137-SHS-BLANK.xlsx" -Force
}

# 4. Register Custom Protocol (sf10-bridge://)
Write-Host "Registering protocol $ProtocolName..."
$RegPath = "HKCU:\Software\Classes\$ProtocolName"
if (Test-Path $RegPath) { Remove-Item $RegPath -Recurse -Force }
New-Item -Path $RegPath -Value "URL:Form 137 Bridge Protocol" -Force | Out-Null
New-ItemProperty -Path $RegPath -Name "URL Protocol" -Value "" -Force | Out-Null
$ShellPath = New-Item -Path "$RegPath\shell\open\command" -Force
Set-Item -Path "$RegPath\shell\open\command" -Value "`"$InstallDir\$ExeName`" `"%1`""

# 5. Add to Windows Startup
Write-Host "Adding to Windows Startup..."
$StartupReg = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
New-ItemProperty -Path $StartupReg -Name "Form137Bridge" -Value "`"$InstallDir\$ExeName`"" -PropertyType String -Force | Out-Null

# 6. Open Firewall Port (Requires Admin, but we can try)
Write-Host "Adding Firewall Rule (Port 8787)..."
try {
    netsh advfirewall firewall add rule name="Form 137 Excel Bridge" dir=in action=allow protocol=TCP localport=8787
}
catch {
    Write-Warning "Failed to add firewall rule. If you are not Admin, please allow port 8787 manually."
}

Write-Host "`nInstallation Complete!" -ForegroundColor Green
Write-Host "Location: $InstallDir"
Write-Host "Protocol: $ProtocolName://"
Write-Host "The bridge will now start automatically with Windows.`n"

# Start the bridge now
Write-Host "Starting bridge..."
Start-Process -FilePath "$InstallDir\$ExeName" -WorkingDirectory $InstallDir
