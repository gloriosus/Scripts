$SoftwarePath = ""
# If newer version of Organizer is out, just paste the name of the installation archive here.
# And do not change the original name of the archive because it is used in pattern matching
$InstallationArchive = "r7-office-organizer-68.5.0.5.ru.win64.zip"
$InstallationArchive -Match "(r7-office-organizer-)(?<version>.*)(.ru.win64.zip)"
$VersionNumber = $Matches.version
$UserName = (Get-CimInstance -Class Win32_ComputerSystem).UserName.Split('\')[1]
$ProfilePath = "C:\Users\$($UserName)\AppData\Roaming\Organizer\Profiles\$($UserName).default"

if (-Not (Test-Path $ProfilePath)){
    New-Item $ProfilePath -ItemType Directory
    powershell.exe -noprofile -executionpolicy bypass -file "$($PSScriptRoot)\Set-InstallsConfig.ps1"
    powershell.exe -noprofile -executionpolicy bypass -file "$($PSScriptRoot)\Set-ProfilesConfig.ps1"
    powershell.exe -noprofile -executionpolicy bypass -file "$($PSScriptRoot)\Set-PrefsConfig.ps1"
}

New-Item "C:\Program Files\R7-Office" -ItemType Directory -Force
Expand-Archive "$($SoftwarePath)\Офисные приложения\Р7-Офис\$($InstallationArchive)" -DestinationPath "C:\Program Files\R7-Office" -Force
Copy-Item -Path "$($PSScriptRoot)\Uninstall-R7Organizer.ps1" -Destination "C:\Windows\Installer"

$Msvc64 = wmic.exe product get name | Select-String -Pattern "(Microsoft Visual C\+\+ 2015 x64)(.*)"
if ($Msvc64.Length -eq 0){
    Start-Process -FilePath "$($SoftwarePath)\Системные компоненты\Microsoft Visual C++ Redistributables\2015\vc_redist.x64.exe" -ArgumentList "/install", "/quiet", "/norestart"
}

$Msvc86 = wmic.exe product get name | Select-String -Pattern "(Microsoft Visual C\+\+ 2015 x86)(.*)"
if ($Msvc86.Length -eq 0){
    Start-Process -FilePath "$($SoftwarePath)\Системные компоненты\Microsoft Visual C++ Redistributables\2015\vc_redist.x86.exe" -ArgumentList "/install", "/quiet", "/norestart"
}

New-Item -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\R7Organizer" -Force
Set-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\R7Organizer" -Name "DisplayIcon" -Value "C:\Program Files\R7-Office\organizer\chrome\icons\default\messengerWindow.ico" -Force
Set-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\R7Organizer" -Name "DisplayName" -Value "Р7-Органайзер" -Force
Set-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\R7Organizer" -Name "DisplayVersion" -Value $VersionNumber -Force
Set-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\R7Organizer" -Name "Publisher" -Value "r7office.ru" -Force
Set-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\R7Organizer" -Name "UninstallString" -Value "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noprofile -executionpolicy bypass -file C:\Windows\Installer\Uninstall-R7Organizer.ps1" -Force

$WshShell = New-Object -ComObject WScript.Shell

$DesktopShortcut = $WshShell.CreateShortcut("C:\Users\$($UserName)\Desktop\Р7-Органайзер.lnk")
$DesktopShortcut.TargetPath = "C:\Program Files\R7-Office\organizer\organizer.exe"
$DesktopShortcut.WorkingDirectory = "C:\Program Files\R7-Office\organizer\"
$DesktopShortcut.IconLocation = "C:\Program Files\R7-Office\organizer\organizer.exe"
$DesktopShortcut.Save()

$StartMenuShortcut = $WshShell.CreateShortcut("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Р7-Офис\Р7-Органайзер.lnk")
$StartMenuShortcut.TargetPath = "C:\Program Files\R7-Office\organizer\organizer.exe"
$StartMenuShortcut.WorkingDirectory = "C:\Program Files\R7-Office\organizer\"
$StartMenuShortcut.IconLocation = "C:\Program Files\R7-Office\organizer\organizer.exe"
$StartMenuShortcut.Save()