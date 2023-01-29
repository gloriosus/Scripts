# It is for uninstallation program on a user computer. Change this to a network path for uninstalling remotely
. "C:\Windows\Installer\Utility.ps1"

Disable-CloseButton
Write-Host "Идет удаление Р7-Офис. Органайзер..."

$DomainWithUsername = (Get-CimInstance -Class Win32_ComputerSystem).Username
$Username = $DomainWithUsername.Split('\')[1]

$Processes = @(
    "organizer"
)

$Paths = @(
    "C:\Program Files\R7-Office\organizer",
    "C:\Users\$($Username)\Desktop\Р7-Органайзер.lnk",
    "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\R7Organizer",
    "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Р7-Офис\Р7-Органайзер.lnk"
)

Close-Processes $Processes | Wait-Process
Start-Sleep -Seconds 5.0
Remove-Paths $Paths