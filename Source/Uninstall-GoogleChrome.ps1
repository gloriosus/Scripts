$RootFolder = (Split-Path -Path $PSScriptRoot -Parent)

Import-Module -Name "$($RootFolder)\Modules\Jarnet.PowerShell.Utility"

$SystemType = (Get-CimInstance -ClassName Win32_ComputerSystem).SystemType
$SystemType -Match "x(?<arch>64|86)-based PC"
$Arch = $Matches.arch

$RegistryPath = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"

if ($Arch -eq "86") {
    $RegistryPath = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
}

$GoogleChromePath = (Get-ChildItem -Path $RegistryPath | Get-ItemProperty | Where-Object DisplayName -like "*Google Chrome*").PSPath
$UninstallString = (Get-ItemProperty -Path $GoogleChromePath).UninstallString
$UninstallString -Match "(?<base>.*)--verbose-logging"
$UninstallString = $Matches.base + "--force-uninstall"
$UninstallString -Match "(?<path>`".*`")\s(?<args>.*)"

Remove-ScService -Name "GoogleChromeElevationService", "gupdate", "gupdatem" -Timeout 30
Stop-Process -Name "setup", "chrome", "GoogleUpdate" -Force -ErrorAction SilentlyContinue | Wait-Process -Timeout 30
Start-Process -FilePath $Matches.path -ArgumentList $Matches.args -ErrorAction SilentlyContinue | Wait-Process -Timeout 180
Unregister-ScheduledTask -TaskName "GoogleUpdateTask*" -Confirm:$false -ErrorAction SilentlyContinue

$UserName = (Get-CimInstance -Class Win32_ComputerSystem).UserName.Split('\')[1]
$Sid = (Get-ChildItem -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" | Get-ItemProperty | Where-Object ProfileImagePath -like "*$($UserName)").PSChildName
Get-ChildItem -Path "Registry::HKEY_USERS\$($Sid)\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" | Get-ItemProperty | Where-Object {(($_.DisplayName -eq "Презентация") -or ($_.DisplayName -eq "YouTube") -or ($_.DisplayName -eq "Gmail") -or ($_.DisplayName -eq "Документы") -or ($_.DisplayName -eq "Таблица") -or ($_.DisplayName -eq "Google Диск"))} | Select-Object PSPath | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

$Paths = @(
    "C:\Program Files\Google",
    "C:\Program Files (x86)\Google",
    "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Google Chrome.lnk",
    "C:\Users\$($UserName)\AppData\Local\Google",
    "C:\Users\$($UserName)\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\Google Chrome.lnk",
    "C:\Users\$($UserName)\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Google Chrome.lnk",
    "C:\Users\$($UserName)\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Приложения Chrome",
    "C:\Users\Public\Desktop\Google Chrome.lnk"
)

Remove-Item -Path $Paths -Recurse -Force -ErrorAction SilentlyContinue