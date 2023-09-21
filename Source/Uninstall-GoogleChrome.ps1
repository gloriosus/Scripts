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

Remove-ScService -Name "GoogleChromeElevationService" -Timeout 30
Stop-Process -Name "setup", "chrome" -Force -ErrorAction SilentlyContinue | Wait-Process -Timeout 30
Invoke-Expression -Command "& $($UninstallString)"

$UserName = (Get-CimInstance -Class Win32_ComputerSystem).UserName.Split('\')[1]
$Paths = @(
    "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Google Chrome.lnk",
    "C:\Users\$($UserName)\AppData\Local\Google",
    "C:\Users\$($UserName)\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\Google Chrome.lnk",
    "C:\Users\$($UserName)\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Google Chrome.lnk",
    "C:\Users\Public\Desktop\Google Chrome.lnk"
)

Remove-Item -Path $Paths -Recurse -Force -ErrorAction SilentlyContinue