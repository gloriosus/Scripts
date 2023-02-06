$SystemType = (Get-CimInstance -ClassName Win32_ComputerSystem).SystemType
$SystemType -Match "x(?<arch>64|86)-based PC"
$Arch = $Matches.arch

$RegistryPath = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"

if ($Arch -eq "86") {
    $RegistryPath = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
}

$GoogleChromePath = (Get-ChildItem -Path $RegistryPath | Get-ItemProperty | Where-Object DisplayName -like "*Google Chrome*").PSPath
$UninstallString = Get-ItemProperty -Path $GoogleChromePath -Name "UninstallString"
$UninstallString -match "`"(?<file>.*)`"(.?)(?<args>.*)"
Start-Process -FilePath $Matches.file -ArgumentList $Matches.args