$UserName = (Get-CimInstance -Class Win32_ComputerSystem).UserName.Split('\')[1]
$ConfigFile = "C:\Users\$($UserName)\AppData\Roaming\Organizer\installs.ini"
# If newer version is out, change the installation ID. It can be obtained by installing R7 on a test machine
$InstallationId = "[EEF607DE8263F914]"

Write-Output $InstallationId | Out-File $ConfigFile -Encoding UTF8
Write-Output "Default=Profiles/$($UserName).default" | Out-File $ConfigFile -Encoding UTF8 -Append
Write-Output "Locked=1" | Out-File $ConfigFile -Encoding UTF8 -Append