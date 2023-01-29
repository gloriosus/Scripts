$UserName = (Get-CimInstance -Class Win32_ComputerSystem).UserName.Split('\')[1]
$ConfigFile = "C:\Users\$($UserName)\AppData\Roaming\Organizer\profiles.ini"
# If newer version is out, change the installation ID. It can be obtained by installing R7 on a test machine
$InstallationId = "[InstallEEF607DE8263F914]"

Write-Output $InstallationId | Out-File $ConfigFile -Encoding UTF8
Write-Output "Default=Profiles/$($UserName).default" | Out-File $ConfigFile -Encoding UTF8 -Append
Write-Output "Locked=1" | Out-File $ConfigFile -Encoding UTF8 -Append
Write-Output "" | Out-File $ConfigFile -Encoding UTF8 -Append
Write-Output "[Profile0]" | Out-File $ConfigFile -Encoding UTF8 -Append
Write-Output "Default=1" | Out-File $ConfigFile -Encoding UTF8 -Append
Write-Output "IsRelative=1" | Out-File $ConfigFile -Encoding UTF8 -Append
Write-Output "Name=$($UserName)" | Out-File $ConfigFile -Encoding UTF8 -Append
Write-Output "Path=Profiles/$($UserName).default" | Out-File $ConfigFile -Encoding UTF8 -Append
Write-Output "" | Out-File $ConfigFile -Encoding UTF8 -Append
Write-Output "[General]" | Out-File $ConfigFile -Encoding UTF8 -Append
Write-Output "StartWithLastProfile=1" | Out-File $ConfigFile -Encoding UTF8 -Append
Write-Output "Version=2" | Out-File $ConfigFile -Encoding UTF8 -Append