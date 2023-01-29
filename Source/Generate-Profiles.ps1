$DomainWithUsername = (Get-CimInstance -Class Win32_ComputerSystem).Username
$Username = $DomainWithUsername.Split('\')[1]
$File = "C:\Users\$($Username)\AppData\Roaming\Organizer\profiles.ini"

# If newer version is out, change the installation ID. It can be obtained by installing R7 on a test machine
Write-Output "[InstallEEF607DE8263F914]" | Out-File $File -Encoding UTF8
Write-Output "Default=Profiles/$($Username).default" | Out-File $File -Encoding UTF8 -Append
Write-Output "Locked=1" | Out-File $File -Encoding UTF8 -Append
Write-Output "" | Out-File $File -Encoding UTF8 -Append
Write-Output "[Profile0]" | Out-File $File -Encoding UTF8 -Append
Write-Output "Default=1" | Out-File $File -Encoding UTF8 -Append
Write-Output "IsRelative=1" | Out-File $File -Encoding UTF8 -Append
Write-Output "Name=$($Username)" | Out-File $File -Encoding UTF8 -Append
Write-Output "Path=Profiles/$($Username).default" | Out-File $File -Encoding UTF8 -Append
Write-Output "" | Out-File $File -Encoding UTF8 -Append
Write-Output "[General]" | Out-File $File -Encoding UTF8 -Append
Write-Output "StartWithLastProfile=1" | Out-File $File -Encoding UTF8 -Append
Write-Output "Version=2" | Out-File $File -Encoding UTF8 -Append