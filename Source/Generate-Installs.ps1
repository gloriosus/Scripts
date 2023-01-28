$UsernameDomain = $(Get-WMIObject -Class Win32_ComputerSystem | Select-Object Username).Username
$Username = $UsernameDomain.Split('\')[1]
$File = "C:\Users\$($Username)\AppData\Roaming\Organizer\installs.ini"

# If newer version is out, change the installation ID. It can be obtained by installing R7 on a test machine
Write-Output "[EEF607DE8263F914]" | Out-File $File -Encoding UTF8
Write-Output "Default=Profiles/$($Username).default" | Out-File $File -Encoding UTF8 -Append
Write-Output "Locked=1" | Out-File $File -Encoding UTF8 -Append