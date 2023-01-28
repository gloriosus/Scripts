# Change to absolute path
. ".\src\Utility.ps1"

$UsernameDomain = $(Get-WMIObject -Class Win32_ComputerSystem | Select-Object Username).Username
$Username = $UsernameDomain.Split('\')[1]
$File = "C:\Users\$($Username)\AppData\Roaming\Organizer\Profiles\$($Username).default\prefs.js"

$Processes = @(
	"organizer"
)

Close-Processes $Processes | Wait-Process -Timeout 10 -ErrorAction Ignore

$Content = Get-Content $File -Force | Where-Object {$_ -ne "user_pref(`"mail.smtpserver.smtp1.authMethod`", 1);"}
$Content = $Content | Where-Object {$_ -ne "user_pref(`"mail.smtpserver.smtp1.authMethod`", 3);"}
$Content = $Content | Where-Object {$_ -ne "user_pref(`"mail.smtpserver.smtp1.username`", `"$($Username)`");"}
Set-Content -Path $File -Value $Content -Force

Write-Output 'user_pref("mail.smtpserver.smtp1.authMethod", 3);' | Out-File $File -Encoding UTF8 -Append
Write-Output "user_pref(`"mail.smtpserver.smtp1.username`", `"$($Username)`");" | Out-File $File -Encoding UTF8 -Append