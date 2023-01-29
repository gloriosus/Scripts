. ".\Utility.ps1"

$Processes = @(
	"organizer"
)

Close-Processes $Processes -ErrorAction SilentlyContinue
Start-Sleep -Seconds 5.0

$DomainWithUsername = (Get-CimInstance -Class Win32_ComputerSystem).Username
$Username = $DomainWithUsername.Split('\')[1]
$File = "C:\Users\$($Username)\AppData\Roaming\Organizer\Profiles\$($Username).default\prefs.js"

if (-Not (Test-Path "C:\Program Files\R7-Office\organizer\organizer.exe")) {
	Exit
}

if (-Not (Test-Path "C:\Users\$($Username)\AppData\Roaming\Organizer\Profiles\$($Username).default\Mail\Local Folders")) {
	Exit
}

if (Test-Path "C:\Users\$($Username)\AppData\Roaming\Organizer\Profiles\$($Username).default\Mail\Local Folders-1") {
	$Content = Get-Content $File -Force | Where-Object { $_ -Ne "user_pref(`"mail.server.server2.directory`", `"C:\\Users\\$($Username)\\AppData\\Roaming\\Organizer\\Profiles\\$($Username).default\\Mail\\Local Folders-1`");" }
	$Content = $Content | Where-Object { $_ -Ne "user_pref(`"mail.server.server2.directory-rel`", `"[ProfD]Mail/Local Folders-1`");" }
	$Content = $Content | Where-Object { $_ -Ne "user_pref(`"mail.server.server2.hostname`", `"Local Folders-1`");" }
	Set-Content -Path $File -Value $Content -Force

	Write-Output "user_pref(`"mail.server.server2.directory`", `"C:\\Users\\$($Username)\\AppData\\Roaming\\Organizer\\Profiles\\$($Username).default\\Mail\\Local Folders`");" | Out-File $File -Encoding UTF8 -Append
	Write-Output "user_pref(`"mail.server.server2.directory-rel`", `"[ProfD]Mail/Local Folders`");" | Out-File $File -Encoding UTF8 -Append
	Write-Output "user_pref(`"mail.server.server2.hostname`", `"Local Folders`");" | Out-File $File -Encoding UTF8 -Append

	Exit
}

if (Test-Path "C:\Users\$($Username)\AppData\Roaming\Organizer\Profiles\$($Username).default\Mail\Local Folders\Корневой уровень файла данных Outlook*") {
	Exit
}

$Trigger = New-ScheduledTaskTrigger -Once -At (Get-Date "1999-01-01 00:00")
$Action = New-ScheduledTaskAction -Execute ".\show-popup.vbs"
Register-ScheduledTask -TaskName "Show-Popup" -Trigger $Trigger -User $DomainWithUsername -Action $Action -ErrorAction SilentlyContinue

$SystemType = (Get-CimInstance -ClassName Win32_ComputerSystem).SystemType
$SystemType -Match "x(?<arch>64|86)-based PC"
$Arch = $Matches.arch

if (Test-Path "C:\Users\$($Username)\Documents\Файлы Outlook\*.pst") {
	foreach ($file in Get-ChildItem -Path "C:\Users\$($Username)\Documents\Файлы Outlook\*.pst")
	{
		Start-Process -FilePath ".\converter\jre1.8.0_321_x$($Arch)\bin\java.exe" -ArgumentList "-jar .\converter\pstconv-0.9.5.jar -i `"C:\Users\$($Username)\Documents\Файлы Outlook\$($file.Name)`" -o `"C:\Users\$($Username)\AppData\Roaming\Organizer\Profiles\$($Username).default\Mail\Local Folders`"" -Wait -ErrorAction SilentlyContinue
	}
}

if (Test-Path "C:\Users\$($Username)\AppData\Local\Microsoft\Outlook\*.pst") {
	foreach ($file in Get-ChildItem -Path "C:\Users\$($Username)\AppData\Local\Microsoft\Outlook\*.pst")
	{
		Start-Process -FilePath ".\converter\jre1.8.0_321_x$($Arch)\bin\java.exe" -ArgumentList "-jar .\converter\pstconv-0.9.5.jar -i `"C:\Users\$($Username)\AppData\Local\Microsoft\Outlook\$($file.Name)`" -o `"C:\Users\$($Username)\AppData\Roaming\Organizer\Profiles\$($Username).default\Mail\Local Folders`"" -Wait -ErrorAction SilentlyContinue
	}
}

Start-ScheduledTask -TaskName "Show-Popup" -ErrorAction SilentlyContinue

$Timeout = 60
$Timer = [Diagnostics.Stopwatch]::StartNew()

while (((Get-ScheduledTask -TaskName "Show-Popup").State -Ne "Ready") -And ($Timer.Elapsed.TotalSeconds -Lt $Timeout)) {
	Write-Verbose -Message "Ожидание завершения выполнения задания..."
	Start-Sleep -Seconds 5.0
}

$Timer.Stop()
Unregister-ScheduledTask -TaskName "Show-Popup" -Confirm:$false -ErrorAction SilentlyContinue