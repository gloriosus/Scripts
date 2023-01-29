$SoftwarePath = ""
$RootFolder = (Split-Path -Path $PSScriptRoot -Parent)

Import-Module -Name "$($RootFolder)\Modules\Jarnet.PowerShell.Utility"

$Processes = @(
	"organizer"
)

Stop-Process -Name $Processes -Force -ErrorAction SilentlyContinue | Wait-Process -Timeout 30

$FullUserName = (Get-CimInstance -Class Win32_ComputerSystem).Username
$UserName = $FullUserName.Split('\')[1]
$ConfigFile = "C:\Users\$($UserName)\AppData\Roaming\Organizer\Profiles\$($UserName).default\prefs.js"

if (-Not (Test-Path "C:\Program Files\R7-Office\organizer\organizer.exe")) {
	Exit
}

if (-Not (Test-Path "C:\Users\$($UserName)\AppData\Roaming\Organizer\Profiles\$($UserName).default\Mail\Local Folders")) {
	Exit
}

if (Test-Path "C:\Users\$($UserName)\AppData\Roaming\Organizer\Profiles\$($UserName).default\Mail\Local Folders-1") {
	$Content = Get-Content $ConfigFile -Force | Where-Object { $_ -Ne "user_pref(`"mail.server.server2.directory`", `"C:\\Users\\$($UserName)\\AppData\\Roaming\\Organizer\\Profiles\\$($UserName).default\\Mail\\Local Folders-1`");" }
	$Content = $Content | Where-Object { $_ -Ne "user_pref(`"mail.server.server2.directory-rel`", `"[ProfD]Mail/Local Folders-1`");" }
	$Content = $Content | Where-Object { $_ -Ne "user_pref(`"mail.server.server2.hostname`", `"Local Folders-1`");" }
	Set-Content -Path $ConfigFile -Value $Content -Force

	Write-Output "user_pref(`"mail.server.server2.directory`", `"C:\\Users\\$($UserName)\\AppData\\Roaming\\Organizer\\Profiles\\$($UserName).default\\Mail\\Local Folders`");" | Out-File $ConfigFile -Encoding UTF8 -Append
	Write-Output "user_pref(`"mail.server.server2.directory-rel`", `"[ProfD]Mail/Local Folders`");" | Out-File $ConfigFile -Encoding UTF8 -Append
	Write-Output "user_pref(`"mail.server.server2.hostname`", `"Local Folders`");" | Out-File $ConfigFile -Encoding UTF8 -Append

	Exit
}

if (Test-Path "C:\Users\$($UserName)\AppData\Roaming\Organizer\Profiles\$($UserName).default\Mail\Local Folders\Корневой уровень файла данных Outlook*") {
	Exit
}

$SystemType = (Get-CimInstance -ClassName Win32_ComputerSystem).SystemType
$SystemType -Match "x(?<arch>64|86)-based PC"
$Arch = $Matches.arch

if (Test-Path "C:\Users\$($UserName)\Documents\Файлы Outlook\*.pst") {
	foreach ($File in Get-ChildItem -Path "C:\Users\$($UserName)\Documents\Файлы Outlook\*.pst")
	{
		Start-Process -FilePath "$($SoftwarePath)\Утилиты\pstconv\jre1.8.0_321_x$($Arch)\bin\java.exe" -ArgumentList "-jar $($SoftwarePath)\Утилиты\pstconv\pstconv-0.9.5.jar -i `"C:\Users\$($UserName)\Documents\Файлы Outlook\$($File.Name)`" -o `"C:\Users\$($UserName)\AppData\Roaming\Organizer\Profiles\$($UserName).default\Mail\Local Folders`"" -Wait -ErrorAction SilentlyContinue
	}
}

if (Test-Path "C:\Users\$($UserName)\AppData\Local\Microsoft\Outlook\*.pst") {
	foreach ($File in Get-ChildItem -Path "C:\Users\$($UserName)\AppData\Local\Microsoft\Outlook\*.pst")
	{
		Start-Process -FilePath "$($SoftwarePath)\Утилиты\pstconv\jre1.8.0_321_x$($Arch)\bin\java.exe" -ArgumentList "-jar $($SoftwarePath)\Утилиты\pstconv\pstconv-0.9.5.jar -i `"C:\Users\$($UserName)\AppData\Local\Microsoft\Outlook\$($File.Name)`" -o `"C:\Users\$($UserName)\AppData\Roaming\Organizer\Profiles\$($UserName).default\Mail\Local Folders`"" -Wait -ErrorAction SilentlyContinue
	}
}

Start-ProcessAsUser -FilePath "$($RootFolder)\Launch\show-popup.vbs" -Argument "Операция конвертации архивов почты завершена. Проверьте наличие архива в Р7-Органайзер" -UserName $FullUserName -ErrorAction SilentlyContinue