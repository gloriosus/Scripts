$UsernameDomain = $(Get-WMIObject -Class Win32_ComputerSystem | Select-Object Username).Username
$Username = $UsernameDomain.Split('\')[1]

$Archives = Get-ChildItem "C:\Users\$($Username)\Documents\Файлы Outlook\*.pst"
foreach ($File in $Archives)
{
	Start-Process -FilePath ".\utility\converter\jre1.8.0_321_x64\bin\java.exe" -ArgumentList "-jar .\utility\converter\pstconv-0.9.4.jar -i `"C:\Users\$($Username)\Documents\Файлы Outlook\$($File.Name)`" -o `"C:\Users\$($Username)\AppData\Roaming\Organizer\Profiles\$($Username).default\Mail\Local Folders`"" -Wait
}

$Archives = Get-ChildItem "C:\Users\$($Username)\AppData\Local\Microsoft\Outlook\*.pst"
foreach ($File in $Archives)
{
	Start-Process -FilePath ".\utility\converter\jre1.8.0_321_x64\bin\java.exe" -ArgumentList "-jar .\utility\converter\pstconv-0.9.4.jar -i `"C:\Users\$($Username)\AppData\Local\Microsoft\Outlook\$($File.Name)`" -o `"C:\Users\$($Username)\AppData\Roaming\Organizer\Profiles\$($Username).default\Mail\Local Folders`"" -Wait
}