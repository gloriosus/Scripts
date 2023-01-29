$SoftwarePath = ""
$InstallationPackage = "r7office_x64_7.2.2.36.exe"
$LicenseFile = ""
$FullUserName = (Get-CimInstance -Class Win32_ComputerSystem).UserName.Split('\')[1]

Start-Process -FilePath "$($SoftwarePath)\Офисные приложения\Р7-Офис\$($InstallationPackage)" -ArgumentList "/SP-", "/VERYSILENT", "/SUPPRESSMSGBOXES", "/NOCANCEL", "/NORESTART", "/ALLUSERS" -Wait
New-Item -Path "C:\ProgramData\R7-Office\License" -ItemType Directory -Force
Copy-Item -Path "$($SoftwarePath)\Офисные приложения\Р7-Офис\$($LicenseFile)" -Destination "C:\ProgramData\R7-Office\License" -Force

$Acl = Get-Acl "C:\ProgramData\R7-Office\License\$($LicenseFile)"
$AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($FullUserName, "FullControl", "Allow")
$Acl.SetAccessRule($AccessRule)
$Acl | Set-Acl "C:\ProgramData\R7-Office\License\$($LicenseFile)"