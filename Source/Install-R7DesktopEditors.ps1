$InstallationExe = "r7office_x64_7.2.2.36.exe"
$LicenseFile = "licensefilename.lickey"
$UsernameDomain = $(Get-WMIObject -Class Win32_ComputerSystem | Select-Object Username).Username

Start-Process -FilePath "pathtoinstallationexe\$($InstallationExe)" -ArgumentList "/SP-", "/VERYSILENT", "/SUPPRESSMSGBOXES", "/NOCANCEL", "/NORESTART", "/ALLUSERS" -Wait
New-Item -Path "C:\ProgramData\R7-Office\License" -ItemType Directory -Force
Copy-Item -Path "pathtolicensefile\$($LicenseFile)" -Destination "C:\ProgramData\R7-Office\License" -Force

$Acl = Get-Acl "C:\ProgramData\R7-Office\License\$($LicenseFile)"
$AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($UsernameDomain, "FullControl", "Allow")
$Acl.SetAccessRule($AccessRule)
$Acl | Set-Acl "C:\ProgramData\R7-Office\License\$($LicenseFile)"