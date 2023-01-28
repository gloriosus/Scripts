$VersionNumber = "68.5.0.5"

# Change to absolute paths
Copy-Item -Path ".\src\Uninstall-R7Organizer.ps1" -Destination "C:\Windows\Installer"
Copy-Item -Path ".\src\Utility.ps1" -Destination "C:\Windows\Installer"

New-Item -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\R7Organizer" -Force
Set-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\R7Organizer" -Name "DisplayName" -Value "Р7-Офис. Органайзер" -Force
Set-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\R7Organizer" -Name "DisplayVersion" -Value $VersionNumber -Force
Set-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\R7Organizer" -Name "Publisher" -Value "r7office.ru" -Force
Set-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\R7Organizer" -Name "UninstallString" -Value "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noprofile -executionpolicy bypass -file C:\Windows\Installer\Uninstall-R7Organizer.ps1" -Force