$Path = "C:\Program Files\Google\Chrome\Application"
$AltPath = "C:\Program Files (x86)\Google\Chrome\Application"

Stop-Process -Name "setup", "chrome", "GoogleUpdate" -Force -ErrorAction SilentlyContinue | Wait-Process -Timeout 30

if (Test-Path -Path $Path) {
    Get-ChildItem -Path $Path -Attributes Directory | Where-Object { $_.Name -Match "(?>\d+\.*){4}" }

    if ($Matches.Count -gt 0) {
        Start-Process -FilePath "$($Path)\$($Matches[0])\Installer\setup.exe" -ArgumentList "--uninstall --channel=stable --system-level --force-uninstall" -ErrorAction SilentlyContinue | Wait-Process -Timeout 180
    }

    Clear-Variable -Name "Matches"
}

if (Test-Path -Path $AltPath) {
    Get-ChildItem -Path $AltPath -Attributes Directory | Where-Object { $_.Name -Match "(?>\d+\.*){4}" }

    if ($Matches.Count -gt 0) {
        Start-Process -FilePath "$($AltPath)\$($Matches[0])\Installer\setup.exe" -ArgumentList "--uninstall --channel=stable --system-level --force-uninstall" -ErrorAction SilentlyContinue | Wait-Process -Timeout 180
    }
}

$Username = (Get-CimInstance -Class Win32_ComputerSystem).Username.Split('\')[1]
$Sid = (Get-ChildItem -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" | Get-ItemProperty | Where-Object ProfileImagePath -like "*$($Username)").PSChildName
Get-ChildItem -Path "Registry::HKEY_USERS\$($Sid)\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" | Get-ItemProperty | Where-Object {(($_.DisplayName -eq "Презентация") -or ($_.DisplayName -eq "YouTube") -or ($_.DisplayName -eq "Gmail") -or ($_.DisplayName -eq "Документы") -or ($_.DisplayName -eq "Таблица") -or ($_.DisplayName -eq "Google Диск"))} | Select-Object PSPath | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

$AdditionalPaths = @(
    "C:\Program Files\Google",
    "C:\Program Files (x86)\Google",
    "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Google Chrome.lnk",
    "C:\Users\$($Username)\AppData\Local\Google",
    "C:\Users\$($Username)\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\Google Chrome.lnk",
    "C:\Users\$($Username)\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Google Chrome.lnk",
    "C:\Users\$($Username)\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Приложения Chrome",
    "C:\Users\Public\Desktop\Google Chrome.lnk"
)

Remove-Item -Path $AdditionalPaths -Recurse -Force -ErrorAction SilentlyContinue