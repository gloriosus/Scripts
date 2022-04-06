# TODO: make downloading different version possible by passing a parameter to a function
$PSVersion = "7.2.2"

# TODO: turn it into the function and move it to separate script module (.psm1)
<#
Go to the latest version "https://github.com/PowerShell/PowerShell/releases/latest",
which turns into "https://github.com/PowerShell/PowerShell/releases/tag/v7.2.2/",
change it to "https://github.com/PowerShell/PowerShell/releases/download/v7.2.2/PowerShell-7.2.2-win-x64.zip"
#>
Invoke-WebRequest "https://github.com/PowerShell/PowerShell/releases/download/v$($PSVersion)/PowerShell-$($PSVersion)-win-x64.zip" -OutFile "PowerShell-$($PSVersion)-win-x64.zip"
Invoke-WebRequest "https://github.com/PowerShell/PowerShell/releases/download/v$($PSVersion)/PowerShell-$($PSVersion)-win-x86.zip" -OutFile "PowerShell-$($PSVersion)-win-x86.zip"

Expand-Archive "PowerShell-$($PSVersion)-win-x64.zip" -DestinationPath "PowerShell-$($PSVersion)-win-x64"
Expand-Archive "PowerShell-$($PSVersion)-win-x86.zip" -DestinationPath "PowerShell-$($PSVersion)-win-x86"

Remove-Item -Path "PowerShell-$($PSVersion)-win-x64.zip"
Remove-Item -Path "PowerShell-$($PSVersion)-win-x86.zip"

# TODO: get only PowerShell files (.ps1)
$scripts = Get-ChildItem -Path "src"

# TODO: consider to change the method of getting OS architecture
# TODO: make a separate script to create bats
foreach ($script in $scripts)
{
	New-Item "$($script.BaseName).bat" -ItemType File -Value "if exist `"C:\Program Files (x86)\`" ( %COMSPEC% /C `"start PowerShell-$($PSVersion)-win-x64\pwsh.exe src\$($script.Name)`" ) else ( %COMSPEC% /C `"start PowerShell-$($PSVersion)-win-x86\pwsh.exe src\$($script.Name)`" )"
}