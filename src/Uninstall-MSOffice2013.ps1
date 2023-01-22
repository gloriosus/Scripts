# Based on the instruction: https://support.microsoft.com/en-us/office/manually-uninstall-office-4e2904ea-25c8-4544-99ee-17696bb3027b?ui=en-us&rs=en-us&ad=us#OfficeVersion=Office_2013

# Change to absolute path
. ".\src\Utility.ps1"

$Processes = @(
	"EXCEL",
	"ONENOTE",
	"ONENOTEM",
	"OUTLOOK",
	"POWERPNT",
	"WINWORD",
	"officeclicktorun",
	"appvshnotify",
	"firstrun",
	"setup*"
)

$Services = @(
	"clicktorunsvc"
)

$ScheduledTasks = @(
	"Office 15 Subscription Heartbeat",
	"Office Automatic Update*",
	"Office Subscription Maintenance"
)

$Paths = @(
	"C:\Program Files\Microsoft Office 15",
	"C:\Program Files (x86)\Microsoft Office 15",
	"C:\ProgramData\Microsoft\ClickToRun",
	"C:\ProgramData\Microsoft\Office\FFPackageLocker",
	"Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\15.0\ClickToRun",
	"Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\15.0\ClickToRun",
	"Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\AppVISV",
	"Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\AppVISV",
	"Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft Office*15*",
	"Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft Office*15*",
	"Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*0FF1CE*",
	"Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*0FF1CE*",
	"Registry::HKEY_CURRENT_USER\Software\Microsoft\Office",
	"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2013"
)

$PathsExpressions = @(
	@{
		Path = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"; 
		Expression = '{$_.InstallLocation -Eq "C:\Program Files\Microsoft Office 15"}'
	},
	@{
		Path = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"; 
		Expression = '{$_.InstallLocation -Eq "C:\Program Files (x86)\Microsoft Office 15"}'
	}
)

Remove-ScheduledTasks $ScheduledTasks
Start-Sleep -Seconds 5.0
Close-Processes $Processes | Wait-Process -Timeout 30 -ErrorAction Ignore
Start-Sleep -Seconds 5.0
Remove-Services $Services
Start-Sleep -Seconds 5.0
Remove-Paths $Paths
Remove-PathsWithExpression $PathsExpressions

msiexec.exe /Uninstall "{50150000-008F-0000-1000-0000000FF1CE}" /NoRestart /Quiet
msiexec.exe /Uninstall "{50150000-007E-0000-0000-0000000FF1CE}" /NoRestart /Quiet
msiexec.exe /Uninstall "{50150000-008C-0000-1000-0000000FF1CE}" /NoRestart /Quiet
msiexec.exe /Uninstall "{90150000-008F-0000-1000-0000000FF1CE}" /NoRestart /Quiet
msiexec.exe /Uninstall "{90150000-008C-0000-0000-0000000FF1CE}" /NoRestart /Quiet
msiexec.exe /Uninstall "{90150000-008C-0419-0000-0000000FF1CE}" /NoRestart /Quiet