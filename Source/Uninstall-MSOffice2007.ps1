# Based on the instruction: https://support.microsoft.com/en-us/office/manually-uninstall-office-4e2904ea-25c8-4544-99ee-17696bb3027b?ui=en-us&rs=en-us&ad=us#OfficeVersion=Office_2007

# Change to absolute path
. ".\src\Utility.ps1"

$Processes = @(
	"EXCEL",
	"INFOPATH",
	"MSACCESS",
	"MSPUB",
	"POWERPNT",
	"WINWORD",
	"OUTLOOK"
)

$Services = @(
	"ose",
	"ose64"
)

$UsernameDomain = $(Get-WMIObject -Class Win32_ComputerSystem | Select-Object Username).Username
$Username = $UsernameDomain.Split('\')[1]

$Paths = @(
	"C:\Program Files (x86)\Common Files\Microsoft Shared\Office12",
	"C:\Program Files (x86)\Common Files\Microsoft Shared\Source Engine",
	"C:\Program Files (x86)\Microsoft Office\Office12",
	"C:\Program Files\Common Files\Microsoft Shared\Office12",
	"C:\Program Files\Common Files\Microsoft Shared\Source Engine",
	"C:\Program Files\Microsoft Office\Office12",
	"C:\MSOCache\All Users\*0FF1CE}-*",
	"C:\Users\$($Username)\AppData\Roaming\Microsoft\Templates\Normal.dotm",
	"C:\Users\$($Username)\AppData\Roaming\Microsoft\Document Building Blocks\*.dotx",
	"C:\ProgramData\Application Data\Microsoft\Office\Data\opa12.dat",
	"Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\12.0",
	"Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\12.0",
	"Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\Delivery\SourceEngine\Downloads\*0FF1CE}-*",
	"Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*0FF1CE*",
	"Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Upgrade Codes\*F01FEC",
	"Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\*F01FEC",
	"Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\12.0",
	"Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\Delivery\SourceEngine\Downloads\*0FF1CE}-*",
	"Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*0FF1CE*",
	"Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Installer\Upgrade Codes\*F01FEC",
	"Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\*F01FEC",
	"Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\ose",
	"Registry::HKEY_CLASSES_ROOT\Installer\Features\*F01FEC",
	"Registry::HKEY_CLASSES_ROOT\Installer\Products\*F01FEC",
	"Registry::HKEY_CLASSES_ROOT\Installer\UpgradeCodes\*F01FEC",
	"Registry::HKEY_CLASSES_ROOT\Installer\Win32Assemblies\*Office12*",
	"Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*Office12*",
	"Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*Office12*",
	"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office"
)

$PathsExpressions = @(
	@{
		Path = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"; 
		Expression = '{$_.UninstallString -Like "*\Office Setup Controller\Setup.exe*"}'
	},
	@{
		Path = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"; 
		Expression = '{$_.UninstallString -Like "*\Office Setup Controller\Setup.exe*"}'
	}
)

Close-Processes $Processes | Wait-Process -Timeout 30 -ErrorAction Ignore
Uninstall-Products "(Microsoft Office)(.*)(2007)(.*)"
Remove-Services $Services
Remove-Paths $Paths
Remove-PathsWithExpression $PathsExpressions