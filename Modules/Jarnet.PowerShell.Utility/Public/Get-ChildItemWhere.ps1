<#
.SYNOPSIS
	Finds full path of subfolders, subkeys and files with the specified 'Where' expression
.DESCRIPTION
	Finds full path of subfolders, subkeys and files with the specified 'Where' expression
.PARAMETER Path
	Specifies a path to the desired folder
.PARAMETER Where
	(Optional) Specifies a 'Where' expression to find child items
.OUTPUTS
	Full path to the found items
.EXAMPLE
	Get-ChildItemWhere -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" -Where "{ `$_.Publisher -eq `"Microsoft Corporation`" }"
#>
function Get-ChildItemWhere {
    [CmdletBinding()]
	param (
		[Parameter(Mandatory = $True, Position = 0)]
		[string[]]$Path,
		[string]$Where
	)
	PROCESS {
		if ([string]::IsNullOrWhiteSpace($Where)) {
			$Where = "{ `$_ }"
		}

		try {
			foreach ($Item in $Path) {
                $FoundPath = Get-ChildItem -Path $Item | Get-ItemProperty | Where-Object (Invoke-Expression $Where)
                foreach ($ItemProperties in $FoundPath) {
                    "$($Item)\$($ItemProperties.PSChildName)"
                }
            }
		} catch [System.Exception] {
			Write-Warning -Message $_.Exception.Message
		}
	}
}