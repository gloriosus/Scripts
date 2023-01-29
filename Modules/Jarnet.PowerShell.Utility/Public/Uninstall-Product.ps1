function Uninstall-Product {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $True, Position = 0)]
		[string]$Pattern
	)
	PROCESS {
		try {
			$FoundProducts = wmic.exe product get name | Select-String -Pattern $Pattern
            foreach ($Product in $FoundProducts) {
                wmic.exe product where "name='$($Product.Line.Trim())'" call uninstall /NoInteractive
            }
		} catch [System.Exception] {
			Write-Warning -Message $_.Exception.Message
		}
	}
}