function Remove-ScService {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $True, Position = 0)]
		[string[]]$Name,
        [ValidateNotNullOrEmpty()]
		[int]$Timeout = 0
	)
	PROCESS {
		try {
			foreach ($Service in $Name) {
                (Stop-Service -Name $Service -Force -PassThru -ErrorAction SilentlyContinue).WaitForStatus("Stopped", (New-TimeSpan -Seconds $Timeout))
                sc.exe delete $Service
            }
		} catch [System.Exception] {
			Write-Warning -Message $_.Exception.Message
		}
	}
}