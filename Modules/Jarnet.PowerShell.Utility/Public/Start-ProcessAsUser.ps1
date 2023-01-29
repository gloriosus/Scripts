<#
.SYNOPSIS
	Start a process within a context of the specified user
.DESCRIPTION
	Allows to start a process within a context of the specified user when a script that uses this function is running within different context
.PARAMETER FilePath
	Specifies a path to the desired executable file (.exe, .bat, .vbs, etc.)
.PARAMETER UserName
	Specifies fullname (with a domain or computer name) of a user in which context the process will start
.OUTPUTS
	No return
.EXAMPLE
	Start-ProcessAsUser -FilePath "C:\Windows\explorer.exe" -UserName "Domain\user-01"
#>
function Start-ProcessAsUser {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $True, Position = 0)]
		[string]$FilePath,
		[string]$Argument = $null,
		[ValidateNotNullOrEmpty()]
		[string]$UserName
	)
	PROCESS {
		try {
			$Trigger = New-ScheduledTaskTrigger -Once -At (Get-Date "1999-01-01 00:00")

			if ([string]::IsNullOrWhiteSpace($Argument)) {
				$Action = New-ScheduledTaskAction -Execute $FilePath
			} else {
				$Action = New-ScheduledTaskAction -Execute $FilePath -Argument $Argument
			}
			
			Register-ScheduledTask -TaskName "Start-ProcessAsUser" -Trigger $Trigger -User $UserName -Action $Action
			Start-ScheduledTask -TaskName "Start-ProcessAsUser"

			$Timeout = 60
			$Timer = [Diagnostics.Stopwatch]::StartNew()

			while (((Get-ScheduledTask -TaskName "Start-ProcessAsUser").State -Ne "Ready") -And ($Timer.Elapsed.TotalSeconds -Lt $Timeout)) {
				Write-Verbose -Message "Waiting for the task completion..."
				Start-Sleep -Seconds 5.0
			}

			$Timer.Stop()
			Unregister-ScheduledTask -TaskName "Start-ProcessAsUser" -Confirm:$false
		} catch [System.Exception] {
			Write-Warning -Message $_.Exception.Message
		}
	}
}