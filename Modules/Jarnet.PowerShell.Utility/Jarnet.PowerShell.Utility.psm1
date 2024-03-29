$Private = @( Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -ErrorAction SilentlyContinue )
$Public  = @( Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -ErrorAction SilentlyContinue )

foreach($import in @($Public + $Private)) {
    try {
        . $import.FullName
    }
    catch {
        Write-Error -Message "Failed to import function $($import.FullName): $_"
    }
}

Export-ModuleMember -Function "Start-ProcessAsUser"
Export-ModuleMember -Function "Disable-CloseButton"
Export-ModuleMember -Function "Get-ChildItemWhere"
Export-ModuleMember -Function "Remove-ScService"
Export-ModuleMember -Function "Uninstall-Product"