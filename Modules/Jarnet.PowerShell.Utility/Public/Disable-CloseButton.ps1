<#
.SYNOPSIS
	Makes the close button of a window disabled
.DESCRIPTION
	Makes the close button of a window disabled to prevent users from closing it
.OUTPUTS
	No return
.EXAMPLE
	Disable-CloseButton
#>
function Disable-CloseButton {
    [CmdletBinding()]
	param ()
	PROCESS {
		try {
            $MethodsCall = @(
                "[DllImport(`"user32.dll`")] public static extern long GetSystemMenu(IntPtr hWnd, bool bRevert);",
                "[DllImport(`"user32.dll`")] public static extern bool EnableMenuItem(long hMenuItem, long wIDEnableItem, long wEnable);",
                "[DllImport(`"user32.dll`")] public static extern long SetWindowLongPtr(long hWnd, long nIndex, long dwNewLong);",
                "[DllImport(`"user32.dll`")] public static extern bool EnableWindow(long hWnd, int bEnable);"
            )
        
            $SC_CLOSE = 0xF060
            $MF_DISABLED = 0x00000002L
        
            Add-Type -MemberDefinition $MethodsCall -Name NativeMethods -Namespace Win32
        
            $PSWindow = Get-Process -Pid $PID
            $hwnd = $PSWindow.MainWindowHandle
        
            $hMenu = [Win32.NativeMethods]::GetSystemMenu($hwnd, 0)
        
            [Win32.NativeMethods]::EnableMenuItem($hMenu, $SC_CLOSE, $MF_DISABLED) | Out-Null
		} catch [System.Exception] {
			Write-Warning -Message $_.Exception.Message
		}
	}
}