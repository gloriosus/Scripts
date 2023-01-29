function Disable-CloseButton {
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

Disable-CloseButton
Write-Host "Идет удаление Р7-Органайзер..."

$UserName = (Get-CimInstance -Class Win32_ComputerSystem).UserName.Split('\')[1]

$Processes = @(
    "organizer"
)

$Paths = @(
    "C:\Program Files\R7-Office\organizer",
    "C:\Users\$($UserName)\Desktop\Р7-Органайзер.lnk",
    "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\R7Organizer",
    "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Р7-Офис\Р7-Органайзер.lnk"
)

Stop-Process -Name $Processes -Force -ErrorAction SilentlyContinue | Wait-Process -Timeout 30
Remove-Item -Path $Paths -Recurse -Force -ErrorAction SilentlyContinue