# Added '-ErrorAction Ignore' just to be sure that Kaspersky will not mark the task as failed

function Close-Processes($processes){
    $result = @()
	foreach ($process in $processes){
		$foundProcess = Get-Process -ProcessName $process -ErrorAction Ignore
        $result += $foundProcess
		if ($foundProcess){
			Stop-Process $foundProcess -Force
		}
	}
    return $result
}

function Remove-Services($services){
    foreach ($service in $services){
        $foundService = Get-Service -Name $service -ErrorAction Ignore
        if ($foundService){
            Stop-Service $foundService -Force
            sc.exe delete $service
        }
    }
}

function Uninstall-Products($namePattern){
    $foundProducts = wmic product get name | Select-String -Pattern $namePattern
    foreach ($product in $foundProducts){
        wmic product where "name='$($product.Line.Trim())'" call uninstall /NoInteractive
    }
}

function Remove-Paths($paths){
    foreach ($path in $paths){
        $isExists = Test-Path $path
        if ($isExists){
            Remove-Item -Path $path -Recurse -Force -ErrorAction Ignore
        }
    }
}

function Remove-PathsWithExpression($pathsExpressions){
    foreach ($pathExp in $pathsExpressions){
        $foundPath = Get-ChildItem -Path $pathExp.Path | Get-ItemProperty | Where-Object (Invoke-Expression $pathExp.Expression)
        foreach ($path in $foundPath){
            Remove-Item -Path "$($pathExp.Path)\$($path.PSChildName)" -Recurse -Force
        }
    }
}

function Remove-ScheduledTasks($scheduledTasks){
    foreach ($task in $scheduledTasks){
        $foundTask = Get-ScheduledTask -TaskName $task -ErrorAction Ignore
        if ($foundTask){
            Unregister-ScheduledTask -TaskName $task -Confirm:$false
        }
    }
}

Function Disable-CloseButton {
    #Calling user32.dll methods for Windows and Menus
    $MethodsCall = '
    [DllImport("user32.dll")] public static extern long GetSystemMenu(IntPtr hWnd, bool bRevert);
    [DllImport("user32.dll")] public static extern bool EnableMenuItem(long hMenuItem, long wIDEnableItem, long wEnable);
    [DllImport("user32.dll")] public static extern long SetWindowLongPtr(long hWnd, long nIndex, long dwNewLong);
    [DllImport("user32.dll")] public static extern bool EnableWindow(long hWnd, int bEnable);
    '

    $SC_CLOSE = 0xF060
    $MF_DISABLED = 0x00000002L


    #Create a new namespace for the Methods to be able to call them
    Add-Type -MemberDefinition $MethodsCall -name NativeMethods -namespace Win32

    $PSWindow = Get-Process -Pid $PID
    $hwnd = $PSWindow.MainWindowHandle

    #Get System menu of windows handled
    $hMenu = [Win32.NativeMethods]::GetSystemMenu($hwnd, 0)

    #Disable X Button
    [Win32.NativeMethods]::EnableMenuItem($hMenu, $SC_CLOSE, $MF_DISABLED) | Out-Null
}