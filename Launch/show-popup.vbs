Dim WshShell, BtnCode
Set WshShell = WScript.CreateObject("WScript.Shell")
BtnCode = WshShell.Popup(WScript.Arguments(0), 0, "Сообщение", 0 + 4096)