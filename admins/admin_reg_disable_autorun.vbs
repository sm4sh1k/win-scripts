'--------------------------------------------------------------------------------------
' Disabling autorun for all drives
' Script tries to get administrative privileges if needed (supposed to be run interactively)
' Author: Valentin Vakhrushev, 2012
'--------------------------------------------------------------------------------------

On Error Resume Next

Set WshShell = WScript.CreateObject("WScript.Shell")

strKey = CreateObject("WScript.Shell").RegRead("HKEY_USERS\s-1-5-19\")
If Err.Number <> 0 Then
	WshShell.RegWrite _
	"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoDriveTypeAutorun", _
	"255", "REG_DWORD"
	
	Set objShell = CreateObject("Shell.Application")
	objShell.ShellExecute "wscript.exe", Chr(34) & _
		WScript.ScriptFullName & Chr(34), "", "runas", 1
	WScript.Quit()
End If

WshShell.RegWrite _
	"HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoDriveTypeAutorun", _
	"255", "REG_DWORD"
WshShell.RegWrite _
	"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoDriveTypeAutorun", _
	"255", "REG_DWORD"

WScript.Echo "Autorun for all drives is disabled."
