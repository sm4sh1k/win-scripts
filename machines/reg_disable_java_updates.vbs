'--------------------------------------------------------------------------------------
' Disabling Java Runtime Environment (JRE) automatic updates
' Author: Valentin Vakhrushev, 2011-2013
'--------------------------------------------------------------------------------------

On Error Resume Next

Const HKEY_LOCAL_MACHINE = &H80000002

Set WshShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

' Determining installation path of newest JRE version
strKeyPath = "SOFTWARE\JavaSoft\Java Runtime Environment"
objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys
strVersion = arrSubKeys(UBound(arrSubKeys))

' Marker file is needed to sure that we did it one time for a version
' When new version of JRE is installed the new marker file is created (appropriate for this version)
strMarkerFile = WshShell.ExpandEnvironmentStrings("%APPDATA%") & "\jre_" & strVersion
If objFSO.FileExists(strMarkerFile) Then WScript.Quit()

strFullKeyPath = "HKLM\" & strKeyPath & "\" & strVersion & "\MSI\"
WshShell.RegWrite strFullKeyPath & "JAVAUPDATE", "0", "REG_SZ"
WshShell.RegWrite strFullKeyPath & "AUTOUPDATECHECK", "0", "REG_SZ"

WshShell.RegWrite "HKLM\SOFTWARE\JavaSoft\Java Update\Policy\EnableJavaUpdate", "0", "REG_DWORD"
WshShell.RegWrite "HKLM\SOFTWARE\JavaSoft\Java Update\Policy\NotifyDownload", "0", "REG_DWORD"
WshShell.RegWrite "HKLM\SOFTWARE\JavaSoft\Java Update\Policy\NotifyInstall", "0", "REG_DWORD"

objFSO.CreateTextFile strMarkerFile, True
