'--------------------------------------------------------------------------------------
' Disabling 2GIS Update Notifier and deleting 2GIS Update Service
' Script is assigned for disabling annoying notifications of 2GIS AutoUpdater
' Author: Valentin Vakhrushev, 2012
'--------------------------------------------------------------------------------------

On Error Resume Next

Set WshShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

strServiceName = "2GISUpdateService"
strUpdProgName = "2GISUpdateService.exe"
strTrayProgName = "2GISTrayNotifier.exe"
strRootRegKey = "HKLM\Software\"
strRunRegKey = "Microsoft\Windows\CurrentVersion\Run\2Gis Update Notifier"
strProgRegKey = "DoubleGIS\Grym\path"

If WshShell.ExpandEnvironmentStrings("%PROGRAMFILES(X86)%") <> "%PROGRAMFILES(X86)%" Then
	strRootRegKey = strRootRegKey & "Wow6432Node\"
End If
strRunRegKey = strRootRegKey & strRunRegKey
strProgRegKey = strRootRegKey & strProgRegKey

strRegKeyData = WshShell.RegRead(strRunRegKey)
strRegKeyData = WshShell.RegRead(strProgRegKey)
If Err.Number <> 0 Then WScript.Quit()

WshShell.RegDelete strRunRegKey
ReturnCode = WshShell.Run("sc delete " & strServiceName, 0)
objFSO.DeleteFile strRegKeyData & strUpdProgName, True
objFSO.DeleteFile strRegKeyData & strTrayProgName, True
