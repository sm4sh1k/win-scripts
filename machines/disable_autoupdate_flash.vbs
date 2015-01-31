'--------------------------------------------------------------------------------------
' Disabling autoupdater for Adobe Flash Player
' Author: Valentin Vakhrushev, 2012
'--------------------------------------------------------------------------------------

On Error Resume Next

Set WSHShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

strTaskName = "Adobe Flash Player Updater"
strFileName = "mms.cfg"

ReturnCode = WSHShell.Run("schtasks /delete /tn " & Chr(34) & strTaskName & Chr(34) & " /f", 0)

If WSHShell.ExpandEnvironmentStrings("%PROGRAMFILES(X86)%") = "%PROGRAMFILES(X86)%" Then
	strFolderPath = WSHShell.ExpandEnvironmentStrings("%SYSTEMROOT%\System32\Macromed\Flash\")
Else
	strFolderPath = WSHShell.ExpandEnvironmentStrings("%SYSTEMROOT%\SysWOW64\Macromed\Flash\")
End If
If Not objFSO.FolderExists(strFolderPath) Then WScript.Quit()
If objFSO.FileExists(strFolderPath & strFileName) Then WScript.Quit()

Set TextStream = objFSO.CreateTextFile(strFolderPath & strFileName, True) 'ANSI
TextStream.WriteLine "AutoUpdateDisable=1"
TextStream.Close
