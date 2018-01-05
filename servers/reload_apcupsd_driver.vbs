'--------------------------------------------------------------------------------------
' Automatic restart the driver and Apcupsd service if the connection between 
' APC UPS and the computer has been lost. Driver reloading is implemented using 
' DevCon program provided by Microsoft. It can be downloaded here:
' https://docs.microsoft.com/ru-ru/windows-hardware/drivers/devtest/devcon
' Current connection status is taken with apcaccess program supplied with Apcupsd.
' Author: Valentin Vakhrushev, 2017
'--------------------------------------------------------------------------------------

On Error Resume Next

Set WshShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objRegExp = CreateObject("VBScript.RegExp")
objRegExp.Multiline = True
objRegExp.Pattern = "^STATUS\s+?:\s+?.*$"

strProgramPath = objFSO.GetParentFolderName(WScript.ScriptFullName) & "\devconx64.exe"
If Not objFSO.FileExists(strProgramPath) Then WScript.Quit()

Set objProgExec = WshShell.Exec("C:\Apcupsd\bin\apcaccess.exe")
Do Until objProgExec.Status
	WScript.Sleep 200
Loop

Set objMatches = objRegExp.Execute(OEMtoANSI(objProgExec.StdOut.ReadAll))
Set objMatch = objMatches.Item(objMatches.Count - 1)

If InStr(objMatch.Value, "COMMLOST") <> 0 Then
	WshShell.LogEvent 4, "Connection with UPS has been lost. Restarting driver and service..."
	WshShell.Run "net stop apcupsd", 0, True
	WshShell.Run strProgramPath & " restart USB\VID_051D*", 0, True
	WshShell.Run "net start apcupsd", 0
End If


Function OEMtoANSI(strString)
	strTemp = vbNullString
	For i = 1 To Len(strString)
		nCode = Asc(Mid(strString, i, 1))
		If nCode >= 128 And nCode <= 175 Then
			strTemp = strTemp + Chr(nCode + 64)
		ElseIf nCode >= 224 And nCode <= 239 Then
			strTemp = strTemp + Chr(nCode + 16)
		Else
			strTemp = strTemp + Mid(strString, i, 1)
		End If
	Next
	OEMtoANSI = strTemp
End Function
