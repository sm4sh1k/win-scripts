'--------------------------------------------------------------------------------------
' Mounting BitLocker encrypted volume when USB flash disk is inserted
' 
' How it works:
' 1) Script is executed on computer start
'    (using Group Policy, Windows registry, scheduled job or something else)
' 2) Script waits when the flash disk with a specified key file is inserted
' 3) Then the encrypted volume is unlocked (using a key file)
' 4) In my circumstances it is needed to restart 'Server' service after that,
'    because of shared folder placed on encrypted volume
' 5) Then script waits 10 minutes when the user extracts flash disk with a key file
' 6) If the user did not extract flash disk during this time then encrypted volume is
'    locked again. This mechanism was added to avoid situation when the user
'    forgets flash disk in a computer
' 7) All messages are written in event log
'
' Script has been tested on Windows 2008 R2
' 
' Author: Valentin Vakhrushev, 2013
'--------------------------------------------------------------------------------------

On Error Resume Next

' Exit if another instance of the script is already running
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\CIMV2")
Set colProc = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE CommandLine LIKE '%" & _
	WScript.ScriptName & "%'")
If colProc.Count > 1 Then WScript.Quit()

Set WshShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
strComputerName = WshShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
bMonitoring = True

' Define a key file name here
strKeyFileSubPath = "1234ABCD-EF123-4567-89AB-CDEF98765432.BEK"
' Type a drive letter of encrypted volume here (in my case it is a disk D:)
strCommand = "manage-bde -unlock D: -RecoveryKey "
strComUnmount = "manage-bde -lock D: -ForceDismount"

' Detect when the flash disk was inserted looking for new drive letters
' Type here drive letters which can be used for appointment to flash disks
colDrives = Split("F: G: H: I: J:")
While bMonitoring = True
	For Each Drive In colDrives
		bMounted = True
		Err.Clear
		Set Drv = objFSO.GetDrive(Drive)
		If Err.Number Or Drv.IsReady = False Or Drv.DriveType <> 1 Then bMounted = False
		If bMounted Then
			If objFSO.FileExists(Drive & "\" & strKeyFileSubPath) Then
				' Mounting encrypted volume
				strCommand = strCommand & Chr(34) & Drive & "\" & strKeyFileSubPath & Chr(34)
				Set objScriptExec = WshShell.Exec(strCommand)
				Do Until objScriptExec.Status
					WScript.Sleep 1000
				Loop
				If objScriptExec.ExitCode = 0 Then
					WshShell.LogEvent 4, "Encrypted volume has been mounted." & vbCrLf, strComputerName
					' Restarting 'Server' service
					Call RestartService()
					' Waiting 10 minutes
					WScript.Sleep 600000
					If Drv.IsReady Then
						Set objScriptExec = WshShell.Exec(strComUnmount)
						Do Until objScriptExec.Status
							WScript.Sleep 1000
						Loop
						WshShell.LogEvent 2, "Encrypted volume has been locked!" & vbCrLf, strComputerName
					End If
				Else
					WshShell.LogEvent 1, "Error mounting encrypted volume has appeared!" & vbCrLf, strComputerName
				End If
				bMonitoring = False
				Exit For
			End If
		End If
		WScript.Sleep 1000
    Next
    WScript.Sleep 3000
Wend

Set objScriptExec = Nothing
WScript.Quit(0)


' Procedure for restart 'Server' service
Sub RestartService()
	strQuery = "Select * from Win32_Service Where Name='LanmanServer'"
	Set colListOfServices = objWMIService.ExecQuery(strQuery)
	For Each objService in colListOfServices
		objService.StopService()
	Next
	Do Until objWMIService.ExecQuery(strQuery & " AND State='Stopped'").Count > 0
		WScript.Sleep 1000
	Loop
	For Each objService in colListOfServices
		objService.StartService()
	Next
End Sub
