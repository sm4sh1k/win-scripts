'--------------------------------------------------------------------------------------
' Automatic and silent computer shutdown (without warning messages)
' 
' Script is intended to run as scheduled job from shared folder. This job can be
' created with Group Policy Settings or some other scripts (or maybe manually). Also
' script uses a file with list of exceptions. If file contains name of current computer
' this computer will not be powered off.
' 	
' Author: Valentin Vakhrushev, 2014
'--------------------------------------------------------------------------------------

On Error Resume Next

Set WshShell = CreateObject("WScript.Shell")
strComputerName = WshShell.ExpandEnvironmentStrings("%COMPUTERNAME%")

' Running script only on workstations (in my case all servers have string 'srv' in their names)
' You can change this algorithm to determine OS name or simply remove next line
If InStr(LCase(strComputerName), "srv") <> 0 Then WScript.Quit()

' Now we try to determine appropriate server name storing settings file (with exceptions list)
' In my case there are few servers with the same shared folder for each branch
' At the main office there is dedicated file server, at branch offices shared folders are placed 
' directly on domain controllers. The logic is simple: if the computer is in the main site of AD, 
' then use server SRV01. If not, use domain controller of your site.
Set objSystemInfo = CreateObject("ADSystemInfo")
If Err.Number <> 0 Then WScript.Quit()
If objSystemInfo.SiteName = "MainOffice" Then
	strServerName = "SRV01"
Else
	Set objDomain = GetObject("LDAP://rootDse")
	strServerName = objDomain.Get("dnsHostName")
	strServerName = Left(strServerName, InStr(strServerName, ".") - 1)
End If

' Path and filename of settings file (with exceptions list)
strSettingsFile = "\\" & strServerName & "\Configs$\Shutdown\except_list.txt"

Set objFSO = CreateObject("Scripting.FileSystemObject")
If Not objFSO.FileExists(strSettingsFile) Then WScript.Quit()

' Reading information from settings file (assume each line is a computer name)
Set TextStream = objFSO.OpenTextFile(strSettingsFile, 1)
While (Not TextStream.AtEndOfStream)
	If LCase(TextStream.ReadLine()) = LCase(strComputerName) Then WScript.Quit()
Wend
TextStream.Close

' Shutdown with writing a message to EventLog
WshShell.LogEvent 4, "Automatic computer shutdown is initiated.", strComputerName
Err.Clear
WshShell.Run "shutdown -s -f -t 10 -c " & Chr(34) & "Automatic computer shutdown" & Chr(34), 0
If Err.Number <> 0 Then
	WshShell.LogEvent 2, "Error while running standard shutdown utility. Trying another way..." & _
		vbCrLf & "Error code: " & Err.Number, strComputerName
	Set colOperatingSystems = GetObject("winmgmts:{(Shutdown)}").ExecQuery("Select * from Win32_OperatingSystem")
	For Each objOperatingSystem in colOperatingSystems
		ObjOperatingSystem.Win32Shutdown(1)
	Next
End If
