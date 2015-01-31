'--------------------------------------------------------------------------------------
' Creating shortcut on desktop for launching 2GIS application
' If application is not installed on local computer then script creates shortcut 
' for application located in shared folder
' Author: Valentin Vakhrushev, 2012
'--------------------------------------------------------------------------------------

On Error Resume Next

Set WSHShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objSystemInfo = CreateObject("ADSystemInfo")
If Err.Number <> 0 Then WScript.Quit()

' Running script only on workstations (in my case all servers have string 'srv' in their names)
' You can change this algorithm to determine OS name or simply remove two next lines
strComputerName = WSHShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
If InStr(LCase(strComputerName), "srv") <> 0 Then WScript.Quit()

' Marker file is needed to sure that we did it one time
strMarkerFile = WSHShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\sc_2gis_created"
' If you want to create shortcut only once, then uncomment next line
' Otherwise shortcut will be updated each time the script is running
'If objFSO.FileExists(strMarkerFile) Then WScript.Quit()

' Setting default program file location according to default installation settings
' This mechanism is universal for 32 and 64 bit OS
If WSHShell.ExpandEnvironmentStrings("%PROGRAMFILES(X86)%") = "%PROGRAMFILES(X86)%" Then
	strWorkDirPath = WSHShell.RegRead("HKLM\SOFTWARE\DoubleGIS\Grym\path")
Else
	strWorkDirPath = WSHShell.RegRead("HKLM\SOFTWARE\Wow6432Node\DoubleGIS\Grym\path")
End If
strExeFilePath = strWorkDirPath & "grym.exe"

' If application is not installed on local computer then script creates shortcut for application in shared folder
If Not objFSO.FileExists(strExeFilePath) Then
	' Now we try to determine appropriate server name where shared folder is located
	' In my case there are few servers with the same shared folder for each branch
	' At the main office there is dedicated file server, at branch offices shared folders are placed directly 
	' on domain controllers. The logic is simple: if the computer is in the main site of AD, then use server SRV01.
	' If not, use domain controller of your site. Additionally at the main site there are two OUs with their own file servers.
	' And if distinguished name of a computer contains specific OU name then it is defined appropriate server name.
	If objSystemInfo.SiteName = "MainOffice" Then
		strServerName = "SRV01"
		If InStr(objSystemInfo.ComputerName, "Security") <> 0 Then strServerName = "SEC01"
		If InStr(objSystemInfo.ComputerName, "Guest Room") <> 0 Then strServerName = "GST01"
	Else
		Set objDomain = GetObject("LDAP://rootDse")
		strServerName = objDomain.Get("dnsHostName")
		strServerName = Left(strServerName, InStr(strServerName, ".") - 1)
	End If
	strWorkDirPath = "\\" & strServerName & "\2GIS\"
	strExeFilePath = strWorkDirPath & "grym.exe"
	If Not objFSO.FileExists(strExeFilePath) Then WScript.Quit()
End If

' Creating shortcut on user's desktop
strDesktopPath = WSHShell.SpecialFolders("Desktop")
Set NewShortcut = WSHShell.CreateShortcut(strDesktopPath & "\2GIS.lnk")
NewShortcut.TargetPath = strExeFilePath
NewShortcut.WorkingDirectory = strWorkDirPath
NewShortcut.WindowStyle = 1
NewShortcut.IconLocation = strExeFilePath & ", 0"
NewShortcut.Save

objFSO.CreateTextFile strMarkerFile, True
