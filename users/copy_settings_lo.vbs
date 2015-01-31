'--------------------------------------------------------------------------------------
' Copying default user profile for LibreOffice
' Default profile is prepared separately just archiving LO folder in user's AppData folder
' Script uses 7-zip to extract folder with profile to user's account directory
' Scripts finely works on both 32 and 64bit OS, tested on Windows XP/7/8/8.1
' Script is designed to run on user logon via group policy
' Author: Valentin Vakhrushev, 2013
'--------------------------------------------------------------------------------------

On Error Resume Next

Set WSHShell = WScript.CreateObject("WScript.Shell")
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

' Don't do this operation on servers (in my case all servers have string 'srv' in their names)
' You can change this algorithm to determine OS name or simply remove two next lines
strComputerName = WSHShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
If InStr(LCase(strComputerName), "srv") <> 0 Then WScript.Quit()

' Marker file is needed to sure that we did it one time
' If you want to recreate profile in the future, just remove this file and run script again
strMarkerFile = WSHShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\conf_lo_copied"
If objFSO.FileExists(strMarkerFile) Then WScript.Quit()

' Now we try to determine appropriate server name storing LO profile
' In my case there are few servers storing archive with the same path to it
' In each branch of our company there is one server with configs
' And there I check if the computer is in the main site of AD, then use server SRV01
' If not, then use domain controller of your site (in my case there is only one DC in each branch office)
Set objSystemInfo = CreateObject("ADSystemInfo")
If Err.Number <> 0 Then WScript.Quit()
If objSystemInfo.SiteName = "MainOffice" Then
	strServerName = "SRV01"
Else
	Set objDomain = GetObject("LDAP://rootDse")
	strServerName = objDomain.Get("dnsHostName")
	strServerName = Left(strServerName, InStr(strServerName, ".") - 1)
End If
' You can set server name directly in the next line and delete all these lines above
' Also change path and name of archive for your needs
strArchiveFilePath = "\\" & strServerName & "\Scripts$\Configs\LibreOffice\lo_profile.7z"
If Not objFSO.FileExists(strArchiveFilePath) Then WScript.Quit()

strConfigFolder = WSHShell.ExpandEnvironmentStrings("%APPDATA%\LibreOffice")
strCommandLineParam = " x " & Chr(34) & strArchiveFilePath & Chr(34) & " -o" & _
	Chr(34) & WSHShell.ExpandEnvironmentStrings("%APPDATA%\") & Chr(34)

' Trying to find installed 7-zip
str7zipFolderPath = WshShell.RegRead("HKEY_USERS\.DEFAULT\Software\7-Zip\Path")
If Err.Number <> 0 Then
	Err.Clear
	str7zipFolderPath = WshShell.RegRead("HKLM\Software\7-Zip\Path")
	If Err.Number <> 0 Then
		If WSHShell.ExpandEnvironmentStrings("%PROGRAMFILES(X86)%") <> "%PROGRAMFILES(X86)%" Then
			str7zipFolderPath = WshShell.RegRead("HKLM\Software\Wow6432Node\7-Zip\Path")
		End If
		Err.Clear
	End If
End If
If Right(str7zipFolderPath, 1) <> "\" Then str7zipFolderPath = str7zipFolderPath & "\"
str7zipFilePath = str7zipFolderPath & "7z.exe"
If Not objFSO.FileExists(str7zipFilePath) Then WScript.Quit()

' Stop all running LO processes
If objFSO.FolderExists(strConfigFolder) Then
	Set objSWbemServicesEx = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set collSWbemObjectSet = objSWbemServicesEx.ExecQuery("SELECT * FROM Win32_Process " & _
		"WHERE Name = 'soffice.bin'", "WQL", 0)
	If collSWbemObjectSet.Count > 0 Then
		For Each objSWbemObjectEx In collSWbemObjectSet
			objSWbemObjectEx.Terminate(0)
			WScript.Sleep 1000
		Next
	End If
	objFSO.DeleteFolder strConfigFolder, True
End If

' Extract archive to user's profile and create marker file
ReturnCode = WSHShell.Run(Chr(34) & str7zipFilePath & Chr(34) & strCommandLineParam, 0, True)
objFSO.CreateTextFile strMarkerFile, True
