'--------------------------------------------------------------------------------------
' Reporting if users have administrative privileges on the local computer
' (excepting users in 'support' group)
' Text report is generated in the shared folder
' Script is designed to run on user logon via group policy
' Author: Valentin Vakhrushev, 2012-06-19
'--------------------------------------------------------------------------------------

On Error Resume Next

' Report for users in this group is not creating
' Users have to be directly in this group to be ignored
strSupportUserGroup = "Support Team"
' Path to shared folder for storing reports
' All authenticated users must have rights to create and change files in this folder
strOutputFolder = "\\srv01\Users\"

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")

Set objSysInfo = CreateObject("ADSystemInfo")
strUserDN = objSysInfo.UserName
Set ObjUser = GetObject("LDAP://" & strUserDN)
If Err.Number <> 0 Then WScript.Quit()

For Each objGroup In ObjUser.Groups
	UserGroups = UserGroups & "[" & objGroup.CN & "]"
Next
If InStr(UserGroups, "[" & strSupportUserGroup & "]") Then WScript.Quit()

strCompName = WshShell.ExpandEnvironmentStrings("%computername%")

strKey = CreateObject("WScript.Shell").RegRead("HKEY_USERS\s-1-5-19\")
If Err.Number = 0 Then
	If objFSO.FolderExists(strOutputFolder) Then
		' Report file is recreated each time script runs
		' It helps to determine has the user administrative privileges now or not
		Set TextStream = objFSO.CreateTextFile(strOutputFolder & strCompName & "_" & _
			ObjUser.sAMAccountName & ".txt", True)
		strText = "Report date: " & Date() & vbCrlf & "Computer: " & strCompName & vbCrlf & _
			"User: " & ObjUser.sAMAccountName & vbCrlf
		TextStream.Write(strText)
		TextStream.Close
	End If
End If


Function InGroup(strGroup)
	InGroup = False
	If InStr(UserGroups, "[" & strGroup & "]") Then
		InGroup = True
	End If
End Function
