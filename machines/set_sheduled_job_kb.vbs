'--------------------------------------------------------------------------------------
' Creating scheduled job for unattended setup of hotfix from shared folder
' The script is suitable for Active Directory environments without WSUS and also
' 	for situations when it is needed to install some application delivered only
' 	with EXE setup on a large quantity of machines
' Designed to run on computer startup via Group Policy politics
' Author: Valentin Vakhrushev, 2013
'--------------------------------------------------------------------------------------

On Error Resume Next

Const HKEY_LOCAL_MACHINE = &H80000002

strPatchName = "KB943729"
strCommand = "\deploy$\HOTFIX\KB943729\Windows-KB943729-x86-RUS.exe"
strParams = "/passive /norestart"
strMsg = "Creating scheduled job (" & strPatchName & ")." & vbCrlf & vbCrlf

' This hotfix is installed only on Windows XP
Set WSHShell = WScript.CreateObject("WScript.Shell")
If InStr(UCase(WshShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProductName")), _
	"WINDOWS XP") = 0 Then WScript.Quit()

' Do not create scheduled job if hotfix is allready installed
If IsPatchInstalled() = True Then WScript.Quit()

' Now we try to determine appropriate server name storing setup executable file
' In my case there are few servers with the same shared folder for each branch
' At the main office there is dedicated file server, at branch offices shared folders are placed directly 
' on domain controllers. The logic is simple: if the computer is in the main site of AD, then use server SRV01.
' If not, use domain controller of your site.
Set objSystemInfo = CreateObject("ADSystemInfo")
If Err.Number <> 0 Then WScript.Quit()
If objSystemInfo.SiteName = "MainOffice" Then
	strServerName = "SRV01"
Else
	Set objDomain = GetObject("LDAP://rootDse")
	strServerName = objDomain.Get("dnsHostName")
	strServerName = Left(strServerName, InStr(strServerName, ".") - 1)
End If
strCommand = "\\" & strServerName & strCommand

' Delete all scheduled jobs with same name and parameters
' This peace of code was added for situations when we want to change time of job launch
' We just delete old scheduled job and create new one
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colScheduledJobs = objWMIService.ExecQuery ("SELECT * FROM Win32_ScheduledJob")
For Each objJob in colScheduledJobs
	If LCase(objJob.Command) = LCase(strCommand & " " & strParams) Then
		Set objInstance = objWMIService.Get("Win32_ScheduledJob.JobID=" & objJob.JobID & "")
		objInstance.Delete
	End If
Next

' Create new scheduled job and write a message to EventLog
Set objNewJob = objWMIService.Get("Win32_ScheduledJob")
errJobCreated = objNewJob.Create (Chr(34) & strCommand & Chr(34) & " " & strParams, RunTime(), _
	False, , , True, intJobID)
If Err.Number <> 0 Then
	strMsg = strMsg & "Error! Scheduled job has not been created." & vbCrlf & _
		"Error code: " & errJobCreated
	WSHShell.LogEvent 1, strMsg, WSHShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
Else
	strMsg = strMsg & "Scheduled job has been created successfully."
	WSHShell.LogEvent 4, strMsg, WSHShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
End If



' Function for launch time generation
Function RunTime()
	' Current time plus ten minutes
	strTemp = Trim(Replace(Left(Right(DateAdd ("n", 10, Now()), 8), 5), ":", ""))
	If Len(strTemp) < 4 Then strTemp = "0" & strTemp
	' Notice for '+480' in next string! This is a time offset from GMT in minutes
	' In my case it equals +8 hours (Beijing time)
	RunTime = "********" & strTemp & "00.000000+480"
End Function

' Function determining is hotfix installed or not
' It searches the name of our hotfix defined at the beginning of this script in variable strPatchName
'  at the list of installed software
Function IsPatchInstalled()
	IsPatchInstalled = False
	Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
	strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
	objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys
	For Each SubKey In arrSubKeys
		strDisplayName = vbNullString
		strSubKey = strKeyPath & "\" & SubKey
		objReg.EnumValues HKEY_LOCAL_MACHINE, strSubKey, arrValueNames, arrValueTypes
		If Not IsNull(arrValueNames) Then
			For Each strValueName In arrValueNames
				If UCase(strValueName) = "DISPLAYNAME" Then
					objReg.GetStringValue HKEY_LOCAL_MACHINE, strSubKey, strValueName, strDisplayName
				End If
			Next
			If InStr(strDisplayName, strPatchName) <> 0 Then
				IsPatchInstalled = True
				Exit For
			End If
		End If
	Next
End Function
