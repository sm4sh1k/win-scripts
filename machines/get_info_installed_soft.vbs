'--------------------------------------------------------------------------------------
' Generating report about installed software
' Script creates text file with list of installed software on current PC
' Designed to run as scheduled job installed with Group Policy Settings
' Author: Valentin Vakhrushev, 2010-2017
'--------------------------------------------------------------------------------------

On Error Resume Next

Const HKEY_LOCAL_MACHINE = &H80000002

Set WshShell = WScript.CreateObject("WScript.Shell")
Set WshNetwork = CreateObject("WScript.Network")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\CIMV2")

' In my environment all reports are placed in shared folder on fileserver
' Type here appropriate path and filename of the report
strOutFile = "\\SRV01\Info$\Soft\" & WshNetwork.ComputerName & ".txt"

Set txtStreamOut = objFSO.OpenTextFile(strOutFile, 2, True)
txtStreamOut.WriteLine "Report date: " & Date()
txtStreamOut.WriteLine "Computer: " & WshNetwork.ComputerName
nCount = 1

Set colSettings = objService.ExecQuery ("Select * from Win32_OperatingSystem")
For Each objOperatingSystem in colSettings
	txtStreamOut.WriteLine "Description: " & objOperatingSystem.Description
	txtStreamOut.WriteLine
Next

Set IPConfigSet = objService.ExecQuery _
	("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")
For Each IPConfig in IPConfigSet
	If Not IsNull(IPConfig.IPAddress) Then 
		For Each IPAddress In IPConfig.IPAddress
			txtStreamOut.WriteLine  "IP address: " & IPAddress
		Next
		txtStreamOut.WriteLine  "MAC address: " & IPConfig.MACAddress
		txtStreamOut.WriteLine
		Exit For
	End If
Next

Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys
For Each subkey In arrSubKeys
	EnumerateValues (strKeyPath & "\" & subkey)
Next

' Get information about 32bit applications on 64bit OS
If WshShell.ExpandEnvironmentStrings("%PROGRAMFILES(X86)%") <> "%PROGRAMFILES(X86)%" Then
	strKeyPath = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
	objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys
	For Each subKey In arrSubKeys
		EnumerateValues (strKeyPath & "\" & subKey)
	Next
End If

txtStreamOut.Close


Sub EnumerateValues(strSubKey)
	objReg.EnumValues HKEY_LOCAL_MACHINE, strSubKey, arrValueNames, arrValueTypes
	For Each strValueName In arrValueNames
		objReg.GetStringValue HKEY_LOCAL_MACHINE, strSubKey, strValueName, strValue
		If UCase(strValueName) = "DISPLAYNAME" Then DisplayName = strValue
	Next
	If DisplayName <> vbNullString Then
		txtStreamOut.WriteLine nCount & vbTab & DisplayName
		nCount = nCount + 1
	End If
End Sub
