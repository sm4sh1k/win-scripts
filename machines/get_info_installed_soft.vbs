'--------------------------------------------------------------------------------------
' Generating report about installed software
' Script creates text file with list of installed software on current PC
' Designed to run as scheduled job installed with Group Policy Settings
' TODO: Add 32bit applications support on 64bit OS
' Author: Valentin Vakhrushev, 2010
'--------------------------------------------------------------------------------------

On Error Resume Next

Const HKEY_LOCAL_MACHINE = &H80000002

Set WshNetwork = CreateObject("WScript.Network")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\CIMV2")

' In my environment all reports are placed in shared folder on fileserver
' Type here appropriate path and filename of the report
strOutFile = "\\SRV01\Info$\Soft\" & WshNetwork.ComputerName & ".txt"

strMsg =  "Report date: " & Date() & vbCrlf & "Computer: " & WshNetwork.ComputerName & vbCrlf
nCount = 1

Set colSettings = objService.ExecQuery ("Select * from Win32_OperatingSystem")
For Each objOperatingSystem in colSettings
	strMsg = strMsg & "Description: " & objOperatingSystem.Description & vbCrlf & vbCrlf
Next

Set IPConfigSet = objService.ExecQuery _
	("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")
For Each IPConfig in IPConfigSet
	If Not IsNull(IPConfig.IPAddress) Then 
		For Each IPAddress In IPConfig.IPAddress
			strMsg = strMsg & "IP address: " & IPAddress & vbCrlf
		Next
		strMsg = strMsg & "MAC address: " & IPConfig.MACAddress & vbCrlf & vbCrlf
		Exit For
	End If
Next

Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys
For Each subkey In arrSubKeys
	EnumerateValues (strKeyPath & "\" & subkey)
Next

Set txtStreamOut = objFSO.OpenTextFile(strOutFile, 2, True)
txtStreamOut.Write strMsg
txtStreamOut.Close


Sub EnumerateValues(strSubKey)
	Dim DisplayName, UninstallString
	objReg.EnumValues HKEY_LOCAL_MACHINE, strSubKey, arrValueNames, arrValueTypes
	
	For Each strValueName In arrValueNames
		objReg.GetStringValue HKEY_LOCAL_MACHINE, strSubKey, strValueName, strValue
		If UCase(strValueName) = "DISPLAYNAME" Then DisplayName = strValue
		If UCase(strValueName) = "UNINSTALLSTRING" Then UninstallString = strValue
	Next
	If DisplayName <> vbNullString Then
		strMsg = strMsg & nCount & vbTab & DisplayName & vbCrlf 
			'& vbNewLine & ">  " & UninstallString & vbCrlf
		nCount = nCount + 1
	End If
End Sub
