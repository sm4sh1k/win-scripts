'--------------------------------------------------------------------------------------
' Recording information about current PC in appropriate Active Directory computer object
' Additionally: updating computer description from Active Directory computer object
' Author: Valentin Vakhrushev, 2011-2013
'--------------------------------------------------------------------------------------

On Error Resume Next

Set objSystemInfo = CreateObject("ADSystemInfo")
If Err.Number <> 0 Then WScript.Quit()

' Getting current IP and MAC addresses
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
Set IPConfigSet = objWMIService.ExecQuery _
	("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")
For Each IPConfig in IPConfigSet
	If Not IsNull(IPConfig.IPAddress) Then 
		For Each IPAddress In IPConfig.IPAddress
			strIP = IPAddress
			Exit For
		Next
		strMAC = IPConfig.MACAddress
		Exit For
	End If
Next

' Recording gathered information to Active Directory object
Set objComp = GetObject("LDAP://" & objSystemInfo.ComputerName)
If objComp.ipHostNumber <> strIP Or objComp.networkAddress <> strMAC Then
	objComp.ipHostNumber = strIP
	objComp.networkAddress = strMAC
	objComp.SetInfo
End If

' Addon: getting computer description from appropriate object in Active Directory
strDescription = objComp.Description
If strDescription = vbNullString Then WScript.Quit()
Set WSHShell = WScript.CreateObject("WScript.Shell")
strComment = WSHShell.RegRead("HKLM\SYSTEM\CurrentControlSet\Services\" & _
	"LanmanServer\Parameters\srvcomment")
If strComment <> strDescription Then
	WSHShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\" & _
		"LanmanServer\Parameters\srvcomment", strDescription, "REG_SZ"
End If
