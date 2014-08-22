'--------------------------------------------------------------------------------------
' Changing AD user properties to allow or deny OWA and OMA access depending on user group
' Author: Valentin 'sm4sh1k', 2012
'--------------------------------------------------------------------------------------

On Error Resume Next

' OWA access will be gained to users in this group, otherwise access will be denied
' Users have to be directly in this group to gain access to OWA
strOWAGroup = "OWA Users"

Set iAdRootDSE = GetObject("LDAP://RootDSE")
strDistinguishedName = iAdRootDSE.Get("defaultNamingContext")
' You can determine direct path to your users OU in AD instead of search at the full domain tree
'strDistinguishedName = "OU=Users,OU=MyBusiness,DC=Domain,DC=local"

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCOmmand.ActiveConnection = objConnection

' We obtain active users with e-mail addresses
objCommand.CommandText = "<LDAP://" & strDistinguishedName & ">;(&(objectCategory=person)" & _
	"(objectClass=user)(mail=*)" & "(!(userAccountControl:1.2.840.113556.1.4.803:=2)))" & _
	";distinguishedName;subtree"

objCommand.Properties("Page Size") = 1000
objCommand.Properties("Timeout") = 30 
objCommand.Properties("Searchscope") = 2 
objCommand.Properties("Cache Results") = False

Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst
Do Until objRecordSet.EOF
	Set objUser=GetObject("LDAP://" & objRecordSet.Fields("distinguishedName").Value)
	strUserGroups = vbNullString
	For Each objGroup In objUser.Groups
		strUserGroups = strUserGroups & "[" & objGroup.Name & "]"
	Next
	If InGroup(strOWAGroup) Then
		If InStr(objUser.protocolSettings, "HTTP§0") <> 0 Then
			objUser.protocolSettings = Replace(objUser.protocolSettings, "HTTP§0", "HTTP§1", 1, 1)
			objUser.SetInfo
		End If
		If objUser.msExchOmaAdminWirelessEnable <> "0" Then
			objUser.msExchOmaAdminWirelessEnable = "0"
			objUser.SetInfo
		End If
	Else
		If InStr(objUser.protocolSettings, "HTTP§0") = 0 Then
			objUser.protocolSettings = "HTTP§0§1§§§§§§"
			objUser.SetInfo
		End If
		If objUser.msExchOmaAdminWirelessEnable <> "7" Then
			objUser.msExchOmaAdminWirelessEnable = "7"
			objUser.SetInfo
		End If
	End If
	objRecordSet.MoveNext
Loop


Function InGroup(strGroup)
	InGroup = False
	If InStr(strUserGroups, "[CN=" & strGroup & "]") Then
		InGroup = True
	End If
End Function
