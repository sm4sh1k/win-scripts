'--------------------------------------------------------------------------------------
' Creating folders for AD users in the shared folder and granting rights for them
' Author: Valentin 'sm4sh1k', 2012
'--------------------------------------------------------------------------------------

Option Explicit
On Error Resume Next

Dim strDistinguishedName, strFolderPath, nUserObjectCount, strMessage, strAccName
Dim objFSO, objConnection, objCommand, objRecordSet, WshNetwork, WshShell

Const ADS_SCOPE_SUBTREE = 2

' Write here path to users OU in your domain
strDistinguishedName = "OU=Users,OU=MyBusiness,DC=Domain,DC=local"
' Write here path to shared folder
strFolderPath = "\\Srv01\UsersShare$"
nUserObjectCount = 0
strMessage = vbNullstring

Set WshShell = WScript.CreateObject("WScript.Shell")
Set WshNetwork = WScript.CreateObject("WScript.Network")
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
If objFSO.FolderExists(strFolderPath) = False Then WScript.Quit()

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection
' We take only enabled user accounts
objCommand.CommandText = "<LDAP://" & strDistinguishedName & ">;(&(objectCategory=person)" & _
	"(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=2)));sAMAccountName;subtree"
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Timeout") = 30 
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
objCommand.Properties("Cache Results") = False

Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst
Do Until objRecordSet.EOF
	strAccName = objRecordSet.Fields("sAMAccountName").Value
	If objFSO.FolderExists(strFolderPath & "\" & strAccName) = False Then
		Err.Clear
		objFSO.CreateFolder strFolderPath & "\" & strAccName
		If Err.Number = 0 Then
			strMessage = strMessage & strAccName & " - " & Set_Security(strAccName, strFolderPath, strAccName, 2)
			nUserObjectCount = nUserObjectCount + 1
		End If
	End If
	objRecordSet.MoveNext
Loop

' Write an EventLog message if we created some folders
If nUserObjectCount > 0 Then
	strMessage = "Shared folders structure was updated at " & Date() & VbCrLf & _
		"Number of created items: " & nUserObjectCount & VbCrLf & VbCrLf & _
		"New folders for users:" & VbCrLf & strMessage
	WshShell.LogEvent 4, strMessage, WshNetwork.ComputerName
End If

Set objFSO = Nothing
Set objConnection = Nothing
Set objCommand = Nothing
Set objRecordSet = Nothing
Set WshNetwork = Nothing
Set WshShell = Nothing
Wscript.Quit(0)


Function Set_Security(strUser, strPath, strFolder, intAccessMask)
	Dim objWMI, objSecSettings, objSD, objItem, objWSNet, arrACE, intResult
	Dim objCollection, objSID, objTrustee, objNewACE, objGroup
	Dim strComputer, strDomain, strUserSID, strResult
	Const strNetDrive = "K:"

	Const ACCESS_ALLOWED = 						0
	Const ACCESS_DENIED = 						1
	
	CONST ALLOW_INHERIT =						33796
	CONST DENY_INHERIT =						37892
	Const SE_OWNER_DEFAULTED =					1
	Const SE_GROUP_DEFAULTED =					2
	Const SE_DACL_PRESENT =						4
	Const SE_DACL_DEFAULTED =					8
	Const SE_SACL_PRESENT =						16
	Const SE_SACL_DEFAULTED =					32
	Const SE_DACL_AUTO_INHERIT_REQ =			256
	Const SE_SACL_AUTO_INHERIT_REQ =			512
	Const SE_DACL_AUTO_INHERITED =				1024
	Const SE_SACL_AUTO_INHERITED =				2048
	Const SE_DACL_PROTECTED =					4096
	Const SE_SACL_PROTECTED =					8192
	Const SE_SELF_RELATIVE =					32768
	
	const ADS_ACEFLAG_FOLDER_ONLY =				0
	const ADS_ACEFLAG_FOLDER_FILES =			1
	const ADS_ACEFLAG_FOLDER_SUBFOLDERS =		2
	const ADS_ACEFLAG_FOLDER_SUBFOLDERS_FILES =	3
	const ADS_ACEFLAG_FILES_ONLY =				9
	const ADS_ACEFLAG_SUBFOLDERS_ONLY =			10
	const ADS_ACEFLAG_SUBFOLDERS_FILES_ONLY =	11
	
	Const FLAG_SYNCHRONIZE = 					1048576
	Const VIEW_FOLDERS_EXECUTE_FILES = 			32
	Const LIST_FOLDER_READ_DATA = 				1
	Const READ_ATTRIBUTES = 					128
	Const READ_ADDITIONAL_ATTRIBUTES = 			8
	Const CREATE_FILES_WRITE_DATA = 			2
	Const CREATE_FOLDERS_APPEND_DATA = 			4
	Const WRITE_ATTRIBUTES = 					256
	Const WRITE_ADDITIONAL_ATTRIBUTES = 		16
	Const DEL_SUBFOLDERS_FILES = 				64
	Const DEL = 								65536
	Const READ_DAC = 							131072
	Const WRITE_DAC = 							262144
	Const WRITE_OWNER = 						524288
	Const ACCESS_SYSTEM_SECURITY = 				16777216
	Const MAXIMUM_ALLOWED = 					33554432
	Const GENERIC_ALL = 						268435456
	Const GENERIC_EXECUTE = 					536870912
	Const GENERIC_WRITE =	 					1073741824
	Const GENERIC_READ = 						2147483648
	
	strResult = vbNullstring
	Set objWSNet = CreateObject("WScript.Network")
	strComputer = objWSNet.ComputerName
	strDomain = objWSNet.UserDomain
	Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\CIMV2")
	Set objCollection = objWMI.ExecQuery("SELECT SID FROM Win32_Account WHERE Name='" & _
		strUser & "' AND Domain='" & strDomain & "'")
	If objCollection.Count > 0 Then
		' beginning of creating new record for ACL
		For Each objItem In objCollection
			strUserSID = objItem.SID
		Next
		Set objSID = objWMI.Get("Win32_SID.SID='" & strUserSID & "'")
		Set objTrustee = objWMI.Get("Win32_Trustee").Spawninstance_()
		objTrustee.Domain = strDomain
		objTrustee.Name = strUser
		objTrustee.SID = objSID.BinaryRepresentation
		objTrustee.SidLength = objSID.SidLength
		objTrustee.SIDString = strUserSID
		Set objSID = Nothing
		Set objNewACE = objWMI.Get("Win32_Ace").Spawninstance_()
		objNewACE.AceType = ACCESS_ALLOWED
		Select Case intAccessMask
			Case 0: objNewACE.AccessMask = VIEW_FOLDERS_EXECUTE_FILES + LIST_FOLDER_READ_DATA + _
						READ_ATTRIBUTES + READ_ADDITIONAL_ATTRIBUTES + READ_DAC + FLAG_SYNCHRONIZE + _
						WRITE_ATTRIBUTES + WRITE_ADDITIONAL_ATTRIBUTES + _
						DEL + CREATE_FILES_WRITE_DATA + CREATE_FOLDERS_APPEND_DATA + DEL_SUBFOLDERS_FILES
			Case 1: objNewACE.AccessMask = VIEW_FOLDERS_EXECUTE_FILES + LIST_FOLDER_READ_DATA + _
						READ_ATTRIBUTES + READ_ADDITIONAL_ATTRIBUTES + READ_DAC + FLAG_SYNCHRONIZE + _
						WRITE_ATTRIBUTES + WRITE_ADDITIONAL_ATTRIBUTES
			Case 2: objNewACE.AccessMask = GENERIC_EXECUTE + GENERIC_WRITE + GENERIC_READ + DEL
			Case Else: objNewACE.AccessMask = VIEW_FOLDERS_EXECUTE_FILES + LIST_FOLDER_READ_DATA + _
						READ_ATTRIBUTES + READ_ADDITIONAL_ATTRIBUTES + READ_DAC + FLAG_SYNCHRONIZE
			End Select
		objNewACE.Trustee = objTrustee
		Set objTrustee = Nothing
		objNewACE.AceFlags = ADS_ACEFLAG_FOLDER_SUBFOLDERS_FILES
		' creating new record for ACL is finished
		objWSNet.MapNetworkDrive strNetDrive, strPath 'Left(strPath, Len(strPath))
		' trying to read security descriptor of a folder
		Set objSecSettings = objWMI.Get("Win32_LogicalFileSecuritySetting.Path='" & _
			strNetDrive & "\\" & strFolder & "'")
		If objSecSettings.GetSecurityDescriptor(objSD) = 0 Then
			' There is algorithm for adding permission to existing permission set and disabling inheritance
			'arrACE = objSD.DACL ' reading array of ACL records
			' adding new record to ACL array
			'ReDim Preserve arrACE(UBound(arrACE) + 1)
			'Set arrACE(UBound(arrACE)) = objNewACE
			'If Not CBool(objSD.ControlFlags And SE_DACL_PROTECTED) Then
				' disable inheritance
				'objSD.ControlFlags = objSD.ControlFlags + SE_DACL_PROTECTED
			'End If
			' And there is algorithm for creating new permission set with inheritance enabled
			' We create folder with inherited permissions and then add only one permission for a user
			ReDim arrACE(0)
			Set arrACE(0) = objNewACE
			objSD.DACL = arrACE
			Set objNewACE = Nothing
			Erase arrACE
			' new record to ACL array is added
			' trying to change security descriptor for a folder
			intResult = objSecSettings.SetSecurityDescriptor(objSD)
			Select Case intResult
				Case 0: strResult = strResult & "Security descriptor has been correctly modified." & vbCrLf
				Case 2: strResult = strResult & "Have no rights to read needed information." & vbCrLf
				Case 9: strResult = strResult & "Have no permissions to change attributes." & vbCrLf
				Case 21: strResult = strResult & "Parameters are not correct!" & vbCrLf
				Case Else: strResult = strResult & "Unknown error." & vbCrLf & _
							"Error code: " & intResult & vbCrLf
			End Select
		Else
			strResult = strResult & "Can't read object's security descriptor: " & _
				UCase(strPath) & "." & vbCrLf
		End If
		Set objSD = Nothing
		Set objSecSettings = Nothing
		objWSNet.RemoveNetworkDrive strNetDrive, True
		Set objWSNet = Nothing
	Else
		strResult = strResult & "Can't find username " & UCase(strUser) & "." & vbCrLf
	End If
	Set objCollection = Nothing
	Set objWMI = Nothing
	Set_Security = strResult
End Function
