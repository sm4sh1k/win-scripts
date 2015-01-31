'--------------------------------------------------------------------------------------
' �������������� �������� ������� 1� ��� ����� ������������� � ������
' ������ ������� ���� ibases.v8i � ������� ������������ � ��������� ����� ������
' �������������� ������������� ��������� � ���������� ����������
' �����: �������� ��������, 2010
'--------------------------------------------------------------------------------------

On Error Resume Next

Const adVarChar = 200
Const MaxCharacters = 255

Set WSHShell = WScript.CreateObject("WScript.Shell")
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set WSHNetwork = WScript.CreateObject("WScript.Network")
Set objSysInfo = CreateObject("ADSystemInfo")
If (Err.Number <> 0) Then WScript.Quit()
strUserDN = objSysInfo.UserName
Set ObjUser = GetObject("LDAP://" & strUserDN)

strProfileDir = WSHShell.ExpandEnvironmentStrings("%APPDATA%\1C\1CEStart")
strProfileFile = "ibases.v8i"
' ���������������� ���� �� ������� ��� 1�
' ������ ����� (���� ����������� ������� ���������): <������> <��������> <���_��> <������_1�>
' ��������: srv10-1c	���� ������		base	8.2
strConfigPath = "\\srv20-dc\Configs$\1C\1cv82.txt"
strMsg = "�������� ������� 1� ��� ������������ " & WSHNetwork.UserName & "." & vbCrlf & vbCrlf
UserGroups = vbNullString

' ���� ������� 1� ��� ������, ������ ��������� ������
If objFSO.FileExists(strProfileDir & "\" & strProfileFile) = True Then WScript.Quit()

' �������� ������ ���� ����� ������������ � Active Directory (� ������ �����������)
For Each objGroup In ObjUser.Groups
	UserGroups = UserGroups & "[" & objGroup.CN & "]"
	GetNested(objGroup)
	Err.Clear
Next

' ������� ��������� ������ ��� ������������� � ������������ �������
If (InGroup("������������ 1�") = False _
	And InGroup("�������������� 1�") = False) _
	Or InGroup("������������ �����������") = True Then
		WScript.Quit()
End If

If objFSO.FolderExists(WSHShell.ExpandEnvironmentStrings("%APPDATA%\1C\")) = False Then
	objFSO.CreateFolder WSHShell.ExpandEnvironmentStrings("%APPDATA%\1C\")
End If
If objFSO.FolderExists(strProfileDir) = False Then
	objFSO.CreateFolder strProfileDir
End If

If (Err.Number <> 0) Then
	strMsg = strMsg & "������! �� ������� ������� ����� �������."
	WSHShell.LogEvent 1, strMsg, WSHNetwork.ComputerName
	WScript.Quit()
End If

' ������� ����������������� ����� � ��������� ����� ������� 1�
If objFSO.FileExists(strConfigPath) = True Then
	Set DataList = CreateObject("ADOR.Recordset")
	DataList.Fields.Append "Srvr", adVarChar, MaxCharacters
	DataList.Fields.Append "Ref", adVarChar, MaxCharacters
	DataList.Fields.Append "Name", adVarChar, MaxCharacters
	DataList.Fields.Append "Ver", adVarChar, MaxCharacters
	DataList.Open
	
	Set TextStream = objFSO.OpenTextFile(strConfigPath, 1)
	While Not TextStream.AtEndOfStream
		strText = TextStream.ReadLine()
		If Len(strText) > 2 Then
			strArray = Split(strText, vbTab)
			DataList.AddNew
			DataList("Srvr") = strArray(0)
			DataList("Ref") = strArray(1)
			DataList("Name") = strArray(2)
			DataList("Ver") = strArray(3)
			DataList.Update
		End If
	Wend
	TextStream.Close
	
	DataList.MoveFirst
	Do Until DataList.EOF
		Call Sub_WriteSection()
		DataList.MoveNext
	Loop
	
	' ��� ��������� ������� � ������ �������
	If Err.Number = 0 Then
		WSHShell.LogEvent 4, strMsg, WSHNetwork.ComputerName
	Else
		strMsg = strMsg & vbCrlf & "������! ��� ������: " & Err.Number & "." & vbCrlf & _
			"��������: " & Err.Description
		WSHShell.LogEvent 1, strMsg, WSHNetwork.ComputerName
	End If
Else
	strMsg = strMsg & "������! ���������������� ���� �� ������."
	WSHShell.LogEvent 1, strMsg, WSHNetwork.ComputerName
End If


' ��������� ��������� ������ ����������������� ����� 1�
Sub Sub_WriteSection()
	Randomize
	Set TextStream = objFSO.OpenTextFile(strProfileDir & "\" & strProfileFile, 8, True)
	TextStream.WriteLine "[" & DataList.Fields.Item("Name") & "]"
	TextStream.WriteLine "Connect=Srvr=" & Chr(34) & DataList.Fields.Item("Srvr") & _
		Chr(34) & ";Ref=" & Chr(34) & DataList.Fields.Item("Ref") & Chr(34) & ";"
	TextStream.WriteLine "ID=" & GenerateID()
	TextStream.WriteLine "OrderInList=" & CStr(Int(1 + (Rnd() * 65535)))
	TextStream.WriteLine "Folder=/"
	TextStream.WriteLine "OrderInTree=" & CStr(Int(1 + (Rnd() * 65535)))
	TextStream.WriteLine "External=0"
	TextStream.WriteLine "ClientConnectionSpeed=Normal"
	TextStream.WriteLine "App=ThickClient"
	TextStream.WriteLine "WA=0"
	TextStream.WriteLine "Version=" & DataList.Fields.Item("Ver")
	TextStream.Close
	strMsg = strMsg & "��������� ������ ��� " & DataList.Fields.Item("Srvr") & "." & vbCrlf
End Sub

' ������� ��� ��������� ����������� ID ���� � �������
Function GenerateID()
	GenerateID = GenDig & GenDig & GenDig & GenDig & GenDig & GenDig & GenDig & GenDig & "-" & _
		GenDig & GenDig & GenDig & GenDig & "-" & GenDig & GenDig & GenDig & GenDig & "-" & _
		GenDig & GenDig & GenDig & GenDig & "-" & GenDig & GenDig & GenDig & GenDig & _
		GenDig & GenDig & GenDig & GenDig & GenDig & GenDig & GenDig & GenDig
End Function

' ������� ��� ��������� ����� � ����������������� �������
Function GenDig()
	Randomize
	strDigit = vbNullString
	nDigit = Int(0 + (Rnd() * 15))
	Select Case nDigit
		Case 10
			strDigit = "a"
		Case 11
			strDigit = "b"
		Case 12
			strDigit = "c"
		Case 13
			strDigit = "d"
		Case 14
			strDigit = "e"
		Case 15
			strDigit = "f"
		Case Else
			strDigit = CStr(nDigit)
	End Select
	GenDig = strDigit
End Function

' ������� ��� �������� ������ �� ������������ � ������
Function InGroup(strGroup)
	InGroup = False
	If InStr(UserGroups, "[" & strGroup & "]") Then
		InGroup = True
	End If
End Function

' ������� ��� ������ ���� ��������� ����� ������������
Function GetNested(objGroup)
	On Error Resume Next
	colMembers = objGroup.GetEx("memberOf")
	For Each strMember in colMembers
		strPath = "LDAP://" & strMember
		Set objNestedGroup = GetObject(strPath)
		UserGroups = UserGroups & "[" & objNestedGroup.CN & "]"
		GetNested(objNestedGroup)
	Next
End Function
