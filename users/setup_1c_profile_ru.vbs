'--------------------------------------------------------------------------------------
' �������������� ��������������/�������� ������� 1� ��� ������������� � ������
' ������ ����������� (��� �������) ���� ibases.v8i � ������� ������������ � 
'  �� ������� ��� ������ �� ����������������� �����
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
strConfigPath = "\\srv11-fs\Configs$\1C\1cv82.txt"
strMsg = "�������������� ������� 1� ��� ������������ " & WSHNetwork.UserName & "." & vbCrlf & vbCrlf
bJustCreated = False
bSameNameRegistered = False
UserGroups = vbNullString
strData = vbNullString
bWriteLog = False

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

' ��������� ������� ����� ������� 1� � ��� ������������� ������� ���
If objFSO.FileExists(strProfileDir & "\" & strProfileFile) = False Then
	strMsg = strMsg & "������� �� ������. ����� ������ ����� ���� �������." & vbCrlf
	If objFSO.FolderExists(WSHShell.ExpandEnvironmentStrings("%APPDATA%\1C\")) = False Then
		objFSO.CreateFolder WSHShell.ExpandEnvironmentStrings("%APPDATA%\1C\")
		WScript.Sleep(100)
	End If
	If objFSO.FolderExists(strProfileDir) = False Then
		objFSO.CreateFolder strProfileDir
		WScript.Sleep(100)
	End If
	objFSO.CreateTextFile strProfileDir & "\" & strProfileFile, True
	If (Err.Number = 0) Then
		bJustCreated = True
		strMsg = strMsg & "����� ���� ������� ������." & vbCrlf
	Else
		strMsg = strMsg & "������! �� ������� ������� ���� �������."
		WSHShell.LogEvent 1, strMsg, WSHNetwork.ComputerName
		WScript.Quit()
	End If
End If

' ������� ����������������� ����� � ���������/��������� ����� ������� 1�
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
	
	' ��������� ���������� �� ������������� ������� 1�
	If bJustCreated = False Then
		Set TextStream = objFSO.OpenTextFile(strProfileDir & "\" & strProfileFile, 1)
		While Not TextStream.AtEndOfStream
			strData = strData & TextStream.ReadLine() & vbCrLf
		Wend
		TextStream.Close
	End If
	
	' ������ ������ � ���� � ���������
	' ���� ��������������� ������ ��� ���� � �������, �� �� ������� �� ��������
	DataList.MoveFirst
	Do Until DataList.EOF
		If bJustCreated = False Then
			strIn = "Connect=Srvr=" & Chr(34) & DataList.Fields.Item("Srvr") & _
				Chr(34) & ";Ref=" & Chr(34) & DataList.Fields.Item("Ref") & Chr(34) & ";"
			If InStr(strData, strIn) = 0 Then
				If InStr(strData, DataList.Fields.Item("Name")) <> 0 Then bSameNameRegistered = True
				Call Sub_WriteSection()
			Else
				strMsg = strMsg & "������ ��� " & DataList.Fields.Item("Srvr") & _
					" ��� ����������." & vbCrlf
			End If
		Else
			Call Sub_WriteSection()
		End If
		DataList.MoveNext
	Loop
	
	' ��� ��������� ������� � ������ �������
	If Err.Number = 0 Then
		If bWriteLog = True Then WSHShell.LogEvent 4, strMsg, WSHNetwork.ComputerName
	Else
		strMsg = strMsg & vbCrlf & "������! ��� ������: " & Err.Number & "." & vbCrlf & _
			"��������: " & Err.Description
		WSHShell.LogEvent 1, strMsg, WSHNetwork.ComputerName
	End If
Else
	strMsg = strMsg & "������! ���������������� ���� �� ������."
	WSHShell.LogEvent 1, strMsg, WSHNetwork.ComputerName
	WScript.Quit()
End If



' ��������� ��������� ������ ����������������� ����� 1�
Sub Sub_WriteSection()
	Randomize
	Set TextStream = objFSO.OpenTextFile(strProfileDir & "\" & strProfileFile, 8, True)
	'TextStream.WriteLine
	If bSameNameRegistered = False Then
		TextStream.WriteLine "[" & DataList.Fields.Item("Name") & "]"
	Else
		TextStream.WriteLine "[" & DataList.Fields.Item("Name") & _
			"_" & GenDig & GenDig & GenDig & GenDig & "]"
	End If
	TextStream.WriteLine "Connect=Srvr=" & Chr(34) & DataList.Fields.Item("Srvr") & _
		Chr(34) & ";Ref=" & Chr(34) & DataList.Fields.Item("Ref") & Chr(34) & ";"
	TextStream.WriteLine "ID=" & GenerateID()
	TextStream.WriteLine "OrderInList=" & CStr(Int(1 + (Rnd() * 55000)))
	TextStream.WriteLine "Folder=/"
	TextStream.WriteLine "OrderInTree=" & CStr(Int(1 + (Rnd() * 55000)))
	TextStream.WriteLine "External=0"
	TextStream.WriteLine "ClientConnectionSpeed=Normal"
	TextStream.WriteLine "App=ThickClient"
	TextStream.WriteLine "WA=0"
	TextStream.WriteLine "Version=" & DataList.Fields.Item("Ver")
	TextStream.Close
	strMsg = strMsg & "��������� ������ ��� " & DataList.Fields.Item("Srvr") & "." & vbCrlf
	bWriteLog = True
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
