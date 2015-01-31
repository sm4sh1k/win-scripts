'--------------------------------------------------------------------------------------
' Автоматическое редактирование/создание профиля 1С для пользователей в домене
' Скрипт редактирует (или создает) файл ibases.v8i в профиле пользователя с 
'  со списком баз данных из конфигурационного файла
' Предполагается использование совместно с групповыми политиками
' Автор: Вахрушев Валентин, 2010
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
' Конфигурационный файл со списком баз 1С
' Формат файла (поля разделяются знаками табуляции): <сервер> <описание> <имя_БД> <версия_1С>
' Например: srv10-1c	База данных		base	8.2
strConfigPath = "\\srv11-fs\Configs$\1C\1cv82.txt"
strMsg = "Редактирование профиля 1С для пользователя " & WSHNetwork.UserName & "." & vbCrlf & vbCrlf
bJustCreated = False
bSameNameRegistered = False
UserGroups = vbNullString
strData = vbNullString
bWriteLog = False

' Получаем список всех групп пользователя в Active Directory (с учетом вложенности)
For Each objGroup In ObjUser.Groups
	UserGroups = UserGroups & "[" & objGroup.CN & "]"
	GetNested(objGroup)
	Err.Clear
Next

' Профиль создается только для пользователей в определенных группах
If (InGroup("Пользователи 1С") = False _
	And InGroup("Администраторы 1С") = False) _
	Or InGroup("Привлеченные специалисты") = True Then
		WScript.Quit()
End If

' Проверяем наличие файла профиля 1С и при необходимости создаем его
If objFSO.FileExists(strProfileDir & "\" & strProfileFile) = False Then
	strMsg = strMsg & "Профиль не найден. Будет создан новый файл профиля." & vbCrlf
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
		strMsg = strMsg & "Новый файл профиля создан." & vbCrlf
	Else
		strMsg = strMsg & "ОШИБКА! Не удалось создать файл профиля."
		WSHShell.LogEvent 1, strMsg, WSHNetwork.ComputerName
		WScript.Quit()
	End If
End If

' Парсинг конфигурационного файла и генерация/изменение файла профиля 1С
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
	
	' Считываем информацию из существующего профиля 1С
	If bJustCreated = False Then
		Set TextStream = objFSO.OpenTextFile(strProfileDir & "\" & strProfileFile, 1)
		While Not TextStream.AtEndOfStream
			strData = strData & TextStream.ReadLine() & vbCrLf
		Wend
		TextStream.Close
	End If
	
	' Запись данных в файл с проверкой
	' Если сгенерированная запись уже есть в профиле, то не создаем ее повторно
	DataList.MoveFirst
	Do Until DataList.EOF
		If bJustCreated = False Then
			strIn = "Connect=Srvr=" & Chr(34) & DataList.Fields.Item("Srvr") & _
				Chr(34) & ";Ref=" & Chr(34) & DataList.Fields.Item("Ref") & Chr(34) & ";"
			If InStr(strData, strIn) = 0 Then
				If InStr(strData, DataList.Fields.Item("Name")) <> 0 Then bSameNameRegistered = True
				Call Sub_WriteSection()
			Else
				strMsg = strMsg & "Запись для " & DataList.Fields.Item("Srvr") & _
					" уже существует." & vbCrlf
			End If
		Else
			Call Sub_WriteSection()
		End If
		DataList.MoveNext
	Loop
	
	' Все сообщения пишутся в журнал событий
	If Err.Number = 0 Then
		If bWriteLog = True Then WSHShell.LogEvent 4, strMsg, WSHNetwork.ComputerName
	Else
		strMsg = strMsg & vbCrlf & "ОШИБКА! Код ошибки: " & Err.Number & "." & vbCrlf & _
			"Описание: " & Err.Description
		WSHShell.LogEvent 1, strMsg, WSHNetwork.ComputerName
	End If
Else
	strMsg = strMsg & "ОШИБКА! Конфигурационный файл не найден."
	WSHShell.LogEvent 1, strMsg, WSHNetwork.ComputerName
	WScript.Quit()
End If



' Процедура генерации секции конфигурационного файла 1С
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
	strMsg = strMsg & "Добавлена запись для " & DataList.Fields.Item("Srvr") & "." & vbCrlf
	bWriteLog = True
End Sub

' Функция для генерации уникального ID базы в профиле
Function GenerateID()
	GenerateID = GenDig & GenDig & GenDig & GenDig & GenDig & GenDig & GenDig & GenDig & "-" & _
		GenDig & GenDig & GenDig & GenDig & "-" & GenDig & GenDig & GenDig & GenDig & "-" & _
		GenDig & GenDig & GenDig & GenDig & "-" & GenDig & GenDig & GenDig & GenDig & _
		GenDig & GenDig & GenDig & GenDig & GenDig & GenDig & GenDig & GenDig
End Function

' Функция для генерации цифры в шестнадцатиричном формате
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

' Функция для проверки входит ли пользователь в группу
Function InGroup(strGroup)
	InGroup = False
	If InStr(UserGroups, "[" & strGroup & "]") Then
		InGroup = True
	End If
End Function

' Функция для поиска всех вложенных групп пользователя
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
