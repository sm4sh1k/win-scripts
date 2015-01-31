'--------------------------------------------------------------------------------------
' Автоматическое создание профиля 1С для пользователей в удаленных филиалах
' Скрипт создает файл ibases.v8i в профиле пользователя с указанной базой данных
' Предполагается использование совместно с групповыми политиками
' +автоматическая конвертация описания базы данных в UTF-8
' Автор: Вахрушев Валентин, 2011
'--------------------------------------------------------------------------------------

On Error Resume Next

' Запуск скрипта на определенном сервере (сервер терминалов)
strServerName = "srv10-ts"
Set WSHNetwork = WScript.CreateObject("WScript.Network")
If LCase(WSHNetwork.ComputerName) <> strServerName Then WScript.Quit()

' Параметры базы данных 1С
strConString = "Connect=Srvr=""srv11-1c"";Ref=""demo"";"
strBaseName = "Комплексная Автоматизация (демо)"
strVersion = "8.2"

Set WSHShell = WScript.CreateObject("WScript.Shell")
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set objSysInfo = CreateObject("ADSystemInfo")
strUserDN = objSysInfo.UserName
Set ObjUser = GetObject("LDAP://" & strUserDN)

strProfileDir = WSHShell.ExpandEnvironmentStrings("%APPDATA%\1C\1CEStart")
strProfileFile = "ibases.v8i"
strMsg = "Создание профиля 1С для пользователя " & WSHNetwork.UserName & "." & vbCrlf & vbCrlf
UserGroups = vbNullString

' Если профиль 1С уже создан, скрипт завершает работу
If objFSO.FileExists(strProfileDir & "\" & strProfileFile) = True Then WScript.Quit()

' Получаем список всех групп пользователя в Active Directory (с учетом вложенности)
For Each objGroup In ObjUser.Groups
	UserGroups = UserGroups & "[" & objGroup.CN & "]"
	GetNested(objGroup)
	Err.Clear
Next

' Профиль создается только для пользователей в определенных группах
If (InGroup("Пользователи терминала") = False _
	And InGroup("Администраторы терминала") = False) _
	Or InGroup("Привлеченные специалисты") = True Then
		WScript.Quit()
End If

If objFSO.FolderExists(WSHShell.ExpandEnvironmentStrings("%APPDATA%\1C\")) = False Then
	objFSO.CreateFolder WSHShell.ExpandEnvironmentStrings("%APPDATA%\1C\")
End If
If objFSO.FolderExists(strProfileDir) = False Then
	objFSO.CreateFolder strProfileDir
End If

If (Err.Number <> 0) Then
	strMsg = strMsg & "ОШИБКА! Не удалось создать папку профиля."
	WSHShell.LogEvent 1, strMsg, WSHNetwork.ComputerName
	WScript.Quit()
End If

' Создание профиля
Call Sub_WriteSection()

' Все сообщения пишутся в журнал событий
If Err.Number = 0 Then
	WSHShell.LogEvent 4, strMsg, WSHNetwork.ComputerName
Else
	strMsg = strMsg & vbCrlf & "ОШИБКА! Код ошибки: " & Err.Number & "." & vbCrlf & _
		"Описание: " & Err.Description
	WSHShell.LogEvent 1, strMsg, WSHNetwork.ComputerName
End If


' Процедура создания профиля
Sub Sub_WriteSection()
	Randomize
	Set TextStream = objFSO.OpenTextFile(strProfileDir & "\" & strProfileFile, 8, True)
	' Конвертируем строку с описанием базы в UTF-8 (иначе в списке баз 1С будут кракозябры)
	TextStream.WriteLine "[" & ConvertString(strBaseName, "WIN", "UTF") & "]"
	TextStream.WriteLine strConString
	TextStream.WriteLine "ID=" & GenerateID()
	TextStream.WriteLine "OrderInList=" & CStr(Int(1 + (Rnd() * 65535)))
	TextStream.WriteLine "Folder=/"
	TextStream.WriteLine "OrderInTree=" & CStr(Int(1 + (Rnd() * 65535)))
	TextStream.WriteLine "External=0"
	TextStream.WriteLine "ClientConnectionSpeed=Normal"
	TextStream.WriteLine "App=ThickClient"
	TextStream.WriteLine "WA=0"
	TextStream.WriteLine "Version=" & strVersion
	TextStream.Close
	strMsg = strMsg & "Добавлена запись " & strBaseName & "." & vbCrlf
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

' Функция для конвертации строк (для корректного отображения кириллицы в списке баз 1С)
Function ConvertString(strInput,strSource,strDest)
   strSource = LCase(strSource)
   strDest = LCase(strDest)
   arrOEM = Split("128;129;130;131;132;133;134;135;136;137;138;139;140;141;142;143;144;145;146;147;148;149;150;151;152;153;154;155;156;157;158;159;160;161;162;163;164;165;166;167;168;169;170;171;172;173;174;175;224;225;226;227;228;229;230;231;232;233;234;235;236;237;238;239;240;241",";")
   arrWin = Split("192;193;194;195;196;197;198;199;200;201;202;203;204;205;206;207;208;209;210;211;212;213;214;215;216;217;218;219;220;221;222;223;224;225;226;227;228;229;230;231;232;233;234;235;236;237;238;239;240;241;242;243;244;245;246;247;248;249;250;251;252;253;254;255;168;184",";")
   arrUTF = Split("208:144;208:145;208:146;208:147;208:148;208:149;208:150;208:151;208:152;208:153;208:154;208:155;208:156;208:157;208:158;208:159;208:160;208:161;208:162;208:163;208:164;208:165;208:166;208:167;208:168;208:169;208:170;208:171;208:172;208:173;208:174;208:175;208:176;208:177;208:178;208:179;208:180;208:181;208:182;208:183;208:184;208:185;208:186;208:187;208:188;208:189;208:190;208:191;209:128;209:129;209:130;209:131;209:132;209:133;209:134;209:135;209:136;209:137;209:138;209:139;209:140;209:141;209:142;209:143;208:129;209:145",";")
   If (strSource = "win" And strDest = "win") Or (strSource = "oem" And strDest = "oem") Or (strSource = "utf" And strDest = "utf") Then
      ConvertString = strInput
      Exit Function
   End If
   If strSource = "win" Then
         arrSrc = arrWin
      ElseIf LCase(strSource) = "oem" Then
         arrSrc = arrOEM
      ElseIf LCase(strSource) = "utf" Then
         arrSrc = arrUTF
      Else
         ConvertString = "Err: The variable strSource isn't true " & strSource
         Exit Function
   End If
   If strDest = "win" Then
         arrDst = arrWin
      ElseIf strDest = "oem" Then
         arrDst = arrOEM
      ElseIf strDest = "utf" Then
         arrDst = arrUTF
      Else
         ConvertString = "Err: The variable strDest isn't true"
         Exit Function
   End If
   Set objDict = CreateObject("Scripting.Dictionary") 
   For n = 0 To UBound(arrSrc)
         objDict.Add arrSrc(n), arrDst(n)
   Next
   If (strSource = "win" And strDest = "oem") Or (strSource = "oem" And strDest = "win") Then
      For n = 1 To Len(strInput)
         If objDict.Item(CStr(Asc(Mid(strInput,n,1)))) <> "" Then
            strSymbol = strSymbol & Chr(objDict.Item(CStr(Asc(Mid(strInput,n,1)))))
         Else
            strSymbol = strSymbol & Mid(strInput,n,1)
         End If
      Next
   ElseIf strSource = "utf" Then
      For n = 1 To Len(strInput)
         If Asc(Mid(strInput,n,1)) = 208 Or Asc(Mid(strInput,n,1)) = 209 Then
            strSymbol = strSymbol & Chr(objDict.Item(CStr(Asc(Left(Mid(strInput,n,2),1)) & ":" & Asc(Right(Mid(strInput,n,2),1)))))
            n = n + 1
         Else
            strSymbol = strSymbol & Mid(strInput,n,1)
         End If
      Next
   ElseIf strDest = "utf" Then
      For n = 1 To Len(strInput)
         If objDict.Item(CStr(Asc(Mid(strInput,n,1)))) <> "" Then
            strSymbol = strSymbol & Chr(Left(objDict.Item(CStr(Asc(Mid(strInput,n,1)))),3)) & Chr(Right(objDict.Item(CStr(Asc(Mid(strInput,n,1)))),3)) 
         Else
            strSymbol = strSymbol & Mid(strInput,n,1)
         End If
      Next
   End If
   Set objDict = Nothing
   ConvertString = strSymbol
End Function
