'--------------------------------------------------------------------------------------
' Монтирование защищенного раздела BitLocker при вставке USB-носителя с ключом
' 
' Описание используемого алгоритма:
' Во время загрузки компьютера запускается скрипт и "ждет", когда будет вставлен
' носитель с ключевым файлом для разблокирования зашифрованного раздела. Когда флешку с
' ключом вставят, происходит автоматическая разблокировка раздела и перезапускается
' служба "Сервер" (поскольку у меня на этом разделе есть расшаренные папки). Далее
' скрипт ждет 10 минут, в течении которых необходимо извлечь флешку. Если по истечении
' этого периода флешка все еще вставлена, то зашифрованный раздел автоматически
' блокируется. Это сделано для того, чтобы сотрудники не оставляли ключевой носитель в 
' компьютере.
' 
' Алгоритм работы:
' 1) Во время загрузки компьютера запускается скрипт
'   (с помощью групповых политик, реестра, назначенных заданий и т.д.)
' 2) Скрипт ждет, когда будет вставлена флешка с определенным ключевым файлом
' 3) Монтируется зашифрованный раздел BitLocker (используя ключ с флешки)
' 4) Поскольку в моей ситуации на зашифрованном разделе находится расшаренная папка,
'    то производится перезапуск службы "Сервер"
' 5) Скрипт ждет 10 минут - в течении этого времени необходимо изъять флешку
' 6) Если через 10 минут флешку не вытащили (забыли), то раздел снова блокируется
'    (сделано для того, чтобы принудить сотрудников извлекать флешку с ключом)
' 7) Все сообщения пишутся в журнал событий
' 
' В скрипте реализован механизм для автоматического расшаривания сетевой папки.
' Но, поскольку сетевая папка восстанавливается при перезапуске службы "Сервер", то
' расшаривать папку повторно не имеет смысла. Данный функционал закомментирован и
' оставлен "на всякий случай".
' Скрипт идеально подойдет для использования в удаленном филиале, т.к. от сотрудников
' требуется только вставить флешку во время загрузки компьютера и через некоторое время
' ее вытащить (и спрятать :)). В случае изъятия компьютера вся важная информация 
' останется защищенной.
' 
' Разработчик: Вахрушев Валентин, 2013
'--------------------------------------------------------------------------------------

On Error Resume Next

Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\CIMV2")
Set colProc = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE CommandLine LIKE '%" & _
	WScript.ScriptName & "%'")
If colProc.Count > 1 Then WScript.Quit()

Set WshShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

strComputerName = WshShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
strKeyFileSubPath = "1234ABCD-EF123-4567-89AB-CDEF98765432.BEK"
strCommand = "manage-bde -unlock D: -RecoveryKey "
strComUnmount = "manage-bde -lock D: -ForceDismount"
'strSharedFolderPath = "D:\Works"
'strSharedFolderName = "Works"
'strSharedFolderInfo = "Рабочая папка"
bMonitoring = True

colDrives = Split("F: G: H: I: J:")
While bMonitoring = True
	For Each Drive In colDrives
		bMounted = True
		Err.Clear
		Set Drv = objFSO.GetDrive(Drive)
		If Err.Number Or Drv.IsReady = False Or Drv.DriveType <> 1 Then bMounted = False
		If bMounted Then
			If objFSO.FileExists(Drive & "\" & strKeyFileSubPath) Then
				' Монтирование зашифрованного раздела
				strCommand = strCommand & Chr(34) & Drive & "\" & strKeyFileSubPath & Chr(34)
				Set objScriptExec = WshShell.Exec(strCommand)
				Do Until objScriptExec.Status
					WScript.Sleep 1000
				Loop
				If objScriptExec.ExitCode = 0 Then
					'Call ShareSec (strSharedFolderPath, strSharedFolderName, strSharedFolderInfo)
					WshShell.LogEvent 4, "Защищенный диск примонтирован." & vbCrLf, strComputerName
					' Перезапуск службы "Сервер"
					Call RestartService()
					' Ожидание (10 минут)
					WScript.Sleep 600000
					If Drv.IsReady Then
						Set objScriptExec = WshShell.Exec(strComUnmount)
						Do Until objScriptExec.Status
							WScript.Sleep 1000
						Loop
						WshShell.LogEvent 2, "Примонтированный диск был отключен!" & vbCrLf, strComputerName
					End If
				Else
					WshShell.LogEvent 1, "Ошибка монтирования диска!" & vbCrLf, strComputerName
				End If
				bMonitoring = False
				Exit For
			End If
		End If
		WScript.Sleep 1000
    Next
    WScript.Sleep 3000
Wend

Set objScriptExec = Nothing
WScript.Quit(0)


' Процедура для перезапуска службы "Сервер"
Sub RestartService()
	strQuery = "Select * from Win32_Service Where Name='LanmanServer'"
	Set colListOfServices = objWMIService.ExecQuery(strQuery)
	For Each objService in colListOfServices
		objService.StopService()
	Next
	Do Until objWMIService.ExecQuery(strQuery & " AND State='Stopped'").Count > 0
		WScript.Sleep 1000
	Loop
	For Each objService in colListOfServices
		objService.StartService()
	Next
End Sub

' Процедура для расшаривания сетевой папки
' Процедура оптимизирована для рабочей группы и открывает полный доступ к папке всем пользователям
Sub ShareSec(strFolderPath, strShareName, strInfo)
	Set Services = GetObject("WINMGMTS:{impersonationLevel=impersonate,(Security)}!\\.\ROOT\CIMV2")
	Set SecDescClass = Services.Get("Win32_SecurityDescriptor")
	Set SecDesc = SecDescClass.SpawnInstance_()

	Set Trustee = Services.Get("Win32_Trustee").SpawnInstance_
	Trustee.Domain = Null
	Trustee.Name = "Все"
	Trustee.Properties_.Item("SID") = Array(1, 1, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0)
	
	Set ACE = Services.Get("Win32_Ace").SpawnInstance_
	ACE.Properties_.Item("AccessMask") = 2032127
	ACE.Properties_.Item("AceFlags") = 3
	ACE.Properties_.Item("AceType") = 0
	ACE.Properties_.Item("Trustee") = Trustee
	SecDesc.Properties_.Item("DACL") = Array(ACE)
	Set Share = Services.Get("Win32_Share")
	Set InParam = Share.Methods_("Create").InParameters.SpawnInstance_()
	InParam.Properties_.Item("Access") = SecDesc
	InParam.Properties_.Item("Description") = strInfo
	InParam.Properties_.Item("Name") = strShareName
	InParam.Properties_.Item("Path") = strFolderPath
	InParam.Properties_.Item("Type") = 0
	Share.ExecMethod_ "Create", InParam
End Sub
