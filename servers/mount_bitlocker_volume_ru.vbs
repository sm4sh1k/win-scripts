'--------------------------------------------------------------------------------------
' ������������ ����������� ������� BitLocker ��� ������� USB-�������� � ������
' 
' �������� ������������� ���������:
' �� ����� �������� ���������� ����������� ������ � "����", ����� ����� ��������
' �������� � �������� ������ ��� ��������������� �������������� �������. ����� ������ �
' ������ �������, ���������� �������������� ������������� ������� � ���������������
' ������ "������" (��������� � ���� �� ���� ������� ���� ����������� �����). �����
' ������ ���� 10 �����, � ������� ������� ���������� ������� ������. ���� �� ���������
' ����� ������� ������ ��� ��� ���������, �� ������������� ������ �������������
' �����������. ��� ������� ��� ����, ����� ���������� �� ��������� �������� �������� � 
' ����������.
' 
' �������� ������:
' 1) �� ����� �������� ���������� ����������� ������
'   (� ������� ��������� �������, �������, ����������� ������� � �.�.)
' 2) ������ ����, ����� ����� ��������� ������ � ������������ �������� ������
' 3) ����������� ������������� ������ BitLocker (��������� ���� � ������)
' 4) ��������� � ���� �������� �� ������������� ������� ��������� ����������� �����,
'    �� ������������ ���������� ������ "������"
' 5) ������ ���� 10 ����� - � ������� ����� ������� ���������� ������ ������
' 6) ���� ����� 10 ����� ������ �� �������� (������), �� ������ ����� �����������
'    (������� ��� ����, ����� ��������� ����������� ��������� ������ � ������)
' 7) ��� ��������� ������� � ������ �������
' 
' � ������� ���������� �������� ��� ��������������� ������������ ������� �����.
' ��, ��������� ������� ����� ����������������� ��� ����������� ������ "������", ��
' ����������� ����� �������� �� ����� ������. ������ ���������� ��������������� �
' �������� "�� ������ ������".
' ������ �������� �������� ��� ������������� � ��������� �������, �.�. �� �����������
' ��������� ������ �������� ������ �� ����� �������� ���������� � ����� ��������� �����
' �� �������� (� �������� :)). � ������ ������� ���������� ��� ������ ���������� 
' ��������� ����������.
' 
' �����������: �������� ��������, 2013
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
'strSharedFolderInfo = "������� �����"
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
				' ������������ �������������� �������
				strCommand = strCommand & Chr(34) & Drive & "\" & strKeyFileSubPath & Chr(34)
				Set objScriptExec = WshShell.Exec(strCommand)
				Do Until objScriptExec.Status
					WScript.Sleep 1000
				Loop
				If objScriptExec.ExitCode = 0 Then
					'Call ShareSec (strSharedFolderPath, strSharedFolderName, strSharedFolderInfo)
					WshShell.LogEvent 4, "���������� ���� �������������." & vbCrLf, strComputerName
					' ���������� ������ "������"
					Call RestartService()
					' �������� (10 �����)
					WScript.Sleep 600000
					If Drv.IsReady Then
						Set objScriptExec = WshShell.Exec(strComUnmount)
						Do Until objScriptExec.Status
							WScript.Sleep 1000
						Loop
						WshShell.LogEvent 2, "���������������� ���� ��� ��������!" & vbCrLf, strComputerName
					End If
				Else
					WshShell.LogEvent 1, "������ ������������ �����!" & vbCrLf, strComputerName
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


' ��������� ��� ����������� ������ "������"
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

' ��������� ��� ������������ ������� �����
' ��������� �������������� ��� ������� ������ � ��������� ������ ������ � ����� ���� �������������
Sub ShareSec(strFolderPath, strShareName, strInfo)
	Set Services = GetObject("WINMGMTS:{impersonationLevel=impersonate,(Security)}!\\.\ROOT\CIMV2")
	Set SecDescClass = Services.Get("Win32_SecurityDescriptor")
	Set SecDesc = SecDescClass.SpawnInstance_()

	Set Trustee = Services.Get("Win32_Trustee").SpawnInstance_
	Trustee.Domain = Null
	Trustee.Name = "���"
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
