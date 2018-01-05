#cs ----------------------------------------------------------------------------------------
About:	Script for remote changing local administrator password on domain computers
Author: Valentin Vakhrushev, 2016
#ce ----------------------------------------------------------------------------------------

#pragma compile(FileDescription, Изменение пароля администратора)
#pragma compile(ProductName, Admin Password Changer)
#pragma compile(ProductVersion, 1.0)
#pragma compile(FileVersion, 1.0.2.0)
#pragma compile(LegalCopyright, Валентин Вахрушев)

#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <GUIConstantsEx.au3>
#include <ListViewConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <FileConstants.au3>
#include <EditConstants.au3>
#include <AD.au3>
#include <Array.au3>

#NoTrayIcon

; Type here OU with Computers accounts
Global $sSearchBase = "OU=Computers,OU=MyBusiness,DC=Domain,DC=local"
Global $sSearchFilter = "(&(objectClass=computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))"

Global $iEventError = 0
Global $oMyError = ObjEvent("AutoIt.Error", "_ErrorHandle")

Global $Form = GUICreate("Изменение пароля администратора", 350, 105, -1, -1)
If @error Then Exit MsgBox(16, "Ошибка", "Ошибка при создании окна приложения." & _
	@CRLF & "Код ошибки: " & @error & ".")
GUISetIcon(@SystemDir & "\shell32.dll", -160)
GUICtrlCreateLabel("Выберите подразделение из списка:", 10, 17, 190, 13)
Global $idComboDiv = GUICtrlCreateCombo("", 207, 15, 130, 100)
GUICtrlCreateLabel("Введите и подтвердите новый пароль:", 10, 50, 200, 13)
Global $idInput1 = GUICtrlCreateInput("", 10, 68, 100, 21, $ES_PASSWORD)
Global $idInput2 = GUICtrlCreateInput("", 120, 68, 100, 21, $ES_PASSWORD)
GUISetFont(11, 700)
Global $idButtonChange = GUICtrlCreateButton("Изменить", 246, 61, 90, 29)
GUISetState(@SW_SHOW)

_GUIDisable()
_AD_Open()
If @error Then Exit MsgBox(16, "Ошибка", "Ошибка подключения к Active Directory." & _
	@CRLF & "Код ошибки: " & @error & ".")

Global $aDivisionOUs = _AD_GetAllOUs($sSearchBase, "\", 0, 1)
If @error > 0 Then Exit MsgBox(16, "Ошибка", "Ошибка чтения объектов Active Directory." & _
	@CRLF & "Код ошибки: " & @error & ".")
For $i = 1 To UBound($aDivisionOUs) - 1
	$aDivisionOUs[$i][0] = StringMid($aDivisionOUs[$i][0], StringInStr($aDivisionOUs[$i][0], "\", 0, -1) + 1)
Next
;_ArrayDisplay($aDivisionOUs, "Divisions")

_AD_Close()
_GUIEnable()

GUICtrlSetData($idComboDiv, _ArrayToString($aDivisionOUs, "", 1, -1, "|", -1, 0), $aDivisionOUs[1][0])


While 1
	Switch GUIGetMsg()
		Case $GUI_EVENT_CLOSE
			ExitLoop
		Case $idButtonChange
			_GUIDisable()
			If StringLen(GUICtrlRead($idInput1)) > 0 And GUICtrlRead($idInput1) = GUICtrlRead($idInput2) Then
				_ChangePasswords(GUICtrlRead($idComboDiv), GUICtrlRead($idInput1))
			Else
				MsgBox(16, "Ошибка", "Пароли не совпадают либо имеют нулевую длину!")
			EndIf
			_GUIEnable()
	EndSwitch
WEnd
Exit


Func _ChangePasswords($sDivision, $sPassword)
	Local $sComputersOU = "", $aComputers, $aResults[0][3]
	Local $sUserName, $oComputer, $oUser, $bError = False
	
	For $i = 1 To UBound($aDivisionOUs) - 1
		If $aDivisionOUs[$i][0] = $sDivision Then
			$sComputersOU = $aDivisionOUs[$i][1]
			ExitLoop
		EndIf
	Next
	If StringLen($sComputersOU) = 0 Then Return MsgBox(16, "Ошибка", _
		"Не удается найти указанное подразделение.")
	
	_AD_Open()
	If @error Then Return MsgBox(16, "Ошибка", "Ошибка подключения к Active Directory." & _
		@CRLF & "Код ошибки: " & @error & ".")
	
	$aComputers = _AD_GetObjectsInOU($sComputersOU, $sSearchFilter, 2, "Name", "Name")
	If @error Then Return MsgBox(16, "Ошибка", "Ошибка извлечения объектов из Active Directory." & _
		@CRLF & "Код ошибки: " & @error & ".")
	
	;_ArrayDisplay($aComputers, "Computers")
	_AD_Close()
	
	Local $aWinPos = WinGetPos("[ACTIVE]")
	ProgressOn("Выполнение...", "", "Изменение паролей администратора...", _
		$aWinPos[2]/2 + $aWinPos[0] - 150, $aWinPos[3]/2 + $aWinPos[1] - 62, 0)
	For $i = 1 To UBound($aComputers) - 1
		ProgressSet(Round(100*$i/(UBound($aComputers) - 1)))
		ReDim $aResults[UBound($aResults) + 1][3]
		$aResults[UBound($aResults) - 1][0] = $aComputers[$i]
		If Ping($aComputers[$i], 200) Then
			$sUserName = _GetAdminUserName($aComputers[$i])
			If Not @error Then
				$aResults[UBound($aResults) - 1][1] = $sUserName
				;$oComputer = ObjGet("WinNT://" & $aComputers[$i] & ",Computer")
				;$oUser = $oComputer.GetObject("User", $sUserName)
				$oUser = ObjGet("WinNT://" & $aComputers[$i] & "/" & $sUserName & ",User")
				$oUser.SetPassword($sPassword)
				If $iEventError Then
					$aResults[UBound($aResults) - 1][2] = "Недостаточно прав"
					$iEventError = 0
					$bError = True
				Else
					$oUser.SetInfo
					If $iEventError Then
						$aResults[UBound($aResults) - 1][2] = "Недостаточно прав"
						$iEventError = 0
						$bError = True
					Else
						$aResults[UBound($aResults) - 1][2] = "Пароль изменен"
					EndIf
				EndIf
			Else
				;$aResults[UBound($aResults) - 1][1] = "-"
				$aResults[UBound($aResults) - 1][2] = "Учетная запись не найдена"
				$bError = True
			EndIf
		Else
			;$aResults[UBound($aResults) - 1][1] = "-"
			$aResults[UBound($aResults) - 1][2] = "Компьютер недоступен"
			$bError = True
		EndIf
	Next
	ProgressOff()
	
	If $bError Then
		MsgBox(48, "Завершено", "Задание завершено с ошибками." & @CRLF & _
			"Не на всех компьютерах в подразделении удалось изменить пароль.")
		_ArrayDisplay($aResults, "Результаты", "", 96, "|", "Компьютер|Пользователь|Результат")
	Else
		MsgBox(64, "Завершено", "Задание успешно завершено." & @CRLF & _
			"Пароли на всех компьютерах в подразделении были изменены.")
	EndIf
EndFunc

Func _GetAdminUserName($sComputerName)
    Local $oWMIService, $oUserAccounts, $oUserAccount
    
	$oWMIService = objGet( "winmgmts:{impersonationLevel=impersonate}!//"  & $sComputerName & "/root/cimv2")
    $oUserAccounts = $oWMIService.ExecQuery("Select Name, SID from Win32_UserAccount WHERE Domain = '" & _
		$sComputerName & "'")
    
	For $oUserAccount In $oUserAccounts
        If StringLeft($oUserAccount.SID, 9) = "S-1-5-21-" And StringRight($oUserAccount.SID, 4) = "-500" Then
            Return $oUserAccount.Name
        Endif
    Next
	
	Return SetError(1, 0, 0)
EndFunc

Func _ErrorHandle()
	;$sHexNumber = Hex($oMyError.number, 8)
	;MsgBox($MB_OK, "", "We intercepted a COM Error !" & @CRLF & _
	;	"Number is: " & $sHexNumber & @CRLF & _
	;	"WinDescription is: " & $oMyError.windescription)
	$iEventError = 1
EndFunc

Func _GUIDisable()
	GUICtrlSetState($idButtonChange, $GUI_DISABLE)
	GUICtrlSetState($idComboDiv, $GUI_DISABLE)
	GUICtrlSetState($idInput1, $GUI_DISABLE)
	GUICtrlSetState($idInput2, $GUI_DISABLE)
EndFunc

Func _GUIEnable()
	GUICtrlSetState($idButtonChange, $GUI_ENABLE)
	GUICtrlSetState($idComboDiv, $GUI_ENABLE)
	GUICtrlSetState($idInput1, $GUI_ENABLE)
	GUICtrlSetState($idInput2, $GUI_ENABLE)
EndFunc
