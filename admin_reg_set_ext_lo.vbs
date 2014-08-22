'--------------------------------------------------------------------------------------
' Registering file associations for LibreOffice
' Script tries to get administrative privileges if needed
' All messages and association descriptions are in russian. Change them if needed
' Author: Valentin 'sm4sh1k', 2013
'--------------------------------------------------------------------------------------

On Error Resume Next

Const HKEY_LOCAL_MACHINE = &H80000002

strKey = CreateObject("WScript.Shell").RegRead("HKEY_USERS\s-1-5-19\")
If Err.Number <> 0 Then
	Set objShell = CreateObject("Shell.Application")
	objShell.ShellExecute "wscript.exe", Chr(34) & _
		WScript.ScriptFullName & Chr(34), "", "runas", 1
	WScript.Quit()
End If

Set WshShell = WScript.CreateObject("WScript.Shell")
strMsg = "Регистрация типов файлов для LibreOffice." & vbCrlf

' Determining registry path to a program depending to OS version (32bit/64bit)
If WSHShell.ExpandEnvironmentStrings("%PROGRAMFILES(X86)%") = "%PROGRAMFILES(X86)%" Then
	strKeyPath = "SOFTWARE\LibreOffice\Layers\LibreOffice"
Else
	strKeyPath = "SOFTWARE\Wow6432Node\LibreOffice\Layers\LibreOffice"
End If

Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys
strWorkDirPath = WSHShell.RegRead("HKLM\" & strKeyPath & "\" & _
	arrSubKeys(UBound(arrSubKeys)) & "\OFFICEINSTALLLOCATION") & "program\"

If Err.Number = 0 And Len(strWorkDirPath) > 0 Then
	WshShell.RegWrite "HKCR\.doc\", "LibreOffice.Doc"
	WshShell.RegWrite "HKCR\LibreOffice.Doc\shell\open\command\", Chr(34) & _
		strWorkDirPath & "swriter.exe" & Chr(34) & " -o " & Chr(34) & "%1" & Chr(34)
	WshShell.RegWrite "HKCR\LibreOffice.Doc\", "Документ Microsoft Word 97-2003"
	WshShell.RegWrite "HKCR\LibreOffice.Doc\DefaultIcon\", strWorkDirPath & "soffice.bin,1"
	
	WshShell.RegWrite "HKCR\.docm\", "LibreOffice.Docm"
	WshShell.RegWrite "HKCR\LibreOffice.Docm\shell\open\command\", Chr(34) & _
		strWorkDirPath & "swriter.exe" & Chr(34) & " -o " & Chr(34) & "%1" & Chr(34)
	WshShell.RegWrite "HKCR\LibreOffice.Docm\", "Документ Microsoft Word"
	WshShell.RegWrite "HKCR\LibreOffice.Docm\DefaultIcon\", strWorkDirPath & "soffice.bin,1"
	
	WshShell.RegWrite "HKCR\.docx\", "LibreOffice.Docx"
	WshShell.RegWrite "HKCR\LibreOffice.Docx\shell\open\command\", Chr(34) & _
		strWorkDirPath & "swriter.exe" & Chr(34) & " -o " & Chr(34) & "%1" & Chr(34)
	WshShell.RegWrite "HKCR\LibreOffice.Docx\", "Документ Microsoft Word"
	WshShell.RegWrite "HKCR\LibreOffice.Docx\DefaultIcon\", strWorkDirPath & "soffice.bin,1"
	
	WshShell.RegWrite "HKCR\.rtf\", "LibreOffice.Rtf"
	WshShell.RegWrite "HKCR\LibreOffice.Rtf\shell\open\command\", Chr(34) & _
		strWorkDirPath & "swriter.exe" & Chr(34) & " -o " & Chr(34) & "%1" & Chr(34)
	WshShell.RegWrite "HKCR\LibreOffice.Rtf\", "Документ RTF"
	WshShell.RegWrite "HKCR\LibreOffice.Rtf\DefaultIcon\", strWorkDirPath & "soffice.bin,1"
	
	WshShell.RegWrite "HKCR\.dot\", "LibreOffice.Dot"
	WshShell.RegWrite "HKCR\LibreOffice.Dot\shell\open\command\", Chr(34) & _
		strWorkDirPath & "swriter.exe" & Chr(34) & " -o " & Chr(34) & "%1" & Chr(34)
	WshShell.RegWrite "HKCR\LibreOffice.Dot\", "Шаблон Microsoft Word 97-2003"
	WshShell.RegWrite "HKCR\LibreOffice.Dot\DefaultIcon\", strWorkDirPath & "soffice.bin,2"
	
	WshShell.RegWrite "HKCR\.dotm\", "LibreOffice.Dotm"
	WshShell.RegWrite "HKCR\LibreOffice.Dotm\shell\open\command\", Chr(34) & _
		strWorkDirPath & "swriter.exe" & Chr(34) & " -o " & Chr(34) & "%1" & Chr(34)
	WshShell.RegWrite "HKCR\LibreOffice.Dotm\", "Шаблон Microsoft Word"
	WshShell.RegWrite "HKCR\LibreOffice.Dotm\DefaultIcon\", strWorkDirPath & "soffice.bin,2"
	
	WshShell.RegWrite "HKCR\.dotx\", "LibreOffice.Dotx"
	WshShell.RegWrite "HKCR\LibreOffice.Dotx\shell\open\command\", Chr(34) & _
		strWorkDirPath & "swriter.exe" & Chr(34) & " -o " & Chr(34) & "%1" & Chr(34)
	WshShell.RegWrite "HKCR\LibreOffice.Dotx\", "Шаблон Microsoft Word"
	WshShell.RegWrite "HKCR\LibreOffice.Dotx\DefaultIcon\", strWorkDirPath & "soffice.bin,2"
	
	WshShell.RegWrite "HKCR\.pps\", "LibreOffice.Pps"
	WshShell.RegWrite "HKCR\LibreOffice.Pps\shell\open\command\", Chr(34) & _
		strWorkDirPath & "simpress.exe" & Chr(34) & " -o " & Chr(34) & "%1" & Chr(34)
	WshShell.RegWrite "HKCR\LibreOffice.Pps\", "Демонстрация Microsoft PowerPoint"
	WshShell.RegWrite "HKCR\LibreOffice.Pps\DefaultIcon\", strWorkDirPath & "soffice.bin,7"
	
	WshShell.RegWrite "HKCR\.ppt\", "LibreOffice.Ppt"
	WshShell.RegWrite "HKCR\LibreOffice.Ppt\shell\open\command\", Chr(34) & _
		strWorkDirPath & "simpress.exe" & Chr(34) & " -o " & Chr(34) & "%1" & Chr(34)
	WshShell.RegWrite "HKCR\LibreOffice.Ppt\", "Презентация Microsoft PowerPoint 97-2003"
	WshShell.RegWrite "HKCR\LibreOffice.Ppt\DefaultIcon\", strWorkDirPath & "soffice.bin,7"
	
	WshShell.RegWrite "HKCR\.pptm\", "LibreOffice.Pptm"
	WshShell.RegWrite "HKCR\LibreOffice.Pptm\shell\open\command\", Chr(34) & _
		strWorkDirPath & "simpress.exe" & Chr(34) & " -o " & Chr(34) & "%1" & Chr(34)
	WshShell.RegWrite "HKCR\LibreOffice.Pptm\", "Презентация Microsoft PowerPoint"
	WshShell.RegWrite "HKCR\LibreOffice.Pptm\DefaultIcon\", strWorkDirPath & "soffice.bin,7"
	
	WshShell.RegWrite "HKCR\.pptx\", "LibreOffice.Pptx"
	WshShell.RegWrite "HKCR\LibreOffice.Pptx\shell\open\command\", Chr(34) & _
		strWorkDirPath & "simpress.exe" & Chr(34) & " -o " & Chr(34) & "%1" & Chr(34)
	WshShell.RegWrite "HKCR\LibreOffice.Pptx\", "Презентация Microsoft PowerPoint"
	WshShell.RegWrite "HKCR\LibreOffice.Pptx\DefaultIcon\", strWorkDirPath & "soffice.bin,7"
	
	WshShell.RegWrite "HKCR\.pot\", "LibreOffice.Pot"
	WshShell.RegWrite "HKCR\LibreOffice.Pot\shell\open\command\", Chr(34) & _
		strWorkDirPath & "simpress.exe" & Chr(34) & " -o " & Chr(34) & "%1" & Chr(34)
	WshShell.RegWrite "HKCR\LibreOffice.Pot\", "Шаблон Microsoft PowerPoint 97-2003"
	WshShell.RegWrite "HKCR\LibreOffice.Pot\DefaultIcon\", strWorkDirPath & "soffice.bin,8"
	
	WshShell.RegWrite "HKCR\.potm\", "LibreOffice.Potm"
	WshShell.RegWrite "HKCR\LibreOffice.Potm\shell\open\command\", Chr(34) & _
		strWorkDirPath & "simpress.exe" & Chr(34) & " -o " & Chr(34) & "%1" & Chr(34)
	WshShell.RegWrite "HKCR\LibreOffice.Potm\", "Шаблон Microsoft PowerPoint"
	WshShell.RegWrite "HKCR\LibreOffice.Potm\DefaultIcon\", strWorkDirPath & "soffice.bin,8"
	
	WshShell.RegWrite "HKCR\.potx\", "LibreOffice.Potx"
	WshShell.RegWrite "HKCR\LibreOffice.Potx\shell\open\command\", Chr(34) & _
		strWorkDirPath & "simpress.exe" & Chr(34) & " -o " & Chr(34) & "%1" & Chr(34)
	WshShell.RegWrite "HKCR\LibreOffice.Potx\", "Шаблон Microsoft PowerPoint"
	WshShell.RegWrite "HKCR\LibreOffice.Potx\DefaultIcon\", strWorkDirPath & "soffice.bin,8"
	
	WshShell.RegWrite "HKCR\.xls\", "LibreOffice.Xls"
	WshShell.RegWrite "HKCR\LibreOffice.Xls\shell\open\command\", Chr(34) & _
		strWorkDirPath & "scalc.exe" & Chr(34) & " -o " & Chr(34) & "%1" & Chr(34)
	WshShell.RegWrite "HKCR\LibreOffice.Xls\", "Лист Microsoft Excel 97-2003"
	WshShell.RegWrite "HKCR\LibreOffice.Xls\DefaultIcon\", strWorkDirPath & "soffice.bin,3"
	
	WshShell.RegWrite "HKCR\.xlsb\", "LibreOffice.Xlsb"
	WshShell.RegWrite "HKCR\LibreOffice.Xlsb\shell\open\command\", Chr(34) & _
		strWorkDirPath & "scalc.exe" & Chr(34) & " -o " & Chr(34) & "%1" & Chr(34)
	WshShell.RegWrite "HKCR\LibreOffice.Xlsb\", "Лист Microsoft Excel"
	WshShell.RegWrite "HKCR\LibreOffice.Xlsb\DefaultIcon\", strWorkDirPath & "soffice.bin,3"
	
	WshShell.RegWrite "HKCR\.xlsm\", "LibreOffice.Xlsm"
	WshShell.RegWrite "HKCR\LibreOffice.Xlsm\shell\open\command\", Chr(34) & _
		strWorkDirPath & "scalc.exe" & Chr(34) & " -o " & Chr(34) & "%1" & Chr(34)
	WshShell.RegWrite "HKCR\LibreOffice.Xlsm\", "Лист Microsoft Excel"
	WshShell.RegWrite "HKCR\LibreOffice.Xlsm\DefaultIcon\", strWorkDirPath & "soffice.bin,3"
	
	WshShell.RegWrite "HKCR\.xlsx\", "LibreOffice.Xlsx"
	WshShell.RegWrite "HKCR\LibreOffice.Xlsx\shell\open\command\", Chr(34) & _
		strWorkDirPath & "scalc.exe" & Chr(34) & " -o " & Chr(34) & "%1" & Chr(34)
	WshShell.RegWrite "HKCR\LibreOffice.Xlsx\", "Лист Microsoft Excel"
	WshShell.RegWrite "HKCR\LibreOffice.Xlsx\DefaultIcon\", strWorkDirPath & "soffice.bin,3"
	
	WshShell.RegWrite "HKCR\.xlt\", "LibreOffice.Xlt"
	WshShell.RegWrite "HKCR\LibreOffice.Xlt\shell\open\command\", Chr(34) & _
		strWorkDirPath & "scalc.exe" & Chr(34) & " -o " & Chr(34) & "%1" & Chr(34)
	WshShell.RegWrite "HKCR\LibreOffice.Xlt\", "Шаблон Microsoft Excel 97-2003"
	WshShell.RegWrite "HKCR\LibreOffice.Xlt\DefaultIcon\", strWorkDirPath & "soffice.bin,4"
	
	WshShell.RegWrite "HKCR\.xltm\", "LibreOffice.Xltm"
	WshShell.RegWrite "HKCR\LibreOffice.Xltm\shell\open\command\", Chr(34) & _
		strWorkDirPath & "scalc.exe" & Chr(34) & " -o " & Chr(34) & "%1" & Chr(34)
	WshShell.RegWrite "HKCR\LibreOffice.Xltm\", "Шаблон Microsoft Excel"
	WshShell.RegWrite "HKCR\LibreOffice.Xltm\DefaultIcon\", strWorkDirPath & "soffice.bin,4"
	
	WshShell.RegWrite "HKCR\.xltx\", "LibreOffice.Xltx"
	WshShell.RegWrite "HKCR\LibreOffice.Xlt\shell\open\command\", Chr(34) & _
		strWorkDirPath & "scalc.exe" & Chr(34) & " -o " & Chr(34) & "%1" & Chr(34)
	WshShell.RegWrite "HKCR\LibreOffice.Xltx\", "Шаблон Microsoft Excel"
	WshShell.RegWrite "HKCR\LibreOffice.Xltx\DefaultIcon\", strWorkDirPath & "soffice.bin,4"
	
	WshShell.RegWrite "HKCR\.vsd\", "LibreOffice.Vsd"
	WshShell.RegWrite "HKCR\LibreOffice.Vsd\shell\open\command\", Chr(34) & _
		strWorkDirPath & "sdraw.exe" & Chr(34) & " -o " & Chr(34) & "%1" & Chr(34)
	WshShell.RegWrite "HKCR\LibreOffice.Vsd\", "Документ Microsoft Visio 2000/XP/2003"
	WshShell.RegWrite "HKCR\LibreOffice.Vsd\DefaultIcon\", strWorkDirPath & "soffice.bin,5"
	
	WshShell.RegWrite "HKCR\.vst\", "LibreOffice.Vst"
	WshShell.RegWrite "HKCR\LibreOffice.Vst\shell\open\command\", Chr(34) & _
		strWorkDirPath & "sdraw.exe" & Chr(34) & " -o " & Chr(34) & "%1" & Chr(34)
	WshShell.RegWrite "HKCR\LibreOffice.Vst\", "Шаблон Microsoft Visio 2000/XP/2003"
	WshShell.RegWrite "HKCR\LibreOffice.Vst\DefaultIcon\", strWorkDirPath & "soffice.bin,5"
	
	strMsg = strMsg & vbCrlf & "Регистрация завершена."
Else
	strMsg = strMsg & vbCrlf & "Не удалось обнаружить путь к директории установки LibreOffice."
End If

WScript.Echo strMsg
