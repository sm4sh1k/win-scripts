'--------------------------------------------------------------------------------------------------
' Creating shortcut in Autostart folder to automatically run LibreOffice Quickstart on user's logon
' If application is not installed on local computer then shortcut is not created
' Author: Valentin Vakhrushev, 2013
'--------------------------------------------------------------------------------------------------

On Error Resume Next

const HKEY_LOCAL_MACHINE = &H80000002

Set WSHShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Running script only on workstations (in my case all servers have string 'srv' in their names)
' You can change this algorithm to determine OS name or simply remove two next lines
strComputerName = WSHShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
If InStr(LCase(strComputerName), "srv") <> 0 Then WScript.Quit()

' Marker file is needed to sure that we did it only one time
strMarkerFile = WSHShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\sc_lo_created"
If objFSO.FileExists(strMarkerFile) Then WScript.Quit()

' Determining appropriate registry key depending on OS version (32bit or 64bit)
If WSHShell.ExpandEnvironmentStrings("%PROGRAMFILES(X86)%") = "%PROGRAMFILES(X86)%" Then
	strKeyPath = "SOFTWARE\LibreOffice\Layers\LibreOffice"
Else
	strKeyPath = "SOFTWARE\Wow6432Node\LibreOffice\Layers\LibreOffice"
End If

' Determining installation path of LibreOffice and location of quickstart.exe file
Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys
strWorkDirPath = WSHShell.RegRead("HKLM\" & strKeyPath & "\" & _
	arrSubKeys(UBound(arrSubKeys)) & "\OFFICEINSTALLLOCATION") & "program\"
strExeFilePath = strWorkDirPath & "quickstart.exe"
If Not objFSO.FileExists(strExeFilePath) Then WScript.Quit()

' Creating shortcut in user's Autostart folder
Set NewShortcut = WSHShell.CreateShortcut(WSHShell.SpecialFolders("Startup") & _
	"\LibreOffice Quickstart.lnk")
NewShortcut.TargetPath = strExeFilePath
NewShortcut.WorkingDirectory = strWorkDirPath
NewShortcut.WindowStyle = 1
NewShortcut.IconLocation = strExeFilePath & ", 0"
NewShortcut.Save

If Err.Number = 0 Then objFSO.CreateTextFile strMarkerFile, True
