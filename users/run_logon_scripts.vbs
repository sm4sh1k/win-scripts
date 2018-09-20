'--------------------------------------------------------------------------------------
' Script for sequentially launching another scripts from the specified network folder
' Some kind of analogue of init.d system in Linux or Group Policy scripts in Windows
' It is assumed that the script will be deployed on remote machines using Ansible
' Author: Valentin Vakhrushev, 2018
'--------------------------------------------------------------------------------------

On Error Resume Next

' Type here appropriate path to the network folder with scripts
strScriptsPath = "\\SRV01\Scripts$\Logon"

Set WSHShell = WScript.CreateObject("WScript.Shell")
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set objShellApp = CreateObject("Shell.Application")

strWSHPath = WSHShell.ExpandEnvironmentStrings("%WINDIR%\System32\wscript.exe")
strPSPath = WSHShell.ExpandEnvironmentStrings("%WINDIR%\System32\WindowsPowerShell\v1.0\powershell.exe")

' Do not launch scripts on servers (on computers with 'SRV' name part)
strComputerName = WSHShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
If InStr(LCase(strComputerName), "srv") <> 0 Then WScript.Quit()

If Not objFSO.FolderExists(strScriptsPath) Then WScript.Quit()

' It is recommended to name scripts with initial numbers (files are taken alphabetically)
Set objItems = objShellApp.NameSpace(strScriptsPath).Items()
For Each objItem In objItems
	strExtension = LCase(Mid(objItem.Path, InStrRev(objItem.Path, ".") + 1))
	' Silently and sequentially running only VBS and PowerShell scripts with 10 seconds delay
	Select Case strExtension
		Case "vbs"
			WSHShell.Run strWSHPath & " " & Chr(34) & objItem.Path & Chr(34), 0, False
			WScript.Sleep 10000
		Case "ps1"
			WSHShell.Run strPSPath & " -file " & Chr(34) & objItem.Path & Chr(34), 0, False
			WScript.Sleep 10000
	End Select
Next
