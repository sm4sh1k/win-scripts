'------------------------------------------------------------------------------------------------
' Configuring Mozilla Firefox profile for proper work in domain based environment
' Modifying prefs.js for correct opening local or network web pages
' http://stackoverflow.com/questions/192080/firefox-links-to-local-or-network-pages-do-not-work
' Designed to run on user logon via Group Policy politics
' Author: Valentin Vakhrushev, 2009
'------------------------------------------------------------------------------------------------

On Error Resume Next

Set WSHShell = WScript.CreateObject("WScript.Shell")
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

' Running script only on workstations (in my case all servers have string 'srv' in their names)
' You can change this algorithm to determine OS name or simply remove two next lines
strComputerName = WSHShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
If InStr(LCase(strComputerName), "srv") <> 0 Then WScript.Quit()

' Marker file is needed to avoid change of file everytime user logins
strMarkerFile = WSHShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\ff_prefs_modified"
If objFSO.FileExists(strMarkerFile) Then WScript.Quit()

' Checking if Firefox profile folder exists
strConfFolder = WSHShell.ExpandEnvironmentStrings("%APPDATA%\Mozilla\Firefox\")
If Not objFSO.FileExists(strConfFolder & "profiles.ini") Then WScript.Quit()

strPrefsFileName = "prefs.js"
' Preference string
strSetString = "user_pref(" & Chr(34) & "capability.policy.default.checkloaduri.enabled" & _
	Chr(34) & ", " & Chr(34) & "allAccess" & Chr(34) & ");"

strText = vbNullString
Set TextStream = objFSO.OpenTextFile(strConfFolder & "profiles.ini", 1)
While Not TextStream.AtEndOfStream
	strText = strText & TextStream.ReadLine() & vbCrlf
Wend
TextStream.Close

If InStr(strText, "IsRelative=") = 0 Then WScript.Quit()
strRelativeValue = Mid(strText, InStr(strText, "IsRelative=") + Len("IsRelative="), 1)

Set objRegExp = CreateObject("VBScript.RegExp")
objRegExp.Global = True
objRegExp.Pattern = "Path=.*?\n"
Set objMatches = objRegExp.Execute(strText)
Set objMatch = objMatches.Item(0)
strOutFolder = Replace(Replace(Mid(objMatch.Value, Len("Path=") + 1), "/", "\"), vbCrlf, "") & "\"
If strRelativeValue = "1" Then strOutFolder = strConfFolder & strOutFolder
If Not objFSO.FileExists(strOutFolder & strPrefsFileName) Then WScript.Quit()

Set TextStream = objFSO.OpenTextFile(strOutFolder & strPrefsFileName, 8, True)
TextStream.WriteLine(strSetString)
TextStream.Close

objFSO.CreateTextFile strMarkerFile, True
