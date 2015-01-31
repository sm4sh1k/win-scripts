'--------------------------------------------------------------------------------------
' Generating HTML with e-mail addresses of Active Directory users
' Author: Valentin Vakhrushev, 2009
'--------------------------------------------------------------------------------------

Const adVarChar = 200
Const MaxCharacters = 255
Const ADS_SCOPE_SUBTREE = 2

StartTime = Now
intUserObjectCount = 1

Set WSHShell = WScript.CreateObject("WScript.Shell")
strOutputFileName = WSHShell.SpecialFolders("Desktop") & "\mail_list.htm"

Set iAdRootDSE = GetObject("LDAP://RootDSE")
strDistinguishedName = iAdRootDSE.Get("defaultNamingContext")
' You can determine direct path to your users OU in AD instead of search at the full domain tree
'strDistinguishedName = "OU=Users,OU=MyBusiness,DC=Domain,DC=local"

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCOmmand.ActiveConnection = objConnection

' We want to obtain only active users with e-mail addresses and filled 'full name' field
objCommand.CommandText = "<LDAP://" & strDistinguishedName & ">;(&(objectCategory=person)" & _
	"(objectClass=user)(mail=*)(givenName=*)" & "(!(userAccountControl:1.2.840.113556.1.4.803:=2)))" & _
	";displayName,mail,telephoneNumber;subtree"

objCommand.Properties("Page Size") = 1000
objCommand.Properties("Timeout") = 30 
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
objCommand.Properties("Cache Results") = False

Set DataList = CreateObject("ADOR.Recordset")
DataList.Fields.Append "displayName", adVarChar, MaxCharacters
DataList.Fields.Append "telephoneNumber", adVarChar, MaxCharacters
DataList.Fields.Append "Mail", adVarChar, MaxCharacters
DataList.Open

' Filling our recordset with user data from Active Directory
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst
Do Until objRecordSet.EOF
	DataList.AddNew
	DataList("displayName") = objRecordSet.Fields("displayName").Value
	If Len(objRecordSet.Fields("telephoneNumber").Value) > 0 Then
		DataList("telephoneNumber") = objRecordSet.Fields("telephoneNumber").Value
	Else
		DataList("telephoneNumber") = "-"
	End If
	DataList("Mail") = objRecordSet.Fields("mail").Value
	DataList.Update
	objRecordSet.MoveNext
Loop

' Creating the backbone of our HTML page
' I have chosen 'windows-1251' charset to correctly display cyrillic symbols
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objOutputFileName = objFSO.OpenTextFile(strOutputFileName, 2, True)
objOutputFileName.Writeline("<html><head><meta charset=" & Chr(34) & "windows-1251" & Chr(34) & ">")
objOutputFileName.Writeline("<meta http-equiv=" & Chr(34) & "Content-Type" & Chr(34) & " content=" & _
	Chr(34) & "text/html; charset=windows-1251" & Chr(34) & ">")
objOutputFileName.Writeline("<meta http-equiv=" & Chr(34) & "Content-Language" & Chr(34) & _
	" content=" & Chr(34) & "ru" & Chr(34) & ">")
objOutputFileName.Writeline("<title>E-mail addresses list of our employees on " & _
	Date() & "</title>")
objOutputFileName.Writeline("<style type=" & Chr(34) & "text/css" & Chr(34) & ">")
objOutputFileName.Writeline("table.solid {border-width: 1px; border-spacing: 0px; border-color: gray;" & _
	" border-style: solid;}")
objOutputFileName.Writeline("table.solid td {border-width: 1px; border-color: gray; border-style: solid;}")
objOutputFileName.Writeline("</style></head>")
objOutputFileName.Writeline("</style></head>")
objOutputFileName.Writeline("<body><h1><center>E-mail addresses list of our employees on " & _
	Date() & "</center></h1>")
objOutputFileName.Writeline("<p align=center>&nbsp;</p>")
objOutputFileName.Writeline("<table class=" & Chr(34) & "solid" & Chr(34) & _
	" align=center border=1 cellpadding=1 cellspacing=0 width=70%>")
objOutputFileName.Writeline("<tr valign=center><td align=center><font size=" & Chr(34) & "3" & _
	Chr(34) & ">Number</font></td><td align=center><font size=" & Chr(34) & "3" & Chr(34) & ">" & _
	"Full Name</font></td><td align=center><font size=" & Chr(34) & "3" & Chr(34) & ">Telephone</font></td>" & _
	"<td align=center><font size=" & Chr(34) & "3" & Chr(34) & ">E-mail</font></td></tr>")
	
' Filling HTML page with data
DataList.Sort = "displayName"
DataList.MoveFirst
Do Until DataList.EOF
	objOutputFileName.Writeline("<tr valign=center><td align=right><font size=" & Chr(34) & "3" & _
		Chr(34) & ">" & intUserObjectCount & "</font></td><td><font size=" & Chr(34) & "3" & _
		Chr(34) & ">" & DataList.Fields.Item("displayName") & "</font></td><td align=center><font size=" & _
		Chr(34) & "3" & Chr(34) & ">" & DataList.Fields.Item("telephoneNumber") & "</font></td><td>" & _
		"<font size=" & Chr(34) & "3" & Chr(34) & "><a href=" & Chr(34) & "mailto:" & _
		DataList.Fields.Item("Mail") & Chr(34) & ">" & DataList.Fields.Item("Mail") & "</a></font></td></tr>")
	intUserObjectCount = intUserObjectCount + 1
	DataList.MoveNext
Loop

objOutputFileName.Writeline("</table></body></html>")
objOutputFileName.Close

EndTime = Now
MsgBox "E-mail list of employees in our company is generated on " & Date() & _
	"." & VbCrLf & VbCrLf & intUserObjectCount-1 & " records found." & VbCrLf & _
	"Time spent: " & DateDiff("s", StartTime, EndTime) & " s." & VbCrLf & VbCrLf & _
	"Document is saved on Desktop with filename  mail_list.htm.", vbInformation, _
	"E-mail list generation"
Wscript.Quit(0)
