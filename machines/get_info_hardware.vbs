'--------------------------------------------------------------------------------------
' Generating report about hardware and Microsoft products installed
' Script creates text file with list of installed hardware and some other stuff:
'  - serial numbers of installed OS and MS Office
'  - connected printers
'  - IP and MAC addresses
'  - shared folders
' Designed to run as scheduled job installed with Group Policy Settings
' Author: Valentin Vakhrushev, 2009-2018
'--------------------------------------------------------------------------------------

On Error Resume Next

Set WshShell = WScript.CreateObject("WScript.Shell")
Set WshNetwork = CreateObject("WScript.Network")
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set objService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\CIMV2")
Set objRegExp = CreateObject("VBScript.RegExp")
' Regular expression for matching IPv4 addresses
objRegExp.Pattern = "^[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}$"

' In my environment all reports are placed in shared folder on fileserver
' Type here appropriate path to folder with stored reports
strOutFolder = "\\SRV01\Info$\"
strOutFile = strOutFolder & WshNetwork.ComputerName & ".txt"

Set txtStreamOut = objFSO.OpenTextFile(strOutFile, 2, True)
If Err.Number <> 0 Then	WScript.Quit
txtStreamOut.WriteLine "Report date: " & Date()
txtStreamOut.WriteLine vbCrLf & "Computer: " & WshNetwork.ComputerName

Set colSettings = objService.ExecQuery ("Select * from Win32_OperatingSystem")
For Each objOperatingSystem in colSettings
	txtStreamOut.WriteLine "Description: " & objOperatingSystem.Description
Next

Set IPConfigSet = objService.ExecQuery _
	("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")
For Each IPConfig in IPConfigSet
	If Not IsNull(IPConfig.IPAddress) Then
		For Each IPAddress In IPConfig.IPAddress
			strIPAddress = IPAddress
			Exit For
		Next
		' We take first interface with IPv4 address
		If objRegExp.Test(strIPAddress) Then
			txtStreamOut.WriteLine "IP address: " & strIPAddress
			txtStreamOut.WriteLine "MAC address: " & IPConfig.MACAddress
			txtStreamOut.WriteLine "Given by DHCP: " & IPConfig.DHCPEnabled
			Exit For
		End If
	End If
Next

For Each objOperatingSystem in colSettings
	txtStreamOut.WriteLine vbCrLf & "Operation system: " & objObject.Caption
	txtStreamOut.WriteLine "Installation date: " & WMIDateStringToDate(objObject.InstallDate)
Next
txtStreamOut.WriteLine "Product key: " & _
	GetKeyWin(WshShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DigitalProductId"))
txtStreamOut.WriteLine "Product ID: " & WshShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT" & _
	"\CurrentVersion\ProductId")
txtStreamOut.WriteLine

Err.Clear
Set objWord = CreateObject("Word.Application")
If Err.Number = 0 Then
	txtStreamOut.WriteLine "Office suite: " _
		& WshShell.RegRead("HKLM\SOFTWARE\Microsoft\Office\" & objWord.Version &_
		"\Registration\" & objWord.ProductCode & "\ProductNameNonQualified")
	txtStreamOut.WriteLine "Product key: " & _
		GetKeyOffice(WshShell.RegRead("HKLM\SOFTWARE\Microsoft\Office\" & objWord.Version &_
		"\Registration\" & objWord.ProductCode & "\DigitalProductId"))
	txtStreamOut.WriteLine "Product ID: " & _
		WshShell.RegRead("HKLM\SOFTWARE\Microsoft\Office\" & objWord.Version &_
		"\Registration\" & objWord.ProductCode & "\ProductID")
    objWord.Quit
	txtStreamOut.WriteLine
Else
    Err.Clear
End If

txtStreamOut.WriteLine "Shared folders:"
intCount = 1
For Each objObject In objService.ExecQuery("SELECT * FROM Win32_Share")
	txtStreamOut.WriteLine intCount & ". Name: " & objObject.Name
	If Len(objObject.Path) > 0 Then 'And InStrRev(objObject.Name, "$") = 0 Then
		txtStreamOut.WriteLine "   Path: " & objObject.Path
	End If
	intCount = intCount + 1
Next

txtStreamOut.WriteLine vbCrLf & "Printers:"
intCount = 1
For Each objObject In objService.ExecQuery("SELECT * FROM Win32_Printer")
	txtStreamOut.WriteLine intCount & ". " & objObject.Name
	intCount = intCount + 1
Next

txtStreamOut.WriteLine vbCrLf & "Hardware:"
txtStreamOut.WriteLine "1. Motherboard"
For Each objObject In objService.ExecQuery("SELECT * FROM Win32_BaseBoard")
	txtStreamOut.WriteLine "Manufacturer: " & objObject.Manufacturer
	txtStreamOut.WriteLine "Model: " & objObject.Product
Next
txtStreamOut.WriteLine vbCrLf & "2. Physical Memory"
For Each objObject In objService.ExecQuery("SELECT * FROM Win32_PhysicalMemory")
	txtStreamOut.WriteLine "Slot " & objObject.DeviceLocator & ": " & (objObject.Capacity/1048576) & "MB"
Next
txtStreamOut.WriteLine vbCrLf & "3. Processor"
For Each objObject In objService.ExecQuery("SELECT * FROM Win32_Processor")
	If InStr(objObject.Name, "Intel Pentium II") = 0 Then
		txtStreamOut.WriteLine "Model: " & Trim(objObject.Name)
	Else
		txtStreamOut.WriteLine "Model: " & Trim(WshShell.RegRead("HKLM\HARDWARE\" & _
			"DESCRIPTION\System\CentralProcessor\0\ProcessorNameString"))
	End If
	txtStreamOut.WriteLine "Clock rate: " & objObject.MaxClockSpeed
	txtStreamOut.WriteLine "Socket: " & objObject.SocketDesignation
Next
txtStreamOut.WriteLine vbCrLf & "4. Video Controller"
For Each objObject In objService.ExecQuery("SELECT * FROM Win32_VideoController")
	txtStreamOut.WriteLine "Model: " & objObject.Caption
	txtStreamOut.WriteLine "Memory: " & (objObject.AdapterRAM/1048576) & "MB"
Next
txtStreamOut.WriteLine vbCrLf & "5. Network Adapter"
For Each objNtw In objService.ExecQuery("SELECT * FROM Win32_NetworkAdapter Where MACAddress=" & _
	Chr(34) & IPConfig.MACAddress & Chr(34))
	txtStreamOut.WriteLine "Model: " & objNtw.Name
	Exit For
Next
txtStreamOut.WriteLine vbCrLf & "6. Hard Drive"
For Each objObject In objService.ExecQuery("SELECT * FROM Win32_DiskDrive")
	nVolumeSize = Int(objObject.Size/1073741824)
	If nVolumeSize > 0 Then
		txtStreamOut.WriteLine intCount & ") Model: " & objObject.Caption
		txtStreamOut.WriteLine "   Capacity: " & nVolumeSize & "GB"
		intCount = intCount + 1
	End If
Next
txtStreamOut.WriteLine vbCrLf & "7. Optical Drive"
For Each objObject In objService.ExecQuery("SELECT * FROM Win32_CDROMDrive")
	txtStreamOut.WriteLine "Model: " & objObject.Caption
Next


' Function for compute serial number from DigitalProductId entry in corresponding registry
' Updated version can retrieve correct serial number for actual versions of Windows
Function GetKeyWin(rpk)
	Const rpkOffset = 52
	szPossibleChars = "BCDFGHJKMPQRTVWXY2346789"
	isWin8 = ((rpk(66) And &HFF) \ 6) And 1
	rpk(66) = (rpk(66) And &HF7) Or ((isWin8 And 2) * 4)
	i = 28
	Do
		dwAccumulator = 0
		j = 14
		Do  
			dwAccumulator = dwAccumulator * 256  
			dwAccumulator = rpk(j + rpkOffset) + dwAccumulator
			rpk(j + rpkOffset) = (dwAccumulator \ 24) And 255
			dwAccumulator = dwAccumulator Mod 24
			j = j - 1
		Loop While j >= 0
		i = i - 1
		szProductKey = Mid(szPossibleChars, dwAccumulator + 1, 1) & szProductKey
		If (((29 - i) Mod 6) = 0) And (i <> -1) Then
			i = i - 1
			szProductKey = "-" & szProductKey
		End If
	Loop While i >= 0
	If isWin8 = 1 Then
		If dwAccumulator = 0 Then
			szProductKey = "N" & szProductKey
		Else
			szProductKey = Mid(szProductKey, 2, dwAccumulator) & "N" & Mid(szProductKey, dwAccumulator + 2)
		End If
	End If
	GetKeyWin = szProductKey
End Function

' Function calculating serial key for Microsoft Office products
' Current version computes correct keys for MS Office 2007 and older
Function GetKeyOffice(rpk)
	Const rpkOffset = 52
	szPossibleChars = "BCDFGHJKMPQRTVWXY2346789"
	i = 28
	Do
		dwAccumulator = 0
		j = 14
		Do  
			dwAccumulator = dwAccumulator * 256  
			dwAccumulator = rpk(j + rpkOffset) + dwAccumulator
			rpk(j + rpkOffset) = (dwAccumulator \ 24) And 255  
			dwAccumulator = dwAccumulator Mod 24
			j = j - 1
		Loop While j >= 0
		i = i - 1
		szProductKey = Mid(szPossibleChars, dwAccumulator + 1, 1) & szProductKey
		If (((29 - i) Mod 6) = 0) And (i <> -1) Then
			i = i - 1
			szProductKey = "-" & szProductKey
		End If
	Loop While i >= 0
	GetKeyOffice = szProductKey
End Function

Function WMIDateStringToDate(dtmInstallDate) 
	WMIDateStringToDate = CDate(Mid(dtmInstallDate, 5, 2) & "/" & Mid(dtmInstallDate, 7, 2) & _
	"/" & Left(dtmInstallDate, 4) & " " & Mid (dtmInstallDate, 9, 2) & ":" & _ 
	Mid(dtmInstallDate, 11, 2) & ":" & Mid(dtmInstallDate, 13, 2))
End Function
