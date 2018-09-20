# Generating report about computer hardware in HTML file
# Script is designed to run as a scheduled job deployed with Group Policy Settings
# Author: Valentin Vakhrushev, 2018
# 
# This script is based on Collect-ServerInfo.ps1 script from Technet Gallery
# writed by Paul Cunningham and distributed under MIT License:
# https://gallery.technet.microsoft.com/scriptcenter/PowerShell-Collect-Server-089f1da3

[CmdletBinding()]
Param()

function Get-ProductKey {
	$map="BCDFGHJKMPQRTVWXY2346789"
	Try {
		$value = (Get-ItemProperty -path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').DigitalProductId[0x34..0x42]
		$isWin8OrNewer = [math]::Floor(($value[14] / 6)) -band 1
		$value[14] = ($value[14] -band 0xF7) -bor (($isWin8OrNewer -band 2) * 4)
		$ProductKey = ""
		for ($i = 24; $i -ge 0; $i--) {
			$r = 0
			for ($j = 14; $j -ge 0; $j--) {
				$r = ($r * 256) -bxor $value[$j]
				$value[$j] = [math]::Floor([double]($r / 24))
				$r = $r % 24
			}
			$ProductKey = $map[$r] + $ProductKey
		}
	}
	Catch {
		$ProductKey = $_.Exception.Message
	}
	if ($isWin8OrNewer) {
		$ProductKey = $ProductKey.Remove(0, 1)
		$ProductKey = $ProductKey.Insert($r, 'N')
	}
	for($i = 5; $i -lt 29; $i = $i + 6) {
		$ProductKey = $ProductKey.Insert($i, '-')
	}
	$ProductKey
}


#Initialize
Write-Verbose "Initializing"

# Type here path to the folder where reports will be stored
$ServerFolder = "\\SRV01\Info$\"
$ComputerName = $env:computername
$Description = (Get-WmiObject Win32_OperatingSystem -ErrorAction STOP).Description
$isVistaOrNewer = @([int]((Get-WmiObject Win32_OperatingSystem).Version -split '\.')[0] -ge 6)

# Process ComputerName
Write-Verbose "=====> Processing $ComputerName <====="
$htmlreport = @()
$htmlbody = @()
#$htmlfile = $ServerFolder + "$($ComputerName)_" + (Get-Date -uformat "%Y-%m-%d") + ".html"
$htmlfile = $ServerFolder + "$($ComputerName).html"
$spacer = "<br />"

# Collect computer system information and convert to HTML fragment
Write-Verbose "Collecting computer system information"
$subhead = "<h3>Computer System Information</h3>"
$htmlbody += $subhead
try {
	$csinfo = Get-WmiObject Win32_ComputerSystem -ErrorAction STOP |
		Select-Object Name,Manufacturer,Model,
					@{Name='Physical Processors';Expression={$_.NumberOfProcessors}},
					@{Name='Logical Processors';Expression={$_.NumberOfLogicalProcessors}},
					@{Name='Total Physical Memory (Gb)';Expression={
						$tpm = $_.TotalPhysicalMemory/1GB;
						"{0:F0}" -f $tpm
					}},
					DnsHostName,Domain
	$htmlbody += $csinfo | ConvertTo-Html -Fragment
	$htmlbody += $spacer
}
catch {
	Write-Warning $_.Exception.Message
	$htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
	$htmlbody += $spacer
}

# Collect operating system information and convert to HTML fragment
Write-Verbose "Collecting operating system information"
$subhead = "<h3>Operating System Information</h3>"
$htmlbody += $subhead
try {
	$osinfo = Get-WmiObject Win32_OperatingSystem -ErrorAction STOP | 
		Select-Object @{Name='Operating System';Expression={$_.Caption}},
					@{Name='Architecture';Expression={$_.OSArchitecture}},
					Version,Organization,
					@{Name='Install Date';Expression={
						$installdate = [datetime]::ParseExact($_.InstallDate.SubString(0,8),"yyyyMMdd",$null);
						$installdate.ToShortDateString()
					}},
					WindowsDirectory,
					@{Name='Product Key';Expression={Get-ProductKey}}
	$htmlbody += $osinfo | ConvertTo-Html -Fragment
	$htmlbody += $spacer
}
catch {
	Write-Warning $_.Exception.Message
	$htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
	$htmlbody += $spacer
}

# Collect pagefile information and convert to HTML fragment
$subhead = "<h3>PageFile Information</h3>"
$htmlbody += $subhead
Write-Verbose "Collecting pagefile information"
try {
	$pagefileinfo = Get-WmiObject Win32_PageFileUsage -ErrorAction STOP |
		Select-Object @{Name='Pagefile Name';Expression={$_.Name}},
					@{Name='Allocated Size (Mb)';Expression={$_.AllocatedBaseSize}}
	$htmlbody += $pagefileinfo | ConvertTo-Html -Fragment
	$htmlbody += $spacer
}
catch {
	Write-Warning $_.Exception.Message
	$htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
	$htmlbody += $spacer
}

# Collect BIOS information and convert to HTML fragment
$subhead = "<h3>BIOS Information</h3>"
$htmlbody += $subhead
Write-Verbose "Collecting BIOS information"
try {
	$biosinfo = Get-WmiObject Win32_Bios -ErrorAction STOP |
		Select-Object Status,Version,Manufacturer,
					@{Name='Release Date';Expression={
						$releasedate = [datetime]::ParseExact($_.ReleaseDate.SubString(0,8),"yyyyMMdd",$null);
						$releasedate.ToShortDateString()
					}},
					@{Name='Serial Number';Expression={$_.SerialNumber}}
	$htmlbody += $biosinfo | ConvertTo-Html -Fragment
	$htmlbody += $spacer
}
catch {
	Write-Warning $_.Exception.Message
	$htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
	$htmlbody += $spacer
}

# Collect processor information and convert to HTML fragment
Write-Verbose "Collecting processor information"
$subhead = "<h3>Processor Information</h3>"
$htmlbody += $subhead
try {
	$processors = @()
	$processorinfo = @(Get-WmiObject Win32_Processor -ErrorAction STOP |
		Select-Object Name,MaxClockSpeed,SocketDesignation)
	foreach ($processor in $processorinfo) {
		$memObject = New-Object PSObject
		$memObject | Add-Member NoteProperty -Name "Name" -Value $processor.Name
		$memObject | Add-Member NoteProperty -Name "Max Speed" -Value $processor.MaxClockSpeed
		$memObject | Add-Member NoteProperty -Name "Socket" -Value $processor.SocketDesignation
		$processors += $memObject
	}
	$htmlbody += $processors | ConvertTo-Html -Fragment
	$htmlbody += $spacer
}
catch {
	Write-Warning $_.Exception.Message
	$htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
	$htmlbody += $spacer
}

# Collect motherboard information and convert to HTML fragment
Write-Verbose "Collecting motherboard information"
$subhead = "<h3>Motherboard Information</h3>"
$htmlbody += $subhead
try {
	$motherboards = @()
	$motherboardinfo = @(Get-WmiObject Win32_BaseBoard -ErrorAction STOP |
		Select-Object Manufacturer,Product)
	foreach ($motherboard in $motherboardinfo) {
		$memObject = New-Object PSObject
		$memObject | Add-Member NoteProperty -Name "Manufacturer" -Value $motherboard.Manufacturer
		$memObject | Add-Member NoteProperty -Name "Model" -Value $motherboard.Product
		$motherboards += $memObject
	}
	$htmlbody += $motherboards | ConvertTo-Html -Fragment
	$htmlbody += $spacer
}
catch {
	Write-Warning $_.Exception.Message
	$htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
	$htmlbody += $spacer
}

# Collect videoadapter information and convert to HTML fragment
Write-Verbose "Collecting videoadapter information"
$subhead = "<h3>Video Adapter Information</h3>"
$htmlbody += $subhead
try {
	$videoadapters = @()
	$videoadapterinfo = @(Get-WmiObject Win32_VideoController -ErrorAction STOP |
		Select-Object Caption,AdapterRAM)
	foreach ($videoadapter in $videoadapterinfo) {
		$memObject = New-Object PSObject
		$memObject | Add-Member NoteProperty -Name "Name" -Value $videoadapter.Caption
		$memObject | Add-Member NoteProperty -Name "Capacity (GB)" -Value ("{0:F0}" -f $videoadapter.AdapterRAM/1GB)
		$videoadapters += $memObject
	}
	$htmlbody += $videoadapters | ConvertTo-Html -Fragment
	$htmlbody += $spacer
}
catch {
	Write-Warning $_.Exception.Message
	$htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
	$htmlbody += $spacer
}

# Collect physical memory information and convert to HTML fragment
Write-Verbose "Collecting physical memory information"
$subhead = "<h3>Physical Memory Information</h3>"
$htmlbody += $subhead
try {
	$memorybanks = @()
	$physicalmemoryinfo = @(Get-WmiObject Win32_PhysicalMemory -ErrorAction STOP |
		Select-Object DeviceLocator,Manufacturer,Speed,Capacity)
	foreach ($bank in $physicalmemoryinfo) {
		$memObject = New-Object PSObject
		$memObject | Add-Member NoteProperty -Name "Device Locator" -Value $bank.DeviceLocator
		$memObject | Add-Member NoteProperty -Name "Manufacturer" -Value $bank.Manufacturer
		$memObject | Add-Member NoteProperty -Name "Speed" -Value $bank.Speed
		$memObject | Add-Member NoteProperty -Name "Capacity (GB)" -Value ("{0:F0}" -f $bank.Capacity/1GB)
		$memorybanks += $memObject
	}
	$htmlbody += $memorybanks | ConvertTo-Html -Fragment
	$htmlbody += $spacer
}
catch {
	Write-Warning $_.Exception.Message
	$htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
	$htmlbody += $spacer
}

# Collect network interface information and convert to HTML fragment
$subhead = "<h3>Network Interface Information</h3>"
$htmlbody += $subhead
Write-Verbose "Collecting network interface information"
try {
	$nics = @()
	# Windows XP does not provide some Win32_NetworkAdapter's properties
	if ($isVistaOrNewer) {
		$nicinfo = @(Get-WmiObject Win32_NetworkAdapter -ErrorAction STOP | Where {$_.PhysicalAdapter} |
			Select-Object Name,AdapterType,MACAddress,
			@{Name='ConnectionName';Expression={$_.NetConnectionID}},
			@{Name='Enabled';Expression={$_.NetEnabled}},
			@{Name='Speed';Expression={$_.Speed/1000000}})
		$nwinfo = @(Get-WmiObject Win32_NetworkAdapterConfiguration -ErrorAction STOP |
			Select-Object Description,DHCPServer,
			@{Name='IpAddress';Expression={$_.IpAddress -join '; '}},
			@{Name='IpSubnet';Expression={$_.IpSubnet -join '; '}},
			@{Name='DefaultIPgateway';Expression={$_.DefaultIPgateway -join '; '}},
			@{Name='DNSServerSearchOrder';Expression={$_.DNSServerSearchOrder -join '; '}})
		foreach ($nic in $nicinfo) {
			$nicObject = New-Object PSObject
			$nicObject | Add-Member NoteProperty -Name "Connection Name" -Value $nic.connectionname
			$nicObject | Add-Member NoteProperty -Name "Adapter Name" -Value $nic.Name
			$nicObject | Add-Member NoteProperty -Name "Type" -Value $nic.AdapterType
			$nicObject | Add-Member NoteProperty -Name "MAC" -Value $nic.MACAddress
			$nicObject | Add-Member NoteProperty -Name "Enabled" -Value $nic.Enabled
			$nicObject | Add-Member NoteProperty -Name "Speed (Mbps)" -Value $nic.Speed
			$ipaddress = ($nwinfo | Where {$_.Description -eq $nic.Name}).IpAddress
			$nicObject | Add-Member NoteProperty -Name "IP Address" -Value $ipaddress
			$nics += $nicObject
		}
	}
	else {
		$nicinfo = @(Get-WmiObject Win32_NetworkAdapter -ErrorAction STOP |
			Where {$_.Manufacturer -ne 'Microsoft' -And $_.PNPDeviceID -notlike 'ROOT\\%' -And $_.MACAddress} |
				Select-Object DeviceID,Name,AdapterType,MACAddress,
				@{Name='ConnectionName';Expression={$_.NetConnectionID}})
		$nwinfo = @(Get-WmiObject Win32_NetworkAdapterConfiguration -ErrorAction STOP |
			Select-Object Index,Description,
			@{Name='IpAddress';Expression={$_.IpAddress -join '; '}})
		foreach ($nic in $nicinfo) {
			$nicObject = New-Object PSObject
			$nicObject | Add-Member NoteProperty -Name "Connection Name" -Value $nic.connectionname
			$nicObject | Add-Member NoteProperty -Name "Adapter Name" -Value $nic.Name
			$nicObject | Add-Member NoteProperty -Name "Type" -Value $nic.AdapterType
			$nicObject | Add-Member NoteProperty -Name "MAC" -Value $nic.MACAddress
			$ipaddress = ($nwinfo | Where {$_.Index -eq $nic.DeviceID}).IpAddress
			$nicObject | Add-Member NoteProperty -Name "IP Address" -Value $ipaddress
			$nics += $nicObject
		}
	}
	$htmlbody += $nics | ConvertTo-Html -Fragment
	$htmlbody += $spacer
}
catch {
	Write-Warning $_.Exception.Message
	$htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
	$htmlbody += $spacer
}

# Collect logical disk information and convert to HTML fragment
$subhead = "<h3>Logical Disk Information</h3>"
$htmlbody += $subhead
Write-Verbose "Collecting logical disk information"
try {
	$diskinfo = Get-WmiObject Win32_LogicalDisk -ErrorAction STOP | 
		Select-Object DeviceID,FileSystem,VolumeName,
		@{Expression={$_.Size /1Gb -as [int]};Label="Total Size (GB)"},
		@{Expression={$_.Freespace / 1Gb -as [int]};Label="Free Space (GB)"}
	$htmlbody += $diskinfo | ConvertTo-Html -Fragment
	$htmlbody += $spacer
}
catch {
	Write-Warning $_.Exception.Message
	$htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
	$htmlbody += $spacer
}

# Collect volume information and convert to HTML fragment
$subhead = "<h3>Volume Information</h3>"
$htmlbody += $subhead
Write-Verbose "Collecting volume information"
try {
	# Windows XP does not provide Win32_Volume WMI object
	if ($isVistaOrNewer) {
		$volinfo = Get-WmiObject Win32_Volume -ErrorAction STOP | 
			Select-Object Label,Name,DeviceID,SystemVolume,
			@{Expression={$_.Capacity /1Gb -as [int]};Label="Total Size (GB)"},
			@{Expression={$_.Freespace / 1Gb -as [int]};Label="Free Space (GB)"}
		$htmlbody += $volinfo | ConvertTo-Html -Fragment
	}
	else {
		$volumes = @()
		$drives = Get-WmiObject Win32_DiskDrive -ErrorAction STOP | 
			Select-Object Caption, DeviceID
		foreach ($drive in $drives) {
			$query = "ASSOCIATORS OF {Win32_DiskDrive.DeviceID=""" +`
				($drive.DeviceID -replace "\\", "\\") +`
				"""} WHERE AssocClass = Win32_DiskDriveToDiskPartition"
			$partitions = Get-WmiObject -Query $query -ErrorAction STOP
			foreach ($partition in $partitions) {
				$query = "ASSOCIATORS OF {Win32_DiskPartition.DeviceID=""" + $partition.DeviceID +`
					"""} WHERE AssocClass = Win32_LogicalDiskToPartition"
				$logicaldisks = Get-WmiObject -Query $query -ErrorAction STOP
				foreach ($logicaldisk in $logicaldisks) {
					$volObject = New-Object PSObject
					$volObject | Add-Member NoteProperty -Name "Label" -Value $logicaldisk.VolumeName
					$volObject | Add-Member NoteProperty -Name "Name" -Value $logicaldisk.Name
					$volObject | Add-Member NoteProperty -Name "Device ID" -Value $drive.DeviceID
					$volObject | Add-Member NoteProperty -Name "Drive Type" -Value $logicaldisk.DriveType
					$volObject | Add-Member NoteProperty -Name "Total Size (GB)" `
						-Value ("{0:F0}" -f $logicaldisk.Size/1GB -as [int])
					$volObject | Add-Member NoteProperty -Name "Free Space (GB)" `
						-Value ("{0:F0}" -f $logicaldisk.FreeSpace/1GB -as [int])
					$volumes += $volObject
				}
			}
		}
		$htmlbody += $volumes | ConvertTo-Html -Fragment
	}
	$htmlbody += $spacer
}
catch {
	Write-Warning $_.Exception.Message
	$htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
	$htmlbody += $spacer
}

# Collect software information and convert to HTML fragment
$subhead = "<h3>Software Information</h3>"
$htmlbody += $subhead
Write-Verbose "Collecting software information"
try {
	$software = Get-WmiObject Win32_Product -ErrorAction STOP | Select-Object Vendor,Name,Version | Sort-Object Vendor,Name
	$htmlbody += $software | ConvertTo-Html -Fragment
	$htmlbody += $spacer 

}
catch {
	Write-Warning $_.Exception.Message
	$htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
	$htmlbody += $spacer
}

# Generate the HTML report and output to file
Write-Verbose "Producing HTML report"
$reportime = Get-Date

#Common HTML head and styles
$htmlhead="<html>
			<style>
			BODY{font-family: Arial; font-size: 8pt;}
			H1{font-size: 20px;}
			H2{font-size: 18px;}
			H3{font-size: 16px;}
			TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
			TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;}
			TD{border: 1px solid black; padding: 5px; }
			td.pass{background: #7FFF00;}
			td.warn{background: #FFE600;}
			td.fail{background: #FF0000; color: #ffffff;}
			td.info{background: #85D4FF;}
			</style>
			<body>
			<h1 align=""center"">Computer Info: $ComputerName</h1>
			<h2 align=""center"">Description: $Description</h2>
			<h3 align=""center"">Generated: $reportime</h3>"
$htmltail = "</body>
		</html>"
$htmlreport = $htmlhead + $htmlbody + $htmltail
$htmlreport | Out-File $htmlfile -Encoding Utf8

#Wrap it up
Write-Verbose "=====> Finished <====="
