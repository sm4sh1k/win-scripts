# Automatic Update for 2GIS Database (parse the download page, download and extract ZIP archive)
# Author: Valentin Vakhrushev, 2017

# Function for sending report via email
# Edit/remove fields to meet your environment needs
function SendMailReport($body) {
	$server = "mail.domain.local"
	$username = "mailuser"
	$password = "mailpass"
	$to = "Support Team <support@domain.com>"
	$from = "2gispudate@server.domain.local"
	$subject = "2GIS Database Update Report"
	
	$secpasswd = ConvertTo-SecureString $password -AsPlainText -Force
	$creds = New-Object System.Management.Automation.PSCredential ($username, $secpasswd)
	
	Send-MailMessage -to $to -from $from -Encoding ([System.Text.Encoding]::UTF8) -subject $subject -body $body -smtpserver $server -credential $creds -usessl
}


# Link to download page with required database file in ZIP archive (version for Linux)
$url = "http://info.2gis.ru/moscow/products/download#skachat-kartu-na-komputer&linux"
# Regular expression with required filename
$regexp = "^.*2GISData_Moscow-[0-9]{2,3}\.orig\.zip$"
# Path to folder with 2GIS shell executable and database file
$outfolder = "C:\Shared\2gis\"
# File mask to be extracted from ZIP archive
$datafile = "*.dgdat"
# 'Everything is OK' message
$message = "2GIS database has been updated to the latest version."

$links = (Invoke-WebRequest -Uri $url -UseBasicParsing).Links.Href

foreach ($link in $links) {
	if ($link -match $regexp) {
		$file = "$PSScriptRoot\" + $link.Substring($link.LastIndexOf("/") + 1)
		$tempdir = "$PSScriptRoot\2GISTemp"
		
		(New-Object System.Net.WebClient).DownloadFile($link, $file)
		
		if (Test-Path -path $tempdir) { Remove-Item -Path $tempdir -Recurse -Force }
		Add-Type -AssemblyName System.IO.Compression.FileSystem
		[System.IO.Compression.ZipFile]::ExtractToDirectory($file, $tempdir)
		# Expand-Archive function works only in PowerShell 5+
		#Expand-Archive -Path $file -DestinationPath $tempdir -Force
		
		if (!(Test-Path -path $outfolder)) { New-Item $outfolder -Type Directory -Force | Out-Null }
		Get-ChildItem -Path $tempdir -Filter $datafile -Recurse -ErrorAction SilentlyContinue -Force | ForEach-Object {
			try {
				Copy-Item -Path $_.FullName -Destination $outfolder -Force
			}
			catch {
				$message = "During update process the error has been occurred:`r`n" + $_.Exception.Message
			}
		}
		
		Write-Host $message
		SendMailReport -body $message
		
		Remove-Item -Path $tempdir -Recurse -Force
		Remove-Item -Path $file -Force
		break
	}
}
