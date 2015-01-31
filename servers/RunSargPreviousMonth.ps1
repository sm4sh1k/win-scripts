# Generating and sending web traffic report for squid proxy-server (using SARG and 7-zip)
# Author: Valentin Vakhrushev, 2014

# Parameters
$date = Get-Date
$numMonthsAgo = -1
$programPath = "C:\sarg\bin\sarg.exe"
$reportsPath = "C:\wwwdocs\squid-reports\"
$sevenzipPath = "C:\Progra~1\7-Zip\7z.exe"
$squidPath = "C:\squid\sbin\squid.exe"

# Report generation (using SARG)
$firstDayOfMonth = Get-Date $date -day 1 -hour 0 -minute 0 -second 0
$lastDayOfMonth = (($firstDayOfMonth).AddMonths(1).AddSeconds(-1))

# Current month minus one equals previous ;)
$firstDayOfMonth = $firstDayOfMonth.AddMonths($numMonthsAgo)
$lastDayOfMonth = $lastDayOfMonth.AddMonths($numMonthsAgo)
 
$params = " -d " + (Get-Date $firstDayOfMonth -uformat "%d/%m/%Y") + "-" + (Get-Date $lastDayOfMonth -uformat "%d/%m/%Y")

Set-Location (Split-Path $programPath)
iex ($programPath + $params)

# Archiving the folder with report (using 7-zip)
$monthName = ''
switch ($firstDayOfMonth.Month) {
	'1' 	{ $monthName = "Jan" }
	'2'		{ $monthName = "Feb" }
	'3'		{ $monthName = "Mar" }
	'4'		{ $monthName = "Apr" }
	'5'		{ $monthName = "May" }
	'6'		{ $monthName = "Jun" }
	'7'		{ $monthName = "Jul" }
	'8'		{ $monthName = "Aug" }
	'9'		{ $monthName = "Sep" }
	'10'	{ $monthName = "Oct" }
	'11'	{ $monthName = "Nov" }
	'12'	{ $monthName = "Dec" }
}

$folderName = $firstDayOfMonth.Year.ToString() + $monthName + $firstDayOfMonth.Day.ToString("00") + '-' + $lastDayOfMonth.Year.ToString() + $monthName + $lastDayOfMonth.Day.ToString("00")
$folderPathName = $reportsPath + $folderName
$params = " a -t7z -mx7 " + $folderName + ".7z " + $folderPathName

Set-Location $reportsPath
If (Test-Path $folderName){
	iex ($sevenzipPath + $params)
}

# Sending archive with report to administrators via e-mail
$server = "smtp.yandex.ru"
$username = "remoteserver12"
$password = "P@$$W0rd"
$to = "Support <support@yourcompany.com>"
$cc = "Administrators <postroot@yourcompany.com>"
$from = "Remote Server <remoteserver12@yandex.ru>"
$subject = "Web traffic report"
$body = Get-Date $firstDayOfMonth -format Y
$attachment = $folderName + ".7z"

$secpasswd = ConvertTo-SecureString $password -AsPlainText -Force
$creds = New-Object System.Management.Automation.PSCredential ($username, $secpasswd)

If (Test-Path ($folderName + ".7z")){
	$body = "Web traffic report is generated for " + $body + '.'
	send-mailmessage -to $to -cc $cc -from $from -Encoding ([System.Text.Encoding]::UTF8) -subject $subject -body $body -smtpserver $server -credential $creds -attachments $attachment
	Remove-Item ($folderName + ".7z")
}
Else {
	$body = "Error! Web traffic report for " + $body + " is not generated!"
	send-mailmessage -to $to -cc $cc -from $from -Encoding ([System.Text.Encoding]::UTF8) -subject $subject -body $body -smtpserver $server -credential $creds
}

write-host $body

# Squid logs rotation
$params = " -n Squid -k rotate"
Set-Location (Split-Path $squidPath)
iex ($squidPath + $params)
