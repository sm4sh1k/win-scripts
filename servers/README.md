Scripts designed to be run on servers as scheduled tasks
==============

## List of scripts

* **ad_change_owa_access_rights.vbs**. Script changes Active Directory user properties to allow or deny Outlook Web Access and Outlook Mobile Access depending on user group. If the user is in the 'OWA Users' group then access is granted, otherwise access is denied. Script makes two things happen: it allow administrator to control OWA/OMA access just putting users in appropriate group and guarantee against configuration mistakes. It is proposed to run this script periodicaly on the server putting it to scheduled tasks.

* **ad_create_private_folders.vbs**. Script creates folders for Active Directory users in the shared folder and grants appropriate rights for them. It is also proposed to run this script automaticaly on the file server. The main goal of this script is to automate procedure of creating private folders for new users. Another script is used for mounting net drive for user after logging in. 

* **RunSargPreviousMonth.ps1**. This is my first PowerShell script :). It sends monthly report of web traffic usage for Squid proxy server with the help of SARG and 7-zip. This script is executed every 1st day of month. It run SARG to generate a report and then 7-zip to archive generated folder. Then it send this archive via e-mail to administrators and execute Squid rotate process. So at the end we have very cheap proxy server with periodically reports. You can also install another free software like *rejik* to restrict access to some undesirable sites. In my opinion for the really remote servers it is more than enough.

* **fb_backup_level0.cmd**, **fb_backup_level1.cmd** and **fb_backup_level2.cmd**. A set of scripts for perfoming incremental everyday backups, weekly backups and full monthly backups of Firebird database using *nbackup* utility. Database snapshots are archived with 7-zip. Old archives are removed automaticaly.

* **mount_bitlocker_volume.vbs**. Script for automatic mounting BitLocker encrypted volume when USB flash disk with key file is inserted. The algorithm is quite simple: we start the file server with encrypted volume, then insert flash stick with key file. Script starts automaticaly via scheduled tasks and determines the flash drive inserted or not. If the key file is found encrypted volume is unlocked. Then operator has to eject the flash drive and put it safely away. Otherwise volume is locked again.

* **reload_apcupsd_driver.vbs**. Script for automatic restart the driver and Apcupsd service if the connection between APC UPS and the computer has been lost. Driver reloading is implemented using DevCon program provided by Microsoft. Script is designed to run via scheduler.

* **Update2GisDatabase.ps1**. Script for automatic update 2GIS database. It parse HTML page to retrieve download link, get ZIP archive with database file from the Web, extract it and copy to specified folder. Also it send report to administrators via email. It is used with 2GIS deployed in the shared folder on a file server. Script is designed to run via scheduler once per month.
