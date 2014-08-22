Useful scripts for Windows system administrators
==============

My collection of scripts for Windows operating system.
Most of them were used to automate administering of Windows-based infrastructure with Active Directory domain.
All these scripts were written by me sometimes using adopted code from various free scripts founded on the Web.

## List of scripts

* **ad_change_owa_access_rights.vbs**. Script changes Active Directory user properties to allow or deny Outlook Web Access and Ooutlook Mobile Access depending on user group. If the user is in the 'OWA Users' group then access is granted, otherwise access is denied. Script makes two things happen: it allow administrator to control OWA/OMA access just putting users in appropriate group and guarantee against configuration mistakes. It is proposed to run this script periodicaly on the server putting it to scheduled tasks.

* **ad_create_private_folders.vbs**. Script creates folders for Active Directory users in the shared folder and grants appropriate rights for them. It is also proposed to run this script automaticaly on the file server. The main goal of this script is to automate procedure of creating private folders for new users. Another script is used for mounting net drive for user after logging in. 

* **admin_reg_disable_autorun.vbs**. This script just disables autorun for all drives. Also script tries to get administrative privileges if needed. It is very useful when running it remotely from the user's desktop.

* **admin_reg_set_ext_lo.vbs**. Registering file associations for LibreOffice. Script tries to determine installation path of LibreOffice and get administrative privileges if needed. All messages and association descriptions are in russian. You can translate them using Google Translate or something...
