Useful scripts for Windows system administrators
==============

My collection of scripts for Windows operating system.
Most of them were used to automate administering of Windows-based infrastructure with Active Directory domain.
All these scripts were written by me sometimes using adopted code from various free scripts founded on the Web.

## List of scripts

* **ad_change_owa_access_rights.vbs**. Script changes Active Directory user properties to allow or deny Outlook Web Access and Ooutlook Mobile Access depending on user group. If the user is in the 'OWA Users' group then access is granted, otherwise access is denied. Script makes two things happen: it allow administrator to control OWA/OMA access just putting users in appropriate group and guarantee against configuration mistakes. It is proposed to run this script periodicaly on the server putting it to scheduled tasks.

* **ad_create_private_folders.vbs**. Script creates folders for Active Directory users in the shared folder and grants appropriate rights for them. It is also proposed to run this script automaticaly on the file server. The main goal of this script is to automate procedure of creating private folders for new users. Another script is used for mounting net drive for user after logging in. 

* **ad_list_all_users_email.vbs**. Script creates HTML file on desktop with the list of e-mail addresses of Active Directory users.

* **ad_mail_forward.hta**. A HTA application for e-mail redirection adjustment. It can read and set e-mail forwarding options for users in Active Directory domain. With this script you can operate with e-mail redirection options much faster then in usual way.

* **ad_usb_access.hta**. Another HTA application. This script helps to adjust access to USB drives on computers in Active Directory domain. Basicaly it just changes parameter in registry on remote computer, but with clean and intuitive interface.

* **ad_write_comp_info.vbs**. Script for recording information about MAC and IP addresses of current PC in appropriate Active Directory computer object. Also script reads description from AD computer object and updates it for current PC. 
* **admin_reg_disable_autorun.vbs**. This script just disables autorun for all drives. Also script tries to get administrative privileges if needed. It is very useful when running it remotely from the user's desktop.

* **admin_reg_set_ext_lo.vbs**. Registering file associations for LibreOffice. Script tries to determine installation path of LibreOffice and get administrative privileges if needed. All messages and association descriptions are in russian. You can translate them using Google Translate or something...

