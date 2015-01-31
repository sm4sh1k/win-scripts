Scripts designed to be run on servers as scheduled tasks
==============

## List of scripts

* **ad_change_owa_access_rights.vbs**. Script changes Active Directory user properties to allow or deny Outlook Web Access and Ooutlook Mobile Access depending on user group. If the user is in the 'OWA Users' group then access is granted, otherwise access is denied. Script makes two things happen: it allow administrator to control OWA/OMA access just putting users in appropriate group and guarantee against configuration mistakes. It is proposed to run this script periodicaly on the server putting it to scheduled tasks.

* **ad_create_private_folders.vbs**. Script creates folders for Active Directory users in the shared folder and grants appropriate rights for them. It is also proposed to run this script automaticaly on the file server. The main goal of this script is to automate procedure of creating private folders for new users. Another script is used for mounting net drive for user after logging in. 
