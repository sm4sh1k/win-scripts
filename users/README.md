Scripts designed to run under user account during logon process
==============

## List of scripts

* **check_admin_rights_users.vbs**. Script creates text report in the shared folder if the current user has administrative privilegies on local machine.

* **copy_settings_lo.vbs**. Script for creating default user profile for LibreOffice. Basicaly it just extracts archive from shared folder to user's AppData folder. But also it makes a lot of associated job such as determining the OS version (32 or 64 bit) or choosing the appropriate file server depending on current Active Directory site or domain controller.

* **create_1c_profile_ru.vbs**, **create_1c_profile_branch_ru.vbs** and **setup_1c_profile_ru.vbs**. These scripts create or adjust user profile for 1C:Enterprise platform. This software is using only in Russian-speaking countries, so all comments in these files are in russian.

* **create_shortcut_2gis.vbs**. Script for creating shortcut to 2GIS application on user's desktop. If application is not installed on local computer then script creates shortcut for application located in shared folder.

* **create_shortcut_lo.vbs**. Script for creating shortcut in Autostart folder to automatically run LibreOffice Quickstart on user's logon.

* **edit_firefox_profile.vbs**. Script adjusts Mozilla Firefox profile for proper work in domain based environment. Basicaly it modifies *prefs.js* file in user's profile for correct opening local or network web pages.

* **run_logon_scripts.vbs**. Script for sequentially and silently launching another scripts from the specified network folder. It is a some kind of analogue of init.d system in Linux or Group Policy scripts in Windows.