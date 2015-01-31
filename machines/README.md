Scripts designed to be run on domain workstations during startup process
==============

## List of scripts

* **ad_write_comp_info.vbs**. Script for recording information about MAC and IP addresses of current PC in appropriate Active Directory computer object. Also script reads description from AD computer object and updates it for current PC.

* **disable_autoupdate_2gis.vbs**. This script is assigned for disabling annoying notifications of 2GIS AutoUpdater. It disables 2GIS Update Notifier and deletes 2GIS Update service.

* **disable_autoupdate_flash.vbs**. Script for disabling Adobe Flash Updater. In Active Directory domain the better choice is to deploy new versions packed to MSI packages or to use WSUS.

* **get_info_hardware.vbs**. Script for generating text report about installed hardware and Microsoft products.

* **get_info_installed_soft.vbs**. Script for generating text report about installed software.

* **reg_disable_java_updates.vbs**. Script for disabling Java Runtime Environment (JRE) automatic updates.

* **set_sheduled_job_kb.vbs**. Script creates scheduled job for unattended setup of hotfix from shared folder. The script is suitable for Active Directory environments without WSUS and also for situations when it is needed to install some application delivered only with EXE setup on a large quantity of machines.

* **shutdown_domain_machine.vbs**. Automatic and silent computer shutdown (without warning messages). Script is intended to run as a scheduled job from the shared folder. This job can be created with Group Policy Settings or some other scripts. Also script uses a file with list of exceptions. If the file contains name of current computer this computer will not be powered off. Another way to do the same is to run similar script from a server. I've chosen this way ;)
