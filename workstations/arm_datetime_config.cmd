@echo off

::Automatic adjustment of time synchronization service for remote workstations with Windows XP
::Script sequentially performs next actions:
::1. Stops time synchronization service
::2. Re-registers time synchronization service (for all options release)
::3. Makes changes in registry for tuning time synchronization service:
::	1) Tough time adjustment if local time differs from server time more than 90 seconds
::	2) Maximum 'back' time correction is 60 seconds
::	3) Maximum 'forward' time correction is 2 hours
::4. Exports settings to registry from file timezones_2014_rus.reg to set correct time zones
::5. Setups current time zone in North Asia Standard Time (GMT +7)
::6. Setups NTP server for time synchronization
::7. Starts Windows time synchronization service
::8. Runs forced time synchronization

net stop w32time
w32tm /unregister
w32tm /unregister
w32tm /register
::Changing name of service. Default name is "Windows Time Service"
::The service name must be written in OEM866 codepage!
sc config w32time displayname= "Служба времени Windows"

reg add "HKLM\SYSTEM\CurrentControlSet\Services\W32Time\Config" /v "MaxNegPhaseCorrection" /t REG_DWORD /d "60" /f
reg add "HKLM\SYSTEM\CurrentControlSet\Services\W32Time\Config" /v "MaxPosPhaseCorrection" /t REG_DWORD /d "6400" /f
reg add "HKLM\SYSTEM\CurrentControlSet\Services\W32Time\Config" /v "MaxAllowedPhaseOffset" /t REG_DWORD /d "90" /f
reg delete "HKLM\SYSTEM\CurrentControlSet\Control\TimeZoneInformation" /v "DisableAutoDaylightTimeSet" /f

regedit.exe /s "timezones_2014_rus.reg"

control timedate.cpl,,/z North Asia Standard Time
::Next comment line does the same work as above
::RunDLL32 shell32.dll,Control_RunDLL timedate.cpl,,/z North Asia Standard Time

net time /setsntp:192.168.10.1
net start w32time
::Or we can set NTP server another way, but only after service start
::w32tm /config /update /syncfromflags:manual /manualpeerlist:192.168.10.1
w32tm /resync
