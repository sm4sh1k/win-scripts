@echo off

echo Disabling unnecessary Windows XP services for remote workstation

sc config FastUserSwitchingCompatibility start= disabled
sc config ShellHWDetection start= disabled
sc config SharedAccess start= disabled
sc config LanmanServer start= disabled
sc config WebClient start= disabled
sc config wuauserv start= disabled
sc config seclogon start= disabled
sc config SSDPSRV start= disabled
sc config helpsvc start= disabled
sc config WZCSVC start= disabled
sc config wscsvc start= disabled
::sc config BITS start= disabled

net stop FastUserSwitchingCompatibility
net stop ShellHWDetection
net stop SharedAccess
net stop LanmanServer
net stop WebClient
net stop wuauserv
net stop seclogon
net stop SSDPSRV
net stop helpsvc
net stop WZCSVC
net stop wscsvc
::net stop BITS
