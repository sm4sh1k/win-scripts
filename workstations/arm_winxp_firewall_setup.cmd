@echo off

echo Setting up firewall for remote workstation (with Windows XP)
echo Allowing only inbound ICMP requests and VNC (RAdmin) connections
echo All outbound connections are allowed due to Windows XP firewall constraints
echo If you want to add some additional parameters use help or read commented lines in the end of this script

sc config SharedAccess start= auto
net start SharedAccess

netsh firewall reset

netsh firewall set icmpsetting type=ALL enable

netsh firewall add portopening TCP 4899 "RAdmin Connection"
netsh firewall add portopening TCP 5800 "VNC Connection (Java)"
netsh firewall add portopening TCP 5900 "VNC Connection"

::netsh firewall set service FileAndPrint
::netsh firewall set service RemoteDesktop enable
::netsh firewall add allowedprogram C:\MyApp.exe "My App" enable
::netsh firewall set logging %WINDIR%\pfirewall.log 4096 disable enable
