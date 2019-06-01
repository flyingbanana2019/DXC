if exist "c:\program files\common files\opsware\agent\wsusscn2.cab" del "c:\program files\common files\opsware\agent\wsusscn2.cab"
Reg query "HKEY_LOCAL_MACHINE\SOFTWARE\McAfee\SystemCore\VSCore\On Access Scanner\McShield\Configuration\Default" | find "softwaredistribution"
Reg query "HKEY_LOCAL_MACHINE\SOFTWARE\McAfee\SystemCore\VSCore\On Access Scanner\McShield\Configuration\high" | find "softwaredistribution"
Reg query "HKEY_LOCAL_MACHINE\SOFTWARE\McAfee\SystemCore\VSCore\On Access Scanner\McShield\Configuration\low" | find "softwaredistribution"
reg query "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\wuauserv" | find /i "DisplayName"
reg query "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\wuauserv" | find /i "start"
echo A "Start REG_DWORD 0x4" indicates the service is disabled! SA Agent will not work correctly.
echo.
net stop wuauserv
if exist %windir%\softwaredistribution del /s /q %windir%\softwaredistribution
net start wuauserv
echo.
echo Also check %windir%\windowsupdate.log for errors.

