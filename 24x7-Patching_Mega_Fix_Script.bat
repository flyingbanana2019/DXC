del "c:\program files\common files\opsware\agent\*.*" /F /S /Q
rmdir "c:\program files\common files\opsware\agent\ogsh.push" /Q

Echo Deleted HPSA cache directory

Reg query "HKEY_LOCAL_MACHINE\SOFTWARE\McAfee\SystemCore\VSCore\On Access Scanner\McShield\Configuration\Default" | find "softwaredistribution"
Reg query "HKEY_LOCAL_MACHINE\SOFTWARE\McAfee\SystemCore\VSCore\On Access Scanner\McShield\Configuration\high" | find "softwaredistribution"
Reg query "HKEY_LOCAL_MACHINE\SOFTWARE\McAfee\SystemCore\VSCore\On Access Scanner\McShield\Configuration\low" | find "softwaredistribution"
reg query "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\wuauserv" | find /i "DisplayName"
reg query "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\wuauserv" | find /i "start"

echo Checking HPSA Agent is correct in registry. A "Start REG_DWORD 0x4" indicates the service is disabled! HPSA Agent will not work correctly and will need investigation.
echo.

net stop wuauserv
if exist %windir%\softwaredistribution del /s /q %windir%\softwaredistribution
net start wuauserv
echo deleted all files from Windows Update Cache
echo Also check %windir%\windowsupdate.log for errors.

if exist "%ProgramFiles%\Opsware\agent\bin\WindowsUpdateAgent-x86.exe" "%ProgramFiles%\Opsware\agent\bin\WindowsUpdateAgent-x86.exe" /quiet /norestart
echo Check and installed x86 Windows Update if required
if exist "%ProgramFiles%\Opsware\agent\bin\WindowsUpdateAgent-x64.exe" "%ProgramFiles%\Opsware\agent\bin\WindowsUpdateAgent-x64.exe" /quiet /norestart
echo Check and installed x64 Windows Update if required
