::
::NAME: WINDOWS-OS-Fix_opsware-1448.bat
::PURPOSE: Resolves the issue with opsware when you try to run a job in HPSA
::SYNTAX:
::  < script >
::USAGE:
::  < script >
::NOTE:
::  Increase HPSA script output to 999
::
::AUTHOR(s): Kazalukian, Walter Leandro <walter.kazalukian@hpe.com>
::DATE WRITTEN: 23 Jun 2016
::MODIFICATION HISTORY:
::  23 Jun 2016 Kazalukian, Walter Leandro
::    - initial release
::

net stop wuauserv 2>nul
del "%ProgramFiles%\Common Files\Opsware\agent\last_sw_inventory" 2>nul
del "%ProgramFiles%\Common Files\Opsware\agent\wsusscn2.cab" 2>nul
del "%ProgramFiles%\Opsware\agent\bin\wusscan.dll" 2>nul
del "%ProgramFiles%\Opsware\agent\bin\qchain.exe" 2>nul
del "%ProgramFiles%\Opsware\agent\bin\mbsacli20.exe" 2>nul
del "%ProgramFiles%\Opsware\agent\bin\WindowsUpdateAgent20-x86.exe" 2>nul
del "%ProgramFiles%\Opsware\agent\bin\WindowsUpdateAgent20-x64.exe" 2>nul
del "%ProgramFiles%\Opsware\agent\bin\wusscan.dll" 2>nul
del "%ProgramFiles%\Opsware\agent\bin\parsembsacli20.exe" 2>nul
del "%ProgramFiles%\Opsware\agent\bin\parsembsacli20_x86.exe" 2>nul
del "%ProgramFiles%\Opsware\agent\bin\parsembsacli20_x64.exe" 2>nul
del /S /Q %WINDIR%\SoftwareDistribution\*.* 2>nul
net start wuauserv 2>nul
"%ProgramFiles%\Opsware\agent\pylibs\cog\bs_hardware.bat" –debug 2>nul
"%ProgramFiles%\Opsware\agent\pylibs\cog\bs_software.bat 2>nul
hostname

exit 0
