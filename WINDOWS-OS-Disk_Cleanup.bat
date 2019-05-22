@echo off

rem ***************************************
rem Made by Esteban Mena * emenasol@hp.com
rem Script based on Disk Clean Up Procedure – L1 - Windows - Clorox.doc document
rem ***************************************


rem ***************************************************************** 
rem Variables 
rem ***************************************************************** 
set scriptname=%computername%-CleaupLog
set logname=%scriptname:.cmd=%.log
set allusers=%SystemDrive%\Documents and Settings\All Users
set YEAR=%DATE:~-4%
set profiledir=Documents and Settings

(
echo.
echo *********************************************************
echo ** Clean Disk Space for %computername%  %date% %time% 
echo ** Disk CleanUp Process Started in Drive %HOMEDRIVE%
echo *********************************************************
echo.


rem ***************************************************************** 
rem Delete all (*.tmp) files. Any protected file is deleted
rem ***************************************************************** 
echo Deleting all "*.tmp" files. Any protected file is deleted
echo.
cd\
del /s *.tmp /q /a h

rem ***************************************************************** 
rem Delete all (*.dmp) files older that 30 days. Any protected file is deleted  
rem ***************************************************************** 
echo Deleting all "*.dmp" files older that 30 days. Any protected file is deleted  
echo.
cd\
forfiles /P "%SystemDrive%" /S /M *.dmp /D -30 /c "cmd /c ECHO @path"
forfiles /P "%SystemDrive%" /S /M *.dmp /D -30 /c "cmd /c DEL @FILE /Q"
echo *********************************************************
echo.


rem ***************************************************************** 
rem Delete all files located on C:\TEMP folder. Any protected file is deleted
rem ***************************************************************** 
echo Deleting all files located on C:\TEMP folder. Any protected file is deleted
echo.
cd\
if exist "%SystemDrive%\Temp" (
cd %SystemDrive%\Temp
dir *.* /b
del *.* /q
) else (echo Path "%SystemDrive%\Temp" not exit)
echo *********************************************************
echo.

rem ***************************************************************** 
rem Delete all files located on C:\WINNT\TEMP or C:\WINDOWS\TEMP. Any protected file is deleted
rem ***************************************************************** 
echo Deleting all files located on C:\WINNT\TEMP or C:\WINDOWS\TEMP. Any protected file is deleted
echo.
cd\
if exist "%SystemRoot%\Temp" (
cd %SystemRoot%\Temp
dir *.* /b
rd /q /s %SystemRoot%\Temp
md %SystemRoot%\Temp
) else (echo Path "%SystemRoot%\Temp" not exit)
echo *********************************************************
echo.

rem ***************************************************************** 
rem Empty Recycle Bin Windows 2003
rem ***************************************************************** 
echo Empty Recycle Bin Windows 2003, XP, 2000
echo.
cd\
if exist "%SystemDrive%\Recycler" (
cd %SystemDrive%\Recycler
dir *.* /b /a h 
rd /q /s %SystemDrive%\Recycler
) else (echo Path "%SystemDrive%\Recycler" not exit)
echo.

rem ***************************************************************** 
rem Empty Recycle Bin Windows 2008
rem ***************************************************************** 
echo Empty Recycle Bin Windows 2008 
echo.
cd\
if exist "%SystemDrive%\$Recycle.Bin" (
cd %SystemDrive%\$Recycle.Bin
dir *.* /b /a h 
rd /q /s %SystemDrive%\$Recycle.Bin
) else (echo Path "%SystemDrive%\$Recycle.Bin" not exit)
echo *********************************************************
echo.

rem ***************************************************************** 
rem Delete “$Uninstall Folders and $Uninstall logs files” older than 6 moths
rem ***************************************************************** 
echo Deleting “$Uninstall Folders and $Uninstall logs files” older than 6 moths
echo.
forfiles /P "%SystemRoot%" /M $NtUninstallKB* /D -180 /c "cmd /c if @isdir==TRUE ECHO @path" 
forfiles /P "%SystemRoot%" /M $NtUninstallKB* /D -180 /c "cmd /c if @isdir==TRUE RMDIR /S /Q @path"
echo.
echo.
echo Deleting $Uninstall logs files older than 6 moths
forfiles /P "%SystemRoot%" /M KB*.log /D -180 /c "cmd /c ECHO @path" 
forfiles /P "%SystemRoot%" /M KB*.log /D -180 /c "cmd /c DEL @FILE /Q"
echo *********************************************************
echo.

rem ***************************************************************** 
rem Delete all files from C:\WINDOWS\ServicePackFiles folder
rem ***************************************************************** 
echo Deleting all files on %SystemRoot%\ServicePackFiles
echo.
cd\
if exist "%SystemRoot%\ServicePackFiles" (
cd %SystemRoot%\ServicePackFiles
dir *.* /b /a h
rd /q /s %SystemRoot%\ServicePackFiles
md %SystemRoot%\ServicePackFiles
) else (echo Path "%SystemRoot%\ServicePackFiles" not exit)
echo *********************************************************
echo.

rem ***************************************************************** 
rem *************************** Antivirus ***************************
rem Delete Antivirus OLD definitions under the directory C:\Program Files\Common Files\Symantec Shared\VirusDefs YYYYMMDD folder
rem ***************************************************************** 
echo Delete Antivirus OLD definitions "XP Machines and 2003 servers"
echo C:\Program Files\Common Files\Symantec Shared\VirusDefs\
echo.
cd\
if exist "%PROGRAMFILES%\Common Files\Symantec Shared\VirusDefs" (
forfiles /P "%PROGRAMFILES%\Common Files\Symantec Shared\VirusDefs" /M %YEAR%* /D -2 /c "cmd /c if @isdir==TRUE ECHO @path" 
forfiles /P "%PROGRAMFILES%\Common Files\Symantec Shared\VirusDefs" /M %YEAR%* /D -2 /c "cmd /c if @isdir==TRUE RMDIR /S /Q @path"
) else (echo Path "%PROGRAMFILES%\Common Files\Symantec Shared\VirusDefs" not exit)
echo *********************************************************
echo.
echo.

rem ***************************************************************** 
rem Delete files from C:\ProgramData\Symantec\Definitions\VirusDefs
rem ***************************************************************** 
echo Deleting Antivirus OLD definitions "2008 servers"
echo.
cd\
if exist "%PROGRAMDATA%\Symantec\Definitions\VirusDefs" (
forfiles /P "%PROGRAMDATA%\Symantec\Definitions\VirusDefs" /M %YEAR%* /D -2 /c "cmd /c if @isdir==TRUE ECHO @path" 
forfiles /P "%PROGRAMDATA%\Symantec\Definitions\VirusDefs" /M %YEAR%* /D -2 /c "cmd /c if @isdir==TRUE RMDIR /S /Q @path"
) else (echo Path "%PROGRAMFILES%\Common Files\Symantec Shared\VirusDefs" not exit)
echo *********************************************************
echo.

rem ***************************************************************** 
rem C:\Documents and Settings\All Users\Application Data\Symantec\LiveUpdate\Downloads\*
rem ***************************************************************** 
echo Deleting Antivirus OLD LiveUpdates Downloaded from C:\Documents and Settings\All Users\Application Data\Symantec\LiveUpdate\Downloads
echo.
cd\
if exist "%allusers%\Application Data\Symantec\LiveUpdate\Downloads" (
forfiles /P "%allusers%\Application Data\Symantec\LiveUpdate\Downloads" /M *.* /D -2 /c "cmd /c ECHO @path" 
forfiles /P "%allusers%\Application Data\Symantec\LiveUpdate\Downloads" /M *.* /D -2 /c "cmd /c DEL @FILE /Q"
) else (echo Path "%allusers%\Application Data\Symantec\LiveUpdate\Downloads" not exit)
echo *********************************************************
echo.

rem ***************************************************************** 
rem Delete files from C:\Documents and Settings\All Users\Application Data\Symantec\Symantec Endpoint Protection, upper of 50 KB
rem ***************************************************************** 
echo Deleting Antivirus OLD Files on NT servers from C:\Documents and Settings\All Users\Application Data\Symantec\Symantec Endpoint Protection
echo.
cd\
if exist "%SystemRoot%\Profiles\All Users\Application Data\Symantec\Norton Antivirus Corporate Edition\7.5" (
forfiles /P "%SystemRoot%\Profiles\All Users\Application Data\Symantec\Norton Antivirus Corporate Edition\7.5" /M *.*  /c "cmd /c IF @fsize GEQ 51200 (ECHO @path)" 
forfiles /P "%SystemRoot%\Profiles\All Users\Application Data\Symantec\Norton Antivirus Corporate Edition\7.5" /M *.*  /c "cmd /c IF @fsize GEQ 51200 (del @file /Q)"
) else (echo Path "%SystemRoot%\Profiles\All Users\Application Data\Symantec\Norton Antivirus Corporate Edition\7.5" not exit)
echo *********************************************************
echo.

rem ***************************************************************** 
rem Delete files from C:\ProgramData\Symantec\Symantec Endpoint Protection\Logs
rem ***************************************************************** 
echo Deleting files from C:\ProgramData\Symantec\Symantec Endpoint Protection\Logs older that 5 day and upper of 50 KB
echo.
cd\
if exist "%PROGRAMDATA%\Symantec\Symantec Endpoint Protection\Logs" (
forfiles /P "%PROGRAMDATA%\Symantec\Symantec Endpoint Protection\Logs" /M *.log /D -5 /c "cmd /c IF @fsize GEQ 51200 (ECHO @path)" 
forfiles /P "%PROGRAMDATA%\Symantec\Symantec Endpoint Protection\Logs" /M *.log /D -5 /c "cmd /c IF @fsize GEQ 51200 (del @file /Q)"
) else (echo Path "%PROGRAMDATA%\Symantec\Symantec Endpoint Protection" not exit)
echo *********************************************************
echo.

rem ************************ Antivirus *********************** 

rem ***************************************************************** 
rem Delete ScanFile and download forders on C:\Windows\SoftwareDistribution
rem ***************************************************************** 
echo Deleting ScanFile folder on %SystemRoot%\SoftwareDistribution\ScanFile
echo.
cd\
if exist "%SystemRoot%\SoftwareDistribution\ScanFile" (
cd %SystemRoot%\SoftwareDistribution\ScanFile
dir *.* /b /a h
del *.* /q /a h
) else (echo Path "%SystemRoot%\SoftwareDistribution\ScanFile" not exit)
echo.

echo Deleting ScanFile folder on %SystemRoot%\SoftwareDistribution\download
echo.
cd\
if exist "%SystemRoot%\SoftwareDistribution\download" (
cd %SystemRoot%\SoftwareDistribution\download
dir *.* /b /a h
del *.* /q /a h
) else (echo Path "%SystemRoot%\SoftwareDistribution\download" not exit)
echo *********************************************************
echo.

rem ***************************************************************** 
rem Delete Wintel L1, L2 and L3 user profiles older than 1 moth
rem ***************************************************************** 
echo Deleting Wintel L1, L2 and L3 user profiles older than 1 moth
echo.
cd\
if exist "%SystemDrive%\%profiledir%" (
cd %SystemDrive%\%profiledir%
for /f "tokens=1" %%i in (C:\script\wintel-users.txt) do (
forfiles /P "%SystemDrive%\%profiledir%" /M %%i /D -30 /c "cmd /c if @isdir==TRUE ECHO @path" 
forfiles /P "%SystemDrive%\%profiledir%" /M %%i /D -30 /c "cmd /c if @isdir==TRUE RMDIR /S /Q @path"
)
) else (echo Path "%SystemDrive%\%profiledir%" not exit)
echo *********************************************************
echo.

echo.
echo %date% %time% Disk CleanUp Process Finished
cd c:\Program Files\HP OpenView\data\bin\instrumentation
df_mon -s -f


)>%logname% 2>>&1