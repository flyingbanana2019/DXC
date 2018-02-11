@echo off
setlocal

set regcmd=%SystemRoot%\system32\reg.exe
set keypath=SOFTWARE\Microsoft\Office\15.0\Excel\Security
set valuename=WorkbookLinkWarnings

:: update current user
set hive=HKCU
set key=%hive%\%keypath%
%regcmd% add "%key%" /v %valuename% /d 0x00000002 /t REG_DWORD /f >nul

:: update all other users on the computer, using a temporary hive
set hive=HKLM\TempHive
set key=%hive%\%keypath%

:: set current directory to "Documents and Settings"
cd /d %USERPROFILE%\..
:: enumerate all folders
for /f "tokens=*" %%i in ('dir /b /ad') do (
if exist ".\%%i\NTUSER.DAT" call :AddRegValue "%%i" ".\%%i\NTUSER.DAT"
)

endlocal
echo.
echo Finished...
echo.
pause

goto :EOF

:AddRegValue
set upd=Y
if /I %1 equ "All Users" set upd=N
if /I %1 equ "LocalService" set upd=N
if /I %1 equ "NetworkService" set upd=N

if %upd% equ Y (
%regcmd% load %hive% %2 >nul 2>&1
%regcmd% add "%key%" /v %valuename% /d 0x00000002 /t REG_DWORD /f >nul 2>&1
%regcmd% unload %hive% >nul 2>&1
)
