::'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
:: Win-RnmLclAccts.bat
::
:: Kirk Aragon                                       Opsware, Inc.  10/6/2004
::'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
::
:: Rename the local Administrator and Guest accounts
::
::
:: Script command syntax:
::
:: This batch file accepts two arguments to rename
:: the Administrator and Guest accounts to EDS
:: standards.  The options are:
::
::     --rename       Renames administrator to
::                    localadmin and guest to
::                    localguest. Disables
::                    localguest.
::     --reset        Renames localadmin to
::                    administrator and localguest
::                    to guest.
::
::'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
::
:: Syntax for the CUsrMgr command:
::
:: Sets a random password to a user.
:: usage: -u UserName [-m \\MachineName] \\ default LocalMachine
::  Resetting Password Function
::       -p Set to a random password
::       -P xxx Sets password to xxx
::  User Functions
::       -r xxx Renames user to xxx
::       -d xxx deletes user xxx
::  Group Functions
::       -rlg xxx yyy Renames local group xxx to yyy
::       -rgg xxx yyy Renames global group xxx to yyy
::       -alg xxx Add user (-u UserName) to local group xxx
::       -agg xxx Add user (-u UserName) to global group xxx
::       -dlg xxx deletes user (-u UserName) from local group xxx
::       -dgg xxx deletes user (-u UserName) from global group xxx
::  SetProperties Functions
::       -c xxx sets Comment to xxx
::       -f xxx sets Full Name to xxx
::       -U xxx sets UserProfile to xxx
::       -n xxx sets LogonScript to xxx
::       -h xxx sets HomeDir to xxx
::
::       -H x   sets HomeDirDrive to x
::
::       +s xxxx sets property xxxx
::       -s xxxx resets property xxxx
::       where xxxx can be any of the following properties:
::              MustChangePassword
::              CanNotChangePassword
::              PasswordNeverExpires
::              AccountDisabled
::              AccountLockout
::              RASUser
::returns 0 on success
::
::
::'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
::
:: Important: you are limited to 9 command line arguments for batch files!!
::
::'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
::
:: Programming notes:
::
::  * We used %SystemDrive% instead of %TEMP% as the target for the copy
::    because the TEMP variable didn't exist in the context of the command
::    prompt session that was started by the OCC. We know that SystemDrive
::    will always exist on NT systems, unless it's deleted on purpose.
::
::  * The %PP% variable is using the 8.3 format name so we don't need to
::    use double-quotes. Plus, it makes the lines shorter.
::
::'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
::
:: 2004-10-06	R.Ingenthron	Added comments and help info.
::				Added some code for error-handling.
::'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
::'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
::
:: 2004-12-29	D.Williams	Added Check for existance of the tool/utility
::				Added Header RUMCOMMAND
::                              Change path of tools/utilities to
::                              %SystemDrive%/EDS
::'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
::'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

@ECHO OFF

SETLOCAL

:: This is just some fancy code to display this file name with an underline.
::
set !Underline=
set !ThisFile=%~nx0
set !Tmp1=%!ThisFile%
set /A len=0
:Loop1
  set !Tmp2=%!Tmp1:~0,1%
  set !Tmp3=%!Tmp1:~1%
  set /A len=len + 1
  if "%!Tmp3%" == ""  goto :EndLoop
  set !Tmp1=%!Tmp3%
  set !Underline=%!Underline%-
  goto :Loop1
:EndLoop
set !Underline=%!Underline%-
echo.
echo.
echo %!ThisFile%
echo %!Underline%
set !Tmp3=
set !Tmp2=
set !Tmp1=
set !ThisFile=
set len=
set !Underline=
ECHO.

If Exist %SystemDrive%\EDS goto BEGIN
mkdir %SystemDrive%\EDS

:BEGIN
SET PP=%SystemDrive%\Progra~1

::
:: Check for the existance of the tool/utility
::
ECHO.
%SystemDrive%\EDS\CUSRMGR.EXE > NUL
if NOT errorlevel == 9009  goto RUNCOMMAND

::
:: Download auditpol.exe from theword
::
:: Set the Python path...
::
::...ECHO.
::...ECHO set pythonpath=%PP%\Loudcloud\blackshadow
set pythonpath=%PP%\Loudcloud\blackshadow

::
:: Download the command...
::
ECHO.
ECHO Getting command file...
::...ECHO %PP%\Loudcloud\lcpython15\pythonw.exe %PYTHONPATH%\coglib\wordclient.pyc --word /packages/any/nt/5.2/CUSRMGR.EXE %SystemDrive%\EDS\CUSRMGR.EXE
%PP%\Loudcloud\lcpython15\pythonw.exe %PYTHONPATH%\coglib\wordclient.pyc --word /packages/any/nt/5.2/CUSRMGR.EXE %SystemDrive%\EDS\CUSRMGR.EXE

::
:: Verify download completed
::
ECHO.
%SystemDrive%\EDS\CUSRMGR.EXE > NUL
ECHO Status of command download: %errorlevel%
if %errorlevel% GTR 1 goto NODOWNLOAD

:RUNCOMMAND

if "%1" == "--reset" goto RESET
if "%1" == "--rename" goto RENAME
if NOT "%1" == "" goto HELP
if "%1" == "" goto HELP


:RENAME
GOTO GUEST


:GUEST
ECHO.
%SystemDrive%\EDS\CUSRMGR.EXE -u GUEST -r LOCALGUEST > NUL
if %errorlevel% GTR 11 goto NOGUEST
%SystemDrive%\EDS\CUSRMGR.EXE -u LOCALGUEST +s AccountDisabled > NUL
ECHO Guest account renamed to LocalGuest and disabled.
GOTO ADMIN


:NOGUEST
ECHO.
ECHO No guest account exists on this computer
ECHO or the guest account has been previously
ECHO renamed.
GOTO ADMIN


:ADMIN
ECHO.
%SystemDrive%\EDS\CUSRMGR.EXE -u ADMINISTRATOR -r LOCALADMIN > NUL
if %errorlevel% GTR 11 goto NOADMIN
%SystemDrive%\EDS\CUSRMGR.EXE -u LOCALADMIN -s PasswordNeverExpires > NUL
ECHO Admin account renamed to LocalAdmin and PasswordNeverExpires
ECHO set off.
GOTO DELETEFILE


:NOADMIN
ECHO.
ECHO The admin account has been renamed
ECHO previously or does not exist.
GOTO DELETEFILE


:HELP
ECHO.
ECHO This batch file accepts two arguments to rename
ECHO the Administrator and Guest accounts to EDS
ECHO standards.  The options are:
ECHO.
ECHO     --rename       Renames administrator to
ECHO                    localadmin and guest to
ECHO                    localguest. Disables
ECHO                    localguest.
ECHO     --reset        Renames localadmin to
ECHO                    administrator and localguest
ECHO                    to guest.
ECHO.
GOTO DELETEFILE


:RESET
ECHO.
%SystemDrive%\EDS\CUSRMGR.EXE -u LOCALGUEST -r GUEST > NUL
if %errorlevel% GTR 11 goto NOLOCALGUEST
%SystemDrive%\EDS\CUSRMGR.EXE -u GUEST +s AccountDisabled > NUL
ECHO The Localguest account has been renamed to Guest
ECHO and disabled.
GOTO LOCALADMIN


:NOLOCALGUEST
ECHO.
ECHO No localguest account exists on this computer
ECHO or the localguest account has been previously
ECHO renamed.
GOTO LOCALADMIN


:LOCALADMIN
ECHO.
%SystemDrive%\EDS\CUSRMGR.EXE -u LOCALADMIN -r ADMINISTRATOR > NUL
if %errorlevel% GTR 11 goto NOLOCALADMIN
%SystemDrive%\EDS\CUSRMGR.EXE -u ADMINISTRATOR +s PasswordNeverExpires > NUL
ECHO LocalAdmin account renamed to Administrator and PasswordNeverExpires
ECHO set on.
GOTO DELETEFILE


:NOLOCALADMIN
ECHO.
ECHO The localadmin account had been previously
ECHO renamed or does not exist.
GOTO DELETEFILE


:NODOWNLOAD
ECHO.
ECHO CUSRMGR.EXE did not download from theword.
ECHO Verify that CUSRMGR.EXE has been uploaded
ECHO as an unknown package type for this version
ECHO of the operating system.
ECHO.
GOTO END


:DELETEFILE
IF EXIST %SystemDrive%\EDS\CUSRMGR.EXE  DEL /F %SystemDrive%\EDS\CUSRMGR.EXE
GOTO END

:END

ENDLOCAL
