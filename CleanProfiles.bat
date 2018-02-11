@echo on
if %os%==Windows_NT goto WINNT
goto NOCON

:WINNT
echo .Using a Windows NT based system
echo ..%computername%

cd /d "c:\users"

echo Deleting Temporary Internet Files
for /d %%a in (*) do del /q /f /s "c:\users\%%a\AppData\Local\Microsoft\Windows\Temporary Internet Files\*.*" >nul 2>&1
echo deleted!

rem echo Deleting Downloads Folder Files
rem del /q /f /s "%USERPROFILE%\Downloads\*.*"
rem echo deleted!
 
echo Deleting Cookies
for /d %%a in (*) do del /q /f /s "%USERPROFILE%\AppData\Roaming\Microsoft\Windows\Cookies\*.* >nul 2>&1
for /d %%a in (*) do del /q /f /s "%USERPROFILE%\AppData\LocalLow\Microsoft\Internet Explorer\DOMStore\*.*" >nul 2>&1
echo deleted!

echo Deleting History
for /d %%a in (*) do del /q /f /s "%USERPROFILE%\AppData\Local\Microsoft\Windows\History\*.*" >nul 2>&1
for /d %%a in (*) do del /q /f /s "%USERPROFILE%\AppData\Local\Microsoft\Internet Explorer\Recovery\Active\*.*" >nul 2>&1
for /d %%a in (*) do del /q /f /s "%USERPROFILE%\AppData\Local\Microsoft\Internet Explorer\Recovery\Last Active\*.*" >nul 2>&1
echo deleted!

echo Deleting Windows Internet Explorer Dat Files
for /d %%a in (*) do del /q /f /s "c:\users\%%a\AppData\Roaming\Microsoft\Windows\PrivacIE\*.*" >nul 2>&1
for /d %%a in (*) do del /q /f /s "c:\users\%%a\AppData\Roaming\Microsoft\Windows\IECompatCache\*.*" >nul 2>&1
for /d %%a in (*) do del /q /f /s "c:\users\%%a\AppData\Roaming\Microsoft\Windows\IETldCache\*.*" >nul 2>&1
echo deleted!

echo Deleting Windows Error Reporting Files
for /d %%a in (*) do del /q /f /s "c:\users\%%a\AppData\Local\Microsoft\Windows\WER\ReportArchive\*.*" >nul 2>&1
echo deleted!

echo Deleting Flash Player Temp Files
for /d %%a in (*) do del /q /f /s "c:\users\%%a\AppData\Roaming\Macromedia\Flash Player\*.*" >nul 2>&1
echo deleted!

echo Deleting Remote Desktop Cache
for /d %%a in (*) do del /q /f /s "%USERPROFILE%\AppData\Local\Microsoft\Terminal Server Client\Cache\*.*" >nul 2>&1
echo deleted!

echo Deleting Profile Temp Files
for /d %%a in (*) do del /q /f /s "c:\users\%%a\AppData\Local\Temp\*.*" >nul 2>&1
echo deleted!

rem echo Delete misc Files in Profile
rem del /q /f /s "%USERPROFILE%\webct_upload_applet.properties"
rem del /q /f /s "%USERPROFILE%\g2mdlhlpx.exe"
rem del /q /f /s "%USERPROFILE%\fred"
rem rmdir /s /q "%USERPROFILE%\temp"
rem rmdir /s /q "%USERPROFILE%\WebEx"
rem rmdir /s /q "%USERPROFILE%\.gimp-2.4"
rem rmdir /s /q "%USERPROFILE%\.realobjects"
rem rmdir /s /q "%USERPROFILE%\.thumbnails"
rem rmdir /s /q "%USERPROFILE%\Bluetooth Software"
rem rmdir /s /q "%USERPROFILE%\Office Genuine Advantage"
rem echo deleted!

echo Deleting FireFox Cache
pushd "%USERPROFILE%\AppData\Local\Mozilla\Firefox\Profiles\*.default\"
del /q /f /s "Cache\*.*"
popd

echo deleted!

echo Deleting User Profile Adobe Temp Files
for /d %%a in (*) do del /q /f /s "c:\users\%%a\AppData\LocalLow\Adobe\Acrobat\9.0\Search\*.*" >nul 2>&1
for /d %%a in (*) do del /q /f /s "c:\users\%%a\AppData\LocalLow\Adobe\Common\Media Cache Files\*.*" >nul 2>&1
for /d %%a in (*) do del /q /f /s "c:\users\%%a\AppData\LocalLow\Adobe\Common\Media Cache\*.*" >nul 2>&1
echo deleted!

echo Deleting User Office Recent Files
for /d %%a in (*) do del /q /f /s "c:\users\%%a\AppData\Roaming\Microsoft\Office\Recent\*.*" >nul 2>&1
echo deleted!

echo Deleting User Office TMP Files
for /d %%a in (*) do del /q /f /s "c:\users\%%a\AppData\Roaming\Microsoft\Office\*.tmp" >nul 2>&1
echo deleted!

goto END

:NOCON
echo Error...Invalid Operating System...
echo Error...No actions were made...
goto END

:END

pause
