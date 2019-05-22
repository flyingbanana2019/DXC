@echo off
REM -----------------------------------------
REM Windows Uptime script
REM Author: Pablo Pagani pablo.pagani@hp.com
REM Created: 2013-07-08 Updated: 2013-07-12
REM -----------------------------------------

REM |****** USAGE *******
REM To use systeminfo method don´t pass any parameter
REM To use wmic method pass "any" parameter
REM systeminfo format is:   System Boot Time:          02/07/2013, 08:02:16 p.m.  (Server 2008 and above)
REM                         System Up Time:            39 Days, 21 Hours, 23 Minutes, 14 Seconds (Server 2003 and before)
REM wmic format is:         LastBootUpTime=20130702200216.490480-180
REM
REM Notes:
REM 1-wmic method runs faster than systeminfo method, and the format is standard for all the platforms.
REM 2-systeminfo method return "System Up Time" or "System Boot Time" depending on the platform. The text and date format depends on the locale.
REM |****** USAGE *******

IF [%1] NEQ [] GOTO WMIC

systeminfo | findstr /i "time:"
GOTO EXIT

:WMIC
wmic OS Get LastBootUpTime /format:list

:EXIT
exit 0