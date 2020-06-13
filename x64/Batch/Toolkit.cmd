@echo off
cd /d "%~dp0"
for %%a in ("%~dp0") do set parent=%%~dpa
for %%a in ("%parent:~0,-1%") do set grandparent=%%~dpa

set SetACL=%grandparent%Configurator\SetACL.exe

title ViPER4Window Fix Toolkit
:list
CLS
echo.
echo. 01. Take Ownership of ViPER4Window Installation Folder
echo. 
echo. 02. Take Ownership of these Registry Key
echo. HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render
echo. HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture
echo. 
echo. 03. Register ViPER4Windows.dll
echo. 
echo. 04. Allow APO with PureSoftApps Signature
echo.
echo. 05. Exit
echo.

CHOICE /C 12345 /M "Enter your choice:"

:: Note - list ERRORLEVELS in decreasing order
IF ERRORLEVEL 5 goto :eof
IF ERRORLEVEL 4 goto :4
IF ERRORLEVEL 3 goto :3
IF ERRORLEVEL 2 goto :2
IF ERRORLEVEL 1 goto :1

:4
call ImportCertificate.cmd
echo Done
pause
goto list 

:3
regsvr32 /s "%parent%ViPER4Windows.dll"
echo Done
pause
goto list

:2
"%SetACL%" -on "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render" -ot reg -actn setowner -ownr "n:Administrators"
"%SetACL%" -on "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render" -ot reg -actn ace -ace "n:Administrators;p:full"
"%SetACL%" -on "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render" -ot reg -actn ace -ace "n:Users;p:full"
"%SetACL%" -on "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture" -ot reg -actn setowner -ownr "n:Administrators"
"%SetACL%" -on "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture" -ot reg -actn ace -ace "n:Administrators;p:full"
"%SetACL%" -on "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture" -ot reg -actn ace -ace "n:Users;p:full"
echo Done
pause
goto list

:1
call Ownership.cmd
echo Done
pause
goto list