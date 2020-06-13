for %%a in ("%~dp0") do set parent=%%~dpa
for %%a in ("%parent:~0,-1%") do set grandparent=%%~dpa
takeown /f "%grandparent%." /r /d n
takeown /f "%grandparent%DriverComm" /r /d n
cacls "%grandparent%." /t /e /p administrators:f
cacls "%grandparent%DriverComm" /t /e /p administrators:f
cacls "%grandparent%." /t /e /p users:f