@ECHO Off
CLS
ECHO "Disabling SMB1...."
powershell.exe -ExecutionPolicy Bypass -file ".\Disable-SMB1.ps1"

for /f %%i in ('powershell.exe -ExecutionPolicy Bypass -file .\Detect-SMB1Enabled.ps1') do set VAR=%%i

ECHO "SMB1 Status: %VAR%"

Pause