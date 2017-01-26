@ECHO OFF
:: .SYNOPSIS
:: Call powershell New-ProvisioningPrinter.ps1 script
:: .PARAMETER: ComputerName
:: Hostname or IP address of remote printserver
:: .PARAMETER: CSVFile
:: CSVFile with printers attributes to create
:: .EXAMPLE
:: NEW-PROVISIONINGPRINTERFROMCSV.cmd PrintServerHost C:\Data\printers.cvs

SETLOCAL
IF [%1]==[] (
ECHO WARNING: Missing SERVER name
SET /p PC="Server Name: "
ECHO.
) ELSE (
SET PC=%1
)

IF [%2]==[] (
ECHO WARNING: Missing CSV File
ECHO WARNING: CSV file should be saved with ';' delimiter and UTF8 Encoding
SET /p CSV="CSV File: "
ECHO.
) ELSE (
SET CSV=%2
)

powershell.exe -ExecutionPolicy Bypass -Command ". .\New-ProvisioningPrinter.ps1; Import-Csv %CSV% -delimiter ';' -Encoding UTF8 | New-ProvisioningPrinter -ComputerName %PC% "

ENDLOCAL



