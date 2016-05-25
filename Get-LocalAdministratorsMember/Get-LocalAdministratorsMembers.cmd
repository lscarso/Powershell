@ECHO OFF
:: NAME: Get-LocalAdministratorsMembers.CMD v1.0
:: DATE: 02/03/20015
:: PURPOSE:  Run Get-LocalAdministratorsMember.ps1 bypassing powershell execution policy.
:: No admin rights is required. 

powershell.exe -ExecutionPolicy Bypass -file Get-LocalAdministratorsMembers.ps1
