# Powershell
Powershell Scripts


# Small/various powershell strings
#### [Import displayName from CSV -> find samAccountName -> export displayName and samAccountName to CSV](https://gist.github.com/lscarso/6c8b55fc3a04657deb740613c33a0e62)
```powershell
Import-Csv C:\Data\Working\users.csv -Delimiter ";" | foreach {Get-ADUser -LDAPFilter "(ObjectClass=User)(anr=$($_.displayname))" -Properties displayname, SamAccountName} | select-Object displayname,SamAccountName | Export-Csv -Delimiter ";" C:\Data\Working\Users_DN_SAM.csv -NoTypeInformation
```
#### Uppercase separate words
```powershell
Import-Csv C:\data\Working\UtentiTest.csv -Delimiter ";" | % {Write-Host (Get-Culture).TextInfo.ToTitleCase(($_."Company").ToLower())}
```
Note that words that are entirely in upper-case are not converted.

#### Find Accounts -> get manager attribute -> export name, displayName, manager.displayName, manager.mail
```powershell
Get-ADUser -filter {name -like "ex*" -and name -notlike "*temp"} -SearchBase "OU=Users Objects,dc=mydomain,dc=local" -properties displayName, manager | select-object name, displayName, @{n='Tutor';e={(get-aduser ($_.Manager) -Properties displayName).displayName}}, @{n='Tutor e-mail Address';e={(get-aduser ($_.Manager) -Properties mail).mail}} | Export-Csv C:\data\Working\externalUsers.csv -Delimiter ';' -NoTypeInformation
```
