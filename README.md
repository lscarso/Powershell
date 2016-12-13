# Powershell
Powershell Scripts


# Small/various powershell strings
#### [Import displayName from CSV -> find samAccountName -> export displayName and samAccountName to CSV](https://gist.github.com/lscarso/6c8b55fc3a04657deb740613c33a0e62)
```powershell
Import-Csv C:\Data\Working\users.csv -Delimiter ";" | foreach {Get-ADUser -LDAPFilter "(ObjectClass=User)(anr=$($_.displayname))" -Properties displayname, SamAccountName} | select-Object displayname,SamAccountName | Export-Csv -Delimiter ";" C:\Data\Working\Users_DN_SAM.csv -NoTypeInformation
```
