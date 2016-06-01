<# 
.SYNOPSIS 
Update CSV adding custom Column

.DESCRIPTION 
Read a CSV and add Column with calculated value 

.NOTES
Edit $CSVinputFile with CSV to read
$CSVinputFile should have "Username" label on column of user-id

Edit $CSVexportFile with CSV to save

Edit "Expression" and "Name" for custom Column Value (Line 29)

Author: Luca Scarsini

#>

#Settings
$CSVinputFile = "C:\Data\Working\Utenti_Tecnici.csv"
$CSVexportFile = "C:\Data\Working\homeUser.csv"

$Table = @()
$Table = import-csv $CSVinputFile -Delimiter ";" 

#Edit Custom Column Expression

#Email Address
#$Table | select *,@{Name="Mail";Expression={$(Get-ADUser $($_.Username).trim() -Properties EmailAddress).EmailAddress}} | export-csv $CSVexportFile -Delimiter ";" -noType

#Home folder
#$Table | select *,@{Name="Home";Expression={$(Get-ADUser $($_.Username).trim() -Properties homedirectory).homedirectory}} | export-csv $CSVexportFile -Delimiter ";" -noType

#Home folder size
$Table | select *,@{Name="Home Size";Expression={"{0:N2}" -f ($(Get-ChildItem $(Get-ADUser $($_.Username).trim() -Properties homedirectory).homedirectory -Recurse | Measure-Object -Property length -sum).sum / 1Mb)}} | export-csv $CSVexportFile -Delimiter ";" -noType
