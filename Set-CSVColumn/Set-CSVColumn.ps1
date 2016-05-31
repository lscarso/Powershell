<# 
.SYNOPSIS 
Update CSV adding custom Column

.DESCRIPTION 
Read a CSV and add Column with calculated value 

.NOTES
Edit $CSVinputFile with CSV to read
Edit $CSVexportFile with CSV to save
Edit "Expression" and "Name" for custom Column Value (Line 24)
#>

#Settings
$CSVinputFile = "C:\Data\Working\Input.csv"
$CSVexportFile = "C:\Data\Working\Export.csv"

$Table = @()
$Table = import-csv $CSVinputFile -Delimiter ";" 

#Edit Custom Column Expression
$Table | select *,@{Name="Mail";Expression={$(Get-ADUser $($_.Username).trim() -Properties EmailAddress).EmailAddress}} | export-csv $CSVexportFile -Delimiter ";" -noType
