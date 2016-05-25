<# 
.SYNOPSIS 
Load reports created by Get-LocalAdministratorsMember.ps1 and create a CSV file.

.DESCRIPTION 
Load reports created by Get-LocalAdministratorsMember.ps1 and create a CSV file.

.EXAMPLE 
Export-LocalAdministratorsReport.ps1

.NOTES 
Update $rpath with the path where save the result
Update $expath with the path + filename where export the data
The file generated is a CSV file
#>

#Settings ###############################
#Add reports path
$rpath = "\\server\share$\_LocalAdmin_Script"
#Add Export path + report filename
$expath = "C:\Data\MyReport.csv"
#########################################

$results = @()
$files=Get-ChildItem "$rpath\*.txt" | % {$_.fullname}

foreach ($file in $files){
    write-host $file
    $data = get-content $file
    $fileattribute = get-item $file

    $props = @{
        date=$fileattribute.LastWriteTime.ToString('MM-dd-yyyy')
        computername=$fileattribute.name.TrimEnd(".txt")
        fileline=$null
    }

    foreach ($line in $data){
        if ($line -like "Domain*" -or $line -like "------*" -or $line -eq ""){
        }
        else {
            $strtemp = ""
            $linedata = $line.Split(" ",[StringSplitOptions]'RemoveEmptyEntries')
            $strtemp = $linedata -join "\"
            $props.fileline = $strtemp
            $results += new-object PsObject -Property $props
        }
    }
}

$results | export-csv -path $expath
