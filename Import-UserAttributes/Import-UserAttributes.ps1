<# 
.SYNOPSIS 
Compare and Import users attributes from a SVC File.

.DESCRIPTION 
Compare and Import users attributes from a SVC File.

.PARAMETER $Whatif 
If present no changes are written to AD (check only)

.PARAMETER $importFile
Path and Filename of CSV with samaccountname and attributes

.EXAMPLE 
.\Import_UserAttributes.ps1 -Attributes manager -importFile C:\Data\Working\usr.csv
Import manager attribute from CSV file specified

.NOTES 
CSV must use ";" as delimiter
CSV must have the attribute name for each row title
The username must be given as simple samaccountname
Rowtitle for the username is "samaccountname"
#>

Param( 
     [string]$importFile = "usr.csv", 
     [switch]$whatif,
     [string[]]$Attributes
     ) 

$arrayattributes = @()

Clear-Variable VAR*, Csv*, ad*
#Get-Variable -Exclude PWD,*Preference, importFile, whatif, Attributes | Remove-Variable -EA 0

# If set to false no changes are written to AD (check only)
$doit = !($whatif)


Function SET-ADattribute ([string]$VARinput){ 
        #Check for special Attributes
        if ($VARinput -eq "Manager"){
            $user.$VARinput = $(get-aduser $user.$VARinput).DistinguishedName
        }
        #Generic Attributes
		if ($adusr.$VARinput) {
            if ($user.$VARinput -eq $adusr.$VARinput) { 
            } else {
			    if ($doit) {
				    if ($user.$VARinput) {
					    set-aduser $adusr -replace @{$VARinput=$user.$VARinput} #-whatif
				    } else {
					    set-aduser $adusr -clear $VARinput #-whatif
				    }
			    }
			    #write-host -fore red "$VARinput :       $($adusr.$VARinput) -> $($user.$VARinput)"
                write-host -fore red "$(($VARinput + ':').padright(30)) $($adusr.$VARinput) -> $($user.$VARinput)" 
                #write-host -fore Red ("{0}: .padright($(30 - $VARinput.Length),' ') {1} -> {2}" -f $VARinput,$($adusr.$VARinput),$($user.$VARinput))
            }
		} else {
            if ($user.$VARinput) {
			    if ($doit) {
				    set-aduser $adusr -replace @{$VARinput=$user.$VARinput} #-whatif
			    }
			    write-host -fore yellow "$VARinput :       * $($user.$VARinput)"
		    } else { }
        }
}


Try{
       
    $nl = ""
    $nl ; $nl ; $nl ; $nl
    write-host -fore gray -back gray "------------------------------------"
    $nl

    # Check Attributes passed on parameter with CSV header
    $CsvAttributes = $(import-csv $importFile -delimiter ";" -Encoding Default | select -first 1 | Get-Member -MemberType Properties).Name
    
    If ($CsvAttributes -ccontains "samaccountname") {
        Write-Host -fore green "samaccountname column is present"
        $nl
    } Else {
        Write-Host -fore red "samaccountname column is not present!"
        Exit
    }

    Foreach ($CsvAtt in $CsvAttributes) {
        if ($CsvAtt -eq "samaccountname"){
        write-host -fore green "Enabled:  $CsvAtt"; $arrayattributes += @($CsvAtt)
        } Else {
            if ($attributes -contains $CsvAtt) { write-host -fore green "Enabled:  $CsvAtt"; $arrayattributes += @($CsvAtt)} else { write-host -fore gray "Disabled: $CsvAtt"}
        }
    }

    $nl
    write-host -fore gray -back gray "------------------------------------"
    $nl ; $nl

    if ($doit) { write-host -fore red -back yellow "CHANGES WILL BE PROCESSED!" } else { write-host -fore cyan "No changes will be made!" }
    $nl ; $nl
    read-host "Press enter to continue"
    $nl

    $users = import-csv $importFile -delimiter ";" -Encoding Default #| select -first 1

    foreach ($user in $users) {
	    $adusr = $NULL
	    $adusr = get-aduser $user.samaccountname -property *
	    write-host -fore gray -back gray "------------------------------------"
	    $nl
	    write-host -fore cyan "$($adusr.samaccountname):"
    
        foreach ($attribute in $arrayattributes){
                Set-ADattribute ($attribute)
        }
	    $nl
    }

    write-host -fore gray -back gray "------------------------------------"
    $nl
}
catch [Exception]{
    $ErrorMsg = "ERROR: $($_.Exception.InnerException)"
    write-host -fore red "$(($_.Exception.ParamName + ':').padright(30)) $ErrorMsg"
}
