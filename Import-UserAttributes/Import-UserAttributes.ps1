<# 
.SYNOPSIS 
Compare and Import users attributes from a SVC File.

.DESCRIPTION 
Compare and Import users attributes from a SVC File.

.PARAMETER $doit 
If set to false no changes are written to AD (check only)

.PARAMETER $VAR* 

.EXAMPLE 

.NOTES 
CSV must use ";" as delimiter
CSV must have the attribute name for each row title
The username must be given as simple samaccountname
Rowtitle for the username is "samaccountname"
#>

Clear-Variable VAR*
#####################################################
# File to import
$importFile = "usr.csv"
#####################################################
# If set to false no changes are written to AD (check only)
$doit = $FALSE
#####################################################
# What attributes to change
# Set to false or comment out to disable the import / check of this attribute
$VARdescription = $TRUE
$VARinfo = $TRUE
$VARdepartment = $TRUE
# $VARemployeenumber = $TRUE
$VARmobilephone = $TRUE
# $VARcompany = $TRUE
# $VARl = $TRUE
$VARtelephonenumber = $TRUE
# $VARdivision = $TRUE
# $VARextensionattribute10 = $TRUE
# $VARc = $TRUE
$VARco = $FALSE
$VARstreetaddress = $TRUE
$VARpostalcode = $TRUE
$VARst = $TRUE
# $VARfacsimileTelephoneNumber = $TRUE
$VARtitle = $TRUE
# $VARextensionAttribute12 = $TRUE
$VARphysicalDeliveryOfficeName = $TRUE
#####################################################

Function SET-ADattribute ([string]$VARinput){ 
	
		if ($adusr.$VARinput) { if ($user.$VARinput -eq $adusr.$VARinput) {  } else {
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
		} else { if ($user.$VARinput) {
			if ($doit) {
				set-aduser $adusr -replace @{$VARinput=$user.$VARinput} #-whatif
			}
			write-host -fore yellow "$VARinput :       * $($user.$VARinput)"
		} else {  } }
	
}


#Inizialize array
$arrayattributes = @()

$nl = ""
$nl ; $nl ; $nl ; $nl
write-host -fore gray -back gray "------------------------------------"
$nl
if ($VARdescription) { write-host -fore green "Enabled:  description"; $arrayattributes += @("description")} else { write-host -fore gray "Disabled: description"}
if ($VARinfo) { write-host -fore green "Enabled:  info"; $arrayattributes += @("info")} else { write-host -fore gray "Disabled: info"}
if ($VARdepartment) { write-host -fore green "Enabled:  department"; $arrayattributes += @("department")} else { write-host -fore gray "Disabled: department"}
if ($VARemployeenumber) { write-host -fore green "Enabled:  employeenumber"; $arrayattributes += @("employeenumber")} else { write-host -fore gray "Disabled: employeenumber"}
if ($VARmobilephone) { write-host -fore green "Enabled:  mobilephone"; $arrayattributes += @("mobilephone")} else { write-host -fore gray "Disabled: mobilephone"}
if ($VARcompany) { write-host -fore green "Enabled:  company"; $arrayattributes += @("company") } else { write-host -fore gray "Disabled: company"}
if ($VARl) { write-host -fore green "Enabled:  l"; $arrayattributes += @("l") } else { write-host -fore gray "Disabled: l"}
if ($VARtelephonenumber) { write-host -fore green "Enabled:  telephonenumber"; $arrayattributes += @("telephonenumber")} else { write-host -fore gray "Disabled: telephonenumber"}
if ($VARdivision) { write-host -fore green "Enabled:  division"; $arrayattributes += @("division")} else { write-host -fore gray "Disabled: division"}
if ($VARextensionattribute10) { write-host -fore green "Enabled:  extensionattribute10"; $arrayattributes += @("extensionattribute10")} else { write-host -fore gray "Disabled: extensionattribute10"}
if ($VARc) { write-host -fore green "Enabled:  c"; $arrayattributes += @("c")} else { write-host -fore gray "Disabled: c"}
if ($VARco) { write-host -fore green "Enabled:  co"; $arrayattributes += @("co")} else { write-host -fore gray "Disabled: co"}
if ($VARstreetaddress) { write-host -fore green "Enabled:  streetaddress"; $arrayattributes += @("streetaddress")} else { write-host -fore gray "Disabled: streetaddress"}
if ($VARpostalcode) { write-host -fore green "Enabled:  postalcode"; $arrayattributes += @("postalcode")} else { write-host -fore gray "Disabled: postalcode"}
if ($VARst) { write-host -fore green "Enabled:  st"; $arrayattributes += @("st")} else { write-host -fore gray "Disabled: st"}
if ($VARfacsimileTelephoneNumber) { write-host -fore green "Enabled:  facsimileTelephoneNumber"; $arrayattributes += @("facsimileTelephoneNumber")} else { write-host -fore gray "Disabled: facsimileTelephoneNumber"}
if ($VARtitle) { write-host -fore green "Enabled:  title" } else { write-host -fore gray "Disabled: title"}
if ($VARextensionAttribute12) { write-host -fore green "Enabled:  extensionAttribute12"; $arrayattributes += @("extensionAttribute12")} else { write-host -fore gray "Disabled: extensionAttribute12"}
if ($VARphysicalDeliveryOfficeName) { write-host -fore green "Enabled:  physicalDeliveryOfficeName"; $arrayattributes += @("physicalDeliveryOfficeName")} else { write-host -fore gray "Disabled: physicalDeliveryOfficeName"}
$nl
write-host -fore gray -back gray "------------------------------------"
$nl ; $nl
if ($doit) { write-host -fore red -back yellow "CHANGES WILL BE PROCESSED!" } else { write-host -fore cyan "No changes will be made!" }
$nl ; $nl
read-host "Press enter to continue"
$nl

$users = import-csv $importFile -delimiter ";" -Encoding Default | select -first 7

foreach ($user in $users) {
	$adusr = $NULL
	$adusr = get-aduser $user.samaccountname -property *
	write-host -fore gray -back gray "------------------------------------"
	$nl
	write-host -fore cyan "$($adusr.samaccountname):"
    
    foreach ($attribute in $arrayattributes){
        if ($attribute) {
            Set-ADattribute ($attribute)
	    }
    }

	$nl
}

write-host -fore gray -back gray "------------------------------------"
$nl
