<# 
.SYNOPSIS 
Apply ACL correction for MS16-072.

.DESCRIPTION 
Check if 'Authenticated Users' group have 'Read' permission on every GPOs.

.PARAMETER $Whatif 
If present no changes are written to GPOs (check only)

.PARAMETER $domain 
Specify the domain where perform the operation. If not present current domain is taken.

.PARAMETER $checkDomainComputer
If present check if 'Domain Computer' group have 'Read' permission and do not add 'Authenticated Users' if true

.PARAMETER $checkPerUserSetting
If present check if 'Per-User setting' is enabled on GPO and add 'Authenticated Users' only if true

.EXAMPLE 
.\Set-GPOACLMS16-072.ps1 -whatif
Check if 'Read' right is given to 'Authenticated Users' on every GPOs, without touch nothing on current domain

.EXAMPLE
.\Set-GPOACLMS16-072.ps1 -domain mydomain.local -checkDomainComputer
Add 'Authenticated Users' with 'Read' permission on all mydomain.local GPOs without touch GPOs that alredy have 'Domain Computer' group with 'Read' permission

.NOTES 
MS16-072 changes the security context with which user group policies are retrieved.
This by-design behavior change protects customers’ computers from a security vulnerability.
Before MS16-072 is installed, user group policies were retrieved by using the user’s security context.
After MS16-072 is installed, user group policies are retrieved by using the computer's security context.

https://support.microsoft.com/en-us/kb/3163622
https://sdmsoftware.com/group-policy-blog/bugs/new-group-policy-patch-ms16-072-breaks-gp-processing-behavior/
https://blogs.technet.microsoft.com/poshchap/2016/06/16/ms16-072-known-issue-use-powershell-to-check-gpos/
#>

Param( 
     [string]$domain=$env:USERDNSDOMAIN, 
     [switch]$whatif,
     [switch]$checkDomainComputer,
     [switch]$checkPerUserSetting
     )

Import-Module GroupPolicy

#$domain = "mydomain.local"
#$checkDomainComputer = $false
#$checkPerUserSetting = $false
#$whatIf = $true

Function Add-AuthUserGPOPermission($GPOSelected, [String]$GPODomain, [Boolean]$doIt){
    Trap {Write-Host "└-- ERROR: $($GPOSelected.DisplayName) - $($Err.Message)" -ForegroundColor Red; Continue}

    If ($doIt){
        Write-Host "└-- WHAT IF: $($GPOSelected.DisplayName) will be modified to grant ‘Authenticated Users’ read access" -ForegroundColor Green
    } Else {
        Set-GPPermissions -Guid $GPOSelected.Id -PermissionLevel GpoRead -TargetName "Authenticated Users" -TargetType Group -DomainName $GPODomain -ErrorVariable Err -ErrorAction SilentlyContinue | Out-Null
        Write-Host "└-- OPERATION: $($GPOSelected.DisplayName) has been modified to grant ‘Authenticated Users’ read access" -ForegroundColor Green
    }
}

Try{
    # Summary of settings
    Write-Host " "
    Write-Host "Selected Domain                         : $domain"
    Write-Host "Check for 'Domain Computers' permission : $checkDomainComputer"
    Write-Host "Check for 'Per-User Setting' GPO        : $checkPerUserSetting"
    Write-Host "----------------------------------------------------------------"
    Write-Host " "
    If ($whatIf) {
        Write-Host "Whatif flag present. Any changes will be performed." -ForegroundColor Green
    } Else {
        Write-Host "Whatif flag not present. Changes will be performed !" -ForegroundColor Red
    }
    Write-Host " "
    Write-Host "----------------------------------------------------------------"
    read-host "Press enter to continue"

    #Get all GPOs in current domain
    $GPOs = Get-GPO -All -domain $domain

    #Check we have GPOs
    if ($GPOs) {
        #Loop through GPOs
        foreach ($GPO in $GPOs) {

            #Nullify $AuthUser & $DomComp
            $AuthUser = $null
            $DomComp = $null
    
            #See if we have an Auth Users perm
            $AuthUser = Get-GPPermissions -Guid $GPO.Id -TargetName “Authenticated Users” -TargetType Group -ErrorAction SilentlyContinue -DomainName $domain
            #See if we have the Domain Computers perm
            $DomComp = Get-GPPermissions -Guid $GPO.Id -TargetName “Domain Computers” -TargetType Group -ErrorAction SilentlyContinue -DomainName $domain

            #Alert if we don’t have an ‘Authenticated Users’ permission
            if (-not $AuthUser) {
                #Now check for ‘Domain Computers’ permission
                if (-not $DomComp) {
                    Write-Host “WARNING: $($GPO.DisplayName) ($($GPO.Id)) does not have an ‘Authenticated Users’ permission or ‘Domain Computers’ permission” # -ForegroundColor Red
                    if ($checkPerUserSetting){
                        if ($GPO.user.DSVersion -gt 0){
                            Add-AuthUserGPOPermission $GPO $domain $whatIf
                        }
                    } Else {
                            Add-AuthUserGPOPermission $GPO $domain $whatIf
                    }
                
                }
                else {
                    Write-Host “INFORMATION: $($GPO.DisplayName) ($($GPO.Id)) does not have an ‘Authenticated Users’ permission but does have a ‘Domain Computers’ permission” -ForegroundColor Yellow
                    If (-not $checkDomainComputer){
                        Add-AuthUserGPOPermission $GPO $domain $whatIf
                    }
                }
            }
            elseif (($AuthUser.Permission -ne “GpoApply”) -and ($AuthUser.Permission -ne “GpoRead”)) {
                #COMMENT OUT THE BELOW LINE TO REDUCE OUTPUT!
                Write-Host “INFORMATION: $($GPO.DisplayName) ($($GPO.Id)) has an ‘Authenticated Users’ permission that isn’t ‘GpoApply’ or ‘GpoRead'” -ForegroundColor Yellow
            }
            else {
                #COMMENT OUT THE BELOW LINE TO REDUCE OUTPUT!
                #Write-Output “INFORMATION: $($GPO.DisplayName) ($($GPO.Id)) has an ‘Authenticated Users’ permission”
            }

        }

    }

}
Catch [Exception]{
    Write-Host "Error:" $_.Exception.Message -ForegroundColor Red
    Write-Host "Script Terminated" -ForegroundColor Red
}
