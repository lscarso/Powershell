<# 
.SYNOPSIS 
Check Local Group Administrators and export the members.

.DESCRIPTION 
Check Local Group Administrators and export the members on a file stored on a share if members are different from $DefaultAdmins user list

.EXAMPLE 
Get-LocalAdministratorsMember.ps1

.NOTES 
Update $DefaultAdmins array with your default members
Update $rpath with the path where save the result
#>

#Settings ############################################
#Add Default Admins Groups
$DefaultAdmins = @("Domain Admins", "Administrator")

#Add reports path
$rpath = "\\server\share$\_LocalAdmin_Script"
######################################################

If (Test-Path $rpath){
    Add-Type -AssemblyName System.DirectoryServices.AccountManagement
    $ctype = [System.DirectoryServices.AccountManagement.ContextType]::Machine
    $context = New-Object -TypeName System.DirectoryServices.AccountManagement.PrincipalContext -ArgumentList $ctype, $env:COMPUTERNAME
    $idtype = [System.DirectoryServices.AccountManagement.IdentityType]::SamAccountName
    $group = [System.DirectoryServices.AccountManagement.GroupPrincipal]::FindByIdentity($context, $idtype, 'Administrators')

    $results = $group.Members |
        select @{N='Domain'; E={$_.Context.Name}}, samaccountName |
        where {$AdminValue = $_.SamAccountName; -not @($DefaultAdmins | ? {$AdminValue -eq $_})}

    If ($results -ne $null){
        $Results | Out-File $rpath\$env:COMPUTERNAME.txt}
}
Else {
    Write-Warning "Path Not Found: $rpath"
}
