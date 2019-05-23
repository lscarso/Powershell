<# 
.SYNOPSIS 
Add passed accounts to Exchange coManaged attribute of selected groups.

.DESCRIPTION 
Add passed accounts to msExchCoManagedByLink attribute.

.PARAMETER Whatif 
If present no changes are written to AD (check only).

.PARAMETER ACLGroup
Groups where add coManaged accounts. You can use wildcard to selct more that one groups. You can pass only one string per time.

.PARAMETER AccountsToAdd
List of accounts to add to coManged attribute.

.EXAMPLE 
Test to add User1 and User2 to coManaged attributes of "TestGroup01" and "TestGroup02"
.\Set-coManager.ps1 -ACLGroup "TestGroup*" -AccountsToAdd "user1", "User2" -Whatif

.EXAMPLE 
Add User1 and User2 to coManaged attributes of "TestGroup01"
.\Set-coManager.ps1 -ACLGroup "TestGroup01" -AccountsToAdd "user1", "User2"

.NOTES 
#>


[CmdletBinding()]
Param (
    [Parameter(Mandatory=$True)]
    [String]$ACLGroup,
    [String[]]$AccountsToAdd = @("DefautAdmin01","DefaultAdmin02"),
    [switch]$whatif
)

foreach ($Account in $AccountsToAdd) {
    $ADdn = (Get-ADUser $Account).DistinguishedName
    Write-Output "Adding $Account to $ACLGroup"
    if ($whatif) {
        Get-ADObject -filter {name -like $ACLGroup} | set-adobject -Add @{msExchCoManagedByLink=$ADdn} -Verbose -WhatIf        
    } Else {
        Get-ADObject -filter {name -like $ACLGroup} | set-adobject -Add @{msExchCoManagedByLink=$ADdn} -Verbose
    }
}
