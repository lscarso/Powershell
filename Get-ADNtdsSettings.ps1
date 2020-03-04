<# 
.SYNOPSIS 
retrive NTDS settings for Active Directory replica
.DESCRIPTION 
retrive NTDS settings for Active Directory replica with site, server and DNS alias making easier replica troubleshooting.
#> 

$Config = (Get-ADRootDSE).configurationNamingContext
$Servers = Get-ADObject -Filter {ObjectClass -eq "Server"} -SearchBase "CN=Sites,$Config" -SearchScope Subtree
foreach ($Server in $Servers){
    $Ntdsa = Get-ADObject -Filter {ObjectClass -eq "nTDSDSA"} -SearchBase "$(($Server).DistinguishedName)" -SearchScope Subtree

    $Ntdsa
}
