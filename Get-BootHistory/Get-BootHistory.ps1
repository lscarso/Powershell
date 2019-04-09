<# 
.SYNOPSIS 
Get-BootHistory retrieves boot up information and powerup, shutdown, restart events from a Computer. 

.DESCRIPTION 
 
.PARAMETER ComputerName 
The Computer name to query. Default: Localhost. 

.EXAMPLE 
Gets the Boot History and uptime from Computer1
Get-BootHistory -ComputerName Computer1
 
.EXAMPLE 
Gets the Boot History from a list of computers in c:\Temp\Computerlist.txt. 
Get-BootHistory.ps1 -ComputerName (Get-Content C:\Temp\Computerlist.txt) 

#> 

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false,
        Position = 0,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
    [Alias("Name")]
    [string[]]$ComputerName = $env:COMPUTERNAME
)

begin { }
process {
    foreach ($Computer in $ComputerName) {
        try {
            #Need to verify that the hostname is valid in DNS
            $hostdns = [System.Net.DNS]::GetHostEntry($Computer)
            
            Write-Host "Connecting to $($hostdns.HostName)..."
            $OS = Get-WmiObject win32_operatingsystem -ComputerName $Computer -ErrorAction Stop
            $BootTime = $OS.ConvertToDateTime($OS.LastBootUpTime)
            $Uptime = $OS.ConvertToDateTime($OS.LocalDateTime) - $boottime
                
            #Uptime Table
            $propHash = @{
                ComputerName = $Computer
                BootTime     = $BootTime
                Uptime       = $Uptime
            }
            $objComputerUptime = New-Object PSOBject -Property $propHash
                            
            #Event 6009 or 1074 Table
            Write-Host "Collecting eventlog..."
            $objComputerEvents = Get-EventLog -logname System -ComputerName $Computer -Newest 3000 | Where-Object { ($_.EventID -eq 6009) -or ($_.EventID -eq 1074) } | Select-Object @{Name = "Time"; Expression = { $_.TimeGenerated } }, @{Name = "ComputerName"; Expression = { $_.MachineName } }, @{Name = "Reason"; Expression = { if ($_.ReplacementStrings[4] -eq 0) { "power on" } else { $_.ReplacementStrings[4] } } }, @{Name = "Message"; Expression = { $_.Message } }
            
            #Display results
            $objComputerUptime | Format-Table -Property ComputerName, BootTime, Uptime
            $objComputerEvents | Format-Table -AutoSize    
        }
        catch [Exception] {
            Write-Output "$computer $($_.Exception.Message)"
            #return
        }
    }
}
end { }
