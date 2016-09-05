<# 
.SYNOPSIS 
Check Active directory Replication status.

.DESCRIPTION 
Run REPADMIN to collect replication status and send report via email.

.PARAMETER from
email address from where you what to send report.

.PARAMETER to
email sddress where send report every time.

.PARAMETER toError
email address where send report if ERRORS are found.

.PARAMETER smtp
smtp server address

.EXAMPLE 
.\Get-ADReplicationStatus.ps1 -from monitoring@mydomain.local -to myself@mydomain.local -toError myself@mydomain.local, myhelpdesk@mydomain.local

.NOTES 
#>

Param( 
     [Parameter(Mandatory)]
     [string]$from, 
     [ValidateScript({($_ -as [System.Net.Mail.MailAddress]).Address -eq $_ -and $_ -ne $null})]
     [string]$to,
     [Parameter(Mandatory)]
     [ValidateScript({($_ -as [System.Net.Mail.MailAddress]).Address -eq $_ -and $_ -ne $null})]
     [string[]]$toError,
     [string[]]$smtp = "smtp.mydomain.local"
     )

$LargestDeltaTreshold = 60

$repadmin = repadmin /replsum

[regex]$regex = '\s+(?<DC>\S+)\s+(?<Delta>\S+)\s+(?<fail>\d{1,2}\s)'

[regex]$regex2 = '\s+(?<FAIL>\d{1,2}\S+)\s\-\s+(?<DC>\S+)'

$repadmin | ForEach-Object {
    if ( $_ -match $regex ) {
        $process = "" | Select-Object DC, Delta, fail
        $process.dc = $matches.dc
        $process.Delta = $matches.Delta
        $process.fail = [int]($matches.fail)
        $VdayTime = 0
        $VHourTime = 0
        $VMinutesTime = 0
        
        if ($process.Delta.contains("d"))  {
		  $VdayTime = [int]($process.Delta.substring(0,2))
		  $VHourTime = [int]($process.Delta.substring(4,2))
		  $VMinutesTime = [int]($process.Delta.substring(8,2))}
		Else {		
		  if ($process.Delta.contains("h"))  {
          $VHourTime = [int]($process.Delta.substring(0,2))
          $VMinutesTime = [int]($process.Delta.substring(4,2))}
          Else {
            if ($process.Delta.contains("m"))  {
                $VMinutesTime = [int]($process.Delta.substring(0,2))}
          }
	    }

        $DeltaMinutes = New-TimeSpan -Days $VdayTime -Hours $VHourTime -Minutes $VMinutesTime

        if (($DeltaMinutes.Minutes -gt $LargestDeltaTreshold) -or ($process.fail -gt 0)) {
            $errorCount = $errorCount + 1
        }

    }

    Elseif ( $_ -match $regex2 ) {
		$errorCount = $errorCount + 1
    }
}
If ($errorCount -ne $null -and $toError -ne $null) {
    $emailsubject = "Daily Forest Replication Status - Some Errors"
    Send-MailMessage -From $from -To $toError -Subject $emailsubject -SmtpServer $smtp -Body ($repadmin | Out-String)}
else {
    $emailsubject = "Daily Forest Replication Status - No Replication Error"
    if ($to -ne $null){
        Send-MailMessage -From $from -To $to -Subject $emailsubject -SmtpServer $smtp -Body ($repadmin | Out-String)
    }
}
