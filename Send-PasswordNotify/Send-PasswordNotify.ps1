<# 
.SYNOPSIS 
Send email to users with approacing password expiration.

.DESCRIPTION 
Load users from selected group and notify via email if the expiration day is less that 7 day and write an eventlog 3001 when sucessfully sent.

.PARAMETER notificationGroup 
group that contain users that what to be notify for password expiration

.PARAMETER emailBodyFile 
Text that users receive via email. This is an optional parameter; if not included, mailbody.txt will be used

.PARAMETER logFile 
if $True write a logfile. Default is $False

.PARAMETER eventLog 
if $True write to eventlog. Default is $True. Check .NOTES for more info

.EXAMPLE 
Read username and email address from group "MyUsers" and send password expiration notification. 
AD_PW_ChangeNotification.ps1 - notificationGroup "MyUsers"

.EXAMPLE 
Read computer names from a file (one name per line) and retrieve their inventory information 
Get-Content c:\names.txt | Get-Inventory.

.NOTES 
Run this to create eventlog and register source:
New-EventLog -LogName "Scripts" -Source Send-PasswordNotify
#>

Param(
    [Parameter(Mandatory=$true)][string]$NotificationGroup,
    [String]$emailbodyFile = "mailbody.txt",
    [Boolean]$logFile = $false,
    [Boolean]$eventLog = $true
    )
# Import Active Directory Module
import-module activedirectory

# Define parameters.
$DaysBeforePWExpires="-7"
$emailbody = get-content $emailbodyFile
$emailsubject = "Password Expiry Notice"
$emailfrom = "helpdesk@yourdomain.com"
$smtpserver = "smtp.yourdomain.com"
# eventlog name where write events
$logName = "Scripts"
# log file name
$logFileName = ".\Send-PasswordNotify.log"

# function to get Password-Duration
function GetPasswordDuration{
    $ThisDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
    $DirectoryRoot = $ThisDomain.GetDirectoryEntry()
 
    $DirectorySearcher = [System.DirectoryServices.DirectorySearcher]$DirectoryRoot
    $DirectorySearcher.Filter = "(objectClass=domainDNS)"
    $DirectorySearchResult = $DirectorySearcher.FindOne()
 
    $MaxPasswordAge = New-Object System.TimeSpan([System.Math]::ABS($DirectorySearchResult.properties["maxpwdage"][0]))
 
    return $MaxPasswordAge.TotalDays
}


Try {
    # Get CurrentDate
    $today = Get-Date
    $logdate = get-date -format yyy-MM-dd

    # load Users
    $Users = Get-ADGroupMember $NotificationGroup | Get-ADUser -Properties *

    # Get Password-Duration
    $PasswordDuration = GetPasswordDuration

    # Write eventLog
    if ($eventLog){
        Write-EventLog -LogName $logName -Source AD_PW_ChangeNotification -EventId 3002 -EntryType Information -Message "Script Started" 
    }
    
    # Write logFile
    If ($logFile){
        $logdateexact = get-date -format "yyy-MM-dd (HH:mm)"
        $logline = "$($logdateexact): Script Started"
        $logline | out-file $logFileName -Append
    }

    foreach ($user in $users){
        
        # Get Expire-Date
        $ExpireDate = $user.passwordlastset.AddDays($PasswordDuration)
        
        # Get Warning-Date
        $WarningDate = $ExpireDate.AddDays($DaysBeforePWExpires)

        # Send E-Mail notification and add to logfile-line ift today = warningdate
        if ($today.Date -eq $WarningDate.Date){
            $emailsubject = "Password Expiry Notice - your password expires at $ExpireDate"
            Send-MailMessage -To $user.mail -Subject $emailsubject -From $emailfrom -Body $emailbody -bodyashtml -SmtpServer $smtpserver
            
            #write eventlog
            if ($eventLog){
                Write-EventLog -LogName $logName -Source Send-PasswordNotify -EventId 3001 -EntryType Information -Message "Mail sent to user $($user.SamAccountName) with this address: $($user.mail)" 
            }
            # write logfile
            if ($logFile){
                $logline = "$($logdate): mail sent to $($user.SamAccountName);$($user.mail)"
                $logline | out-file $logFileName -Append
            }
        }
    }
}
catch [Exception]{
    if ($eventLog){
        Write-EventLog -LogName $logName -Source Send-PasswordNotify -EventId 1001 -EntryType Error -Message "Error: $_.Exception.Message"
    }
    if ($logFile){
        $logline = "$($logdate): Error: $_.Exception.Message"
        $logline | out-file $logFileName -Append
    }
}
