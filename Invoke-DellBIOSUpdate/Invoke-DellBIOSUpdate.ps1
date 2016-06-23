#Key
$KeyFile = "\\computer01\BIOS\Settings.bin"
# $Key = New-Object Byte[] 16   # You can use 16, 24, or 32 for AES
# [Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($Key)
# $Key | out-file $KeyFile

#Pwd
$PasswordFile = "\\computer01\BIOS\Data.bin"
$Key = Get-Content $KeyFile
# read-host -assecurestring | ConvertFrom-SecureString -key $Key | Out-File $PasswordFile




$ScriptFolder = Split-Path -Parent $MyInvocation.MyCommand.Definition
$Model = $((Get-WmiObject -Class Win32_ComputerSystem).Model).Trim()
$BIOSVersion = ((Get-WMIObject -Class Win32_BIOS).SMBIOSBIOSVersion)

if(Test-Path -Path $ScriptFolder\$model){
    Write-Host "Model repository Found!"
    $BIOSUpdateFile = Get-ChildItem -Path $ScriptFolder\$Model
    $BIOSUpdateFileVersion = $BIOSUpdateFile.ToString() -replace ($BIOSUpdateFile.Extension,"")
    
    #Get the actual BIOS Update File Version from the file name
    switch ($Model) {
	    "OptiPlex 7040" {$BIOSUpdateFileVersion = $BIOSUpdateFileVersion.Substring($BIOSUpdateFileVersion.Length -5)}
		default {$BIOSUpdateFileVersion = $BIOSUpdateFileVersion.Substring($BIOSUpdateFileVersion.Length -3)}
	}
    #$BIOSUpdateFileVersion = $BIOSUpdateFileVersion.Substring($BIOSUpdateFileVersion.Length -3)                   

    Write-Host "Current BIOS version   : $BIOSVersion"
    Write-Host "Repository BIOS version: $BIOSUpdateFileVersion`n"

    #Compare the BIOS File Version with the currently installed version to determine next steps
	switch ($BIOSVersion.CompareTo($BIOSUpdateFileVersion)) {
	    0 {
		    Write-Host "BIOS Version is up to date`n"
			break
		}
		1 {
			Write-Host "BIOS Version is newer than supplied BIOS version`n"
			break
		}
		default {
		    Write-Output "BIOS Update Needed. Attempting BIOS Flash Operation..."
            Write-Output "Loading Admin Credential..."
            
            $username = "service01"
            #$password = cat $KeyFile | convertto-securestring
            #$Cred = new-object -typename System.Management.Automation.PSCredential -argumentlist ($username, $password)
            $Cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, (Get-Content $PasswordFile | ConvertTo-SecureString -Key $key)

            #Invoke-Expression $ScriptFolder\$Model\$BIOSUpdateFile " /quiet"
            $objStartInfo = New-Object System.Diagnostics.ProcessStartInfo
            $objStartInfo.FileName = "$ScriptFolder\$Model\$BIOSUpdateFile"
            $objStartInfo.WorkingDirectory = "$ScriptFolder\$Model\"
            $objStartInfo.Arguments = "-noreboot -nopause -forceit"
            #$objStartInfo.CreateNoWindow = $true
            $objStartinfo.verb = “RunAs”
            $objStartInfo.UserName = $Cred.Username
            $objStartInfo.Password = $Cred.Password
            $objStartInfo.UseShellExecute = $false
            [System.Diagnostics.Process]::Start($objStartInfo) #| Out-Null
			
            break	
		}
    }



#    if($BIOSVersion.CompareTo($BIOSUpdateFileVersion) -eq 0){
#        Write-Output "BIOS Version is up to date"
#    }else{
#        Try{
#            Write-Output "BIOS Update Needed. Attempting BIOS Flash Operation..."
#            Write-Output "Loading Admin Credential..."
#            
#            $username = "srvc8011@d400.mh.grp"
#            #$password = cat $KeyFile | convertto-securestring
#            #$Cred = new-object -typename System.Management.Automation.PSCredential -argumentlist ($username, $password)
#            $Cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, (Get-Content $PasswordFile | ConvertTo-SecureString -Key $key)
#
#            #Invoke-Expression $ScriptFolder\$Model\$BIOSUpdateFile " /quiet"
#            $objStartInfo = New-Object System.Diagnostics.ProcessStartInfo
#            $objStartInfo.FileName = "$ScriptFolder\$Model\$BIOSUpdateFile"
#            $objStartInfo.WorkingDirectory = "$ScriptFolder\$Model\"
#            $objStartInfo.Arguments = "-noreboot -nopause -forceit"
#            #$objStartInfo.CreateNoWindow = $true
#            $objStartinfo.verb = “RunAs”
#            $objStartInfo.UserName = $Cred.Username
#            $objStartInfo.Password = $Cred.Password
#            $objStartInfo.UseShellExecute = $false
#            [System.Diagnostics.Process]::Start($objStartInfo) #| Out-Null
#        }
#        Catch{[Exception]
#            Write-Output "Failed: $_"
#        }
#    }            

    Write-Output "End Dell BIOS Update Operation"
}
else
{
    Write-Output "Model Not Supported"
}
