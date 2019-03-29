Function Copy-WithProgress
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Source,
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Destination,
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=0)]
        [switch]$Silent,
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=0)]
        [switch]$Force
       
    )
$TotalCopied = 0
$TotalSkipped = 0

$Source=$Source.tolower()

$Filelist=get-childitem -path $source -Recurse
$Total=$Filelist.count
$Position=0
    foreach ($File in $Filelist)
    { 
        $Filename=$File.Fullname.tolower().replace($Source,'') 
        $DestinationFile=($Destination+$Filename).replace('\\','\')
        Write-Progress -Activity "Copying data from $source to $Destination" -Status "Copying Files: $($File.FullName)" -PercentComplete (($Position/$total)*100)
        if ($File -is [System.IO.DirectoryInfo] -and $(Test-path -path $DestinationFile))
        {
            # Folder already present     
        }
        Else
        {       
            if (Test-Path -Path $DestinationFile){
                #if ($(Get-ItemProperty -Path $File.FullName).LastWriteTime -eq $(Get-ItemProperty -Path $DestinationFile).LastWriteTime) {
                #if ($(Compare-Object $File $(Get-ItemProperty -Path $DestinationFile) -Property LastWriteTime -PassThru) -eq $null) {
                if ($File.LastWriteTime -eq $(Get-ItemProperty -Path $DestinationFile).LastWriteTime) {
                    # File alredy present with the same LastWriteTime
                    Write-Verbose "SKIPPED: $($File.Fullname)"
                    $TotalSkipped++
                }
                Else
                {
                    if ($Force)
                    {
                        Copy-Item -path $File.FullName -Destination $DestinationFile -Force
                    }
                    Else
                    {
                        Copy-Item -path $File.FullName -Destination $DestinationFile
                    }
                    Write-Verbose "$($File.Fullname) => $DestinationFile"
                    $TotalCopied++
                }
            }
            Else 
            {
                if ($Force)
                {
                    Copy-Item -path $File.FullName -Destination $DestinationFile -Force
                }
                Else
                {
                    Copy-Item -path $File.FullName -Destination $DestinationFile
                }
                Write-Verbose "$($File.Fullname) => $DestinationFile"
                $TotalCopied++
            }
        }
            
        $Position++
    }
if (!$Silent)
{
    Write-Output ""
    Write-Output "Total Files: $Total"
    Write-Output "Total Files Copied: $TotalCopied"
    Write-Output "Total Files Skipped: $TotalSkipped"
}

}
