function ReadAttribute {
    $DWGAttr = $Null
    Try{$DWGAttr = $photo.GetPropertyItem($args[0])}
    Catch{$DWGAttr = $Null;}
    Finally{Write-Output $DWGAttr}
}

function ConvertToString {
    Try{$DWGstr = (new-object System.Text.UTF8Encoding).GetString($args[0])}
    Catch{$DWGstr = $null;}
    Finally{Write-Output $DWGstr}
}

function ConvertToNumber {
    $First =$args[0].value[0] + 256 * $args[0].value[1] + 65536 * $args[0].value[2] + 16777216 * $args[0].value[3]
    $Second = $args[0].value[4] + 256 * $args[0].value[5] + 65536 * $args[0].value[6] + 16777216 * $args[0].value[7]
    if ($first -gt 2147483648) {$first = $first  - 4294967296}
    if ($Second -gt 2147483648) {$Second= $Second - 4294967296}
    if ($Second -eq 0) {$Second= 1}
    if (($first -eq 1) -and ($Second -ne 1)) {
        write-output ("1/" + $Second)
    } 
    else {write-output ($first / $second)}
}

function Get-Exif {
    <# 
    .SYNOPSIS 
    
    .DESCRIPTION 
    
    .PARAMETER
    
    .EXAMPLE 
    
    .NOTES
    #>

    Param(
        [Parameter(Mandatory=$True,
        ValueFromPipelineByPropertyName=$True)]
        [string]$path
        ) 

    # Load the System.Drawing DLL before doing any operations
    [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") > $null
    # And System.Text if reading any of the string fields
    [System.Reflection.Assembly]::LoadWithPartialName("System.Text") > $null

    # Create ExifData Collection
    $colExifData = @()

    # Create an Image object
    $photo = [System.Drawing.Image]::FromFile($path)

    # Read out the date taken (string)
    $colExifData = New-Object System.Object
    $colExifData | Add-Member -type ExifProperty -Name DateTaken -Value (ConvertToString(ReadAttribute(0x9003).Value))

    # ISO (unsigned short integer)
    $isoProperty = ReadAttribute(0x8827)
    if ($isoProperty -eq $null){
        $iso = $null
    }
    Else {
        $iso = [System.BitConverter]::ToUInt16($isoProperty.Value, 0)
    }
    $colExifData | Add-Member -Type ExifProperty -Name Iso -Value $iso
    


    }