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

# Start


$code = {
    # Create an Image object
    $photo = [System.Drawing.Image]::FromFile($_.FullName)

    # Read out the date taken (string)
    $dateProperty = ReadAttribute(0x9003)
    $dateTaken = ConvertToString($dateProperty.Value)
    
    $timestamp = Get-Date ($dateTaken.substring(0,10).replace(":","/") + " " + $dateTaken.substring(11))

    $extension = $_.Extension
    $counter = 0
    while ($true) {
        $filename = '{0:yyyyMMdd}' -f $timestamp + "-" + $_.BaseName + $( if ($counter) { "-$counter" } else { '' }) + $extension
        $filepath = Join-Path (Split-Path $_.FullName) $filename
        if (Test-Path $filepath) {
            $counter++
        }
        else {
            $filename
            break
        }
    }
    # Dispose of the Image once we're done using it
    $photo.Dispose()
}


# Load the System.Drawing DLL before doing any operations
[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") > $null
# And System.Text if reading any of the string fields
[System.Reflection.Assembly]::LoadWithPartialName("System.Text") > $null

If ($args[0] -eq $null) {
    Write-Host "Usage: Rename-Exif [path]"
    Exit
}
Else {
    If ((Test-Path $args[0]) -ne $true) {
        Write-Host "Path not found"
        exit
    }
}

dir $args[0] | Rename-Item -NewName $code
