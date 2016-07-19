<# 
.SYNOPSIS 
Get EXIF attributes from image

.DESCRIPTION 
Display EXIF attributes of an image 

.PARAMETER $Path
Path of image

.EXAMPLE
Get-Exif.ps1 C:\MyPictures\Image01.jpg
#>

#$strtime = (makestring $foo.GetPropertyItem(36867))

#$DataScatto = ([datetime]::ParseExact($strtime.substring(0,10).replace(":","/"),"yyyy/MM/dd",$null)).ToShortDateString()
#$OraScatto = ([datetime]$strtime.substring(11))

#Write-Host $strtime.substring(0,10).replace(":","/")
#Write-Host $strtime.substring(11)
#Write-Host "New File= $DataScatto"
#Write-Host "New File= $OraScatto"



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

function ConvertToNumber {$First =$args[0].value[0] + 256 * $args[0].value[1] + 65536 * $args[0].value[2] + 16777216 * $args[0].value[3] ;$Second=$args[0].value[4] + 256 * $args[0].value[5] + 65536 * $args[0].value[6] + 16777216 * $args[0].value[7] ; 
if ($first -gt 2147483648) {$first = $first  - 4294967296} ;if ($Second -gt 2147483648) {$Second= $Second - 4294967296} ; if ($Second -eq 0) {$Second= 1} ; 
if (($first –eq 1) -and ($Second -ne 1)) {write-output ("1/" + $Second)} else {write-output ($first / $second)}}


If ($args[0] -eq $null) {
    Write-Host "Usage: Get-Exif [image path]"
    Exit
}
Else {
    If ((Test-Path $args[0]) -ne $true) {
        Write-Host "File not found"
        exit
    }
}

# Load the System.Drawing DLL before doing any operations
[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") > $null
# And System.Text if reading any of the string fields
[System.Reflection.Assembly]::LoadWithPartialName("System.Text") > $null

$filename = $args[0]

# Create an Image object
$photo = [System.Drawing.Image]::FromFile($filename)

# Read out the date taken (string)
$dateProperty = ReadAttribute(0x9003)
$dateTaken = ConvertToString($dateProperty.Value)

# ISO (unsigned short integer)
$isoProperty = ReadAttribute(0x8827)
if ($isoProperty -eq $null){
    $iso = $null
}
Else {
    $iso = [System.BitConverter]::ToUInt16($isoProperty.Value, 0)
}

# Title
$TitleProperty = ReadAttribute(0x010e)
$Title = ConvertToString($TitleProperty.Value)

# Author
$AuthorProperty = ReadAttribute(0x013b)
$Author = ConvertToString($AuthorProperty.Value)

# Maker
$makerProperty = ReadAttribute(0x010f)
$maker = ConvertToString($makerProperty.Value)

# Model
$modelProperty = ReadAttribute(0x0110)
$model = ConvertToString($modelProperty.Value)

# Orientation
$orientationProperty = ReadAttribute(0x0112)
if ($orientationProperty -eq $null){
    $orientation = $null
}
Else {
    $orientation = [System.BitConverter]::ToUInt16($orientationProperty.Value, 0)
}

# Width resolution
$xResProperty = ReadAttribute(0x011a)
if ($xResProperty -eq $null){
    $xRes = $null
}
Else {
    $xRes = [System.BitConverter]::ToUInt16($xResProperty.Value, 0)
}

# Height resolution
$yResProperty = ReadAttribute(0x011b)
if ($yResProperty -eq $null){
    $yRes = $null
}
Else {
    $yRes = [System.BitConverter]::ToUInt16($yResProperty.Value, 0)
}

# Resolution unit
$resUnitProperty = ReadAttribute(0x0128)
if ($resUnitProperty -eq $null){
    $resUnit = $null
}
Else {
    $resUnit = [System.BitConverter]::ToUInt16($resUnitProperty.Value, 0)
}

# Exposure time
$exposureTimeProperty = ReadAttribute(0x829a)
if ($exposureTimeProperty -eq $null){
    $exposureTime = $null
}
Else {
    $exposureTime = ConvertToNumber($exposureTimeProperty)
}

# F-Number
$fNumberProperty = ReadAttribute(0x829d)
if ($fNumberProperty -eq $null){
    $fNumber = $null
}
Else {
    $fNumber = ConvertToNumber($fNumberProperty)
}

# Exposure compensation
$expCompProperty = ReadAttribute(0x9204)
if ($expCompProperty -eq $null){
    $expComp = $null
}
Else {
    $expComp = ConvertToNumber($expCompProperty)
}

# Metering mode
$meteringProperty = ReadAttribute(0x9207)
if ($meteringProperty -eq $null){
    $metering = $null
}
Else {
    $metering = [System.BitConverter]::ToUInt16($meteringProperty.Value, 0)
}

# Flash mode
$flashProperty = ReadAttribute(0x9209)
if ($flashProperty -eq $null){
    $flash = $null
}
Else {
    $flash = [System.BitConverter]::ToUInt16($flashProperty.Value, 0)
}

# Focal lenght
$focalProperty = ReadAttribute(0x920a)
if ($focalProperty -eq $null){
    $focal = $null
}
Else {
    $focal = ConvertToNumber($focalProperty)
}

# Color space
$colorProperty = ReadAttribute(0xa001)
if ($colorProperty -eq $null){
    $color = $null
}
Else {
    $color = [System.BitConverter]::ToUInt16($colorProperty.Value, 0)
}

# Width
$xPixelProperty = ReadAttribute(0xa002)
if ($xPixelProperty -eq $null){
    $xPixel = $null
}
Else {
    $xPixel = [System.BitConverter]::ToUInt16($xPixelProperty.Value, 0)
}

# Height
$yPixelProperty = ReadAttribute(0xa003)
if ($yPixelProperty -eq $null){
    $yPixel = $null
}
Else {
    $yPixel = [System.BitConverter]::ToUInt16($yPixelProperty.Value, 0)
}

# Source
$sourceFileProperty = ReadAttribute(0xa300)
if ($sourceFileProperty -eq $null){
    $sourceFile = $null
}
Else {
    $sourceFile = $sourceFileProperty.Value
}

# Exposure Mode
$expModeProperty = ReadAttribute(0xa402)
if ($expModeProperty -eq $null){
    $expMode = $null
}
Else {
    $expMode = [System.BitConverter]::ToUInt16($expModeProperty.Value, 0)
}

# White Balance
$whiteBalanceProperty = ReadAttribute(0xa403)
if ($whiteBalanceProperty -eq $null){
    $whiteBalance = $null
}
Else {
    $whiteBalance = [System.BitConverter]::ToUInt16($whiteBalanceProperty.Value, 0)
}

# Gain control
$gainCtrProperty = ReadAttribute(0xa407)
if ($gainCtrProperty -eq $null){
    $gainCtr = $null
}
Else {
    $gainCtr = [System.BitConverter]::ToUInt16($gainCtrProperty.Value, 0)
}

# Contrast
$contrastProperty = ReadAttribute(0xa408)
if ($contrastProperty -eq $null){
    $contrast = $null
}
Else {
    $contrast = [System.BitConverter]::ToUInt16($contrastProperty.Value, 0)
}

# Saturation
$saturationProperty = ReadAttribute(0xa409)
if ($saturationProperty -eq $null){
    $saturation = $null
}
Else {
    $saturation = [System.BitConverter]::ToUInt16($saturationProperty.Value, 0)
}

# Sharpness
$sharpnessProperty = ReadAttribute(0xa40a)
if ($sharpnessProperty -eq $null){
    $sharpness = $null
}
Else {
    $sharpness = [System.BitConverter]::ToUInt16($sharpnessProperty.Value, 0)
}

# Subject distance mode
$subjectDistProperty = ReadAttribute(0xa40c)
if ($subjectDistProperty -eq $null){
    $subjectDist = $null
}
Else {
    $subjectDist = [System.BitConverter]::ToUInt16($subjectDistProperty.Value, 0)
}

# Exposure program
$ExpProgProperty = ReadAttribute(0x8822)
if ($ExpProgProperty -eq $null){
    $ExpProg = $null
}
Else {
    $ExpProg = [System.BitConverter]::ToUInt16($ExpProgProperty.Value, 0)
}

# Subject distance
$SubjDistProperty = ReadAttribute(0x9206)
if ($SubjDistProperty -eq $null){
    $SubjDist = $null
}
Else {
    $SubjDist = [System.BitConverter]::ToUInt16($SubjDistProperty.Value, 0)
}

# Light source
$LightSourceProperty = ReadAttribute(0x9208)
if ($LightSourceProperty -eq $null){
    $LightSource = $null
}
Else {
    $LightSource = [System.BitConverter]::ToUInt16($LightSourceProperty.Value, 0)
}

# Scene type
$SceneTypeProperty = ReadAttribute(0xa407)
if ($SceneTypeProperty -eq $null){
    $SceneType = $null
}
Else {
    $SceneType = [System.BitConverter]::ToUInt16($SceneTypeProperty.Value, 0)
}

# Focal Lenght 35mm eq
$Focal35Property = ReadAttribute(0xa405)
if ($Focal35Property -eq $null){
    $Focal35 = $null
}
Else {
    $Focal35 = [System.BitConverter]::ToUInt16($Focal35Property.Value, 0)
}

# Brightness
$BrightnessProperty = ReadAttribute(0x9203)
if ($BrightnessProperty -eq $null){
    $Brightness = $null
}
Else {
    $Brightness = ConvertToNumber($BrightnessProperty)
}

# Lens maker
$LensMakerProperty = ReadAttribute(0xa433)
$LensMaker = ConvertToString($LensMakerProperty.Value)

# Lens model
$LensModelProperty = ReadAttribute(0xa434)
$LensModel = ConvertToString($LensModelProperty.Value)

# Dispose of the Image once we're done using it
$photo.Dispose()


# Display attibute
Write-Host "Image ------------------------------------------" -foregroundcolor DarkGreen
Write-Host "Pixel X Dimension = " $xPixel
Write-Host "Pixel Y Dimension = " $yPixel
Write-Host "X Resolution = " $xRes "dpi"
Write-Host "Y Resolution = " $yRes "dpi"
#Switch ($resUnit){
#    1 {"Resolution Unit = None"}
#    2 {"Resolution Unit = Inches"}
#    3 {"Resolution Unit = Centimeters"}
#}

Switch ($color){
    1 {"Color Space = sRGB"}
    2 {"Color Space = Adobe RGB"}
    default {"Color Space = "}
}

Write-Host "Data = " $dateTaken
Write-Host "Title = " $Title
Write-Host "Author = " $Author

Switch ($sourceFile){
    1 {"File Source = Film Scanner"}
    2 {"File Source = Reflection print Scanner"}
    3 {"File Source = Digital Camera"}
    default {"File Source = "}
}

Write-Host "Camera------------------------------------------" -foregroundcolor DarkGreen
Write-Host "Maker = " $maker
Write-Host "Model = " $model
Write-Host "Lens Maker = " $LensMaker
Write-Host "Lens Model = " $LensModel

If ($fNumberProperty -eq $null){
    Write-Host "F-Number = "}
Else {
    Write-Host ("F-Number = f/" + ("{0:N1}" -f $fNumber))
}

If ($exposureTimeProperty -eq $null){
    Write-Host "Shutter Speed = "}
Else {
    Write-Host "Shutter Speed = " $exposureTime "Sec."
}

If ($iso -eq 0){
    Write-Host "ISO = "}
Else {
    Write-Host "ISO = " $iso
}

Write-Host "Focal Lenght = " ("{0:N0}" -f $focal) "mm"
Write-Host "Focal Lenght 35mm = " $Focal35

Switch ($subjectDist){
    0 {"Subject Distance Range = Unknown"}
    1 {"Subject Distance Range = Macro"}
    2 {"Subject Distance Range = Close"}
    3 {"Subject Distance Range = Distant"}
    default  {"Subject Distance Range = "}
}

Write-Host "Subject Distance = " $SubjDist

$hexflash = "{0:X0}" -f $flash
switch ($hexflash){
    0 {"Flash = No Flash"}
    1 {"Flash = Fired"}
    5 {"Flash = Fired, Return not detected"}
    7 {"Flash = Fired, Return detected"}
    8 {"Flash = On, Did not fire"}
    9 {"Flash = On, Fired"}
    D {"Flash = On, Return not detected"}
    F {"Flash = On, Return detected"}
    10 {"Flash = Off, Did not fire"}
    14 {"Flash = Off, Did not fire, Return not detected"}
    18 {"Flash = Auto, Did not fire"}
    19 {"Flash = Auto, Fired"}
    1D {"Flash = Auto, Fired, Return not detected"}
    1F {"Flash = Auto, Fired, Return detected"}
    20 {"Flash = No flash function"}
    30 {"Flash = Off, No flash function"}
    41 {"Flash = Fired, Red-eye reduction"}
    45 {"Flash = Fired, Red-eye reduction, Return not detected"}
    47 {"Flash = Fired, Red-eye reduction, Return detected"}
    49 {"Flash = On, Red-eye reduction"}
    4D {"Flash = On, Red-eye reduction, Return not detected"}
    4F {"Flash = On, Red-eye reduction, Return detected"}
    50 {"Flash = Off, Red-eye reduction"}
    58 {"Flash = Auto, Did not fire, Red-eye reduction"}
    59 {"Flash = Auto, Fired, Red-eye reduction"}
    5D {"Flash = Auto, Fired, Red-eye reduction, Return not detected"}
    5F {"Flash = Auto, Fired, Red-eye reduction, Return detected"}
    default {"Flash = "}
}

Switch ($orientation){
    1 {"Orientation = Horizontal"}
    2 {"Orientation = Mirror Horizontal"}
    3 {"Orientation = Rotate 180°"}
    4 {"Orientation = Mirror Vertical"}
    5 {"Orientation = Mirror Horizontal & Rotate 180°"}
    6 {"Orientation = Rotate 90°clockwise"}
    7 {"Orientation = Mirror Horizontal & Rotate 90°clockwise"}
    8 {"Orientation = Rotate 270°clockwise"}
    default {"Orientation = "}
}

Write-Host "Advanced photo----------------------------------" -foregroundcolor DarkGreen

Switch ($contrast){
    0 {"Contrast = Normal"}
    1 {"Contrast = Low"}
    2 {"Contrast = High"}
    default  {"Contrast = "}
}

Switch ($saturation){
    0 {"Saturation = Normal"}
    1 {"Saturation = Low"}
    2 {"Saturation = High"}
    default  {"Saturation = "}
}

Switch ($sharpness){
    0 {"Sharpness = Normal"}
    1 {"Sharpness = Soft"}
    2 {"Sharpness = Hard"}
    default  {"Sharpness = "}
}

Write-Host "Brightness = " $Brightness

Switch ($whiteBalance){
    0 {"White Balance = Auto"}
    1 {"White Balance = Manual"}
    default {"White Balance = "}
}

Switch ($metering){
    0 {"Metering Mode = Unknow"}
    1 {"Metering Mode = Avarage"}
    2 {"Metering Mode = Center-weighted Avarage"}
    3 {"Metering Mode = Spot"}
    4 {"Metering Mode = Multi-spot"}
    5 {"Metering Mode = Multi-segment"}
    6 {"Metering Mode = Partial"}
    255 {"Metering Mode = Other"}
    defaul {"Metering Mode = "}
}

Switch ($expMode){
    0 {"Exposure Mode = Auto"}
    1 {"Exposure Mode = Manual"}
    2 {"Exposure Mode = Auto Bracket"}
    default {"Exposure Mode = "}
}

Switch ($ExpProg){
    1 {"Exposure Program = Manual"}
    2 {"Exposure Program = Program AE"}
    3 {"Exposure Program = Aperture-priority AE"}
    4 {"Exposure Program = Shutter-priority AE"}
    5 {"Exposure Program = Creative"}
    6 {"Exposure Program = Action"}
    7 {"Exposure Program = Portrait"}
    8 {"Exposure Program = Landscape"}
    9 {"Exposure Program = Bulb"}
    default {"Exposure Program = "}
}

Write-Host "Exposure Bias = " ("{0:N1}" -f $expComp) "step"

Switch ($SceneType){
    0 {"Scene Type = Standard"}
    1 {"Scene Type = Landscape"}
    2 {"Scene Type = Portrait"}
    3 {"Scene Type = Night"}
    default {"Scene Type = "}
}

Switch ($gainCtr){
    0 {"Gain Control = None"}
    1 {"Gain Control = Low gain up"}
    2 {"Gain Control = Hight gain up"}
    3 {"Gain Control = low gain down"}
    4 {"Gain Control = Hight gain down"}
    default {"Gain Control = "}
}

Switch ($LightSource){
    1 {"Light Source = Daylight"}
    2 {"Light Source = Fluorescent"}
    3 {"Light Source = Tungsten"}
    4 {"Light Source = Flash"}
    9 {"Light Source = Fine Weather"}
    10 {"Light Source = Cloudy"}
    11 {"Light Source = Shade"}
    12 {"Light Source = Daylight Fluorescent"}
    13 {"Light Source = Day White Fluorescent"}
    14 {"Light Source = Cool White Fluorescent"}
    15 {"Light Source = White Fluorescent"}
    16 {"Light Source = Warm White Fluorescent"}
    255 {"Light Source = Other"}
    default {"Light Source = "}
}
