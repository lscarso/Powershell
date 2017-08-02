# Functions from http://blog.cincura.net/233463-renaming-files-based-on-exif-data-in-powershell/
function GetTakenData([object]$image) {
    try {
        return $image.GetPropertyItem(36867).Value
    }   
    catch {
        return $null
    }
}

function Get-ExifData([string]$file)
 {
    [Reflection.Assembly]::LoadFile('C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.Drawing.dll') | Out-Null
    $image = New-Object System.Drawing.Bitmap -ArgumentList $file
    try {
        $takenData = GetTakenData($image)
        if ($takenData -eq $null) {
            return $null
        }
        $takenValue = [System.Text.Encoding]::Default.GetString($takenData, 0, $takenData.Length - 1)
        $taken = [DateTime]::ParseExact($takenValue, 'yyyy:MM:dd HH:mm:ss', $null)
        return $taken
    }
    finally {
        $image.Dispose()
    }
}


# http://blogs.technet.com/b/heyscriptingguy/archive/2012/06/01/use-powershell-to-modify-file-access-time-stamps.aspx
Function Set-FileTimeStamps
{
 Param (
    [Parameter(mandatory=$true)]
    [string]$path,
    [datetime]$date = (Get-Date))
    $file = Get-ChildItem -Path $path
    $file.CreationTime = $date
    $file.LastAccessTime = $date
    $file.LastWriteTime = $date
} 

gci "C:\YourJpegs\*.jpg" | foreach {
    Write-Host "$_`t->`t" -ForegroundColor Cyan -NoNewLine
    $date = (Get-ExifData $_.FullName)
    if ($date -eq $null) {
        Write-Host '{ No ''Date Taken'' in Exif }' -ForegroundColor Cyan    
        return
    }
    Set-FileTimeStamps $_ $date
    Write-Host "Set file timestamp to $date" -ForegroundColor Cyan
}
