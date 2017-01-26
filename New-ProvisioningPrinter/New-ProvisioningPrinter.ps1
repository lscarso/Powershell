function New-ProvisioningPrinter {
    <# 
    .SYNOPSIS 
    Create Printer.

    .DESCRIPTION 
    Create Printer port and Printer.

    .PARAMETER PrinterName
    Name of printer. Share name will be the same.

    .PARAMETER PrinterAddress
    IP address/HostName of printer. Port name will be the same.

    .PARAMETER PrinterLocation
    Location attribute of print queue.

    .PARAMETER PrinterComment
    Comment attribute of print queue.

    .PARAMETER PrinterDriver
    Driver to use on Printer queue. Driver should be already installed on printserver.

    .PARAMETER ComputerName
    IP address/HostName of printeserver where install print queue.

    .EXAMPLE 
    Import-Csv "C:\Data\Working\printers.csv" -delimiter ";" -Encoding UTF8 | New-ProvisioningPrinter -ComputerName PRINTSRV   

    .NOTES 
    Preinstall driver on Server:
    pnputil -i -a "c:\test\drivers\printerdrive.inf"
    Add-PrinterDriver -ComputerName PRINTSRV -Name "RICOH PCL6 UniversalDriver V4.12"

    CSV must use ";" as delimiter
    CVS Mandatory Column:
        PrinterName;PrinterAddress;
    CSV Not Mandatory Column:
        PrinterLocation;PrinterComment
    CSV Example:
        PrinterName;PrinterAddress;PrinterLocation;PrinterComment

    To change Default settings:
    Set-PrintConfiguration -ComputerName $ComputerName –PrinterName $PrinterName -PaperSize A4 -Color $False
    Get-PrintConfiguration -PrinterName $PrinterName -ComputerName $ComputerName
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True,
        ValueFromPipelineByPropertyName=$True)]
        [String]$PrinterName,
        [Parameter(Mandatory=$True,
        ValueFromPipelineByPropertyName=$True)]
        [String]$PrinterAddress,
        [Parameter(ValueFromPipelineByPropertyName=$True)]
        [String]$PrinterLocation,
        [Parameter(ValueFromPipelineByPropertyName=$True)]
        [String]$PrinterComment,
        [Parameter(ValueFromPipelineByPropertyName=$True)]
        [String]$PrinterDriver = "RICOH PCL6 UniversalDriver V4.12",
        [Parameter(Mandatory=$True,
        ValueFromPipelineByPropertyName=$True)]
        [String]$ComputerName
        )

    Begin{
    }
    Process{
        Write-Output "Creating $PrinterName :"
        Write-Output "- Creating Port $PrintAddress"
        Add-PrinterPort -ComputerName $ComputerName -Name $PrinterAddress -PrinterHostAddress $PrinterAddress -SNMPCommunity "public" -SNMP 1
        Write-Output "- Creating Shared Printer $PrintName"
        Add-Printer -ComputerName $ComputerName -Name $PrinterName -DriverName $PrinterDriver -port $PrinterAddress -Shared -ShareName $PrinterName –Published -Location $PrinterLocation -Comment $PrinterComment 
    }
    End{
    }
}

function New-ProvisioningDriver {
    <# 
    .SYNOPSIS 
    Add Printer Driver to driverStore.

    .DESCRIPTION 
    Add Printer Driver to driverStore.

    .PARAMETER PrinterDriver
    Driver Name. Check inf content.

    .PARAMETER DriverPath
    inf file of driver to install.

    .EXAMPLE 
    .\Add-ProvisioningDriver.ps1 -PrinterName "RICOH PCL6 UniversalDriver V4.12" -DriverPath C:\Temp\Prtinters\Ricoh_UniDrv_pcl6\uni.inf  

    .NOTES 
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True,
        ValueFromPipelineByPropertyName=$True)]
        [String]$PrinterDriver = "RICOH PCL6 UniversalDriver V4.12",
        [Parameter(Mandatory=$True,
        ValueFromPipelineByPropertyName=$True)]
        [String]$DriverPath
        )

    if (!(get-printerdriver -ea 0 $PrinterDriver)){
        pnputil -i -a $DriverPath
        add-printerdriver -name $PrinterDriver
    }

    #if (!(get-printerport -ea 0 "Main7835")){
    #    add-printerport -name "Main7835" -printerhostaddress "192.168.0.252"}
    #if (!(get-printer -ea 0 "Main Color Copier")){
    #    add-printer -name "Main Color Copier" -drivername "Xerox WorkCentre 7835 PS" -portname "Main7835"}
}
