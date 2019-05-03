<#
 
.SYNOPSIS
This script generates an html/csv report containing update compliance info for members of WSUS Update Groups. 
 
.PARAMETER group
Array of WSUS groups to report. "*" means every groups on WSUS

.PARAMETER WsusServer
Hostname or IP address of WSUS server

.PARAMETER exportToCSV
Export to CSV File. -reporPath is mandatory with this switch

.PARAMETER exportToHTML
Export to HTML File. -reporPath is mandatory with this switch

.PARAMETER reportPath
Path (folders + filename) where save the report. Mandatory with -exportToCSV and -exportToHTML

.PARAMETER exportToSharepoint
Save REPORT-Wsus.html to a Sharepoint Room. -Url is mandatory 

.PARAMETER Url
URL pointing to a SharePoint document library (omit the '/forms/default.aspx' portion).

.EXAMPLE
.\wsus_reportGroup.ps1 -group "Exchange Servers"
Provide a group and get a Screen report. More Groups separated by ","

.EXAMPLE
.\wsus_reportGroup.ps1 -group "Exchange Servers", "DC Servers" | ft server, security, failed
Provide a list of groups and get a screen report. More Groups separated by ","
 
.EXAMPLE
.\wsus_reportGroup.ps1
Don't provide a group and get a report for all WSUS Update Groups.

.EXAMPLE
.\wsus_reportGroup.ps1 -group "Test" -exportToCsv -reportPath C:\Test\report.csv

.EXAMPLE
.\wsus_reportGroup.ps1 -group "Test" -exportToHtml -reportPath C:\Test\report.html

.EXAMPLE
.\wsus_reportGroup.ps1 -group "Test" -exportToSharepoint -Url https://room01.sharepoint.net/01/Wiki/

 
.NOTES
A. The following is based off Joey Piccola wsus_reportGroup.ps1 (http://www.joeypiccola.com/tech/2014/9/14/web-based-dashboards-with-wsus-2012)
B. The HTML generation leverages Cookie.Monster's HTMLTable.ps1 script (http://gallery.technet.microsoft.com/scriptcenter/PowerShell-HTML-Notificatio-e1c5759d)
C. I should use more comments. 
 
#>

Param( 
    [cmdletBinding()]
    [parameter(ParameterSetName = "Shared")]
    [parameter(ParameterSetName = "CSV")]
    [parameter(ParameterSetName = "HTML")]
    [parameter(ParameterSetName = "Publish")]
    [string[]]$group = "*",

    [parameter(ParameterSetName = "Shared")]
    [parameter(ParameterSetName = "CSV")]
    [parameter(ParameterSetName = "HTML")]
    [parameter(ParameterSetName = "Publish")]
    [string]$wsusServer = "DEABGSWD22",
    
    [parameter(ParameterSetName = "CSV")]
    [switch]$exportToCsv,

    [parameter(Mandatory=$true,ParameterSetName = "CSV")]
    [parameter(Mandatory=$true,ParameterSetName = "HTML")]
    [ValidateScript({Test-path (Split-Path $_)})]
    [string]$reportPath,

    [parameter(ParameterSetName = "HTML")]
    [switch]$exportToHtml,
    
    [parameter(ParameterSetName = "Publish")]
    [switch]$exportToSharepoint,

    [parameter(Mandatory=$true,ParameterSetName = "Publish")]
    [System.Uri]$Url
)

begin
{
    function ConvertTo-PropertyValue {
        <#
        .SYNOPSIS
        Convert an object with various properties into an array of property, value pairs 
        
        .DESCRIPTION
        Convert an object with various properties into an array of property, value pairs

        If you output reports or other formats where a table with one long row is poorly formatted, this is a quick way to create a table of property value pairs.

        There are other ways you could do this.  For example, I could list all noteproperties from Get-Member results and return them.
        This function will keep properties in the same order they are provided, which can often be helpful for readability of results.

        .PARAMETER inputObject
        A single object to convert to an array of property value pairs.

        .PARAMETER leftheader
        Header for the left column.  Default:  Property

        .PARAMETER rightHeader
        Header for the right column.  Default:  Value

        .PARAMETER memberType
        Return only object members of this membertype.  Default:  Property, NoteProperty, ScriptProperty

        .EXAMPLE
        get-process powershell_ise | convertto-propertyvalue

        I want details on the powershell_ise process.
            With this command, if I output this to a table, a csv, etc. I will get a nice vertical listing of properties and their values
            Without this command, I get a long row with the same info

        .EXAMPLE
        #This example requires and demonstrates using the New-HTMLHead, New-HTMLTable, Add-HTMLTableColor, ConvertTo-PropertyValue and Close-HTML functions.
        
        #get processes to work with
            $processes = Get-Process
        
        #Build HTML header
            $HTML = New-HTMLHead -title "Process details"

        #Add CPU time section with top 10 PrivateMemorySize processes.  This example does not highlight any particular cells
            $HTML += "<h3>Process Private Memory Size</h3>"
            $HTML += New-HTMLTable -inputObject $($processes | sort PrivateMemorySize -Descending | select name, PrivateMemorySize -first 10)

        #Add Handles section with top 10 Handle usage.
        $handleHTML = New-HTMLTable -inputObject $($processes | sort handles -descending | select Name, Handles -first 10)

            #Add highlighted colors for Handle count
                
                #build hash table with parameters for Add-HTMLTableColor.  Argument and AttrValue will be modified each time we run this.
                $params = @{
                    Column = "Handles" #I'm looking for cells in the Handles column
                    ScriptBlock = {[double]$args[0] -gt [double]$args[1]} #I want to highlight if the cell (args 0) is greater than the argument parameter (arg 1)
                    Attr = "Style" #This is the default, don't need to actually specify it here
                }

                #Add yellow, orange and red shading
                $handleHTML = Add-HTMLTableColor -HTML $handleHTML -Argument 1500 -attrValue "background-color:#FFFF99;" @params
                $handleHTML = Add-HTMLTableColor -HTML $handleHTML -Argument 2000 -attrValue "background-color:#FFCC66;" @params
                $handleHTML = Add-HTMLTableColor -HTML $handleHTML -Argument 3000 -attrValue "background-color:#FFCC99;" @params
        
            #Add title and table
            $HTML += "<h3>Process Handles</h3>"
            $HTML += $handleHTML

        #Add process list containing first 10 processes listed by get-process.  This example does not highlight any particular cells
            $HTML += New-HTMLTable -inputObject $($processes | select name -first 10 ) -listTableHead "Random Process Names"

        #Add property value table showing details for PowerShell ISE
            $HTML += "<h3>PowerShell Process Details PropertyValue table</h3>"
            $processDetails = Get-process powershell_ise | select name, id, cpu, handles, workingset, PrivateMemorySize, Path -first 1
            $HTML += New-HTMLTable -inputObject $(ConvertTo-PropertyValue -inputObject $processDetails)

        #Add same PowerShell ISE details but not in property value form.  Close the HTML
            $HTML += "<h3>PowerShell Process Details object</h3>"
            $HTML += New-HTMLTable -inputObject $processDetails | Close-HTML

        #write the HTML to a file and open it up for viewing
            set-content C:\test.htm $HTML
            & 'C:\Program Files\Internet Explorer\iexplore.exe' C:\test.htm

        .FUNCTIONALITY
        General Command
        #> 
        [cmdletbinding()]
        param(
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true,
                ValueFromRemainingArguments=$false)]
            [PSObject]$InputObject,
            
            [validateset("AliasProperty", "CodeProperty", "Property", "NoteProperty", "ScriptProperty",
                "Properties", "PropertySet", "Method", "CodeMethod", "ScriptMethod", "Methods",
                "ParameterizedProperty", "MemberSet", "Event", "Dynamic", "All")]
            [string[]]$memberType = @( "NoteProperty", "Property", "ScriptProperty" ),
                
            [string]$leftHeader = "Property",
                
            [string]$rightHeader = "Value"
        )

        begin{
            #init array to dump all objects into
            $allObjects = New-Object System.Collections.ArrayList

        }
        process{
            #if we're taking from pipeline and get more than one object, this will build up an array
            [void]$allObjects.add($inputObject)
        }

        end{
            #use only the first object provided
            $allObjects = $allObjects[0]

            #Get properties.  Filter by memberType.
            $properties = $allObjects.psobject.properties | Where-Object{$memberType -contains $_.memberType} | Select-Object -ExpandProperty Name

            #loop through properties and display property value pairs
            foreach($property in $properties){

                #Create object with property and value
                $temp = "" | Select-Object $leftHeader, $rightHeader
                $temp.$leftHeader = $property.replace('"',"")
                $temp.$rightHeader = try { $allObjects | Select-Object -ExpandProperty $temp.$leftHeader -erroraction SilentlyContinue } catch { $null }
                $temp
            }
        }
    }

    function New-HTMLHead {
        <#
        .SYNOPSIS
            Returns HTML including internal style sheet

        .DESCRIPTION
            Returns HTML including internal style sheet

        .PARAMETER cssPath
            If specified, contents of this file are embedded in an internal style sheet via <style> tags
            
            Note:  If you include your own CSS, please note that the New-HTMLTable function looks for 'odd' and 'even' class names.  Note .odd and .even defitions in $HTMLStyle.

        .PARAMETER title
            If specified, title to add in the head section

        .EXAMPLE
        #This example requires and demonstrates using the New-HTMLHead, New-HTMLTable, Add-HTMLTableColor, ConvertTo-PropertyValue and Close-HTML functions.
        
        #get processes to work with
            $processes = Get-Process
        
        #Build HTML header
            $HTML = New-HTMLHead -title "Process details"

        #Add CPU time section with top 10 PrivateMemorySize processes.  This example does not highlight any particular cells
            $HTML += "<h3>Process Private Memory Size</h3>"
            $HTML += New-HTMLTable -inputObject $($processes | sort PrivateMemorySize -Descending | select name, PrivateMemorySize -first 10)

        #Add Handles section with top 10 Handle usage.
        $handleHTML = New-HTMLTable -inputObject $($processes | sort handles -descending | select Name, Handles -first 10)

            #Add highlighted colors for Handle count
                
                #build hash table with parameters for Add-HTMLTableColor.  Argument and AttrValue will be modified each time we run this.
                $params = @{
                    Column = "Handles" #I'm looking for cells in the Handles column
                    ScriptBlock = {[double]$args[0] -gt [double]$args[1]} #I want to highlight if the cell (args 0) is greater than the argument parameter (arg 1)
                    Attr = "Style" #This is the default, don't need to actually specify it here
                }

                #Add yellow, orange and red shading
                $handleHTML = Add-HTMLTableColor -HTML $handleHTML -Argument 1500 -attrValue "background-color:#FFFF99;" @params
                $handleHTML = Add-HTMLTableColor -HTML $handleHTML -Argument 2000 -attrValue "background-color:#FFCC66;" @params
                $handleHTML = Add-HTMLTableColor -HTML $handleHTML -Argument 3000 -attrValue "background-color:#FFCC99;" @params
        
            #Add title and table
            $HTML += "<h3>Process Handles</h3>"
            $HTML += $handleHTML

        #Add process list containing first 10 processes listed by get-process.  This example does not highlight any particular cells
            $HTML += New-HTMLTable -inputObject $($processes | select name -first 10 ) -listTableHead "Random Process Names"

        #Add property value table showing details for PowerShell ISE
            $HTML += "<h3>PowerShell Process Details PropertyValue table</h3>"
            $processDetails = Get-process powershell_ise | select name, id, cpu, handles, workingset, PrivateMemorySize, Path -first 1
            $HTML += New-HTMLTable -inputObject $(ConvertTo-PropertyValue -inputObject $processDetails)

        #Add same PowerShell ISE details but not in property value form.  Close the HTML
            $HTML += "<h3>PowerShell Process Details object</h3>"
            $HTML += New-HTMLTable -inputObject $processDetails | Close-HTML

        #write the HTML to a file and open it up for viewing
            set-content C:\test.htm $HTML
            & 'C:\Program Files\Internet Explorer\iexplore.exe' C:\test.htm

        .FUNCTIONALITY
            General Command

        #>
        [cmdletbinding(DefaultParameterSetName="String")]    
        param(
            
            [Parameter(ParameterSetName='File')]
            [validatescript({test-path $_ -pathtype leaf})]$cssPath = $null,
            
            [Parameter(ParameterSetName='String')]
            [string]$style = "<style>
                        body {
                            color:#333333;
                            font-family:Calibri,Tahoma,arial,verdana;
                            font-size: 11pt;
                        }
                        h1 {
                            text-align:center;
                        }
                        h2 {
                            border-top:1px solid #666666;
                        }
                        table {
                            border-collapse:collapse;
                        }
                        th {
                            text-align:left;
                            font-weight:bold;
                            color:#eeeeee;
                            background-color:#333333;
                            border:1px solid black;
                            padding:5px;
                        }
                        td {
                            padding:5px;
                            border:1px solid black;
                        }
                        .odd { background-color:#ffffff; }
                        .even { background-color:#dddddd; }
                    </style>",
            
            [string]$title = $null
        )

        #add css from file if specified
        if($cssPath){$style = "<style>$(get-content $cssPath | out-string)</style>"}

        #Return HTML
        @"
        <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
            <html xmlns="http://www.w3.org/1999/xhtml">
                <head>
                    $(if($title){"<title>$title</title>"})
                        $style
                </head>
                <body>

"@
    }

    function New-HTMLTable {
        <# 
        .SYNOPSIS 
        Create an HTML table from an input object
    
        .DESCRIPTION 
        Create an HTML table from an input object
    
        .PARAMETER  InputObject 
        One or more objects (ie. (Get-process | select Name,Company) 
    
        .PARAMETER Properties
        If specified, limit table to these specific properties in the order specified.
        
        .PARAMETER setAlternating
        Add CSS class = odd or even to each row.  True by default.  Be sure your CSS includes odd and even definitions

        .PARAMETER listTableHead
        If a list is provided, use this parameter to specify the list header (PowerShell uses * by default)

        .EXAMPLE
        #This example requires and demonstrates using the New-HTMLHead, New-HTMLTable, Add-HTMLTableColor, ConvertTo-PropertyValue and Close-HTML functions.
        
        #get processes to work with
            $processes = Get-Process
        
        #Build HTML header
            $HTML = New-HTMLHead -title "Process details"

        #Add CPU time section with top 10 PrivateMemorySize processes.  This example does not highlight any particular cells
            $HTML += "<h3>Process Private Memory Size</h3>"
            $HTML += New-HTMLTable -inputObject $($processes | sort PrivateMemorySize -Descending | select name, PrivateMemorySize -first 10)

        #Add Handles section with top 10 Handle usage.
        $handleHTML = New-HTMLTable -inputObject $($processes | sort handles -descending | select Name, Handles -first 10)

            #Add highlighted colors for Handle count
                
                #build hash table with parameters for Add-HTMLTableColor.  Argument and AttrValue will be modified each time we run this.
                $params = @{
                    Column = "Handles" #I'm looking for cells in the Handles column
                    ScriptBlock = {[double]$args[0] -gt [double]$args[1]} #I want to highlight if the cell (args 0) is greater than the argument parameter (arg 1)
                    Attr = "Style" #This is the default, don't need to actually specify it here
                }

                #Add yellow, orange and red shading
                $handleHTML = Add-HTMLTableColor -HTML $handleHTML -Argument 1500 -attrValue "background-color:#FFFF99;" @params
                $handleHTML = Add-HTMLTableColor -HTML $handleHTML -Argument 2000 -attrValue "background-color:#FFCC66;" @params
                $handleHTML = Add-HTMLTableColor -HTML $handleHTML -Argument 3000 -attrValue "background-color:#FFCC99;" @params
        
            #Add title and table
            $HTML += "<h3>Process Handles</h3>"
            $HTML += $handleHTML

        #Add process list containing first 10 processes listed by get-process.  This example does not highlight any particular cells
            $HTML += New-HTMLTable -inputObject $($processes | select name -first 10 ) -listTableHead "Random Process Names"

        #Add property value table showing details for PowerShell ISE
            $HTML += "<h3>PowerShell Process Details PropertyValue table</h3>"
            $processDetails = Get-process powershell_ise | select name, id, cpu, handles, workingset, PrivateMemorySize, Path -first 1
            $HTML += New-HTMLTable -inputObject $(ConvertTo-PropertyValue -inputObject $processDetails)

        #Add same PowerShell ISE details but not in property value form.  Close the HTML
            $HTML += "<h3>PowerShell Process Details object</h3>"
            $HTML += New-HTMLTable -inputObject $processDetails | Close-HTML

        #write the HTML to a file and open it up for viewing
            set-content C:\test.htm $HTML
            & 'C:\Program Files\Internet Explorer\iexplore.exe' C:\test.htm

        .NOTES 
        Props to Zachary Loeber for the idea: http://gallery.technet.microsoft.com/scriptcenter/Colorize-HTML-Table-Cells-2ea63acd

        I believe that .Net 3.5 is a requirement for using the Linq libraries

        .FUNCTIONALITY
        General Command
        #> 
        [CmdletBinding()] 
        param ( 
            [Parameter( Position=0,
                        Mandatory=$true, 
                        ValueFromPipeline=$true)]
            [PSObject[]]$InputObject,

            [Parameter( Mandatory=$false, 
                        ValueFromPipeline=$false)]
            [string[]]$Properties,
            
            [Parameter( Mandatory=$false, 
                        ValueFromPipeline=$false)]
            [bool]$setAlternating = $true,

            [Parameter( Mandatory=$false, 
                        ValueFromPipeline=$false)]
            [string]$listTableHead = $null

            )
        
        BEGIN { 
            #requires -version 2.0
            add-type -AssemblyName System.xml.linq | out-null
            $Objects = New-Object System.Collections.ArrayList
        } 
    
        PROCESS { 

            #Loop through inputObject, add to collection.  Filter properties if specified.
            foreach($object in $inputObject){
                if($Properties){ [void]$Objects.add(($object | Select-Object $Properties)) }
                else{ [void]$Objects.add( $object )}
            }

        } 
    
        END { 

            # Convert our data to x(ht)ml  
            $xml = [System.Xml.Linq.XDocument]::Parse("$($Objects | ConvertTo-Html -Fragment)")
            
            #replace * as table head if specified.  Note, this should only be done for a list...
            if($listTableHead){
                $xml = [System.Xml.Linq.XDocument]::parse( $xml.Document.ToString().replace("<th>*</th>","<th>$listTableHead</th>") )
            }

            if($setAlternating){
                #loop through descendents.  If their index is even mark with class even, odd with class odd.
                foreach($descendent in $($xml.Descendants("tr"))){
                    if(($descendent.NodesBeforeSelf() | Measure-Object).count % 2 -eq 0){
                        $descendent.SetAttributeValue("class", "even") 
                    }
                    else{
                        $descendent.SetAttributeValue("class", "odd") 
                    }
                }
            }
            #Provide full HTML or just the table depending on param
            $xml.Document.ToString()
        }
    }

    function Add-HTMLTableColor {
        <# 
        .SYNOPSIS 
        Colorize cells or rows in an HTML table, or add other inline CSS
    
        .DESCRIPTION 
        Colorize cells or rows in an HTML table, or add other inline CSS
    
        .PARAMETER  HTML 
        HTML string to work with

        .PARAMETER  Column 
        If specified, the column you want to modify.  This is case sensitive

        .PARAMETER  Argument 
        If Column is specified, this argument can be used to compare with current cell.

        .PARAMETER ScriptBlock
        If Column is specified, used to evaluate whether to colorize a cell.  If the scriptblock returns $true the cell will be colorized.
    
        $args[0] is the existing cell value in the table
        $args[1] is your Argument parameter

        Examples:
            {[string]$args[0] -eq [string]$args[1]} #existing cell value equals Argument.  This is the default
            {[double]$args[0] -gt [double]$args[1]} #existing cell value is greater than Argument.

        Use strong typesetting if possible.
    
        .PARAMETER  Attr 
        If Column is specified, the attribute to change should ColumnValue be found in the Column specified or if the ScriptBlock is true.  Default:  Style
    
        .PARAMETER  AttrValue 
        If Column is specified, the attribute value to set when the ColumnValue is found in the Column specified or if the ScriptBlock is true.
        
        Example: "background-color:#FFCC99;" 
    
        .PARAMETER WholeRow
        If specified, and Column is specified, set the Attr and AttrValue for the entire row, not just a cell.

        .EXAMPLE
        #This example requires and demonstrates using the New-HTMLHead, New-HTMLTable, Add-HTMLTableColor, ConvertTo-PropertyValue and Close-HTML functions.
        
        #get processes to work with
            $processes = Get-Process
        
        #Build HTML header
            $HTML = New-HTMLHead -title "Process details"

        #Add CPU time section with top 10 PrivateMemorySize processes.  This example does not highlight any particular cells
            $HTML += "<h3>Process Private Memory Size</h3>"
            $HTML += New-HTMLTable -inputObject $($processes | sort PrivateMemorySize -Descending | select name, PrivateMemorySize -first 10)

        #Add Handles section with top 10 Handle usage.
        $handleHTML = New-HTMLTable -inputObject $($processes | sort handles -descending | select Name, Handles -first 10)

            #Add highlighted colors for Handle count
                
                #build hash table with parameters for Add-HTMLTableColor.  Argument and AttrValue will be modified each time we run this.
                $params = @{
                    Column = "Handles" #I'm looking for cells in the Handles column
                    ScriptBlock = {[double]$args[0] -gt [double]$args[1]} #I want to highlight if the cell (args 0) is greater than the argument parameter (arg 1)
                    Attr = "Style" #This is the default, don't need to actually specify it here
                }

                #Add yellow, orange and red shading
                $handleHTML = Add-HTMLTableColor -HTML $handleHTML -Argument 1500 -attrValue "background-color:#FFFF99;" @params
                $handleHTML = Add-HTMLTableColor -HTML $handleHTML -Argument 2000 -attrValue "background-color:#FFCC66;" @params
                $handleHTML = Add-HTMLTableColor -HTML $handleHTML -Argument 3000 -attrValue "background-color:#FFCC99;" @params
        
            #Add title and table
            $HTML += "<h3>Process Handles</h3>"
            $HTML += $handleHTML

        #Add process list containing first 10 processes listed by get-process.  This example does not highlight any particular cells
            $HTML += New-HTMLTable -inputObject $($processes | select name -first 10 ) -listTableHead "Random Process Names"

        #Add property value table showing details for PowerShell ISE
            $HTML += "<h3>PowerShell Process Details PropertyValue table</h3>"
            $processDetails = Get-process powershell_ise | select name, id, cpu, handles, workingset, PrivateMemorySize, Path -first 1
            $HTML += New-HTMLTable -inputObject $(ConvertTo-PropertyValue -inputObject $processDetails)

        #Add same PowerShell ISE details but not in property value form.  Close the HTML
            $HTML += "<h3>PowerShell Process Details object</h3>"
            $HTML += New-HTMLTable -inputObject $processDetails | Close-HTML

        #write the HTML to a file and open it up for viewing
            set-content C:\test.htm $HTML
            & 'C:\Program Files\Internet Explorer\iexplore.exe' C:\test.htm

        .EXAMPLE
        # Table with the 20 most recent events, highlighting error and warning rows

            #gather 20 events from the system log and pick out a few properties
            $events = Get-EventLog -LogName System -Newest 20 | select TimeGenerated, Index, EntryType, UserName, Message

        #Create the HTML table without alternating rows, colorize Warning and Error messages, highlighting the whole row.
            $eventTable = $events | New-HTMLTable -setAlternating $false |
                Add-HTMLTableColor -Argument "Warning" -Column "EntryType" -AttrValue "background-color:#FFCC66;" -WholeRow |
                Add-HTMLTableColor -Argument "Error" -Column "EntryType" -AttrValue "background-color:#FFCC99;" -WholeRow

        #Build the HTML head, add an h3 header, add the event table, and close out the HTML
            $HTML = New-HTMLHead
            $HTML += "<h3>Last 20 System Events</h3>"
            $HTML += $eventTable | Close-HTML

        #test it out
            set-content C:\test.htm $HTML
            & 'C:\Program Files\Internet Explorer\iexplore.exe' C:\test.htm

        .NOTES 
        Props to Zachary Loeber and Jaykul for the idea and help:
        http://gallery.technet.microsoft.com/scriptcenter/Colorize-HTML-Table-Cells-2ea63acd
        http://stackoverflow.com/questions/4559233/technique-for-selectively-formatting-data-in-a-powershell-pipeline-and-output-as

        I believe that .Net 3.5 is a requirement for using the Linq libraries
        
        .FUNCTIONALITY
        General Command
        #> 
        [CmdletBinding()] 
        param ( 
            [Parameter( Mandatory=$true,  
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$false)]  
            [string]$HTML,
            
            [Parameter( Mandatory=$false, 
                        ValueFromPipeline=$false)]
            [String]$Column="Name",
            
            [Parameter( Mandatory=$false,
                        ValueFromPipeline=$false)]
            $Argument=0,
            
            [Parameter( ValueFromPipeline=$false)]
            [ScriptBlock]$ScriptBlock = {[string]$args[0] -eq [string]$args[1]},
            
            [Parameter( ValueFromPipeline=$false)]
            [String]$Attr = "style",
            
            [Parameter( Mandatory=$true, 
                        ValueFromPipeline=$false)] 
            [String]$AttrValue,
            
            [Parameter( Mandatory=$false, 
                        ValueFromPipeline=$false)] 
            [switch]$WholeRow=$false

            )
        
            #requires -version 2.0
            add-type -AssemblyName System.xml.linq | out-null

            # Convert our data to x(ht)ml  
            $xml = [System.Xml.Linq.XDocument]::Parse($HTML)   
            
            #Get column index.  try th with no namespace first, then default namespace provided by convertto-html
            try{ 
                $columnIndex = (($xml.Descendants("th") | Where-Object { $_.Value -eq $Column }).NodesBeforeSelf() | Measure-Object).Count 
            }
            catch { 
                Try {
                    $columnIndex = (($xml.Descendants("{http://www.w3.org/1999/xhtml}th") | Where-Object { $_.Value -eq $Column }).NodesBeforeSelf() | Measure-Object).Count
                }
                Catch {
                    Throw "Error:  Namespace incorrect."
                }
            }

            #if we got the column index...
            if($columnIndex -as [double] -ge 0){
                
                #take action on td descendents matching that index
                switch($xml.Descendants("td") | Where-Object { ($_.NodesBeforeSelf() | Measure-Object).Count -eq $columnIndex })
                {
                    #run the script block.  If it is true, set attributes
                    {$(Invoke-Command $ScriptBlock -ArgumentList @($_.Value, $Argument))} { 
                        
                        #mark the whole row or just a cell depending on param
                        if ($WholeRow)  { 
                            $_.Parent.SetAttributeValue($Attr, $AttrValue) 
                        } 
                        else { 
                            $_.SetAttributeValue($Attr, $AttrValue) 
                        }
                    }
                }
            }
            
            #return the XML
            $xml.Document.ToString() 
    }

    function Close-HTML {
        <# 
        .SYNOPSIS 
        Close out the body and html tags
    
        .DESCRIPTION 
        Close out the body and html tags
    
        .PARAMETER  HTML 
        HTML string to work with

        .PARAMETER Decode
        If specified, run HTML string through HtmlDecode

        .EXAMPLE
        #This example requires and demonstrates using the New-HTMLHead, New-HTMLTable, Add-HTMLTableColor, ConvertTo-PropertyValue and Close-HTML functions.
        
        #get processes to work with
            $processes = Get-Process
        
        #Build HTML header
            $HTML = New-HTMLHead -title "Process details"

        #Add CPU time section with top 10 PrivateMemorySize processes.  This example does not highlight any particular cells
            $HTML += "<h3>Process Private Memory Size</h3>"
            $HTML += New-HTMLTable -inputObject $($processes | sort PrivateMemorySize -Descending | select name, PrivateMemorySize -first 10)

        #Add Handles section with top 10 Handle usage.
        $handleHTML = New-HTMLTable -inputObject $($processes | sort handles -descending | select Name, Handles -first 10)

            #Add highlighted colors for Handle count
                
                #build hash table with parameters for Add-HTMLTableColor.  Argument and AttrValue will be modified each time we run this.
                $params = @{
                    Column = "Handles" #I'm looking for cells in the Handles column
                    ScriptBlock = {[double]$args[0] -gt [double]$args[1]} #I want to highlight if the cell (args 0) is greater than the argument parameter (arg 1)
                    Attr = "Style" #This is the default, don't need to actually specify it here
                }

                #Add yellow, orange and red shading
                $handleHTML = Add-HTMLTableColor -HTML $handleHTML -Argument 1500 -attrValue "background-color:#FFFF99;" @params
                $handleHTML = Add-HTMLTableColor -HTML $handleHTML -Argument 2000 -attrValue "background-color:#FFCC66;" @params
                $handleHTML = Add-HTMLTableColor -HTML $handleHTML -Argument 3000 -attrValue "background-color:#FFCC99;" @params
        
            #Add title and table
            $HTML += "<h3>Process Handles</h3>"
            $HTML += $handleHTML

        #Add process list containing first 10 processes listed by get-process.  This example does not highlight any particular cells
            $HTML += New-HTMLTable -inputObject $($processes | select name -first 10 ) -listTableHead "Random Process Names"

        #Add property value table showing details for PowerShell ISE
            $HTML += "<h3>PowerShell Process Details PropertyValue table</h3>"
            $processDetails = Get-process powershell_ise | select name, id, cpu, handles, workingset, PrivateMemorySize, Path -first 1
            $HTML += New-HTMLTable -inputObject $(ConvertTo-PropertyValue -inputObject $processDetails)

        #Add same PowerShell ISE details but not in property value form.  Close the HTML
            $HTML += "<h3>PowerShell Process Details object</h3>"
            $HTML += New-HTMLTable -inputObject $processDetails | Close-HTML

        #write the HTML to a file and open it up for viewing
            set-content C:\test.htm $HTML
            & 'C:\Program Files\Internet Explorer\iexplore.exe' C:\test.htm

        .FUNCTIONALITY
        General Command
        #>

        [cmdletbinding()]
        param(
            [Parameter( Mandatory=$true,  
                ValueFromPipeline=$true,
                ValueFromPipelineByPropertyName=$false)]  
            [string]$HTML,

            [switch]$Decode
        )
        #Thanks to ‏@ashyoungblood!
        if($Decode)
        {
            Add-Type -AssemblyName System.Web
            $HTML = [System.Web.HttpUtility]::HtmlDecode($HTML)
        }
        "$HTML </body></html>"
    }

    function Publish-File {
        param (
            [parameter( Mandatory = $true, HelpMessage="URL pointing to a SharePoint document library (omit the '/forms/default.aspx' portion)." )]
            [System.Uri]$Url,
            [parameter( Mandatory = $true, ValueFromPipeline = $true, HelpMessage="One or more files to publish. Use 'dir' to produce correct object type." )]
            [System.IO.FileInfo[]]$FileName,
            [system.Management.Automation.PSCredential]$Credential
        )
        $wc = new-object System.Net.WebClient
        if ( $Credential ) { $wc.Credentials = $Credential }
        else { $wc.UseDefaultCredentials = $true }
        $FileName | ForEach-Object {
            $DestUrl = "{0}{1}{2}" -f $Url.ToString().TrimEnd("/"), "/", $_.Name
            Write-Verbose "$( get-date -f s ): Uploading file: $_"
            $wc.UploadFile( $DestUrl , "PUT", $_.FullName )
            Write-Verbose "$( get-date -f s ): Upload completed"
        }
        
    }
}

process
{
    try{
        Import-Module servermanager
        #Check if Windows feature is installed- if not attempts to install Windows backup and tools features
        $FeatureStatus=(Get-WindowsFeature UpdateServices-Ui).Installed
        If($FeatureStatus -eq $False){
            Add-windowsfeature -Name UpdateServices-Ui -OutVariable results
            #Confirm sucessful feature installation
            foreach($result in $results){
                If($result.success){Write-Host "Windows Software Update Services installed successfully on $env:Computername" -ForegroundColor "Green" }
                Else{ Write-Host "Windows Software Update Services installation failed on $env:Computername" -ForegroundColor "Red"}
            }
        }


        # wsus server info
        #$wsusServer = "DEABGSWD22"
        $wsusPort = "8530"
        # window for how long a computer has gone without reporting before it's flagged to appear as a differnt color in the html table
        $thirtydaysago = (get-date).adddays(-30)
    
        # set up our connection 
        [void][reflection.assembly]::loadwithpartialname("microsoft.updateservices.administration")
        $wsus = [microsoft.updateservices.administration.adminproxy]::getupdateserver($wsusServer,$false,$wsusPort)
    
        $UpdateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
        $UpdateScope.ApprovedStates = [Microsoft.UpdateServices.Administration.ApprovedStates]::LatestRevisionApproved
        $ComputerScope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope
        $compArray = @()
        
        # define how GetComputerTargetGroups works and whether or not to exclude a specific parent group
        #$UpdateGroups = $WSUS.GetComputerTargetGroups() | Where {$_.name -like $group -and $_.name -notlike "INSERT PARTENT NAME HERE"}
        If ($group -eq '*') {
            Write-Output 'Collecting data for all WSUS Groups'
            $UpdateGroups = $WSUS.GetComputerTargetGroups()    
        }
        else {
            Write-Output "Searching WSUS Groups: $group"
            $UpdateGroups = $WSUS.GetComputerTargetGroups() | Where-Object {$_.name -and $group -eq $_.name}    
        }
        If ($UpdateGroups -eq $null) {
            Write-Warning "No WSUS Groups found."
        }

        $HTML = New-HTMLHead 
        $UpdateGroups = $UpdateGroups | Sort-Object -Property name
        foreach ($UpdateGroup in $UpdateGroups)
        {   
            Write-Output "Collecting data for $($UpdateGroup.name)"
            $UpdateGroupMembers = $wsus.getComputerTargetGroup($UpdateGroup.Id).GetComputerTargets()
            if ($UpdateGroupMembers.Count -gt 0)
            {
                $groupClean = $UpdateGroup.name.Replace(" ","")
                $compArray = @()
                foreach ($UpdateGroupMember in $UpdateGroupMembers)
                {
                    $needed = 0
                    $downloaded = 0
                    $notinstalled = 0
                    $su = 0
                    $cu = 0
                    $nu = 0 # <-- not used
                    $other = 0
                    $name = $null
                    
                    $compSummary = $wsus.GetSummariesPerComputerTarget($updatescope,$computerscope) | ?{$_.ComputerTargetID -eq $UpdateGroupMember.id}
                    $memberFQDN = $UpdateGroupMember.fulldomainname
                    $name = $memberFQDN.Split(".")[0]
                    $compObj = New-Object PSObject
                    $compObj | Add-Member -MemberType NoteProperty -Name Server -Value $UpdateGroupMember.FullDomainName
    
                    $neededUpdates = ($WSUS.GetComputerTargetbyname($memberFQDN)).GetUpdateInstallationInfoPerUpdate() | `
                    Where-Object{($_.UpdateApprovalAction -eq "install") -and (($_.UpdateInstallationState -eq "downloaded") -or ($_.UpdateInstallationState -eq "notinstalled"))}
                    
                    if ($neededUpdates -ne $null)
                    {
                        foreach ($update in $neededUpdates)
                        {
                            $updateMeta = $wsus.GetUpdate([Guid]$update.updateid)
                            $needed++
                            if ($updateMeta.UpdateClassificationTitle -eq "Updates") {$nu++}  # <-- not used
                            elseif ($updateMeta.UpdateClassificationTitle -eq "Security Updates") {$su++}
                            elseif ($updateMeta.UpdateClassificationTitle -eq "Critical Updates") {$cu++}
                            else {$other++} # <-- not used
                            
                            if (($update.UpdateInstallationState -eq "downloaded") -or ($update.UpdateInstallationState -eq "notinstalled"))
                            {
                                $notinstalled++
                            }
                        }
                    }
    
                    $compObj | Add-Member -MemberType NoteProperty -Name Installed -Value $compSummary.InstalledCount
                    $compObj | Add-Member -MemberType NoteProperty -Name "Not Installed" -Value $notinstalled
                    $compObj | Add-Member -MemberType NoteProperty -Name Critical -Value $cu
                    $compObj | Add-Member -MemberType NoteProperty -Name Security -Value $su
                    $compObj | Add-Member -MemberType NoteProperty -Name Failed -Value $compSummary.failedCount
                    $compObj | Add-Member -MemberType NoteProperty -Name "Not Applicable" -Value $compSummary.NotApplicableCount
                    $compObj | Add-Member -MemberType NoteProperty -Name "Pending Reboot" -Value $compSummary.InstalledPendingRebootCount
                    $compObj | Add-Member -MemberType NoteProperty -Name "Last Reported Status Time" -Value $UpdateGroupMember.LastReportedStatusTime
                    $compObj | Add-Member -MemberType NoteProperty -Name "Last Sync Time" -Value $UpdateGroupMember.LastSyncTime
                    $compObj | Add-Member -MemberType NoteProperty -Name "Last Sync Result" -Value $UpdateGroupMember.LastSyncResult
                    $compObj | Add-Member -MemberType NoteProperty -Name "Operating System" -Value $UpdateGroupMember.OSDescription
                    $compArray += $compObj
                }
                
                # time to get fancy, params hash for easier reading
                $paramsNotInstalled = @{ 
                    Column = "Not Installed"
                    ScriptBlock = {[double]$args[0] -ge [double]$args[1]}  
                    Attr = "Style"
                }
                $paramsPendingReboot = @{ 
                    Column = "Pending Reboot"
                    ScriptBlock = {[double]$args[0] -gt [double]$args[1]} 
                    Attr = "Style"
                }
                $paramsLastSyncResult = @{ 
                    Column = "Last Sync Result"
                    ScriptBlock = {$args[0] -ne $args[1]} 
                    Attr = "Style"
                }
                $paramsFailed = @{ 
                    Column = "Failed"
                    ScriptBlock = {[double]$args[0] -gt [double]$args[1]} 
                    Attr = "Style"
                }           
                $paramsLastReportedStatusTime = @{ 
                    Column = "Last Reported Status Time"
                    ScriptBlock = {[datetime]::Parse($args[0]) -le $args[1]} 
                    Attr = "Style"
                }           
                
                # begin generating the HTML Table and define the default sort order
                $compTable = $compArray | Sort-Object -Property "Not Installed","Server" -Descending | New-HTMLTable -setAlternating $true
                $HTML += "<h3>Update Group: $($UpdateGroup.name) ($($UpdateGroupMembers.count)x)</h3>"
                $HTML += "<h4>Last Updated: $(get-date)</h4>"
                $compTable = Add-HTMLTableColor -HTML $compTable -Argument 20 -attrValue "background-color:#FFF284;" @paramsNotInstalled
                $compTable = Add-HTMLTableColor -HTML $compTable -Argument 40 -attrValue "background-color:#FFCB2F;" @paramsNotInstalled
                $compTable = Add-HTMLTableColor -HTML $compTable -Argument 60 -attrValue "background-color:#FF5353;" @paramsNotInstalled    
                $compTable = Add-HTMLTableColor -HTML $compTable -Argument 0 -attrValue "background-color:#8CD1E6;" @paramsPendingReboot
                $compTable = Add-HTMLTableColor -HTML $compTable -Argument "Succeeded" -attrValue "background-color:#9669FE;" @paramsLastSyncResult
                $compTable = Add-HTMLTableColor -HTML $compTable -Argument 0 -attrValue "background-color:#88AC76;" @paramsFailed
                $compTable = Add-HTMLTableColor -HTML $compTable -Argument $thirtydaysago -attrValue "background-color:#FF00FF;" @paramsLastReportedStatusTime
                
                $HTML += $compTable
                $CSV += $compArray
            }
        }
        $HTML = $HTML | close-html
        
        #Export to CSV
        if ($exportToCsv) {
            Write-Output "Saving $reportPath"
            $CSV | Sort-Object -Property "Not Installed","Server" -Descending | Export-Csv -Path $reportPath -NoTypeInformation
        }

        #Export to HTML
        if ($exportToHtml) {
            Write-Output "Saving $reportPath"
            set-content $reportPath $HTML
        }

        #Export to Sharepoint
        if ($exportToSharepoint) {
            Write-Output "Saving WSUSReport.html to Sharepoint: $Url"
            Set-Content '.\Working\REPORT-Wsus.html' $HTML
            Get-ChildItem ".\Working\REPORT-Wsus.html" | Publish-File -Url $Url
        }

        return $CSV | Sort-Object -Property "Not Installed","Server" -Descending

    }
    Catch{
        Write-Warning $_
    }
}

end{
}