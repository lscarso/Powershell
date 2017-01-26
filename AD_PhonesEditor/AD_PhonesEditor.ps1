<# 
.SYNOPSIS 
Edit phones attribute on Active direcory Users Objects.

.DESCRIPTION 
Graphical APP for telephoneNumber and MobilePhone attibutes management. 

.NOTES
Edit $ADSearchBase for filtering users
Edit $logFileName for save modified record log
Author: Luca Scarsini

#>

# Script configuration
$ADSearchBase = "OU=My_Users,DC=My_Domain,DC=local" 
$logFileName = "C:\Temp\PhoneEditor.log"

Add-Type -AssemblyName presentationframework
#Add-Type -AssemblyName System.Windows.Forms
#Add-Type -AssemblyName System.DirectoryServices.AccountManagement
#Import-Module ActiveDirectory


[xml]$xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:sys="clr-namespace:System;assembly=mscorlib"
    Title="Phone Editor" Width="550" Height="500" 
    FontSize="13" WindowStartupLocation="CenterScreen"
    >  
  
  <Grid
    FocusManager.FocusedElement="{Binding ElementName=txtSearch}">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto" />
      <RowDefinition Height="*" />
    </Grid.RowDefinitions>

    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="200" />
      <ColumnDefinition Width="*" />
    </Grid.ColumnDefinitions>

    <Grid Grid.Row="0" Margin="5">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto" />
            <ColumnDefinition Width="*" />
        </Grid.ColumnDefinitions>      
        <TextBlock Grid.Column="0" Margin="5" Text="Search for:" VerticalAlignment="Center"/>
        <TextBox Grid.Column="1" Margin="5" x:Name="txtSearch" />
    </Grid>
    <ListBox x:Name="listbox" Grid.Row="1" Margin="5" DisplayMemberPath="DisplayName"> 
        <ListBox.ItemContainerStyle>
            <Style TargetType="{x:Type ListBoxItem}">
            </Style>
        </ListBox.ItemContainerStyle>
    </ListBox>
    <Grid Grid.Row="0" Grid.Column="1" Margin="5">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto" />
            <ColumnDefinition Width="*" />
        </Grid.ColumnDefinitions>
        <TextBlock Grid.Row="0" Grid.Column="0" Margin="10,5,5,5" Text = "Last operation Status:" />
        <TextBox Grid.Row="0" Grid.Column="1" Margin="5,5,10,5" x:Name="txtStatus" IsReadOnly="True" HorizontalContentAlignment="Right" Focusable="False" /> 

    </Grid>
    <Grid Grid.Row="1" Grid.Column="1">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="*" />            
        </Grid.RowDefinitions>
        <Separator Grid.Row="0" Margin="5" VerticalAlignment="Top" />  
            <Grid Grid.Row="1" Margin="5">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto" />
                    <RowDefinition Height="Auto" />
                    <RowDefinition Height="Auto" />
                    <RowDefinition Height="*" />
                 </Grid.RowDefinitions>        
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto" />
                    <ColumnDefinition Width="*" />
                </Grid.ColumnDefinitions>      
                <TextBlock Grid.Row="0" Grid.Column="0" Margin="10" Text = "Telephone Number:" />
                <TextBox Grid.Row="0" Grid.Column="1" Margin="10" x:Name="txtTelephone" Text = "{Binding ElementName=listbox,Path=SelectedItem.telephoneNumber,Mode=OneWay}" />
                <TextBlock Grid.Row="1" Grid.Column="0" Margin="10" Text = "Mobile Number:" />
                <TextBox Grid.Row="1" Grid.Column="1" Margin="10" x:Name="txtMobile" Text = "{Binding ElementName=listbox,Path=SelectedItem.MobilePhone,Mode=OneWay}"/>
                <Button Grid.Row="2" Grid.Column="1" Margin="10" x:Name="saveButton" Content="Save" />
                <Button Grid.Row="3" Grid.Column="1" Margin="10" x:Name="exportMobileButton" Content="Export Mobile Phones" VerticalAlignment="Bottom" />
            </Grid>
    </Grid>
  </Grid>
</Window>
"@


$reader=(New-Object System.Xml.XmlNodeReader $xaml)
$Window=[Windows.Markup.XamlReader]::Load( $reader )

#Connect to Controls
$textbox1 = $Window.FindName('txtSearch')
$listbox = $Window.FindName('listbox')
$telephoneBox = $Window.FindName('txtTelephone')
$mobileBox = $Window.FindName('txtMobile')
$saveButton = $Window.FindName('saveButton')
$txtStatus = $Window.FindName('txtStatus')
$exportMobileButton = $Window.FindName('exportMobileButton')

$Window.Add_Loaded({
    #Have to have something initially in the collection
    $Global:observableCollection = New-Object System.Collections.ObjectModel.ObservableCollection[System.Object]
    $listbox.ItemsSource = $observableCollection
    $observableCollection.Clear()


    $objDomain = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$ADSearchBase")
    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
    $objSearcher.SearchRoot = $objDomain
    $objSearcher.PageSize = 1000
    $strFilter = "(&(objectCategory=person)(objectClass=user))"
    $objSearcher.Filter = $strFilter
        
    $colProplist = @("displayName", "telephoneNumber", "Mobile", "distinguishedName")
    foreach ($i in $colPropList){$objSearcher.PropertiesToLoad.Add($i)}

    $objSearcher.FindAll() | Select @{e={$_.Properties.item("displayname")};n='DisplayName'},@{e={$_.Properties.item("telephoneNumber")};n='telephoneNumber'},@{e={$_.Properties.item("Mobile")};n='MobilePhone'},@{e={$_.Properties.item("distinguishedName")};n='distinguishedName'} | Sort-Object DisplayName | ForEach {
        $observableCollection.Add($_)
    }
})

#Events
$textbox1.Add_TextChanged({
    $filterText = $textbox1.Text + "*"
    $observableCollection = @($observableCollection | Where-Object {$_.DisplayName -like $filterText})
    $listbox.itemsSource = $observableCollection 
    #[System.Windows.Data.CollectionViewSource]::GetDefaultView( $Listbox.ItemsSource ).Refresh()
}) 
$saveButton.Add_Click({
    try {
        $SetTelephoneNumber = $telephoneBox.Text
        $SetMobilePhone = $mobileBox.Text
        $SelectedUser = $listbox.SelectedItem.distinguishedName
        $idOperator = [Environment]::UserName
        $logdate = get-date -format yyy-MM-dd
            
            $objUser = New-Object DirectoryServices.DirectoryEntry("LDAP://$SelectedUser")
            if ($SetTelephoneNumber.Length -lt 1){
                $SetTelephoneNumber = $null
                $objUser.PutEx(1, "telephoneNumber", 0)
            } Else {
                if ($SetTelephoneNumber.StartsWith("+")){
                    # Number on correct form
                }
                Else {
                    $SetTelephoneNumber = '+39' + $SetTelephoneNumber
                }
            $objUser.Put("telephoneNumber", $SetTelephoneNumber)
            }
            if ($SetMobilePhone.Length -lt 1){
                $SetMobilePhone = $null
                $objUser.PutEx(1, "mobile", 0)
            } Else {
                if ($SetMobilePhone.StartsWith("+")){
                    # Number on correct form
                }
                Else {
                    $SetMobilePhone = '+39' + $SetMobilePhone
                }
                $objUser.Put("mobile", $SetMobilePhone)
            }
            
            $objUser.SetInfo()
            $txtStatus.Text = "Saved $($listbox.SelectedItem.DisplayName)"
    }
    catch [System.SystemException] {
            $msg = $Error[0].Exception.Message
            [System.Windows.MessageBox]::Show($msg, 'Upating Status', 'OK', 'Error')
            $txtStatus.Text = "ERROR on $($listbox.SelectedItem.DisplayName)"
    }
    finally{
        $objUser = New-Object DirectoryServices.DirectoryEntry("LDAP://$SelectedUser")
        $tmpObject = $observableCollection[$observableCollection.IndexOf(($listbox.SelectedItem))]
        $tmpObject.telephoneNumber = [string]$objUser.telephoneNumber
        $tmpObject.mobilePhone = [string]$objUser.Mobile
        $index = $observableCollection.IndexOf(($listbox.SelectedItem))
        $observableCollection.RemoveAt($index)
        $observableCollection.Insert($index, $tmpObject)
        $listbox.SelectedItem = $listbox.Items.GetItemAt($index)

        #$observableCollection[$observableCollection.IndexOf(($listbox.SelectedItem))].telephoneNumber = [string]$objUser.telephoneNumber
        #$observableCollection[$observableCollection.IndexOf(($listbox.SelectedItem))].mobilePhone = [string]$objUser.Mobile
        
        $logline = $logdate + ";" + $SelectedUser + ";" + $SetTelephoneNumber + ";" + $SetMobilePhone + ";" + $idOperator
        $logline | out-file $logFileName -Append
    }
    
})

$exportMobileButton.Add_Click({
    try {
        $folderPicker = New-Object Microsoft.Win32.SaveFileDialog
        $folderPicker.Title = "Export CSV"
        $folderPicker.FileName = "ExportedMobile"
        $folderPicker.DefaultExt = ".csv" 
        $folderPicker.Filter = "CSV Documents (.csv)|*.csv"
        $folderPicker.InitialDirectory = [Environment]::GetFolderPath('Desktop')
        if($folderPicker.ShowDialog() -eq $true)
        {
            $exportPathFile = $folderPicker.FileName
            if ($exportPathFile -ne $null){
                $observableCollection | Where-Object {$_.MobilePhone -ne $null} | Select -Property 'DisplayName', 'MobilePhone'  | Export-Csv -Path ($exportPathFile) -NoTypeInformation -Delimiter ';'
            }
        }
    }
    Catch [System.SystemException] {
        $msg = $Error[0].Exception.Message
        [System.Windows.MessageBox]::Show($msg, 'Exporting Status', 'OK', 'Error')
    }
})


#$Window.ShowDialog() | Out-Null
$null = $window.Dispatcher.InvokeAsync{$window.ShowDialog()}.Wait()
