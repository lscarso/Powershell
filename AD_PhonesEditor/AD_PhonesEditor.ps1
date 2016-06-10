<# 
.SYNOPSIS 
Edit phones attribute on Active direcory Users Objects.

.DESCRIPTION 
Graphical APP for telephoneNumber and MobilePhone attibutes management. 

.NOTES
Edit $ADSearchBase for filtering users
Edit $logFileName for save modified record log

#>

# Script configuration
$ADSearchBase = "OU=My_Users,DC=My_Domain,DC=local"
$logFileName = "C:\Temp\PhoneEditor.log"

Add-Type -AssemblyName presentationframework
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
            <RowDefinition Height="Auto" />            
        </Grid.RowDefinitions>
        <Separator Grid.Row="0" Margin="5" VerticalAlignment="Top" />  
            <Grid Grid.Row="1" Margin="5">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto" />
                    <RowDefinition Height="Auto" />
                    <RowDefinition Height="Auto" />
                </Grid.RowDefinitions>        
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto" />
                    <ColumnDefinition Width="*" />
                </Grid.ColumnDefinitions>      
                <TextBlock Grid.Row="0" Grid.Column="0" Margin="10" Text = "Telephone Number:" />
                <TextBox Grid.Row="0" Grid.Column="1" Margin="10" x:Name="txtTelephone" Text = "{Binding ElementName=listbox,Path=SelectedItem.telephoneNumber}" />
                <TextBlock Grid.Row="1" Grid.Column="0" Margin="10" Text = "Mobile Number:" />
                <TextBox Grid.Row="1" Grid.Column="1" Margin="10" x:Name="txtMobile" Text = "{Binding ElementName=listbox,Path=SelectedItem.MobilePhone}"/>
                <Button Grid.Row="2" Grid.Column="1" Margin="10" x:Name="saveButton" Content="Save" />
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
    $SetTelephoneNumber = $telephoneBox.Text
    $SetMobilePhone = $mobileBox.Text
    $SelectedUser = $listbox.SelectedItem.distinguishedName
    $idOperator = [Environment]::UserName
    $logdate = get-date -format yyy-MM-dd
        
        $objUser = New-Object DirectoryServices.DirectoryEntry("LDAP://$SelectedUser")
        if ($SetTelephoneNumber.Length -lt 1){
            $SetTelephoneNumber = $null
            $objUser.PutEx(1, "telephoneNumber", 0)
        } Else {$objUser.Put("telephoneNumber", $SetTelephoneNumber)}
        if ($SetMobilePhone.Length -lt 1){
            $SetMobilePhone = $null
            $objUser.PutEx(1, "mobile", 0)
        } Else {$objUser.Put("mobile", $SetMobilePhone)}
        
        $objUser.SetInfo()

        $logline = $logdate + ";" + $SelectedUser + ";" + $SetTelephoneNumber + ";" + $SetMobilePhone + ";" + $idOperator
        
        if( -not $?){
            $msg = $Error[0].Exception.Message
            [windows.forms.messagebox]::Show("Error: $msg", "Updating Status")
        }
        Else {
            $txtStatus.Text = "Saved $($listbox.SelectedItem.DisplayName)"
            $logline | out-file $logFileName -Append
        }
})


$Window.ShowDialog() | Out-Null
