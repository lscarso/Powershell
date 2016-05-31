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
$ADSearchBase = "OU=100_Personal,OU=800_OM,OU=100_Users,OU=KION Objects,DC=d400,DC=mh,DC=grp"
$logFileName = "\\DEFRKIM0196.d400.mh.grp\Sys_Data1$\IT8000_Software\02_ReceptionScript\Log\PhoneEditor.log"

Add-Type -AssemblyName presentationframework
#Add-Type -AssemblyName System.DirectoryServices.AccountManagement
Import-Module ActiveDirectory


[xml]$xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:sys="clr-namespace:System;assembly=mscorlib"
    Title="STILL Italy - Phone Editor" Width="550" Height="500" 
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
    Get-ADUser -Filter * -SearchScope Subtree -SearchBase $ADSearchBase -Properties DisplayName, telephoneNumber, MobilePhone | Sort-Object DisplayName | Select-Object DisplayName, telephoneNumber, MobilePhone | ForEach {
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
    $SelectedUser = $listbox.SelectedItem.DisplayName
    $idOperator = [Environment]::UserName
    $logdate = get-date -format yyy-MM-dd
        if ($SetTelephoneNumber.Length -lt 1){
            $SetTelephoneNumber = $null
        }
        if ($SetMobilePhone.Length -lt 1){
            $SetMobilePhone = $null
        }
        $logline = $logdate + ";" + $SelectedUser + ";" + $SetTelephoneNumber + ";" + $SetMobilePhone + ";" + $idOperator
        Get-ADUser -filter {DisplayName -eq $SelectedUser} | Set-ADUser -OfficePhone $SetTelephoneNumber -MobilePhone $SetMobilePhone
        if( -not $?){
            $msg = $Error[0].Exception.Message
            [windows.forms.messagebox]::Show("Error: $msg", "Updating Status")
        }
        Else {
            $txtStatus.Text = "Saved $SelectedUser"
            $logline | out-file $logFileName -Append
        }
})


$Window.ShowDialog() | Out-Null
