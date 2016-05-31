<# 
.SYNOPSIS 
Edit phones attribute on Active direcory Users Objects.

.DESCRIPTION 
Graphical APP for telephoneNumber and MobilePhone attibutes management. 

.NOTES
Edit $ADSearchBase for filtering users
Author: Luca Scarsini

#>
 
# Script configuration
$ADSearchBase = "OU=My_Users,DC=My_Domain,DC=local"
 
#----------------------------------------------
# Generated Form Function
#----------------------------------------------
function GenerateForm {
 
    #----------------------------------------------
    #region Import Assemblies
    #----------------------------------------------
    [void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
    [void][reflection.assembly]::Load("System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
    [void][reflection.assembly]::Load("mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
    [void][reflection.assembly]::Load("System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
    #endregion
    
    Import-Module ActiveDirectory 
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $form1 = New-Object System.Windows.Forms.Form
    $btnSave = New-Object System.Windows.Forms.Button
    $btnExport = New-Object System.Windows.Forms.Button
    
    $label3 = New-Object System.Windows.Forms.Label
    $label2 = New-Object System.Windows.Forms.Label
    $label1 = New-Object System.Windows.Forms.Label

    $MobilePhone = New-Object System.Windows.Forms.TextBox
    $telephoneNumber = New-Object System.Windows.Forms.TextBox

    $listbox1 = New-Object System.Windows.Forms.ListBox
    
    $InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
 
    $FormEvent_Load={
            $results = Get-ADUser -Filter * -SearchScope Subtree -SearchBase $ADSearchBase -Properties DisplayName, telephoneNumber, MobilePhone
             
            foreach ($result in $results) {
                $listbox1.Items.Add($result.DisplayName)
            }
    }
     
    $handler_listbox1_SelectedIndexChanged={
            $entry = Get-ADUser -filter {DisplayName -eq $listbox1.Text} -Properties DisplayName, telephoneNumber, MobilePhone
      
            $telephoneNumber.Text = $entry.telephoneNumber
            $MobilePhone.Text = $entry.MobilePhone
    }
    
    $handler_btnSave_Click={
            # Save changes in AD
            $SetTelephoneNumber = $telephoneNumber.Text
            $SetMobilePhone = $MobilePhone.Text

            if ($SetTelephoneNumber.Length -lt 1){
                $SetTelephoneNumber = $null
            }
            if ($SetMobilePhone.Length -lt 1){
                $SetMobilePhone = $null
            }
            
            $entry = Get-ADUser -filter {DisplayName -eq $listbox1.Text} | Set-ADUser -OfficePhone $SetTelephoneNumber -MobilePhone $SetMobilePhone
            if( -not $?){
                $msg = $Error[0].Exception.Message
                [System.Windows.Forms.MessageBox]::Show("Error: $msg", "Updating Status")
            }
    }
    
    $handle_btnExport_Click={
        # Export Telephone Number CSV
        $objShell = new-object -com shell.application
        $objFolder = $objShell.NameSpace("Home")
        $namedfolder = $objShell.BrowseForFolder(0,"Select the folder where you what to export the list:",0,5)
        Get-ADGroupMember O-008000-OM_Users | where-object {$_.DistinguishedName -like "*$ADSearchBase"} | Get-ADUser -Properties DisplayName, telephoneNumber, Mobile, l | Where-Object {$_.telephoneNumber -ne $null -or $_.Mobile -ne $null} | Sort-Object l | Select-Object @{n="Nome";e={$_.DisplayName}}, @{n="Telefono Fisso";e={$_.telephoneNumber}}, @{n="Cellulare";e={$_.Mobile}}, @{n="Sede";e={$_.l}} | Export-Csv -Delimiter ";" -Path "$($namedfolder.self.path)\ElencoTelefonico.csv" -NoTypeInformation | out-null
        if( -not $?){
            $msg = $Error[0].Exception.Message
            [System.Windows.Forms.MessageBox]::Show("Error: $msg", "Exporting Status")
        }Else{
            [System.Windows.Forms.MessageBox]::Show("File created on $($namedfolder.self.path)", "Exporting Status")
        }
    }
     
    $Form_StateCorrection_Load={
            #Correct the initial state of the form to prevent the .Net maximized form issue
            $form1.WindowState = $InitialFormWindowState
    }
     
    #----------------------------------------------
    #region Generated Form Code
    #----------------------------------------------
    #
    # form1
    #
    $form1.Controls.Add($btnSave)
    $form1.Controls.Add($btnExport)
    $form1.Controls.Add($label3)
    $form1.Controls.Add($label2)
    $form1.Controls.Add($label1)
    $form1.Controls.Add($MobilePhone)
    $form1.Controls.Add($telephoneNumber)
    $form1.Controls.Add($listbox1)
    $form1.Text = "AD Phones Attribute Editor v1.0"
    $form1.Name = "form1"
    $form1.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $form1.ClientSize = New-Object System.Drawing.Size(606,465)
    $form1.add_Load($FormEvent_Load)
    #
    # btnSave
    #
    $btnSave.TabIndex = 3
    $btnSave.Name = "btnSave"
    $btnSave.Size = New-Object System.Drawing.Size(120,23)
    $btnSave.UseVisualStyleBackColor = $True
    $btnSave.Text = "Save changes"
    $btnSave.Location = New-Object System.Drawing.Point(474,100)
    $btnSave.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $btnSave.add_Click($handler_btnSave_Click)
    #
    # btnExport
    #
    $btnExport.TabIndex = 4
    $btnExport.Name = "btnExport"
    $btnExport.Size = New-Object System.Drawing.Size(120,23)
    $btnExport.UseVisualStyleBackColor = $True
    $btnExport.Text = "Export CSV"
    $btnExport.Location = New-Object System.Drawing.Point(474,430)
    $btnExport.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $btnExport.add_Click($handle_btnExport_Click)
    #
    # custom2
    #
    $MobilePhone.Size = New-Object System.Drawing.Size(300,20)
    $MobilePhone.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $MobilePhone.Name = "custom2"
    $MobilePhone.Location = New-Object System.Drawing.Point(294,59)
    $MobilePhone.TabIndex = 2
    #
    # custom1
    #
    $telephoneNumber.Size = New-Object System.Drawing.Size(300,20)
    $telephoneNumber.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $telephoneNumber.Name = "custom1"
    $telephoneNumber.Location = New-Object System.Drawing.Point(294,33)
    $telephoneNumber.TabIndex = 1
    #
    # label Mobile Phone
    #
    $label3.TabIndex = 18
    $label3.Size = New-Object System.Drawing.Size(115,19)
    $label3.Text = "Mobile Phone"
    $label3.Location = New-Object System.Drawing.Point(173,63)
    $label3.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $label3.Name = "label3"
    #
    # label Office Phone
    #
    $label2.TabIndex = 2
    $label2.Size = New-Object System.Drawing.Size(100,19)
    $label2.Text = "Office Phone"
    $label2.Location = New-Object System.Drawing.Point(173,36)
    $label2.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $label2.Name = "label2"
    #
    # listbox User List
    #
    $listbox1.FormattingEnabled = $True
    $listbox1.Size = New-Object System.Drawing.Size(155,420)
    $listbox1.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $listbox1.Name = "listbox1"
    $listbox1.Location = New-Object System.Drawing.Point(12,33)
    $listbox1.Sorted = $True
    $listbox1.TabIndex = 0
    $listbox1.add_SelectedIndexChanged($handler_listbox1_SelectedIndexChanged)
    #
    # label Users List
    #
    $label1.TabIndex = 1
    $label1.Size = New-Object System.Drawing.Size(174,23)
    $label1.Text = "Please select a user:"
    $label1.Location = New-Object System.Drawing.Point(12,9)
    $label1.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation 
    $label1.Name = "label1"
    #endregion Generated Form Code
 
    #----------------------------------------------
 
    #Save the initial state of the form
    $InitialFormWindowState = $form1.WindowState
    #Init the OnLoad event to correct the initial state of the form
    $form1.add_Load($Form_StateCorrection_Load)
    #Show the Form
    return $form1.ShowDialog()
 
} #End Function
 


#Create the form
GenerateForm | Out-Null
