#Imports the AD-module, comment/uncomment if/when nessecary:
Import-Module ActiveDirectory

Add-Type -AssemblyName System.Windows.Forms
    
$secGrpForm = New-Object System.Windows.Forms.Form #- WHY DOES IT OPEN TWICE ?? fml
$secGrpForm.Text = "Create a Security Group - ADUC-PS v.0.1"
$secGrpForm.Size = New-Object System.Drawing.Size(800, 600)
$secGrpForm.StartPosition = "CenterScreen"

#The label and textbox for secGrp name
$secGrpLabel1 = New-Object System.Windows.Forms.Label
$secGrpLabel1.Text = "Security groups name:"
$secGrpLabel1.Location = New-Object System.Drawing.Point(10, 30)
$secGrpLabel1.AutoSize = $true
$secGrpForm.Controls.Add($secGrpLabel1)
$secGrpTextBox1 = New-Object System.Windows.Forms.TextBox
$secGrpTextBox1.Location = New-Object System.Drawing.Point(10, 50)
$secGrpTextBox1.Size = New-Object System.Drawing.Size(200, 20)
$secGrpForm.Controls.Add($secGrpTextBox1)
        
#The label and combobox for secGrp scope
$secGrpLabel1 = New-Object System.Windows.Forms.Label
$secGrpLabel1.Text = "Select the groups scope:"
$secGrpLabel1.Location = New-Object System.Drawing.Point(10, 70)
$secGrpLabel1.AutoSize = $true
$secGrpForm.Controls.Add($secGrpLabel1)
$secComboBox = New-Object System.Windows.Forms.ComboBox
$secComboBox.Location = New-Object System.Drawing.Point(10, 100)
$secComboBox.Width = 200
$secComboBox.Items.AddRange(@("Global", "Universal", "Domain local"))
$secGrpForm.Controls.Add($secComboBox)
$selectedItem = $secComboBox.SelectedItem

#Create the combobox for selection of OU. Import the CSV data exported in the main script when first launched
$importFile = Join-Path -Path $PSScriptRoot -ChildPath "Data\OU_structure.csv"

if (Test-Path -Path $importFile) {
    $data = Import-Csv -Path $importFile

    if ($data) {

        $secGrpLabel2 = New-Object System.Windows.Forms.Label
        $secGrpLabel2.Text = "Select which OU to create the security group in"
        $secGrpLabel2.Location = New-Object System.Drawing.Point(10, 250)
        $secGrpLabel2.AutoSize = $true
        $comboBox2 = New-Object System.Windows.Forms.ComboBox
        $comboBox2.Location = New-Object System.Drawing.Point(10, 270)
        $comboBox2.Size = New-Object System.Drawing.Size(200, 30)
        $comboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

        foreach ($item in $data) {
            $comboBox.Items.Add($item.Name)
        }

        $secGrpForm.Controls.Add($comboBox2)
        $secGrpForm.Controls.Add($secGrpLabel2)
    } else {
        [System.Windows.Forms.MessageBox]::Show("The CSV file is empty or does not have valid data.", "Error")
    }
} else {
    [System.Windows.Forms.MessageBox]::Show("The CSV file was not found at: $importFile", "Error")
}

#OK-button creating a the group
$secOkButton = New-Object System.Windows.Forms.Button
$secOkButton.Text = "Create"
$secOkButton.Location = '390,220'
$secOkButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$secGrpForm.Controls.Add($secOkButton)
$secGrpForm.AcceptButton = $secOkButton

if ($secGrpForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $userInput5 = $secGrpTextBox1.Text
    Write-Host "$userInput5"
    
#@TO-DO.. Make a better OU-solution.. Just throwing in some parts of a older script and slightly modifying it as a start ...
    New-ADGroup $userInput5 -SamAccountName $userInput5 -GroupCategory Security`
    -GroupScope $selectedItem -DisplayName -$userInput5 -Path $secGrpOU
    $Group = Get-ADGroup -Filter {Name -eq $GroupName} 
    if ($Group) { 
        if ($Dept -eq "IT") { 
            Move-ADObject -Identity $Group.DistinguishedName -TargetPath "OU=IT-Groups,OU=IT,OU=fjellstad.local,DC=fjellstad,DC=local" 
            Write-Host "Gruppen er blitt flyttet til IT-Groups-OU'en!" 
        } elseif ($Dept -eq "Management") { 
            Move-ADObject -Identity $Group.DistinguishedName -TargetPath "OU=MGM-Groups,OU=Management,OU=fjellstad.local,DC=fjellstad,DC=local" 
            Write-Host "Gruppen er blitt flyttet til MGM-Groups-OU'en!" 
        } elseif ($Dept -eq "Auditors") { 
            Move-ADObject -Identity $Group.DistinguishedName -TargetPath "OU=AUD-Groups,OU=Auditors,OU=fjellstad.local,DC=fjellstad,DC=local" 
            Write-Host "Gruppen ble flyttet til AUD-Groups-OU'en!" 
        } elseif ($Dept -eq "Technical") { 
            Move-ADObject -Identity $Group.DistinguishedName -TargetPath "OU=TEC-Groups,OU=Technical,OU=fjellstad.local,DC=fjellstad,DC=local" 
            Write-Host "Gruppen ble flyttet til TEC-Groups-OU'en!" 
        } else { 
            Write-Host "Ukjent avdeling! Gruppen som ble opprettet er nå blitt slettet igjen :( Prøv på nytt!" 
            Remove-ADGroup $GroupName -Confirm:$False 
        } 
    } 
}       
$secGrpForm.ShowDialog()