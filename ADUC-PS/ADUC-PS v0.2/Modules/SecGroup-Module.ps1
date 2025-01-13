# Import the AD module
Import-Module ActiveDirectory

Add-Type -AssemblyName System.Windows.Forms

#Create the form
$secGrpForm = New-Object System.Windows.Forms.Form
$secGrpForm.Text = "Create a Security Group - ADUC-PS v.0.2"
$secGrpForm.Size = New-Object System.Drawing.Size(800, 600)
$secGrpForm.StartPosition = "CenterScreen"

# Security group name input
$secGrpLabel1 = New-Object System.Windows.Forms.Label
$secGrpLabel1.Text = "Security group's name:"
$secGrpLabel1.Location = New-Object System.Drawing.Point(10, 30)
$secGrpLabel1.AutoSize = $true
$secGrpForm.Controls.Add($secGrpLabel1)

$secGrpTextBox1 = New-Object System.Windows.Forms.TextBox
$secGrpTextBox1.Location = New-Object System.Drawing.Point(10, 50)
$secGrpTextBox1.Size = New-Object System.Drawing.Size(200, 20)
$secGrpForm.Controls.Add($secGrpTextBox1)

# Group scope selection
$secGrpLabel2 = New-Object System.Windows.Forms.Label
$secGrpLabel2.Text = "Select the group's scope:"
$secGrpLabel2.Location = New-Object System.Drawing.Point(10, 70)
$secGrpLabel2.AutoSize = $true
$secGrpForm.Controls.Add($secGrpLabel2)

$secComboBox = New-Object System.Windows.Forms.ComboBox
$secComboBox.Location = New-Object System.Drawing.Point(10, 100)
$secComboBox.Width = 200
$secComboBox.Items.AddRange(@("Global", "Universal", "DomainLocal"))
$secComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$secGrpForm.Controls.Add($secComboBox)

# OU selection from CSV
$importFile = Join-Path -Path $PSScriptRoot -ChildPath "Data\OU_structure.csv"

if (Test-Path -Path $importFile) {
    $data = Import-Csv -Path $importFile
    if ($data) {
        $secGrpLabel3 = New-Object System.Windows.Forms.Label
        $secGrpLabel3.Text = "Select which OU to create the security group in:"
        $secGrpLabel3.Location = New-Object System.Drawing.Point(10, 250)
        $secGrpLabel3.AutoSize = $true
        $secGrpForm.Controls.Add($secGrpLabel3)

        $comboBox2 = New-Object System.Windows.Forms.ComboBox
        $comboBox2.Location = New-Object System.Drawing.Point(10, 270)
        $comboBox2.Size = New-Object System.Drawing.Size(200, 30)
        $comboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
        foreach ($item in $data) {
            $comboBox2.Items.Add($item.Name)
        }
        $secGrpForm.Controls.Add($comboBox2)
    } else {
        [System.Windows.Forms.MessageBox]::Show("The CSV file is empty or does not have valid data.", "Error")
    }
} else {
    [System.Windows.Forms.MessageBox]::Show("The CSV file was not found at: $importFile", "Error")
}

# Create button
$secOkButton = New-Object System.Windows.Forms.Button
$secOkButton.Text = "Create"
$secOkButton.Location = '390,220'
$secOkButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$secGrpForm.Controls.Add($secOkButton)
$secGrpForm.AcceptButton = $secOkButton

if ($secGrpForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $groupName = $secGrpTextBox1.Text
    $selectedScope = $secComboBox.SelectedItem
    $selectedOU = $comboBox2.SelectedItem

    if (-not $groupName -or -not $selectedScope -or -not $selectedOU) {
        Write-Host "Error: All fields must be filled out." -ForegroundColor Red
        exit
    }

    #Convert scope to AD-compatible values.. Apparently this was needed at one point. Keeping it.
    $scopeMapping = @{
        "Global"       = "Global"
        "Universal"    = "Universal"
        "DomainLocal"  = "DomainLocal"
    }
    $adScope = $scopeMapping[$selectedScope]

#Convert the selected OU to its Distinguished Name
$selectedOU = $data | Where-Object { $_.Name -eq $comboBox2.SelectedItem } | Select-Object -ExpandProperty DistinguishedName

if (-not $selectedOU) {
    Write-Host "Error: Could not resolve the selected OU to a Distinguished Name." -ForegroundColor Red
    exit
}

#Creates the group
try {
    New-ADGroup -Name $groupName -SamAccountName $groupName -GroupCategory Security -GroupScope $adScope -DisplayName $groupName -Path $selectedOU
    Write-Host "Security group '$groupName' created successfully in OU '$selectedOU'."
} catch {
    Write-Host "Error creating group: $_" -ForegroundColor Red
}

}
