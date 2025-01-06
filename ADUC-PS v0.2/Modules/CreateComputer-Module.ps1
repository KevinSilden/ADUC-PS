#Imports the AD-module, comment/uncomment if/when nessecary:
Import-Module ActiveDirectory

Add-Type -AssemblyName System.Windows.Forms
    
$compObjForm = New-Object System.Windows.Forms.Form #- WHY DOES IT OPEN TWICE ?? fml
$compObjForm.Text = "Create a Computer Object - ADUC-PS v.0.1"
$compObjForm.Size = New-Object System.Drawing.Size(800, 600)
$compObjForm.StartPosition = "CenterScreen"

#The label and textbox for compGroup name
$compObjLabel1 = New-Object System.Windows.Forms.Label
$compObjLabel1.Text = "Computers name:"
$compObjLabel1.Location = New-Object System.Drawing.Point(10, 30)
$compObjLabel1.AutoSize = $true
$compObjForm.Controls.Add($compObjLabel1)
$compObjTextBox1 = New-Object System.Windows.Forms.TextBox
$compObjTextBox1.Location = New-Object System.Drawing.Point(10, 50)
$compObjTextBox1.Size = New-Object System.Drawing.Size(200, 20)
$compObjForm.Controls.Add($compObjTextBox1)

#Create the combobox for selection of OU. Import the CSV data exported in the main script when first launched. Need to silently import the OU-path as well and pass this when creating object
$importFileComputer = Join-Path -Path $PSScriptRoot -ChildPath "Data\OU_structure.csv"

if (Test-Path -Path $importFileComputer) {
    $data = Import-Csv -Path $importFileComputer

    if ($data) {

        $compObjLabel2 = New-Object System.Windows.Forms.Label
        $compObjLabel2.Text = "Select which OU to create the computer object in"
        $compObjLabel2.Location = New-Object System.Drawing.Point(10, 250)
        $compObjLabel2.AutoSize = $true
        $comboBox = New-Object System.Windows.Forms.ComboBox
        $comboBox.Location = New-Object System.Drawing.Point(10, 270)
        $comboBox.Size = New-Object System.Drawing.Size(200, 30)
        $comboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

        foreach ($item in $data) {
            $comboBox.Items.Add($item.Name)
        }

        $compObjForm.Controls.Add($comboBox)
        $compObjForm.Controls.Add($compObjLabel2)

    } else {
        [System.Windows.Forms.MessageBox]::Show("The CSV file is empty or does not have valid data.", "Error")
    }
} else {
    [System.Windows.Forms.MessageBox]::Show("The CSV file was not found at: $importFile", "Error")
}

#OK-button creating the computer object
$compOkButton = New-Object System.Windows.Forms.Button
$compOkButton.Text = "Create"
$compOkButton.Location = '390,220'
$compOkButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$compObjForm.Controls.Add($compOkButton)
$compObjForm.AcceptButton = $compOkButton

#'TO-DO - create a solution for OU's...

if ($compObjForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $userInput6 = $secGrpTextBox1.Text
    Write-Host "$userInput6"
    New-ADComputer $userInput6
}       
$compObjForm.ShowDialog()