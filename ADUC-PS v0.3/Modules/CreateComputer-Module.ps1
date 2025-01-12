Import-Module ActiveDirectory
Add-Type -AssemblyName System.Windows.Forms

#Create the form and input field
$compObjForm = New-Object System.Windows.Forms.Form
$compObjForm.Text = "Create a Computer Object - ADUC-PS v.0.2"
$compObjForm.Size = New-Object System.Drawing.Size(800, 600)
$compObjForm.StartPosition = "CenterScreen"
$compObjLabel1 = New-Object System.Windows.Forms.Label
$compObjLabel1.Text = "Computers name:"
$compObjLabel1.Location = New-Object System.Drawing.Point(10, 30)
$compObjLabel1.AutoSize = $true
$compObjForm.Controls.Add($compObjLabel1)

$compObjTextBox1 = New-Object System.Windows.Forms.TextBox
$compObjTextBox1.Location = New-Object System.Drawing.Point(10, 50)
$compObjTextBox1.Size = New-Object System.Drawing.Size(200, 20)
$compObjForm.Controls.Add($compObjTextBox1)

#OU selection from .csv-file
$importFile = Join-Path -Path $PSScriptRoot -ChildPath "Data\OU_structure.csv"

if (Test-Path -Path $importFile) {
    $data = Import-Csv -Path $importFile
    if ($data) {
        $comoObjLabel2 = New-Object System.Windows.Forms.Label
        $comoObjLabel2.Text = "Select which OU to create the Computer Object in:"
        $comoObjLabel2.Location = New-Object System.Drawing.Point(10, 250)
        $comoObjLabel2.AutoSize = $true
        $compObjForm.Controls.Add($comoObjLabel2)

        $compObjComboBox = New-Object System.Windows.Forms.ComboBox
        $compObjComboBox.Location = New-Object System.Drawing.Point(10, 270)
        $compObjComboBox.Size = New-Object System.Drawing.Size(200, 30)
        $compObjComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
        foreach ($item in $data) {
            $compObjComboBox.Items.Add($item.Name)
        }
        $compObjForm.Controls.Add($compObjComboBox)
    } else {
        [System.Windows.Forms.MessageBox]::Show("The CSV file is empty or does not have valid data.", "Error")
    }
} else {
    [System.Windows.Forms.MessageBox]::Show("The CSV file was not found at: $importFile", "Error")
}

#Create-button
$compObjOkButton = New-Object System.Windows.Forms.Button
$compObjOkButton.Text = "Create"
$compObjOkButton.Location = '390,220'
$compObjOkButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$compObjForm.Controls.Add($compObjOkButton)
$compObjForm.AcceptButton = $compObjOkButton

if ($compObjForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $compName = $compObjTextBox1.Text
    $selectedOU = $compObjComboBox.SelectedItem

 #Check if the form inputs are filled
if (-not $compName -or -not $selectedOU) {
    Write-Host "Error: All fields must be filled out." -ForegroundColor Red
    exit
}

#Remove invalid characters
$userInput1 = $compName -replace '[\\\/\[\]:;|=,+*?<>]', ''
if (-not $userInput1) {
    Write-Host "Error: One or more input fields are empty."
    return
}

#Convert the selected OU to its Distinguished Name(copy-paste from user/secgroup)
$selectedOU = $data | Where-Object { $_.Name -eq $compObjComboBox.SelectedItem } | Select-Object -ExpandProperty DistinguishedName

if (-not $selectedOU) {
    Write-Host "Error: Could not resolve the selected OU to a Distinguished Name." -ForegroundColor Red
    exit
}

#Creation of the computer object
try {
    New-ADComputer -Name $userInput1 -SamAccountName $userInput1 -DisplayName $userInput1 -Path $selectedOU
    Write-Host "Computer Object '$compName' created successfully in OU '$selectedOU'."
} catch {
    Write-Host "Error creating computer object: $_" -ForegroundColor Red
}

}
