#Imports the AD-module, comment/uncomment if/when nessecary:
Import-Module ActiveDirectory

Add-Type -AssemblyName System.Windows.Forms
    
    $userForm = New-Object System.Windows.Forms.Form
    $userForm.Text = "Create a user - ADUC-PS v.0.1"
    $userForm.Size = New-Object System.Drawing.Size(800, 600)
    $userForm.StartPosition = "CenterScreen"

    #The label and textbox for users first name
    $userLabel1 = New-Object System.Windows.Forms.Label
    $userLabel1.Text = "Users first name:"
    $userLabel1.Location = New-Object System.Drawing.Point(10, 30)
    $userLabel1.AutoSize = $true
    $userForm.Controls.Add($userLabel1)
    $userTextBox1 = New-Object System.Windows.Forms.TextBox
    $userTextBox1.Location = New-Object System.Drawing.Point(10, 50)
    $userTextBox1.Size = New-Object System.Drawing.Size(200, 20)
    $userForm.Controls.Add($userTextBox1)

    #The label and textbox for users last name
    $userLabel2 = New-Object System.Windows.Forms.Label
    $userLabel2.Text = "Users last name:"
    $userLabel2.Location = New-Object System.Drawing.Point(10, 70)
    $userLabel2.AutoSize = $true
    $userForm.Controls.Add($userLabel2)
    $userTextBox2 = New-Object System.Windows.Forms.TextBox
    $userTextBox2.Location = New-Object System.Drawing.Point(10, 90)
    $userTextBox2.Size = New-Object System.Drawing.Size(200, 20)
    $userForm.Controls.Add($userTextBox2)

    #The label and textbox for users UPN
    $userLabel3 = New-Object System.Windows.Forms.Label
    $userLabel3.Text = "Enter domain(login for user, not local):"
    $userLabel3.Location = New-Object System.Drawing.Point(10, 110)
    $userLabel3.AutoSize = $true
    $userForm.Controls.Add($userLabel3)
    $userTextBox3 = New-Object System.Windows.Forms.TextBox
    $userTextBox3.Location = New-Object System.Drawing.Point(10, 130)
    $userTextBox3.Size = New-Object System.Drawing.Size(200, 20)
    $userForm.Controls.Add($userTextBox3)

    #The label and textbox for users password
    $userLabel4 = New-Object System.Windows.Forms.Label
    $userLabel4.Text = "Users password for first logon:"
    $userLabel4.Location = New-Object System.Drawing.Point(10, 150)
    $userLabel4.AutoSize = $true
    $userForm.Controls.Add($userLabel4)
    $userTextBox4 = New-Object System.Windows.Forms.TextBox
    $userTextBox4.Location = New-Object System.Drawing.Point(10, 170)
    $userTextBox4.Size = New-Object System.Drawing.Size(200, 20)
    $userForm.Controls.Add($userTextBox4)

# @TO-DO: This has no use atm. add later.
    #The label and textbox for users department.
    $userLabel5 = New-Object System.Windows.Forms.Label
    $userLabel5.Text = "Users department:"
    $userLabel5.Location = New-Object System.Drawing.Point(10, 190)
    $userLabel5.AutoSize = $true
    $userForm.Controls.Add($userLabel5)
    $userTextBox5 = New-Object System.Windows.Forms.TextBox
    $userTextBox5.Location = New-Object System.Drawing.Point(10, 210)
    $userTextBox5.Size = New-Object System.Drawing.Size(200, 20)
    $userForm.Controls.Add($userTextBox5)

#Creates the combobox for selection of OU. Import the CSV data exported in the main script when first launched
$importFile = Join-Path -Path $PSScriptRoot -ChildPath "Data\OU_structure.csv"

if (Test-Path -Path $importFile) {
    $data = Import-Csv -Path $importFile

    if ($data) {
        # Create a label for the ComboBox
        $userLabel6 = New-Object System.Windows.Forms.Label
        $userLabel6.Text = "Select which OU to create the user in"
        $userLabel6.Location = New-Object System.Drawing.Point(10, 250)
        $userLabel6.AutoSize = $true

        # Create the ComboBox
        $comboBox = New-Object System.Windows.Forms.ComboBox
        $comboBox.Location = New-Object System.Drawing.Point(10, 270)
        $comboBox.Size = New-Object System.Drawing.Size(200, 30)
        $comboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

        # Populate ComboBox with Name, stores DistinguishedName in a hashtable
        $ouMapping = @{}
        foreach ($item in $data) {
            $comboBox.Items.Add($item.Name)
            $ouMapping[$item.Name] = $item.DistinguishedName
        }

        # Event to retrieve DistinguishedName of the selected OU
        $comboBox.Add_SelectedIndexChanged({
            $selectedOUName = $comboBox.SelectedItem
            if ($ouMapping.ContainsKey($selectedOUName)) {
                $selectedDistinguishedName = $ouMapping[$selectedOUName]
                Write-Host "Selected Distinguished Name: $selectedDistinguishedName"
            }
        })

        # Add the controls to the form
        $userForm.Controls.Add($comboBox)
        $userForm.Controls.Add($userLabel6)
    } else {
        [System.Windows.Forms.MessageBox]::Show("The CSV file is empty or does not have valid data.", "Error")
    }
} else {
    [System.Windows.Forms.MessageBox]::Show("The CSV file was not found at: $importFile", "Error")
}


    #Creates OK-button, pressing OK will run the input into separate variables
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "Create user"
    $okButton.Location = '390,220'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $userForm.Controls.Add($okButton)
    $userForm.AcceptButton = $okButton

    if ($userForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $userInput = $textBox1.Text
        $userInput2 = $textBox2.Text
        $userInput3 = $textBox3.Text
        $userUPN = $userInput1.ToLower()+"."+$userInput2.ToLower()+"@"+$userInput3.ToLower()
        $userInput4 = $textBox4.Text
        
        #Converting password to secure string:
        $SecurePassword = ConvertTo-SecureString $userInput4 -AsPlainText -Force
        
        #Used for searching after the creation of the user
        $SAMACCName = $userInput1.ToLower()+"."+$userInput2.ToLower()
        
        #Creating the AD-user:
        New-ADUser -Name "$userInput1 $userInput2" -GivenName $userInput -Surname $userInput2 -SamAccountName $SAMACCName -UserPrincipalName $userUPN -Enabled $true -AccountPassword $SecurePassword -ChangePasswordAtLogon = $true `
    
        #@TO-DO
        ######################
        #Checking for the existence of the user created, searching everything based on SAM account name
        $User = Get-ADUser -Filter {SAMAccountName -eq $SAMACCName}
        #Moves the user to the desired OU
         if ($user) {
             Write-Host "User found: $($User.DistinguishedName)"
             Write-Host "User has been moved to $OrgUnit !"
             Move-ADObject -Identity $user.DistinguishedName -TargetPath $selectedDistinguishedName
         } else {
             Write-Host "User $SAMACCName not found in the directory."
         }
        ######################

        #Not a real confirmation atm..
        [System.Windows.Forms.MessageBox]::Show("User have been created!")
    
        #Testing
        Write-Host "Users first name is: $userInput1"
        Write-Host "Users last name is: $userInput2"
        Write-Host "Users UPN is: $UserUPN"
        Write-Host "Users first password is: $userInput4"

        $userForm.Close()
        Exit
}