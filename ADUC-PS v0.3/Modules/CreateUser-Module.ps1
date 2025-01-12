#Imports the AD-module, comment/uncomment if/when nessecary:
Import-Module ActiveDirectory

Add-Type -AssemblyName System.Windows.Forms
    
    $userForm = New-Object System.Windows.Forms.Form
    $userForm.Text = "ADUC-PS v.0.3 - Create a user"
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

    #The label and textbox for users UPN, create a .csv-file and populate the file with what's in the textbox
    $userLabel3 = New-Object System.Windows.Forms.Label
    $userLabel3.Text = "Enter domain(after @):"
    $userLabel3.Location = New-Object System.Drawing.Point(10, 110)
    $userLabel3.AutoSize = $true
    $userForm.Controls.Add($userLabel3)
    $userTextBox3 = New-Object System.Windows.Forms.TextBox
    $userTextBox3.Location = New-Object System.Drawing.Point(10, 130)
    $userTextBox3.Size = New-Object System.Drawing.Size(200, 20)
    $userForm.Controls.Add($userTextBox3)

    $csvFilePath = Join-Path -Path $PSScriptRoot -ChildPath "Data\Domain.csv"

    # Ensure the CSV file exists and has a header
    if (-not (Test-Path $csvFilePath)) {
        "UPN" | Out-File -FilePath $csvFilePath -Encoding UTF8
    }

    # Read the first existing UPN value (if any) from the CSV file
    $currentUPN = ""
    if ((Get-Content $csvFilePath | Measure-Object).Count -gt 1) {
        $currentUPN = (Get-Content $csvFilePath)[1]  # Get the first line after the header
    }

    # Populate the textbox with the current UPN value, overwrite if changed
    $userTextBox3.Text = $currentUPN  # Assign the initial value after adding the TextBox to the form
    $userTextBox3.Add_TextChanged({
        $upn = $userTextBox3.Text
        "UPN" | Out-File -FilePath $csvFilePath -Encoding UTF8  # Write header
    if ($upn -ne "") {
        $upn | Out-File -FilePath $csvFilePath -Encoding UTF8 -Append
        }
    })

##############################

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
        $userLabel6 = New-Object System.Windows.Forms.Label
        $userLabel6.Text = "Select which OU to create the user in"
        $userLabel6.Location = New-Object System.Drawing.Point(10, 250)
        $userLabel6.AutoSize = $true
        $comboBox = New-Object System.Windows.Forms.ComboBox
        $comboBox.Location = New-Object System.Drawing.Point(10, 270)
        $comboBox.Size = New-Object System.Drawing.Size(200, 30)
        $comboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    
        #Populate ComboBox with Name, stores DistinguishedName in a hashtable
        $ouMapping = @{}
        foreach ($item in $data) {
            $comboBox.Items.Add($item.Name)
            $ouMapping[$item.Name] = $item.DistinguishedName
        }

        Write-Host "Populating ComboBox with the following data:"
        foreach ($item in $data) {
            Write-Host "Name: $($item.Name), DistinguishedName: $($item.DistinguishedName)"
        }
    
        #Retrieve DistinguishedName of the selected OU
        $comboBox.Add_SelectedIndexChanged({
            $selectedOUName = $comboBox.SelectedItem
            Write-Host "Selected OU Name: $selectedOUName"
            if ($ouMapping.ContainsKey($selectedOUName)) {
                $global:selectedDistinguishedName = $ouMapping[$selectedOUName]
                Write-Host "Selected Distinguished Name: $global:selectedDistinguishedName"
            } else {
                Write-Host "Error: Selected OU not found in the mapping."
            }
        })
    
        $userForm.Controls.Add($comboBox)
        $userForm.Controls.Add($userLabel6)
    } else {
        [System.Windows.Forms.MessageBox]::Show("The CSV file is empty or does not have valid data.", "Error")
    }
}

#Creates the combobox for selection of security group. Import the CSV data exported in the main script when first launched
$importFile2 = Join-Path -Path $PSScriptRoot -ChildPath "Data\SecGroups.csv"

if (Test-Path -Path $importFile2) {
    $data = Import-Csv -Path $importFile2

    if ($data) {
        $userLabel6 = New-Object System.Windows.Forms.Label
        $userLabel6.Text = "Select which Security Group the user is a member of"
        $userLabel6.Location = New-Object System.Drawing.Point(10, 290)
        $userLabel6.AutoSize = $true
        $comboBox2 = New-Object System.Windows.Forms.ComboBox
        $comboBox2.Location = New-Object System.Drawing.Point(10, 310)
        $comboBox2.Size = New-Object System.Drawing.Size(200, 30)
        $comboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    
        #Populate ComboBox with Name, stores DistinguishedName in a hashtable
        $secGroupMapping = @{}
        foreach ($item in $data) {
            $comboBox2.Items.Add($item.Name)
            $secGroupMapping[$item.Name] = $item.DistinguishedName
        }

        Write-Host "Populating ComboBox with the following data:"
        foreach ($item in $data) {
            Write-Host "Name: $($item.Name), DistinguishedName: $($item.DistinguishedName)"
        }
    
        $userForm.Controls.Add($comboBox2)
        $userForm.Controls.Add($userLabel6)
    } else {
        [System.Windows.Forms.MessageBox]::Show("The CSV file is empty or does not have valid data.", "Error")
    }
}

#Creates OK-button, pressing OK will run the input into separate variables
$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "Create user"
$okButton.Location = '390,220'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$userForm.Controls.Add($okButton)
$userForm.AcceptButton = $okButton

if ($userForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    # Retrieve input values
    $userInput1 = $userTextBox1.Text.Trim()
    $userInput2 = $userTextBox2.Text.Trim()
    $userInput3 = $userTextBox3.Text.Trim()
    $userUPN = "$($userInput1.ToLower()).$($userInput2.ToLower())@$($userInput3.ToLower())"
    $userInput4 = $userTextBox4.Text.Trim()

    # Sanitize inputs to remove invalid characters
    $userInput1 = $userInput1 -replace '[\\\/\[\]:;|=,+*?<>]', ''
    $userInput2 = $userInput2 -replace '[\\\/\[\]:;|=,+*?<>]', ''
    $userUPN = $userUPN -replace '[\\\/\[\]:;|=,+*?<>]', ''
    $SAMACCName = "$($userInput1.ToLower()).$($userInput2.ToLower())"

    # Validate inputs
    if (-not $userInput1 -or -not $userInput2 -or -not $userInput3 -or -not $userInput4) {
        Write-Host "Error: One or more input fields are empty."
        return
    }

    if ($userUPN -notmatch '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
        Write-Host "Error: Invalid UPN format: $userUPN"
        return
    }

    # Secure password
    $SecurePassword = ConvertTo-SecureString $userInput4 -AsPlainText -Force

    # Check for conflicts
    $existingUser = Get-ADUser -Filter {SamAccountName -eq $SAMACCName -or UserPrincipalName -eq $userUPN}
    if ($existingUser) {
        Write-Host "Error: A user with the same SAM Account Name or UPN already exists."
        return
    }

    # Create the AD user
    try {
        New-ADUser -Name "$userInput1 $userInput2" -GivenName $userInput1 -Surname $userInput2 `
            -SamAccountName $SAMACCName -UserPrincipalName $userUPN `
            -Enabled $true -AccountPassword $SecurePassword -ChangePasswordAtLogon $true
        Write-Host "User created successfully: $userUPN"
    } catch {
        Write-Host "Error creating user: $_"
        return
    }

    # Validate the selected OU
    if (-not $global:selectedDistinguishedName) {
        Write-Host "Error: No OU selected or '$global:selectedDistinguishedName' is null."
        return
    }

    # Move the user to the selected OU
    $User = Get-ADUser -Filter {SamAccountName -eq $SAMACCName}
    if ($User) {
        Write-Host "User found: $($User.DistinguishedName)"
        Write-Host "Moving user to: $global:selectedDistinguishedName"
        Move-ADObject -Identity $User.DistinguishedName -TargetPath $global:selectedDistinguishedName
    } else {
        Write-Host "Error: User '$SAMACCName' not found in the directory."
    }
    [System.Windows.Forms.MessageBox]::Show("User has been created, moved to the OU and added to the Security Group!")

    #Add the user to the chosen Security Group after creation and moving it
    $selectedSecurityGroup = $secGroupMapping[$comboBox2.SelectedItem]
    if ($selectedSecurityGroup) {
        Add-ADGroupMember -Identity $selectedSecurityGroup -Members $SAMACCName
        Write-Host "User $SAMACCName has been added to the group $selectedSecurityGroup"
    } else {
        Write-Host "Error: No valid security group selected."
}


}
