<!-- Top right panel -->
<Grid Grid.Row="0" Grid.Column="1" Background="Black">
    <TextBlock x:Name="topRightUserTitlePanel" Text="Create a user" FontWeight="Bold" Foreground="White" 
        HorizontalAlignment="Center" VerticalAlignment="Center"
        FontSize="24" Visibility="Collapsed"/>

<Grid Grid.Row="0" Grid.Column="1" Background="Black">
    <TextBlock x:Name="topRightCsvImportTitle" Text="Import User(s) with .csv" FontWeight="Bold" Foreground="White" 
        HorizontalAlignment="Center" VerticalAlignment="Center"
        FontSize="24" Visibility="Collapsed"/>

<Grid Grid.Row="0" Grid.Column="1" Background="Black">
    <TextBlock x:Name="topRightSecGroupTitle" Text="Create a Security Group" FontWeight="Bold" Foreground="White" 
        HorizontalAlignment="Center" VerticalAlignment="Center"
        FontSize="24" Visibility="Collapsed"/>

<Grid Grid.Row="0" Grid.Column="1" Background="Black">
    <TextBlock x:Name="topRightCompObjectTitle" Text="Create a Computer Object" FontWeight="Bold" Foreground="White" 
        HorizontalAlignment="Center" VerticalAlignment="Center"
        FontSize="24" Visibility="Collapsed"/>






#Imports the AD-module, uncomment if nessecary:
#Import-Module ActiveDirectory

#Define the OU path:    @TO-DO: Try to incorporate this in the GUI somehow.. Like another file or something.
$secGrpOU = "OU=,DC=,DC="
$singleUserOU = "OU=,DC=,DC="
$csvOU = "OU=,DC=,DC="

#Load the Windows Forms assembly
Add-Type -AssemblyName System.Windows.Forms

#Creates the Form(GUI-window), centers it on screen.
$form = New-Object System.Windows.Forms.Form
$form.Text = "ADUC-PS v.0.1"
$form.Size = New-Object System.Drawing.Size(500, 300)
$form.StartPosition = "CenterScreen"

#Text @ the top of the window
$label = New-Object System.Windows.Forms.Label
$label.Text = "What you wanna do?"
$label.Location = New-Object System.Drawing.Point(10, 10)
$label.AutoSize = $true
$form.Controls.Add($label)

#Create a user button on first form
$createUserButton = New-Object System.Windows.Forms.Button
$createUserButton.Text = "Create a user"
$createUserButton.Location = '10,30'
$createUserButton.Size = New-Object System.Drawing.Size(81, 20)
$form.Controls.Add($createUserButton)

#Create a user function

$createUserButton.Add_Click({
    $userForm = New-Object System.Windows.Forms.Form
    $userForm.Text = "Create a user - ADUC-PS v.0.1"
    $userForm.Size = New-Object System.Drawing.Size(500, 300)
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

    #The label and textbox for users department. @TO-DO: This has no use atm. add later.
    $userLabel5 = New-Object System.Windows.Forms.Label
    $userLabel5.Text = "Users department:"
    $userLabel5.Location = New-Object System.Drawing.Point(10, 190)
    $userLabel5.AutoSize = $true
    $userForm.Controls.Add($userLabel5)
    $userTextBox5 = New-Object System.Windows.Forms.TextBox
    $userTextBox5.Location = New-Object System.Drawing.Point(10, 210)
    $userTextBox5.Size = New-Object System.Drawing.Size(200, 20)
    $userForm.Controls.Add($userTextBox5)

    #Creates OK-button, pressing OK will run the input into separate variables
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "Create user"
    $okButton.Location = '390,220'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $userForm.Controls.Add($okButton)
    $userForm.AcceptButton = $okButton

    $userForm.ShowDialog()

    if ($userForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $userInput = $textBox1.Text
        $userInput2 = $textBox2.Text
        $userInput3 = $textBox3.Text
        $userUPN = $userInput1.ToLower()+"."+$userInput2.ToLower()+"@"+$userInput3.ToLower()
        $userInput4 = $textBox4.Text
        
    #Converting password to secure string:
    $SecurePassword = ConvertTo-SecureString $userInput4 -AsPlainText -Force
        
    #Used for searching after creating of the user
    $SAMACCName = $userInput1.ToLower()+"."+$userInput2.ToLower()
        
    #Creating the AD-user:
    New-ADUser -Name "$userInput1 $userInput2" `
    -GivenName $userInput `
    -Surname $userInput2 `
    -SamAccountName $SAMACCName`
    -UserPrincipalName $userUPN `
    -Enabled $true `
    -AccountPassword $SecurePassword `
    -ChangePasswordAtLogon = $true
        
    #Checking for the existence of the user created, searching everything based on SAM account name
    $User = Get-ADUser -Filter {SAMAccountName -eq $SAMACCName}
        
#@TO-DO - Find a better solution when it comes to OU..
        #Moves the user to the desired OU
        if ($user) {
            Write-Host "User found: $($User.DistinguishedName)"
            Write-Host "User has been moved to $OrgUnit !"
            Move-ADObject -Identity $user.DistinguishedName `
            -TargetPath $singleUserOU
        } else {
            Write-Host "User $SAMACCName not found in the directory."
        }
    [System.Windows.Forms.MessageBox]::Show("User have been created!")

    #Testing
    Write-Host "Users first name is: $userInput1"
    Write-Host "Users last name is: $userInput2"
    Write-Host "Users UPN is: $UserUPN"
    Write-Host "Users first password is: $userInput4"
    }
})

    #Import button for .csv-files
    $importCSVButton = New-Object System.Windows.Forms.Button
    $importCSVButton.Text = "Import user(s) using .csv"
    $importCSVButton.Location = '10,60'
    $importCSVButton.Size = New-Object System.Drawing.Size(140, 20)
    $form.Controls.Add($importCSVButton)

    $importCSVButton.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Title = "Select a CSV file"
        $openFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
        $openFileDialog.Filter = "CSV Files (*.csv)|*.csv"
        $openFileDialog.FilterIndex = 2
    
        #Check whether a file was chosen or not
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedFile = $openFileDialog.FileName
            [System.Windows.Forms.MessageBox]::Show("User(s) have been created!")
    
            $Import = Import-Csv $selectedFile

            Foreach ($user in $Import) {
                New-ADUser -Name $user.Name -GivenName $user.FirstName -Surname `
                $user.LastName -Path $csvOU -AccountPassword $password -ChangePasswordAtLogon $True `
                -Enabled $True
            }        
            } else {
                [System.Windows.Forms.MessageBox]::Show("No file selected.")
            }
        })
    
#Create a security group
$createSecGrpButton = New-Object System.Windows.Forms.Button
$createSecGrpButton.Text = "Create a security group"
$createSecGrpButton.Location = '10,90'
$createSecGrpButton.Size = New-Object System.Drawing.Size(140, 20)
$form.Controls.Add($createSecGrpButton)

$createSecGrpButton.Add_Click({
    $secGrpForm = New-Object System.Windows.Forms.Form #- WHY DOES IT OPEN TWICE ?? fml
    $secGrpForm.Text = "Create a Security Group - ADUC-PS v.0.1"
    $secGrpForm.Size = New-Object System.Drawing.Size(500, 300)
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
    $secGrpLabel1.Location = New-Object System.Drawing.Point(10, 30)
    $secGrpLabel1.AutoSize = $true
    $secGrpForm.Controls.Add($secGrpLabel1)
    $secComboBox = New-Object System.Windows.Forms.ComboBox
    $secComboBox.Location = New-Object System.Drawing.Point(10, 100)
    $secComboBox.Width = 200
    $secComboBox.Items.AddRange(@("Global", "Universal", "Domain local"))
    $secGrpForm.Controls.Add($secComboBox)
    $selectedItem = $secComboBox.SelectedItem

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
})
$form.ShowDialog()