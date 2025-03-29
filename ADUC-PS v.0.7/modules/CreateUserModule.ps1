$topRightUserTitlePanel.Visibility = [System.Windows.Visibility]::Visible
$bottomRightUserCreationPanel.Visibility = [System.Windows.Visibility]::Visible
$topRightCsvImportTitle.Visibility = [System.Windows.Visibility]::Collapsed
$bottomRightCsvImportPanel.Visibility = [System.Windows.Visibility]::Collapsed
$topRightSecGroupTitle.Visibility = [System.Windows.Visibility]::Collapsed
$bottomRightSecGroupPanel.Visibility = [System.Windows.Visibility]::Collapsed
$topRightCompObjectTitle.Visibility = [System.Windows.Visibility]::Collapsed
$bottomRightCompObjectPanel.Visibility = [System.Windows.Visibility]::Collapsed
$topRightOrgUnitTitle.Visibility = [System.Windows.Visibility]::Collapsed
$bottomRightOrgUnitPanel.Visibility = [System.Windows.Visibility]::Collapsed
$bottomRightLandingPagePanel.Visibility = [System.Windows.Visibility]::Collapsed

# Populating OU user combobox. Maps to distingquishedname.
if (Test-Path -Path $csvOUSelectionFilePath) {
    $data = Import-Csv -Path $csvOUSelectionFilePath
        if ($data) {
            $ouMapping = @{}
            foreach ($item in $data) {
                $cbUserOU.Items.Add($item.Name)
                $ouMapping[$item.Name] = $item.DistinguishedName
            }
                Write-Host "Populating ComboBox with the following data:"
                foreach ($item in $data) {
                Write-Host "Name: $($item.Name), DistinguishedName: $($item.DistinguishedName)"
            }
            # Add SelectionChanged handler, check if null or empty
            $cbUserOU.Add_SelectionChanged({
            if (-not $ouMapping) {
                Write-Host "Error: The OU mapping dictionary is null or not initialized."
            return
            }
            # Mapping selected OU, checks if valid
            $selectedOUName = $cbUserOU.SelectedItem
            if ($selectedOUName -and $ouMapping.ContainsKey($selectedOUName)) {
                $global:selectedDistinguishedName = $ouMapping[$selectedOUName]
                Write-Host "Selected Distinguished Name: $global:selectedDistinguishedName"
            } else {
                Write-Host "Error: Invalid or missing OU selection."
            }
        })
    }
}

# Populating Create User Security Groups combobox. Maps to distingquishedname.
if (Test-Path -Path $csvSecGroupFilePath) {
    $secGroupData = Import-Csv -Path $csvSecGroupFilePath
        if ($secGroupData) {
            $secGroupMapping = @{}
        foreach ($item in $secGroupData) {
            $cbUserSecGroup.Items.Add($item.Name)
            $secGroupMapping[$item.Name] = $item.DistinguishedName
        }

    $cbUserSecGroup.Add_SelectionChanged({
        if (-not $secGroupMapping) {
        Write-Host "Error: The Security Group mapping dictionary is null or not initialized."
        return
        }
    $selectedSecGroupName = $cbUserSecGroup.SelectedItem
    if ($selectedSecGroupName -and $secGroupMapping.ContainsKey($selectedSecGroupName)) {
    $global:selectedSecurityGroup = $secGroupMapping[$selectedSecGroupName]
    Write-Host "Selected Security Group DN: $global:selectedSecurityGroup"
    } else {
    Write-Host "Error: Invalid or missing Security Group selection."
    }
    })
    } else {
    Write-Host "No data found in Security Group CSV file."
    }
    } else {
    Write-Host "Security Group CSV file not found at: $csvSecGroupFilePath"
}

# Create user button - executes loads of commands.
$window.FindName("btnCreateUserAction").Add_Click({
    $UserFirstName = $txtUserFirstName.Text.Trim()
    $UserLastName = $txtUserLastName.Text.Trim()
    $domain = $txtUPN.Text.Trim()
    $userUPN = "$($UserFirstName.ToLower()).$($UserLastName.ToLower())@$($domain.ToLower())"

    # Replace non valid characters
    $UserFirstName = $UserFirstName -replace '[\\\/\[\]:;|=,+*?<>]', ''
    $txtUserLastName = $UserLastName -replace '[\\\/\[\]:;|=,+*?<>]', ''
    $userUPN = $userUPN -replace '[\\\/\[\]:;|=,+*?<>]', ''
    $SAMACCName = "$($UserFirstName.ToLower()).$($UserLastName.ToLower())"

    # Validate the inputs
    if (-not $txtUserFirstName -or -not $txtUserLastName -or -not $txtUPN -or -not $txtPassword) {
        Write-Host "Error: One or more input fields are empty."
        return
    }
    if ($userUPN -notmatch '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
        Write-Host "Error: Invalid UPN format: $userUPN"
        return
    }

    $SecurePassword = ConvertTo-SecureString $txtPassword -AsPlainText -Force

    # Checks whether the user already exists
    $existingUser = Get-ADUser -Filter {SamAccountName -eq $SAMACCName -or UserPrincipalName -eq $userUPN}
    if ($existingUser) {
        Write-Host "Error: A user with the same SAM Account Name or UPN already exists."
        return
    }

    # Creates the AD user
    try {
        New-ADUser -Name "$UserFirstName $UserLastName" -GivenName $UserFirstName -Surname $UserLastName `
            -SamAccountName $SAMACCName -UserPrincipalName $userUPN `
            -Enabled $true -AccountPassword $SecurePassword -ChangePasswordAtLogon $true
        Write-Host "User created successfully: $UPN"
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

    # Add the user to the chosen Security Group after creation and moving it
    $selectedSecurityGroup = $secGroupMapping[$cbUserSecGroup.SelectedItem]
    if ($selectedSecurityGroup) {
        Add-ADGroupMember -Identity $selectedSecurityGroup -Members $SAMACCName
        Write-Host "User $SAMACCName has been added to the group $selectedSecurityGroup"
    } else {
        Write-Host "Error: No valid security group selected."
    }
})