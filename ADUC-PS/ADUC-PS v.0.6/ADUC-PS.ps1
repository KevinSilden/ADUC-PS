Import-Module ActiveDirectory

#Needs Windows.Forms for .csv-import
Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase, System.Windows.Forms

#Hides the PowerShell window
Add-Type -Name Window -Namespace Win32 -MemberDefinition @"
[DllImport("user32.dll")]
public static extern bool ShowWindow(int hWnd, int nCmdShow);
[DllImport("kernel32.dll")]
public static extern int GetConsoleWindow();
"@
$consolePtr = [Win32.Window]::GetConsoleWindow()
[Win32.Window]::ShowWindow($consolePtr, 0) # 0 = Hide, 5 = Show

#Creating folders for data-files, and datafiles excluding builtin OU's and built-in sec-groups.
$directoryPath = "$PSScriptRoot\Data"

if (-not (Test-Path $directoryPath)) {
    mkdir -Name "Data" -Path $PSScriptRoot
}

$csvOUSelectionFilePath = Join-Path -Path $PSScriptRoot -ChildPath "Data\OU_structure.csv"
$csvSecGroupFilePath = Join-Path -Path $PSScriptRoot -ChildPath "Data\SecGroups.csv"

Get-ADOrganizationalUnit -Filter * |
Where-Object { $_.Name -notlike "Builtin" } |
Select-Object Name, DistinguishedName |
Export-Csv -Path $csvOUSelectionFilePath -NoTypeInformation

$csvSecGroupFilePath = Join-Path -Path $PSScriptRoot -ChildPath "Data\SecGroups.csv"

#Collection of names that are built in, excluding these from exporting.
$builtinGroupNames = @(
    "Administrators", "Users", "Guests", "Domain Admins", "Enterprise Admins",
    "Schema Admins", "Cert Publishers", "Group Policy Creator Owners", 
    "DnsAdmins", "Server Operators", "Account Operators", "Print Operators",
    "Backup Operators", "Read-Only Domain Controllers", "Pre-Windows 2000 Compatible Access"
)

#Exclude groups based on their name or distinguished name
Get-ADGroup -Filter * |
Where-Object {
    $_.GroupCategory -eq "Security" -and # Include only security groups
    $_.DistinguishedName -notlike "*CN=Builtin,*" -and
    $_.DistinguishedName -notlike "*CN=Users,*" -and
    $_.Name -notin $builtinGroupNames
} |
Select-Object Name, DistinguishedName |
Export-Csv -Path $csvSecGroupFilePath -NoTypeInformation

#Define the path of the .xaml
$xamlFilePath = Join-Path -Path $PSScriptRoot -ChildPath "GUI\layout.xaml"

#Load the XAML
$xamlContent = Get-Content -Path $xamlFilePath -Raw
[xml]$xamlData = $xamlContent
$reader = New-Object System.Xml.XmlNodeReader $xamlData
$window = [Windows.Markup.XamlReader]::Load($reader)

#Accessing elements - ADUC-PS title
$TopLeftTitlePanel = $window.FindName("TopLeftTitlePanel")

#Accessing elements - Create a User
$btnCreateUser = $window.FindName("btnCreateUser")
$topRightUserTitlePanel = $window.FindName("topRightUserTitlePanel")
$bottomRightUserCreationPanel = $window.FindName("bottomRightUserCreationPanel")
$txtUserFirstName = $window.FindName("txtUserFirstName")
$txtUserLastName = $window.FindName("txtUserLastName")
$txtPassword = $window.FindName("txtPassword")
$txtUPN = $window.FindName("txtUPN")
$cbUserOU = $window.FindName("cbUserOU")

#Accessing elements - Import user(s) from .csv
$btnImportCSV = $window.FindName("btnImportCSV")
$topRightCsvImportTitle = $window.FindName("topRightCsvImportTitle")
$bottomRightCsvImportPanel = $window.FindName("bottomRightCsvImportPanel")

#Accessing elements - Create a Security Group
$btnCreateSecGroup = $window.FindName("btnCreateSecGroup")
$topRightSecGroupTitle = $window.FindName("topRightSecGroupTitle")
$bottomRightSecGroupPanel = $window.FindName("bottomRightSecGroupPanel")
$cbUserSecGroup = $window.FindName("cbUserSecGroup")
$cbSecGroupOU = $window.FindName("cbSecGroupOU")
$cbGroupScope = $Window.FindName("cbGroupScope")

#Accessing elements - Create a Computer Object
$btnCreateCompObject = $window.FindName("btnCreateCompObject")
$topRightCompObjectTitle = $window.FindName("topRightCompObjectTitle")
$bottomRightCompObjectPanel = $window.FindName("bottomRightCompObjectPanel")
$cbCompObjOU = $window.FindName("cbCompObjOU")

#Import OU-csv-file and check for header
$csvOUFilePath = Join-Path -Path $PSScriptRoot -ChildPath "Data\UPN.csv"
if (-not (Test-Path $csvOUFilePath)) {
    "UPN" | Out-File -FilePath $csvOUFilePath -Encoding UTF8
}
#Reads the first existing UPN value (if any) from the CSV file
$currentUPN = ""
if ((Get-Content $csvOUFilePath | Measure-Object).Count -gt 1) {
    $currentUPN = (Get-Content $csvOUFilePath)[1]  # Get the first line after the header
}
#Populates the textbox with the current UPN value, overwrites if changed
$txtUPN.Text = $currentUPN
$txtUPN.Add_TextChanged({
    $upn = $txtUPN.Text
    "UPN" | Out-File -FilePath $csvOUFilePath -Encoding UTF8
if ($upn -ne "") {
    $upn | Out-File -FilePath $csvOUFilePath -Encoding UTF8 -Append
    }
})

#Populating OU user combobox. Maps to distingquishedname.
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
            #Add SelectionChanged handler, check if null or empty
            $cbUserOU.Add_SelectionChanged({
            if (-not $ouMapping) {
                Write-Host "Error: The OU mapping dictionary is null or not initialized."
            return
            }
            #Mapping selected OU, checks if valid
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

#Populating Create User Security Groups combobox. Maps to distingquishedname.
if (Test-Path -Path $csvSecGroupFilePath) {
    $secGroupData = Import-Csv -Path $csvSecGroupFilePath
        if ($secGroupData) {
            $secGroupMapping = @{}
        foreach ($item in $secGroupData) {
            $cbUserSecGroup.Items.Add($item.Name)
            $secGroupMapping[$item.Name] = $item.DistinguishedName
        }
    # Add SelectionChanged handler for Security Group ComboBox
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

#Populating combobox in create Security Group from .csv
$importFile = Join-Path -Path $PSScriptRoot -ChildPath "Data\OU_structure.csv"
if (Test-Path -Path $importFile) {
    $data = Import-Csv -Path $importFile
    foreach ($item in $data) {
        if ($item.Name) {
            $cbSecGroupOU.Items.Add($item.Name)
        }
    }

    if ($cbSecGroupOU.Items.Count -eq 0) {
        [System.Windows.MessageBox]::Show("The CSV file does not contain valid 'Name' entries.", "Error")
    }
} else {
    [System.Windows.MessageBox]::Show("The CSV file was not found at: $importFile", "Error")
}

#Populating combobox in Create Computer from .csv
if (Test-Path -Path $importFile) {
    $data = Import-Csv -Path $importFile
    foreach ($item in $data) {
        if ($item.Name) {
            $cbCompObjOU.Items.Add($item.Name)
        }
    }

    if ($cbCompObjOU.Items.Count -eq 0) {
        [System.Windows.MessageBox]::Show("The CSV file does not contain valid 'Name' entries.", "Error")
    }
} else {
    [System.Windows.MessageBox]::Show("The CSV file was not found at: $importFile", "Error")
}

#Function to update the font size based on window width
function UpdateFontSize {
    $windowWidth = $window.ActualWidth
    # Ensure the font size is between 12 and 48
    $newFontSize = [math]::Max([math]::Min($windowWidth / 15, 48), 12)

    $TopLeftTitlePanel.FontSize = $newFontSize
    $topRightUserTitlePanel.FontSize = $newFontSize
    $topRightCsvImportTitle.FontSize = $newFontSize
    $topRightSecGroupTitle.FontSize = $newFontSize
    $topRightCompObjectTitle.FontSize = $newFontSize
}


#Hook into the SizeChanged event to dynamically update font size
$window.add_SizeChanged({
    UpdateFontSize
})

#Initial font size update
UpdateFontSize

#"Event listener" on btnCreateUser
$btnCreateUser.Add_Click({
    $topRightUserTitlePanel.Visibility = [System.Windows.Visibility]::Visible
    $bottomRightUserCreationPanel.Visibility = [System.Windows.Visibility]::Visible
    $topRightCsvImportTitle.Visibility = [System.Windows.Visibility]::Collapsed
    $bottomRightCsvImportPanel.Visibility = [System.Windows.Visibility]::Collapsed
    $topRightSecGroupTitle.Visibility = [System.Windows.Visibility]::Collapsed
    $bottomRightSecGroupPanel.Visibility = [System.Windows.Visibility]::Collapsed
    $topRightCompObjectTitle.Visibility = [System.Windows.Visibility]::Collapsed
    $bottomRightCompObjectPanel.Visibility = [System.Windows.Visibility]::Collapsed
})

#"Event listener" on btnImportCSV
$btnImportCSV.Add_Click({
    $topRightUserTitlePanel.Visibility = [System.Windows.Visibility]::Collapsed
    $bottomRightUserCreationPanel.Visibility = [System.Windows.Visibility]::Collapsed
    $topRightCsvImportTitle.Visibility = [System.Windows.Visibility]::Visible
    $bottomRightCsvImportPanel.Visibility = [System.Windows.Visibility]::Visible
    $topRightSecGroupTitle.Visibility = [System.Windows.Visibility]::Collapsed
    $bottomRightSecGroupPanel.Visibility = [System.Windows.Visibility]::Collapsed
    $topRightCompObjectTitle.Visibility = [System.Windows.Visibility]::Collapsed
    $bottomRightCompObjectPanel.Visibility = [System.Windows.Visibility]::Collapsed
})

#"Event listener" on btnCreateSecGroup
$btnCreateSecGroup.Add_Click({
    $topRightUserTitlePanel.Visibility = [System.Windows.Visibility]::Collapsed
    $bottomRightUserCreationPanel.Visibility = [System.Windows.Visibility]::Collapsed
    $topRightCsvImportTitle.Visibility = [System.Windows.Visibility]::Collapsed
    $bottomRightCsvImportPanel.Visibility = [System.Windows.Visibility]::Collapsed
    $topRightSecGroupTitle.Visibility = [System.Windows.Visibility]::Visible
    $bottomRightSecGroupPanel.Visibility = [System.Windows.Visibility]::Visible
    $topRightCompObjectTitle.Visibility = [System.Windows.Visibility]::Collapsed
    $bottomRightCompObjectPanel.Visibility = [System.Windows.Visibility]::Collapsed
})

#"Event listener" on btnCreateCompObject
$btnCreateCompObject.Add_Click({
    $topRightUserTitlePanel.Visibility = [System.Windows.Visibility]::Collapsed
    $bottomRightUserCreationPanel.Visibility = [System.Windows.Visibility]::Collapsed
    $topRightCsvImportTitle.Visibility = [System.Windows.Visibility]::Collapsed
    $bottomRightCsvImportPanel.Visibility = [System.Windows.Visibility]::Collapsed
    $topRightSecGroupTitle.Visibility = [System.Windows.Visibility]::Collapsed
    $bottomRightSecGroupPanel.Visibility = [System.Windows.Visibility]::Collapsed
    $topRightCompObjectTitle.Visibility = [System.Windows.Visibility]::Visible
    $bottomRightCompObjectPanel.Visibility = [System.Windows.Visibility]::Visible
})

#############################################################################################
#"Nested" buttons below
#Create user button - executes loads of commands.
$window.FindName("btnCreateUserAction").Add_Click({
    $UserFirstName = $txtUserFirstName.Text.Trim()
    $UserLastName = $txtUserLastName.Text.Trim()
    $domain = $txtUPN.Text.Trim()
    $userUPN = "$($UserFirstName.ToLower()).$($UserLastName.ToLower())@$($domain.ToLower())"

    #Replace non valid characters
    $UserFirstName = $UserFirstName -replace '[\\\/\[\]:;|=,+*?<>]', ''
    $txtUserLastName = $UserLastName -replace '[\\\/\[\]:;|=,+*?<>]', ''
    $userUPN = $userUPN -replace '[\\\/\[\]:;|=,+*?<>]', ''
    $SAMACCName = "$($UserFirstName.ToLower()).$($UserLastName.ToLower())"

    #Validate the inputs
    if (-not $txtUserFirstName -or -not $txtUserLastName -or -not $txtUPN -or -not $txtPassword) {
        Write-Host "Error: One or more input fields are empty."
        return
    }
    if ($userUPN -notmatch '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
        Write-Host "Error: Invalid UPN format: $userUPN"
        return
    }
    #I have no clue why it went here, but hey, if it ain't broke don't fix it
    $SecurePassword = ConvertTo-SecureString $txtPassword -AsPlainText -Force

    #Checks whether the user already exists
    $existingUser = Get-ADUser -Filter {SamAccountName -eq $SAMACCName -or UserPrincipalName -eq $userUPN}
    if ($existingUser) {
        Write-Host "Error: A user with the same SAM Account Name or UPN already exists."
        return
    }

    #Creates the AD user
    try {
        New-ADUser -Name "$UserFirstName $UserLastName" -GivenName $UserFirstName -Surname $UserLastName `
            -SamAccountName $SAMACCName -UserPrincipalName $userUPN `
            -Enabled $true -AccountPassword $SecurePassword -ChangePasswordAtLogon $true
        Write-Host "User created successfully: $UPN"
    } catch {
        Write-Host "Error creating user: $_"
        return
    }

    #Validate the selected OU
    if (-not $global:selectedDistinguishedName) {
        Write-Host "Error: No OU selected or '$global:selectedDistinguishedName' is null."
        return
    }

    #Move the user to the selected OU
    $User = Get-ADUser -Filter {SamAccountName -eq $SAMACCName}
    if ($User) {
        Write-Host "User found: $($User.DistinguishedName)"
        Write-Host "Moving user to: $global:selectedDistinguishedName"
        Move-ADObject -Identity $User.DistinguishedName -TargetPath $global:selectedDistinguishedName
    } else {
        Write-Host "Error: User '$SAMACCName' not found in the directory."
    }

    #Confirmation message
    [System.Windows.Forms.MessageBox]::Show("User has been created, moved to the OU and added to the Security Group!")

    #Add the user to the chosen Security Group after creation and moving it
    $selectedSecurityGroup = $secGroupMapping[$cbUserSecGroup.SelectedItem]
    if ($selectedSecurityGroup) {
        Add-ADGroupMember -Identity $selectedSecurityGroup -Members $SAMACCName
        Write-Host "User $SAMACCName has been added to the group $selectedSecurityGroup"
    } else {
        Write-Host "Error: No valid security group selected."
    }
})

#Import csv-button for creating users.. Probably broken atm. Not tested. At all. Basically a placeholder atm.
$window.FindName("btnImportAction").Add_Click({

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
        [System.Windows.Forms.MessageBox]::Show("No file selected. No users created :(")
    }    

})

#Button that creates a Security Group
$window.FindName("btnCreateSecGroupAction").Add_Click({
    $groupName = $window.FindName("txtSecGroupName").Text
    $selectedScope = $window.FindName("cbGroupScope").SelectedItem.Content
    $selectedOU = $cbSecGroupOU.SelectedItem

    if (-not $groupName -or -not $selectedScope -or -not $selectedOU) {
        [System.Windows.MessageBox]::Show("All fields must be filled out.", "Error")
        return
    }

    #Converts the values in combobox to AD-values.
    $scopeMapping = @{
        "Global"       = "Global"
        "Universal"    = "Universal"
        "DomainLocal" = "DomainLocal"
    }
    $adScope = $scopeMapping[$selectedScope]

    #Convert the selected OU to Distinguished Name
    $selectedOU = $data | Where-Object { $_.Name -eq $selectedOU } | Select-Object -ExpandProperty DistinguishedName

    if (-not $selectedOU) {
        [System.Windows.MessageBox]::Show("Could not resolve the selected OU to a Distinguished Name.", "Error")
        return
    }

    #Creates the Security Group
    try {
        New-ADGroup -Name $groupName -SamAccountName $groupName -GroupCategory Security -GroupScope $adScope -DisplayName $groupName -Path $selectedOU
        [System.Windows.MessageBox]::Show("Security group '$groupName' created successfully.", "Success")
    } catch {
        [System.Windows.MessageBox]::Show("Error creating group: $_", "Error")
    }

})

$window.FindName("btnCreateComputerAction").Add_Click({

    $compName = $window.FindName("txtCompName").Text
    $selectedCompOU = $cbCompObjOU.SelectedItem
    
    #Check if the form inputs are filled
    if (-not $compName -or -not $selectedCompOU) {
        Write-Host "Error: All fields must be filled out." -ForegroundColor Red
        exit
    }
    
    #Convert the selected OU to its Distinguished Name(copy-paste from user/secgroup)
    $selectedCompOU = $data | Where-Object { $_.Name -eq $cbCompObjOU.SelectedItem } | Select-Object -ExpandProperty DistinguishedName
    
    if (-not $selectedCompOU) {
        Write-Host "Error: Could not resolve the selected OU to a Distinguished Name." -ForegroundColor Red
        [System.Windows.MessageBox]::Show("Error: Could not resolve the selected OU to a Distinguished Name.")
        exit
    }
    
    #Creation of the computer object
    try {
        New-ADComputer -Name $compName -SamAccountName $compName -DisplayName $compName -Path $selectedCompOU
        Write-Host "Computer Object '$compName' created successfully in OU '$selectedCompOU'."
        [System.Windows.MessageBox]::Show("Computer Object '$compName' created successfully in OU '$selectedCompOU'.")
    } catch {
        Write-Host "Error creating computer object: $_" -ForegroundColor Red
        [System.Windows.MessageBox]::Show("Error creating computer object: $_")
    }

})

#@TO-DO: Create Organizational Units
#@TO-DO: Make .csv-import functional
#@TO-DO: Is it even possible to add printer objects and map them..?

$window.ShowDialog()

