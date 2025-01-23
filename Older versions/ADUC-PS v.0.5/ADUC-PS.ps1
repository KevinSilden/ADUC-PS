Import-Module ActiveDirectory

#Needs Forms in order for bulk import to function
Add-Type -AssemblyName System.Windows.Forms

#Hides the PowerShell-window
Add-Type -Name Window -Namespace Win32 -MemberDefinition @"
[DllImport("user32.dll")]
public static extern bool ShowWindow(int hWnd, int nCmdShow);
[DllImport("kernel32.dll")]
public static extern int GetConsoleWindow();
"@
$consolePtr = [Win32.Window]::GetConsoleWindow()
[Win32.Window]::ShowWindow($consolePtr, 0) # 0 = Hide, 5 = Show

##########
##########
##########
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


##########
##########
##########
#Buttons and styling of the GUI
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        
Title="ADUC-PS v.0.5" Height="600" Width="800" WindowStartupLocation="CenterScreen"
    Background="Black">
    
    <Grid Margin="10">
    <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
    </Grid.RowDefinitions>

    <Grid.ColumnDefinitions>
        <ColumnDefinition Width="400" x:Name="LeftColumn"/>
        <ColumnDefinition Width="*" x:Name="RightColumn"/>
    </Grid.ColumnDefinitions>

    <!-- Title left grid-->
    <TextBlock Text="ADUC-PS v0.5" FontSize="20" FontWeight="Bold" Grid.Row="0" Grid.Column="0"
        HorizontalAlignment="Center" VerticalAlignment="Top" Margin="0,10" Foreground="White"/>

    <!-- Left grid -->
    <Grid Grid.Row="1" Grid.Column="0" HorizontalAlignment="Left" VerticalAlignment="Top" Width="400">
        <StackPanel x:Name="ButtonsPanel">
            <Button Name="btnCreateUser" Content="Create a User" Height="30" Margin="5"/>
            <Button Name="btnImportCSV" Content="Import Users using .CSV" Height="30" Margin="5"/>
            <Button Name="btnCreateSecGroup" Content="Create a Security Group" Height="30" Margin="5"/>
            <Button Name="btnCreateComputer" Content="Create Computer Object" Height="30" Margin="5"/>
            <TextBlock Text="silden.it" FontSize="20" FontWeight="Bold" HorizontalAlignment="Center"
                       Margin="0,10" Foreground="White"/>
        </StackPanel>
    </Grid>

    <!-- Right grid -->
    <!-- Expanded Section for User creation -->
            <Grid Grid.Row="1" Grid.Column="1" Background="Black" x:Name="createUserExpandedPanel" Visibility="Collapsed">
        <!-- Section Title -->
            <TextBlock Text="Create a User" FontSize="20" FontWeight="Bold"
                Grid.Row="1" Grid.Column="1" VerticalAlignment="Top" HorizontalAlignment="Center"
                Margin="0,10" Foreground="White"/>
               
        <!-- Defining rows -->
            <Grid Margin="0,40,0,0">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

        <!-- "Enter users first name" label and textbox -->
            <TextBlock Text="First Name:" FontSize="14" Foreground="White" 
                Grid.Row="0" Margin="0,10,0,5"/>
            <TextBox Name="txtUserFirstName" Height="25" Background="White" 
                Foreground="Black" BorderBrush="Gray" Grid.Row="1"/>

        <!-- "Enter users last name" label and textbox -->
            <TextBlock Text="Last name:" FontSize="14" Foreground="White"
                Grid.Row="2" Margin="0,10,0,5"/>
            <TextBox Name="txtUserLastName" Height="25" Background="White"
                Foreground="Black" BorderBrush="Gray" Grid.Row="3"/>

        <!-- "Enter domain" label and textbox -->
            <TextBlock Text="Domain(change if nessecary):" FontSize="14" Foreground="White"
                Grid.Row="4" MarginE="0,10,0,5"/>
            <TextBox Name="txtUPN" Height="25" Background="White"
                Foreground="Black" BorderBrush="Gray" Grid.Row="5"/>

        <!-- "Enter password" label and textbox -->
            <TextBlock Text="Users password for first logon:" FontSize="14" Foreground="White"
                Grid.Row="6" Margin="0,10,0,5"/>
            <TextBox Name="txtPassword" Height="25" Background="White"
                Foreground="Black" BorderBrush="Gray" Grid.Row="7"/>

        <!-- "Users OU" label and combobox -->
            <TextBlock Text="Choose users OU" FontSize="14" Foreground="White" 
                Grid.Row="8" Margin="0,10,0,5"/>
            <ComboBox Name="cbUserOU" Background="Black" Foreground="Black" BorderBrush="Gray"
                Grid.Row="9" Margin="0,10,0,5"/>

        <!-- "Select security group" label and combobox -->
            <TextBlock Text="Choose security group" FontSize="14" Foreground="White" 
                Grid.Row="10" Margin="0,10,0,5"/>
            <ComboBox Name="cbUserSecGroup" Background="Black" Foreground="Black" BorderBrush="Gray"
                Grid.Row="11" Margin="0,10,0,5"/>

        <!-- &nbsp lol -->
            <TextBlock Text=" " FontSize="14" Foreground="White" 
                Grid.Row="12" Margin="0,10,0,5"/>
                    
        <!-- Create user button section -->
            <StackPanel Grid.Row="13" HorizontalAlignment="Center" VerticalAlignment="Bottom" Width="300">
                <Button Name="btnCreateUserAction" Content="Create user" Height="30" Margin="5"/>
            </StackPanel>
        </Grid>

</Grid>

        <!-- Expanded section for csv import-->
            <Grid Grid.Row="1" Grid.Column="1" Background="Black" x:Name="csvImportExpandedPanel" Visibility="Collapsed">
        <!-- Section Title -->
            <TextBlock Text="Import user(s) from .csv" FontSize="20" FontWeight="Bold" 
                Grid.Row="2" VerticalAlignment="Top" HorizontalAlignment="Center" 
                Margin="0,10" Foreground="White"/>
               
        <!-- Defining rows -->
            <Grid Margin="0,40,0,0">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/> <!-- Row for Import user(s) Label -->
                <RowDefinition Height="Auto"/> <!-- yadayada -->
                <RowDefinition Height="Auto"/> <!-- yadayada -->
                <RowDefinition Height="Auto"/> <!-- Text -->
                <RowDefinition Height="Auto"/> <!-- Text -->
                <RowDefinition Height="Auto"/> <!-- Import user(s) Button -->
            </Grid.RowDefinitions>

            <TextBlock HorizontalAlignment="Center" Text="This may or may not import users when you choose a file" FontSize="14" Foreground="White" 
                Grid.Row="3" Margin="0,10,0,5"/>
        
            <TextBlock HorizontalAlignment="Center" Text="This functionality isn't complete yet" FontSize="14" Foreground="White" 
                Grid.Row="4" Margin="0,10,0,5"/>

        <!-- CSV Button Section -->
            <StackPanel Grid.Row="5" HorizontalAlignment="Center" VerticalAlignment="Bottom" Width="300">
                <Button Name="btnImportAction" Content="Import user(s) from .csv" Height="30" Margin="5"/>
            </StackPanel>
        </Grid>
    </Grid>

        <!-- Expanded Section for Security Group Creation -->
            <Grid Grid.Row="1" Grid.Column="1" Background="Black" x:Name="SecGroupExpandedPanel" Visibility="Collapsed">
        <!-- Title -->
            <TextBlock Text="Create a Security Group" FontSize="20" FontWeight="Bold" 
                Grid.Row="0" VerticalAlignment="Top" HorizontalAlignment="Center"
                Margin="0,10" Foreground="White"/>
               
        <!-- Row definitions -->
        <Grid Margin="0,40,0,0">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/> <!-- Row for "Enter Security Group" Label -->
                <RowDefinition Height="Auto"/> <!-- Row for TextBox -->
                <RowDefinition Height="Auto"/> <!-- Row for "Select OU" Label -->
                <RowDefinition Height="Auto"/> <!-- Row for ComboBox -->
                <RowDefinition Height="Auto"/> <!-- Row, blank space -->
                <RowDefinition Height="Auto"/> <!-- Row for TextBox -->
                <RowDefinition Height="Auto"/> <!-- Row for ComboBox -->
                <RowDefinition Height="Auto"/> <!-- Row for create sec grp button -->
            </Grid.RowDefinitions>

        <!-- "Enter Security Group Name" label and texbox-->
            <TextBlock Text="Enter Security Group name:" FontSize="14" Foreground="White" 
                Grid.Row="0" Margin="0,10,0,5"/>
            <TextBox Name="txtSecGroupName" Height="25" Background="White" 
                Foreground="Black" BorderBrush="Gray" Grid.Row="1"/>

        <!-- "Select OU" label and combobox -->
            <TextBlock Text="Select OU:" FontSize="14" Foreground="White" 
                Grid.Row="2" Margin="0,10,0,5"/>
            <ComboBox Name="cbSecGroupOU" Background="White" 
                Foreground="Black" BorderBrush="Gray" Grid.Row="3"/>

        <!-- "Select group scope" text and combobox -->
            <TextBlock Text="Select Group Scope:" FontSize="14" Foreground="Gray" Margin="0,10,0,5" Grid.Row="4"/>
            <ComboBox Name="cbGroupScope" Background="Black" Foreground="Black" BorderBrush="Gray" Grid.Row="5">
                <ComboBoxItem Content="Global"/>
                <ComboBoxItem Content="Universal"/>
                <ComboBoxItem Content="Domain Local"/>
            </ComboBox>

        <!-- &nbsp lol -->
            <TextBlock Text=" " FontSize="14" Foreground="White" 
                Grid.Row="6" Margin="0,10,0,5"/>

        <!-- Create security group button section -->
            <StackPanel Grid.Row="7" HorizontalAlignment="Center" VerticalAlignment="Bottom" Width="300">
                <Button Name="btnCreateSecGroupAction" Content="Create Security Group" Height="30" Margin="5"/>
            </StackPanel>
        </Grid>
    </Grid>

        <!-- Expanded section for Computer Object -->
            <Grid Grid.Row="1" Grid.Column="1" Background="Black" x:Name="compObjExpandedPanel" Visibility="Collapsed">
        <!-- Title -->
            <TextBlock Text="Create a Computer Object" FontSize="20" FontWeight="Bold" 
                Grid.Row="0" VerticalAlignment="Top" HorizontalAlignment="Center" 
                Margin="0,10" Foreground="White"/>
               
        <!-- Row definitions -->
            <Grid Margin="0,40,0,0">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/> <!-- Row for "Enter Computer Object Name" Label -->
                <RowDefinition Height="Auto"/> <!-- Row for TextBox -->
                <RowDefinition Height="Auto"/> <!-- Row for "Select OU" Label -->
                <RowDefinition Height="Auto"/> <!-- Row for blank space -->
                <RowDefinition Height="Auto"/> <!-- Row for ComboBox -->
            </Grid.RowDefinitions>

        <!-- "Enter Computer Object Name" label and textbox -->
            <TextBlock Text="Enter Computer Object name:" FontSize="14" Foreground="White" 
                Grid.Row="0" Margin="0,10,0,5"/>
            <TextBox Name="txtCompName" Height="25" Background="White" 
                Foreground="Black" BorderBrush="Gray" Grid.Row="1"/>

        <!-- Select OU -->
            <TextBlock Text="Select OU:" FontSize="14" Foreground="White" 
                Grid.Row="2" Margin="0,10,0,5"/>
            <ComboBox Name="cbCompObjOU" Background="White" 
                Foreground="Black" BorderBrush="Gray" Grid.Row="3"/>

        <!-- &nbsp lol -->
            <TextBlock Text=" " FontSize="14" Foreground="White" 
                Grid.Row="5" Margin="0,10,0,5"/>

        <!-- Create computer object button section -->
            <StackPanel Grid.Row="6" HorizontalAlignment="Center" VerticalAlignment="Bottom" Width="300">
                <Button Name="btnCreateComputerAction" Content="Create Computer Object" Height="30" Margin="5"/>
            </StackPanel>
        </Grid>
    </Grid>
</Grid>

</Window>
"@

##########
##########
##########
#Load XAML
Add-Type -AssemblyName PresentationFramework
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

##########
##########
##########
#Accessing elements
$leftColumn = $window.FindName("LeftColumn")
$rightColumn = $window.FindName("RightColumn")
$compObjExpandedPanel = $window.FindName("compObjExpandedPanel")
$SecGroupExpandedPanel = $window.FindName("SecGroupExpandedPanel")
$createUserExpandedlPanel = $window.FindName("createUserExpandedPanel")
$csvImportExpandedPanel = $window.FindName("csvImportExpandedPanel")
$txtUserFirstName = $window.FindName("txtUserFirstName")
$txtUserLastName = $window.FindName("txtUserLastName")
$txtPassword = $window.FindName("txtPassword")
$txtUPN = $window.FindName("txtUPN")
$cbUserOU = $window.FindName("cbUserOU")
$cbUserSecGroup = $window.FindName("cbUserSecGroup")
$cbSecGroupOU = $window.FindName("cbSecGroupOU")
$cbCompObjOU = $window.FindName("cbCompObjOU")

##########
##########
##########
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

##########
##########
##########
# Creates the animation for expanding the window width. Some of this is probably redundant. Lost track. I'll try to clean up later.
$widthAnimation = New-Object System.Windows.Media.Animation.DoubleAnimation
$widthAnimation.From = 800
$widthAnimation.To = 800
$widthAnimation.Duration = New-Object System.Windows.Duration([System.TimeSpan]::FromSeconds(0.3))
[System.Windows.Media.Animation.Storyboard]::SetTargetProperty($widthAnimation, "(Width)")
[System.Windows.Media.Animation.Storyboard]::SetTarget($widthAnimation, $window)

##########
##########
##########
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

##########
##########
##########
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

##########
##########
##########
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

##########
##########
##########
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

##########
##########
##########
#Buttons on "main window" - animation
$window.FindName("btnCreateUser").Add_Click({
    $window.BeginAnimation([System.Windows.FrameworkElement]::WidthProperty, $widthAnimation)
    $compObjExpandedPanel.Visibility = 'Collapsed'
    $SecGroupExpandedPanel.Visibility = 'Collapsed'
    $csvImportExpandedPanel.Visibility = 'Collapsed'
    $createUserExpandedlPanel.Visibility = 'Visible'
})

$window.FindName("btnCreateSecGroup").Add_Click({
    $window.BeginAnimation([System.Windows.FrameworkElement]::WidthProperty, $widthAnimation)
    $compObjExpandedPanel.Visibility = 'Collapsed'
    $SecGroupExpandedPanel.Visibility = 'Visible'
    $csvImportExpandedPanel.Visibility = 'Collapsed'
    $createUserExpandedlPanel.Visibility = 'Collapsed'
})

$window.FindName("btnImportCSV").Add_Click({
    $window.BeginAnimation([System.Windows.FrameworkElement]::WidthProperty, $widthAnimation)
    $compObjExpandedPanel.Visibility = 'Collapsed'
    $SecGroupExpandedPanel.Visibility = 'Collapsed'
    $createUserExpandedlPanel.Visibility = 'Collapsed'
    $csvImportExpandedPanel.Visibility = 'Visible'
})

$window.FindName("btnCreateComputer").Add_Click({
    $window.BeginAnimation([System.Windows.FrameworkElement]::WidthProperty, $widthAnimation)
    $compObjExpandedPanel.Visibility = 'Visible'
    $SecGroupExpandedPanel.Visibility = 'Collapsed'
    $createUserExpandedlPanel.Visibility = 'Collapsed'
    $csvImportExpandedPanel.Visibility = 'Collapsed'
})

##########
########## "Nested" buttons below
##########

##########
##########
##########
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

##########
##########
##########
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

##########
##########
##########
#Button that creates a Security Group
$window.FindName("btnCreateSecGroupAction").Add_Click({
    $groupName = $window.FindName("txtSecGroupName").Text #VSCode says I do not use this, but oh I do..
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
        "Domain Local" = "DomainLocal"
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

##########
##########
##########
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

#Draws the GUI
$window.ShowDialog()