$topRightUserTitlePanel.Visibility = [System.Windows.Visibility]::Collapsed
$bottomRightUserCreationPanel.Visibility = [System.Windows.Visibility]::Collapsed
$topRightCsvImportTitle.Visibility = [System.Windows.Visibility]::Visible
$bottomRightCsvImportPanel.Visibility = [System.Windows.Visibility]::Visible
$topRightSecGroupTitle.Visibility = [System.Windows.Visibility]::Collapsed
$bottomRightSecGroupPanel.Visibility = [System.Windows.Visibility]::Collapsed
$topRightCompObjectTitle.Visibility = [System.Windows.Visibility]::Collapsed
$bottomRightCompObjectPanel.Visibility = [System.Windows.Visibility]::Collapsed
$topRightOrgUnitTitle.Visibility = [System.Windows.Visibility]::Collapsed
$bottomRightOrgUnitPanel.Visibility = [System.Windows.Visibility]::Collapsed
$bottomRightLandingPagePanel.Visibility = [System.Windows.Visibility]::Collapsed

# Import csv-button for creating users.. Probably broken atm. Old script thrown in.
$window.FindName("btnImportAction").Add_Click({

    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = "Select a CSV file"
    $openFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    $openFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    $openFileDialog.FilterIndex = 2
    
    # Check whether a file was chosen or not
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
