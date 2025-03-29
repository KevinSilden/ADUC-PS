$topRightUserTitlePanel.Visibility = [System.Windows.Visibility]::Collapsed
$bottomRightUserCreationPanel.Visibility = [System.Windows.Visibility]::Collapsed
$topRightCsvImportTitle.Visibility = [System.Windows.Visibility]::Collapsed
$bottomRightCsvImportPanel.Visibility = [System.Windows.Visibility]::Collapsed
$topRightSecGroupTitle.Visibility = [System.Windows.Visibility]::Collapsed
$bottomRightSecGroupPanel.Visibility = [System.Windows.Visibility]::Collapsed
$topRightCompObjectTitle.Visibility = [System.Windows.Visibility]::Visible
$bottomRightCompObjectPanel.Visibility = [System.Windows.Visibility]::Visible
$topRightOrgUnitTitle.Visibility = [System.Windows.Visibility]::Collapsed
$bottomRightOrgUnitPanel.Visibility = [System.Windows.Visibility]::Collapsed
$bottomRightLandingPagePanel.Visibility = [System.Windows.Visibility]::Collapsed

# Populating combobox in Create Computer from .csv
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

# Button that create Computer Object
$window.FindName("btnCreateComputerAction").Add_Click({

    $compName = $window.FindName("txtCompName").Text
    $selectedCompOU = $cbCompObjOU.SelectedItem
    
    # Check if the form inputs are filled
    if (-not $compName -or -not $selectedCompOU) {
        Write-Host "Error: All fields must be filled out." -ForegroundColor Red
        exit
    }
    
    # Convert the selected OU to its Distinguished Name(copy-paste from user/secgroup)
    $selectedCompOU = $data | Where-Object { $_.Name -eq $cbCompObjOU.SelectedItem } | Select-Object -ExpandProperty DistinguishedName
    
    if (-not $selectedCompOU) {
        Write-Host "Error: Could not resolve the selected OU to a Distinguished Name." -ForegroundColor Red
        [System.Windows.MessageBox]::Show("Error: Could not resolve the selected OU to a Distinguished Name.")
        exit
    }
    
    # Creation of the computer object
    try {
        New-ADComputer -Name $compName -SamAccountName $compName -DisplayName $compName -Path $selectedCompOU
        Write-Host "Computer Object '$compName' created successfully in OU '$selectedCompOU'."
        [System.Windows.MessageBox]::Show("Computer Object '$compName' created successfully in OU '$selectedCompOU'.")
    } catch {
        Write-Host "Error creating computer object: $_" -ForegroundColor Red
        [System.Windows.MessageBox]::Show("Error creating computer object: $_")
    }

})
