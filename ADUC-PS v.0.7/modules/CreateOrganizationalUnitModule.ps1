$topRightUserTitlePanel.Visibility = [System.Windows.Visibility]::Collapsed
$bottomRightUserCreationPanel.Visibility = [System.Windows.Visibility]::Collapsed
$topRightCsvImportTitle.Visibility = [System.Windows.Visibility]::Collapsed
$bottomRightCsvImportPanel.Visibility = [System.Windows.Visibility]::Collapsed
$topRightSecGroupTitle.Visibility = [System.Windows.Visibility]::Collapsed
$bottomRightSecGroupPanel.Visibility = [System.Windows.Visibility]::Collapsed
$topRightCompObjectTitle.Visibility = [System.Windows.Visibility]::Collapsed
$bottomRightCompObjectPanel.Visibility = [System.Windows.Visibility]::Collapsed
$topRightOrgUnitTitle.Visibility = [System.Windows.Visibility]::Visible
$bottomRightOrgUnitPanel.Visibility = [System.Windows.Visibility]::Visible
$bottomRightLandingPagePanel.Visibility = [System.Windows.Visibility]::Collapsed

# Populate combobox in Organizational Unit
if (Test-Path -Path $importFile) {
    $data = Import-Csv -Path $importFile
    foreach ($item in $data) {
        if ($item.Name) {
            $cbOrgUnitOU.Items.Add($item.Name)
        }
    }

    if ($cbOrgUnitOU.Items.Count -eq 0) {
        [System.Windows.MessageBox]::Show("The CSV file does not contain valid 'Name' entries.", "Error")
    }
} else {
    [System.Windows.MessageBox]::Show("The CSV file was not found at: $importFile", "Error")
}


# Button that creates Organizational Unit
$window.FindName("btnCreateOrgUnitAction").Add_Click({

    $orgUnitName = $window.FindName("txtOrgUnitName").Text
    $selectedOrgUnitOU = $cbOrgUnitOU.SelectedItem
    
    # Check if the form inputs are filled
    if (-not $orgUnitName -or -not $selectedOrgUnitOU) {
        Write-Host "Error: All fields must be filled out." -ForegroundColor Red
        exit
    }
    
    # Convert the selected OU to its Distinguished Name(copy-paste from user/secgroup)
    $selectedOrgUnitOU = $data | Where-Object { $_.Name -eq $cbOrgUnitOU.SelectedItem } | Select-Object -ExpandProperty DistinguishedName
    
    if (-not $selectedOrgUnitOU) {
        Write-Host "Error: Could not resolve the selected OU to a Distinguished Name." -ForegroundColor Red
        [System.Windows.MessageBox]::Show("Error: Could not resolve the selected OU to a Distinguished Name.")
        exit
    }
    
    #Creation of the Organizational Unit
    try {
        New-ADOrganizationalUnit -Name $orgUnitName -DisplayName $orgUnitName -Path $selectedOrgUnitOU
        Write-Host "Organizational Unit '$orgUnitName' created successfully in OU '$selectedOrgUnitOU'."
        [System.Windows.MessageBox]::Show("Organizational Unit '$orgUnitName' created successfully in OU '$selectedOrgUnitOU'.")
    } catch {
        Write-Host "Error creating Organizational Unit $_" -ForegroundColor Red
        [System.Windows.MessageBox]::Show("Error creating Organizational Unit: $_")
    }

})
