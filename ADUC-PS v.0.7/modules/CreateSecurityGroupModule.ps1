$topRightUserTitlePanel.Visibility = [System.Windows.Visibility]::Collapsed
$bottomRightUserCreationPanel.Visibility = [System.Windows.Visibility]::Collapsed
$topRightCsvImportTitle.Visibility = [System.Windows.Visibility]::Collapsed
$bottomRightCsvImportPanel.Visibility = [System.Windows.Visibility]::Collapsed
$topRightSecGroupTitle.Visibility = [System.Windows.Visibility]::Visible
$bottomRightSecGroupPanel.Visibility = [System.Windows.Visibility]::Visible
$topRightCompObjectTitle.Visibility = [System.Windows.Visibility]::Collapsed
$bottomRightCompObjectPanel.Visibility = [System.Windows.Visibility]::Collapsed
$topRightOrgUnitTitle.Visibility = [System.Windows.Visibility]::Collapsed
$bottomRightOrgUnitPanel.Visibility = [System.Windows.Visibility]::Collapsed
$bottomRightLandingPagePanel.Visibility = [System.Windows.Visibility]::Collapsed

# Populating combobox in create Security Group from .csv
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

# Button that creates a Security Group
$window.FindName("btnCreateSecGroupAction").Add_Click({
    $groupName = $window.FindName("txtSecGroupName").Text
    $selectedScope = $window.FindName("cbGroupScope").SelectedItem.Content
    $selectedOU = $cbSecGroupOU.SelectedItem

    if (-not $groupName -or -not $selectedScope -or -not $selectedOU) {
        [System.Windows.MessageBox]::Show("All fields must be filled out.", "Error")
        return
    }

    # Converts the values in combobox to AD-values.
    $scopeMapping = @{
        "Global"       = "Global"
        "Universal"    = "Universal"
        "DomainLocal" = "DomainLocal"
    }
    $adScope = $scopeMapping[$selectedScope]

    # Convert the selected OU to Distinguished Name
    $selectedOU = $data | Where-Object { $_.Name -eq $selectedOU } | Select-Object -ExpandProperty DistinguishedName

    if (-not $selectedOU) {
        [System.Windows.MessageBox]::Show("Could not resolve the selected OU to a Distinguished Name.", "Error")
        return
    }

    # Creates the Security Group
    try {
        New-ADGroup -Name $groupName -SamAccountName $groupName -GroupCategory Security -GroupScope $adScope -DisplayName $groupName -Path $selectedOU
        [System.Windows.MessageBox]::Show("Security group '$groupName' created successfully.", "Success")
    } catch {
        [System.Windows.MessageBox]::Show("Error creating group: $_", "Error")
    }

})
