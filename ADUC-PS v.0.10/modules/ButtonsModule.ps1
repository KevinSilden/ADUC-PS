# "Event listeners" for the different buttons
$btnCreateUser.Add_Click({
    Invoke-Command -ScriptBlock { & ".\modules\CreateUserModule.ps1" }
})

$btnImportCSV.Add_Click({
    Invoke-Command -ScriptBlock { & ".\modules\CreateUsersFromCSVModule.ps1" }
})

$btnCreateSecGroup.Add_Click({
    Invoke-Command -ScriptBlock { & ".\modules\CreateSecurityGroupModule.ps1" }
})

$btnCreateCompObject.Add_Click({
    Invoke-Command -ScriptBlock { & ".\modules\CreateComputerObjectModule.ps1" }
})

$btnCreateOrgUnit.Add_Click({
    Invoke-Command -ScriptBlock { & ".\modules\CreateOrganizationalUnitModule.ps1" }
})

$window.FindName("btnDisplayOUStructure").Add_Click({
  
    # Check if the file exists
    if (-Not (Test-Path $xamlHierarchicalOUFilePath)) {
        [System.Windows.MessageBox]::Show("The XAML file for the OU hierarchy does not exist at the specified path.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return
    }
    
    try {
        # Load the XAML file
        $xamlContent2 = Get-Content -Path $xamlHierarchicalInput -Raw
        [xml]$xamlData = $xamlContent2
        $reader = New-Object System.Xml.XmlNodeReader $xamlData
    
        # Create the new window
        $window2 = [Windows.Markup.XamlReader]::Load($reader)
    
        # Show the new window
        $window2.ShowDialog()
    }
    catch {
        # Handle errors in XAML loading
        [System.Windows.MessageBox]::Show("Failed to load the XAML file. Error: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
})
    