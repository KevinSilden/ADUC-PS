#Import the AD module
Import-Module ActiveDirectory

#CSV File Path
$importFile = Join-Path -Path $PSScriptRoot -ChildPath "Data\OU_structure.csv"

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Create security group" Height="400" Width="400" WindowStartupLocation="CenterScreen"
        Background="Black">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>

        <!-- Title -->
        <TextBlock Text="Create a security group" FontSize="20" FontWeight="Bold" Grid.Row="0" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="0,10" Foreground="White"/>

        <StackPanel Grid.Row="1" HorizontalAlignment="Center" VerticalAlignment="Top" Width="300">
            <!-- Label for the TextBox -->
            <TextBlock Text="Enter Security Group name:" FontSize="14" Foreground="White" Margin="0,10,0,5"/>
            <TextBox Name="txtGroupName" Width="280" Height="25" Background="White" Foreground="Black" BorderBrush="Gray"/>

            <TextBlock Text="Select Group Scope:" FontSize="14" Foreground="Gray" Margin="0,10,0,5"/>
            <ComboBox Name="cbGroupScope" Width="280" Background="Black" Foreground="Black" BorderBrush="Gray">
                <ComboBoxItem Content="Global"/>
                <ComboBoxItem Content="Universal"/>
                <ComboBoxItem Content="Domain Local"/>
            </ComboBox>

            <TextBlock Text="Select OU:" FontSize="14" Foreground="Gray" Margin="0,10,0,5"/>
            <ComboBox Name="cbGroupScope2" Width="280" Background="Black" Foreground="Black" BorderBrush="Gray"/>

        </StackPanel>

        <!-- Buttons Section -->
        <StackPanel Grid.Row="1" HorizontalAlignment="Center" VerticalAlignment="Bottom" Width="300">
            <Button Name="btnCreateGroup" Content="Create Group" Height="30" Margin="5"/>
        </StackPanel>
    </Grid>
</Window>
"@

#Load XAML
Add-Type -AssemblyName PresentationFramework
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

$cbGroupScope2 = $window.FindName("cbGroupScope2")

#Data from CSV into combobox
if (Test-Path -Path $importFile) {
    $data = Import-Csv -Path $importFile
    foreach ($item in $data) {
        if ($item.Name) {
            $cbGroupScope2.Items.Add($item.Name)
        }
    }

    if ($cbGroupScope2.Items.Count -eq 0) {
        [System.Windows.MessageBox]::Show("The CSV file does not contain valid 'Name' entries.", "Error")
    }
} else {
    [System.Windows.MessageBox]::Show("The CSV file was not found at: $importFile", "Error")
}

#Button do button stuff
$window.FindName("btnCreateGroup").Add_Click({
    $groupName = $window.FindName("txtGroupName").Text
    $selectedScope = $window.FindName("cbGroupScope").SelectedItem.Content
    $selectedOU = $cbGroupScope2.SelectedItem

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

$window.ShowDialog()