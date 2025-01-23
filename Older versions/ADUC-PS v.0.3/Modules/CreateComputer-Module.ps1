#Import the AD module
Import-Module ActiveDirectory

#Hides the PowerShell-window
Add-Type -Name Window -Namespace Win32 -MemberDefinition @"
[DllImport("user32.dll")]
public static extern bool ShowWindow(int hWnd, int nCmdShow);
[DllImport("kernel32.dll")]
public static extern int GetConsoleWindow();
"@
$consolePtr = [Win32.Window]::GetConsoleWindow()
[Win32.Window]::ShowWindow($consolePtr, 0) # 0 = Hide, 5 = Show

#CSV File Path
$importFile = Join-Path -Path $PSScriptRoot -ChildPath "Data\OU_structure.csv"

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="ADUC-PS v.0.3 - Create Computer Object" Height="400" Width="400" WindowStartupLocation="CenterScreen"
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
        <TextBlock Text="Create a Computer Object" FontSize="20" FontWeight="Bold" Grid.Row="0" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="0,10" Foreground="White"/>
        <StackPanel Grid.Row="1" HorizontalAlignment="Center" VerticalAlignment="Top" Width="300">
            <!-- Label for the TextBox -->
            <TextBlock Text="Enter Computer Object name:" FontSize="14" Foreground="White" Margin="0,10,0,5"/>
            <TextBox Name="txtCompName" Width="280" Height="25" Background="White" Foreground="Black" BorderBrush="Gray"/>

            <TextBlock Text="Select OU:" FontSize="14" Foreground="Gray" Margin="0,10,0,5"/>
            <ComboBox Name="cbGroupScope2" Width="280" Background="Black" Foreground="Black" BorderBrush="Gray"/>

        </StackPanel>

        <!-- Buttons Section -->
        <StackPanel Grid.Row="1" HorizontalAlignment="Center" VerticalAlignment="Bottom" Width="300">
            <Button Name="btnCreateComp" Content="Create Computer Object" Height="30" Margin="5"/>
        </StackPanel>
    </Grid>
</Window>
"@

#Load XAML
Add-Type -AssemblyName PresentationFramework
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

$cbGroupScope2 = $window.FindName("cbGroupScope2")
$txtCompName = $window.FindName("txtCompName")

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
$window.FindName("btnCreateComp").Add_Click({

    $compName = $txtCompName.Text
    $selectedOU = $cbGroupScope2.SelectedItem
    
     #Check if the form inputs are filled
    if (-not $compName -or -not $selectedOU) {
        Write-Host "Error: All fields must be filled out." -ForegroundColor Red
        exit
    }
    
    #Remove invalid characters
    $userInput1 = $compName -replace '[\\\/\[\]:;|=,+*?<>]', ''
    if (-not $userInput1) {
        Write-Host "Error: One or more input fields are empty."
        return
    }
    
    #Convert the selected OU to its Distinguished Name(copy-paste from user/secgroup)
    $selectedOU = $data | Where-Object { $_.Name -eq $cbGroupScope2.SelectedItem } | Select-Object -ExpandProperty DistinguishedName
    
    if (-not $selectedOU) {
        Write-Host "Error: Could not resolve the selected OU to a Distinguished Name." -ForegroundColor Red
        [System.Windows.MessageBox]::Show("Error: Could not resolve the selected OU to a Distinguished Name.")
        exit
    }
    
    #Creation of the computer object
    try {
        New-ADComputer -Name $userInput1 -SamAccountName $userInput1 -DisplayName $userInput1 -Path $selectedOU
        Write-Host "Computer Object '$compName' created successfully in OU '$selectedOU'."
        [System.Windows.MessageBox]::Show("Computer Object '$compName' created successfully in OU '$selectedOU'.")
    } catch {
        Write-Host "Error creating computer object: $_" -ForegroundColor Red
        [System.Windows.MessageBox]::Show("Error creating computer object: $_")
    }
    
    }
)

$window.ShowDialog()