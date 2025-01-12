#Imports the Active Directory module
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

#Export the OU structure and security groups
$outputFileOU = Join-Path -Path $PSScriptRoot -ChildPath "Modules\Data\OU_structure.csv"
$outputFileGroups = Join-Path -Path $PSScriptRoot -ChildPath "Modules\Data\SecGroups.csv"
Get-ADOrganizationalUnit -Filter * | Select-Object Name, DistinguishedName | Export-Csv -Path $outputFileOU -NoTypeInformation
Get-ADGroup -Filter * | Select-Object Name, DistinguishedName | Export-Csv -Path $outputFileGroups -NoTypeInformation

#Fancystuffs
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="ADUC-PS v.0.2" Height="400" Width="400" WindowStartupLocation="CenterScreen"
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
        <TextBlock Text="ADUC-PS v0.3" FontSize="20" FontWeight="Bold" Grid.Row="0" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="0,10" Foreground="White"/>

        <!-- Buttons Section -->
        <StackPanel Grid.Row="1" HorizontalAlignment="Center" VerticalAlignment="Top" Width="300">
            <Button Name="btnCreateUser" Content="Create a User" Height="30" Margin="5"/>
            <Button Name="btnImportCSV" Content="Import Users using .CSV" Height="30" Margin="5"/>
            <Button Name="btnCreateSecGrp" Content="Create a Security Group" Height="30" Margin="5"/>
            <Button Name="btnCreateComputer" Content="Create Computer Object" Height="30" Margin="5"/>
            <Button Name="btnOpenOUFile" Content="View OU Structure" Height="30" Margin="5"/>

        <TextBlock Text="silden.it" FontSize="20" FontWeight="Bold" Grid.Row="0" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="0,10" Foreground="White"/>
        </StackPanel>
    </Grid>
</Window>
"@

#Load XAML
Add-Type -AssemblyName PresentationFramework
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

#Buttons do button stuff
$window.FindName("btnCreateUser").Add_Click({
    $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath "Modules\CreateUser-Module.ps1"
    Start-Process "powershell.exe" -ArgumentList "-ExecutionPolicy Bypass", "-NoExit", "-WindowStyle Hidden", "-Command", "& '$scriptPath'"
})

$window.FindName("btnImportCSV").Add_Click({
    $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath "Modules\CSV-Module.ps1"
    Start-Process "powershell.exe" -ArgumentList "-ExecutionPolicy Bypass", "-NoExit", "-WindowStyle Hidden", "-Command", "& '$scriptPath'"
})

$window.FindName("btnCreateSecGrp").Add_Click({
    $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath "Modules\SecGroup-Module.ps1"
    Start-Process "powershell.exe" -ArgumentList "-ExecutionPolicy Bypass", "-NoExit", "-WindowStyle Hidden", "-Command", "& '$scriptPath'"
})

$window.FindName("btnCreateComputer").Add_Click({
    $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath "Modules\CreateComputer-Module.ps1"
    Start-Process "powershell.exe" -ArgumentList "-ExecutionPolicy Bypass", "-NoExit", "-WindowStyle Hidden", "-Command", "& '$scriptPath'"
})

$window.FindName("btnOpenOUFile").Add_Click({
    Start-Process -FilePath "notepad.exe" -ArgumentList $outputFileOU
})

#Shows the GUI, kills powershell to clean the minimized instances
$window.ShowDialog()
Stop-Process -Name "powershell" -Force