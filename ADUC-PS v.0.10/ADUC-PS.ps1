Import-Module ActiveDirectory

# Needs Windows.Forms for .csv-import @TO-DO ..... no good functionality atm
Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase, System.Windows.Forms

# Hides the PowerShell window
Add-Type -Name Window -Namespace Win32 -MemberDefinition @"
[DllImport("user32.dll")]
public static extern bool ShowWindow(int hWnd, int nCmdShow);
[DllImport("kernel32.dll")]
public static extern int GetConsoleWindow();
"@
$consolePtr = [Win32.Window]::GetConsoleWindow()
[Win32.Window]::ShowWindow($consolePtr, 0) # 0 = Hide, 5 = Show

# Creating folders for data-files, and datafiles excluding builtin OU's and built-in sec-groups.
$directoryPath = "$PSScriptRoot\Data"

if (-not (Test-Path $directoryPath)) {
    mkdir -Name "Data" -Path $PSScriptRoot
}

$csvOUSelectionFilePath = Join-Path -Path $PSScriptRoot -ChildPath "Data\OU_structure.csv"

# Generate XAML for displaying the second window
$outputXamlFilePath = Join-Path -Path $PSScriptRoot -ChildPath "GUI\OU_Hierarchy.xaml"

# Export exising OU's as well as the "root"
$rootDistinguishedName = (Get-ADDomain).DistinguishedName
$domainName = ($rootDistinguishedName -replace "DC=", "") -replace ",", "."

# Array for exporting both the "root" as well as other existing OU's.
$data = @()

$data += [PSCustomObject]@{
    Name             = "$domainName(root)"
    DistinguishedName = $rootDistinguishedName
}

$data += Get-ADOrganizationalUnit -Filter * |
    Where-Object { $_.Name -notlike "Builtin" } |
    Select-Object Name, DistinguishedName

$data | Export-Csv -Path $csvOUSelectionFilePath -NoTypeInformation

$csvSecGroupFilePath = Join-Path -Path $PSScriptRoot -ChildPath "Data\SecGroups.csv"

# Collection of names that are built in, excluding these from exporting.
$builtinGroupNames = @(
    "Administrators", "Users", "Guests", "Domain Admins", "Enterprise Admins",
    "Schema Admins", "Cert Publishers", "Group Policy Creator Owners", 
    "DnsAdmins", "Server Operators", "Account Operators", "Print Operators",
    "Backup Operators", "Read-Only Domain Controllers", "Pre-Windows 2000 Compatible Access"
)

# Exclude groups based on their name or distinguished name
Get-ADGroup -Filter * |
Where-Object {
    $_.GroupCategory -eq "Security" -and # Include only security groups
    $_.DistinguishedName -notlike "*CN=Builtin,*" -and
    $_.DistinguishedName -notlike "*CN=Users,*" -and
    $_.Name -notin $builtinGroupNames
} |
Select-Object Name, DistinguishedName |
Export-Csv -Path $csvSecGroupFilePath -NoTypeInformation

# Define the path of the .xaml used for GUI
$xamlFilePath = Join-Path -Path $PSScriptRoot -ChildPath "GUI\layout.xaml"

# Get the root distinguished name
$rootDistinguishedName = (Get-ADDomain).DistinguishedName
$domainName = ($rootDistinguishedName -replace "DC=", "") -replace ",", "."

# Load the XAML for main GUI
$xamlContent = Get-Content -Path $xamlFilePath -Raw
[xml]$xamlData = $xamlContent
$reader = New-Object System.Xml.XmlNodeReader $xamlData
$window = [Windows.Markup.XamlReader]::Load($reader)

# Accessing elements - ADUC-PS title + landing page, buttn
$TopLeftTitlePanel = $window.FindName("TopLeftTitlePanel")
$bottomRightLandingPagePanel = $window.FindName("bottomRightLandingPagePanel")
$bottomRightLandingPagePanel.Visibility = [System.Windows.Visibility]::Visible
$btnGit = $window.FindName("btnGit")
$btnGit.Add_Click({
    Start-Process "https://github.com/KevinSilden/ADUC-PS"
})

# Accessing elements - Create a User
$btnCreateUser = $window.FindName("btnCreateUser")
$topRightUserTitlePanel = $window.FindName("topRightUserTitlePanel")
$bottomRightUserCreationPanel = $window.FindName("bottomRightUserCreationPanel")
$txtUserFirstName = $window.FindName("txtUserFirstName")
$txtUserLastName = $window.FindName("txtUserLastName")
$txtPassword = $window.FindName("txtPassword")
$txtUPN = $window.FindName("txtUPN")
$cbUserOU = $window.FindName("cbUserOU")

# Accessing elements - Import user(s) from .csv
$btnImportCSV = $window.FindName("btnImportCSV")
$topRightCsvImportTitle = $window.FindName("topRightCsvImportTitle")
$bottomRightCsvImportPanel = $window.FindName("bottomRightCsvImportPanel")

# Accessing elements - Create a Security Group
$btnCreateSecGroup = $window.FindName("btnCreateSecGroup")
$topRightSecGroupTitle = $window.FindName("topRightSecGroupTitle")
$bottomRightSecGroupPanel = $window.FindName("bottomRightSecGroupPanel")
$cbUserSecGroup = $window.FindName("cbUserSecGroup")
$cbSecGroupOU = $window.FindName("cbSecGroupOU")
$cbGroupScope = $Window.FindName("cbGroupScope")

# Accessing elements - Create a Computer Object
$btnCreateCompObject = $window.FindName("btnCreateCompObject")
$topRightCompObjectTitle = $window.FindName("topRightCompObjectTitle")
$bottomRightCompObjectPanel = $window.FindName("bottomRightCompObjectPanel")
$cbCompObjOU = $window.FindName("cbCompObjOU")

# Accessing elements - Create an Organizational Unit
$btnCreateOrgUnit = $window.FindName("btnCreateOrgUnit")
$topRightOrgUnitTitle = $window.FindName("topRightOrgUnitTitle")
$bottomRightOrgUnitPanel = $window.FindName("bottomRightOrgUnitPanel")
$cbOrgUnitOU = $window.FindName("cbOrgUnitOU")

# Accessing elements - Show OU structure
$btnDisplayOUStructure = $window.FindName("btnDisplayOUStructure")

# Import OU-csv-file and check for header.
$csvOUFilePath = Join-Path -Path $PSScriptRoot -ChildPath "Data\UPN.csv"
if (-not (Test-Path $csvOUFilePath)) {
    "UPN" | Out-File -FilePath $csvOUFilePath -Encoding UTF8
}

# Reads the first existing UPN value (if any) from the CSV file
$currentUPN = ""
if ((Get-Content $csvOUFilePath | Measure-Object).Count -gt 1) {
    $currentUPN = (Get-Content $csvOUFilePath)[1]
}

# Populates the textbox with the current UPN value, overwrites if changed
$txtUPN.Text = $currentUPN
$txtUPN.Add_TextChanged({
    $upn = $txtUPN.Text
    "UPN" | Out-File -FilePath $csvOUFilePath -Encoding UTF8
if ($upn -ne "") {
    $upn | Out-File -FilePath $csvOUFilePath -Encoding UTF8 -Append
    }
})

# Running the script responsible for defining the intial XAML-file for OU structure
Invoke-Command -ScriptBlock { & ".\modules\TreeviewOUStructureModule.ps1" }

# Running script responsible for updating font size according to screen size
Invoke-Command -ScriptBlock { & ".\modules\FontUpdateModule.ps1" }

# Run the script responsible for the buttons on the landing page
Invoke-Command -ScriptBlock { & ".\modules\ButtonsModule.ps1" }

# Need to define this here, even though it worked earlier when the buttons were in the main script - and this was inside the "display OU"-button...
$xamlHierarchicalInput = Join-Path -Path $PSScriptRoot -ChildPath "GUI\OU_hierarchy.xaml"

$window.ShowDialog()