#Imports the AD-module, comment/uncomment if/when nessecary:
Import-Module ActiveDirectory

#Hides the powershell window on the main script
Add-Type -Name Window -Namespace Win32 -MemberDefinition @"
[DllImport("user32.dll")]
public static extern bool ShowWindow(int hWnd, int nCmdShow);
[DllImport("kernel32.dll")]
public static extern int GetConsoleWindow();
"@

$consolePtr = [Win32.Window]::GetConsoleWindow()
[Win32.Window]::ShowWindow($consolePtr, 0) # 0 = Hide, 5 = Show

#Export the OU structure to file, exports name and distinguished name
$outputFile = Join-Path -Path $PSScriptRoot -ChildPath "Modules\Data\OU_structure.csv"
Get-ADOrganizationalUnit -Filter * | Select-Object Name, DistinguishedName | Export-Csv -Path $outputFile -NoTypeInformation

#Load the Windows Forms assembly
Add-Type -AssemblyName System.Windows.Forms

#Creates the Form(GUI-window), centers it on screen.
$form = New-Object System.Windows.Forms.Form
$form.Text = "ADUC-PS v.0.1"
$form.Size = New-Object System.Drawing.Size(800, 600)
$form.StartPosition = "CenterScreen"

#Define the OU-structure:
$label = New-Object System.Windows.Forms.Label
$label.Text = "View your OU-structure"
$label.Location = New-Object System.Drawing.Point(280, 10)
$label.AutoSize = $true
$form.Controls.Add($label)
$buttonOU = New-Object System.Windows.Forms.Button
$buttonOU.Location = '280,30'
$buttonOU.Text = "Open"
$buttonOU.Size = New-Object System.Drawing.Size(40, 20)
$form.Controls.Add($buttonOU)
$buttonOU.Add_Click({
    Start-Process -FilePath "notepad.exe" -ArgumentList $outputFile
})

#Text @ the top of the window
$label = New-Object System.Windows.Forms.Label
$label.Text = "What you wanna do?"
$label.Location = New-Object System.Drawing.Point(10, 10)
$label.AutoSize = $true
$form.Controls.Add($label)

#Create a user button on first form
$createUserButton = New-Object System.Windows.Forms.Button
$createUserButton.Text = "Create a user"
$createUserButton.Location = '10,30'
$createUserButton.Size = New-Object System.Drawing.Size(81, 20)
$form.Controls.Add($createUserButton)
$createUserButton.Add_Click({
    $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath "Modules\CreateUser-Module.ps1"
    Write-Host "Starting script with -Command"
    Start-Process "powershell.exe" -ArgumentList "-ExecutionPolicy Bypass", "-NoExit", "-WindowStyle Hidden", "-Command", "& '$scriptPath'"
})

#CSV button
$importCSVButton = New-Object System.Windows.Forms.Button
$importCSVButton.Text = "Import user(s) using .csv"
$importCSVButton.Location = '10,60'
$importCSVButton.Size = New-Object System.Drawing.Size(140, 20)
$form.Controls.Add($importCSVButton)
$importCSVButton.Add_Click({
    $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath "Modules\CSV-Module.ps1"
    Write-Host "Starting script with -Command"
    Start-Process "powershell.exe" -ArgumentList "-ExecutionPolicy Bypass", "-NoExit", "-WindowStyle Hidden", "-Command", "& '$scriptPath'"
})

#Security group button
$createSecGrpButton = New-Object System.Windows.Forms.Button
$createSecGrpButton.Text = "Create a security group"
$createSecGrpButton.Location = '10,90'
$createSecGrpButton.Size = New-Object System.Drawing.Size(140, 20)
$form.Controls.Add($createSecGrpButton)
$createSecGrpButton.Add_Click({
    $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath "Modules\SecGroup-Module.ps1"
    Write-Host "Starting script with -Command"
    Start-Process "powershell.exe" -ArgumentList "-ExecutionPolicy Bypass", "-NoExit", "-WindowStyle Hidden", "-Command", "& '$scriptPath'"
})

#Computer object-button
$createComputerButton = New-Object System.Windows.Forms.Button
$createComputerButton.Text = "Create computer object"
$createComputerButton.Location = '10,120'
$createComputerButton.Size = New-Object System.Drawing.Size(140, 20)
$form.Controls.Add($createComputerButton)
$createComputerButton.Add_Click({
    $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath "Modules\CreateComputer-Module.ps1"
    Write-Host "Starting script with -Command"
    Start-Process "powershell.exe" -ArgumentList "-ExecutionPolicy Bypass", "-NoExit", "-WindowStyle Hidden", "-Command", "& '$scriptPath'"
})

$form.ShowDialog()
Stop-Process -Name "powershell" -Force