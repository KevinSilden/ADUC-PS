#Imports the AD-module, comment/uncomment if/when nessecary:
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
     
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Title = "Select a CSV file"
$openFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
$openFileDialog.Filter = "CSV Files (*.csv)|*.csv"
$openFileDialog.FilterIndex = 2

#Check whether a file was chosen or not
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