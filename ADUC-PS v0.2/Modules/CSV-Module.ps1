#Imports the AD-module, comment/uncomment if/when nessecary:
Import-Module ActiveDirectory

Add-Type -AssemblyName System.Windows.Forms
        
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
    [System.Windows.Forms.MessageBox]::Show("No file selected.")
}