function GenerateTreeViewItems {
    param (
        [string]$ParentDN,
        [int]$Level = 0
    )

    $ous = Get-ADOrganizationalUnit -Filter * | Where-Object {
        $_.DistinguishedName -like "*,$ParentDN"
    } | Sort-Object Name

    $treeViewItems = @()
    foreach ($ou in $ous) {
        $treeViewItem = "<TreeViewItem Header='$($ou.Name)'>"

        $childItems = GenerateTreeViewItems -ParentDN $ou.DistinguishedName -Level ($Level + 1)
        $treeViewItem += $childItems
        $treeViewItem += "</TreeViewItem>"

        $treeViewItems += $treeViewItem
    }
    return $treeViewItems -join "`n"
}

# Building the start of the XAML (OU treeview), necessary in order to generate the file which builds upon the extracted data
$xamlHeader = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="OU Hierarchy Viewer (CLOSE OUT IN ORDER TO ACCESS MAIN GUI AGAIN)" Height="400" Width="600"
        WindowStartupLocation="CenterScreen">
    <Grid>
        <TreeView Name="OU_TreeView" Margin="10">
"@

$xamlFooter = @"
        </TreeView>
    </Grid>
</Window>
"@

# Generating items in the treeview
$treeViewContent = "<TreeViewItem Header='$domainName (root)'>"
$treeViewContent += GenerateTreeViewItems -ParentDN $rootDistinguishedName
$treeViewContent += "</TreeViewItem>"
$xamlContent = $xamlHeader + $treeViewContent + $xamlFooter
Set-Content -Path $outputXamlFilePath -Value $xamlContent -Encoding UTF8
Write-Output "OU hierarchy has been exported to $outputXamlFilePath"