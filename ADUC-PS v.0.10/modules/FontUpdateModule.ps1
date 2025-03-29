# Function to update the font size based on window width
function UpdateFontSize {
    $windowWidth = $window.ActualWidth
    # Ensure the font size is between 12 and 48
    $newFontSize = [math]::Max([math]::Min($windowWidth / 15, 48), 12)

    $TopLeftTitlePanel.FontSize = $newFontSize
    $topRightUserTitlePanel.FontSize = $newFontSize
    $topRightCsvImportTitle.FontSize = $newFontSize
    $topRightSecGroupTitle.FontSize = $newFontSize
    $topRightCompObjectTitle.FontSize = $newFontSize
}

$window.add_SizeChanged({
    UpdateFontSize
})
UpdateFontSize