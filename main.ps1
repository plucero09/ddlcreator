# Load PresentationFramework for WPF
Add-Type -AssemblyName PresentationFramework

# Function to connect to Exchange Online
function Connect-ExchangeOnline {
    $credential = Get-Credential
    Connect-ExchangeOnline -Credential $credential
}

# Function to get available user properties
function Get-UserProperties {
    # This function needs to be expanded to fetch actual user properties from Entra ID
    return @('Property1', 'Property2', 'Property3')
}

# Function to create GUI
function Show-GUI {
    # Create WPF elements here
    $window = New-Object System.Windows.Window
    $window.Title = "Dynamic Distribution List Creator"
    $window.Width = 400
    $window.Height = 300

    # Add other GUI elements (Dropdown, Text Fields, Buttons) here

    # Show the GUI
    $window.ShowDialog() | Out-Null
}

# Main script execution
Connect-ExchangeOnline
Show-GUI

# You will need to add the logic for each GUI element's functionality, such as:
# - Fetching and displaying user properties
# - Handling filter input and previews
# - Exporting preview data
# - Creating the dynamic distribution list based on filters

# Create ComboBox for attribute selection
$comboBox = New-Object System.Windows.Controls.ComboBox
$comboBox.ItemsSource = Get-UserProperties()
$comboBox.Add_SelectionChanged({ 
    # Add event logic here if needed
})
$window.AddChild($comboBox)

# Create TextBox for attribute value input
$textBox = New-Object System.Windows.Controls.TextBox
$window.AddChild($textBox)

# Create Preview Button
$previewButton = New-Object System.Windows.Controls.Button
$previewButton.Content = "Preview"
$previewButton.Add_Click({
    # Call a function to handle the preview logic
    Preview-Users -Attribute $comboBox.SelectedItem -Value $textBox.Text
})
$window.AddChild($previewButton)

function Preview-Users {
    param($Attribute, $Value)
    # Query Exchange Online and filter users based on the attribute and value
    # Display the results in a new window or a dedicated area in the existing GUI
}

# Create Export Button
$exportButton = New-Object System.Windows.Controls.Button
$exportButton.Content = "Export"
$exportButton.Add_Click({
    # Export logic here
})
$window.AddChild($exportButton)

# Create Distribution List Button
$createListButton = New-Object System.Windows.Controls.Button
$createListButton.Content = "Create List"
$createListButton.Add_Click({
    # Prompt for list name and create the list
})
$window.AddChild($createListButton)

$stackPanel = New-Object System.Windows.Controls.StackPanel
$stackPanel.Children.Add($comboBox)
$stackPanel.Children.Add($textBox)
$stackPanel.Children.Add($previewButton)
$stackPanel.Children.Add($exportButton)
$stackPanel.Children.Add($createListButton)

$window.Content = $stackPanel
