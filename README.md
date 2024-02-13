# Load Outlook interop assembly
Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"

# Create Outlook application object
$outlook = New-Object -ComObject Outlook.Application

# Get the Inbox folder
$inbox = $outlook.Session.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)

# Get the rules using reflection
$rulesProperty = $inbox.GetType().GetProperty("Rules")
$rules = $rulesProperty.GetValue($inbox, $null)

# Loop through each rule
foreach ($rule in $rules) {
    # Use reflection to get rule properties
    $ruleType = $rule.GetType()
    $conditions = $ruleType.InvokeMember("Conditions", [System.Reflection.BindingFlags]::GetProperty, $null, $rule, $null)
    $actions = $ruleType.InvokeMember("Actions", [System.Reflection.BindingFlags]::GetProperty, $null, $rule, $null)
    $exceptions = $ruleType.InvokeMember("Exceptions", [System.Reflection.BindingFlags]::GetProperty, $null, $rule, $null)
    $name = $ruleType.InvokeMember("Name", [System.Reflection.BindingFlags]::GetProperty, $null, $rule, $null)

    # Output rule data
    Write-Output "Rule Name: $name"
    Write-Output "Conditions: $conditions"
    Write-Output "Actions: $actions"
    Write-Output "Exceptions: $exceptions"
}
