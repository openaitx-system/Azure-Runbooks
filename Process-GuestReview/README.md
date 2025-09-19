# Guest User SharePoint Sync

This Azure Automation runbook synchronizes guest users from Azure AD to a SharePoint list for guest governance and management purposes.

## Overview

The Guest User SharePoint Sync runbook connects to Microsoft Graph API using Azure Automation's System-Assigned Managed Identity, retrieves all guest users from your Azure AD tenant, and synchronizes their information to a SharePoint list. This enables organizations to maintain a comprehensive view of guest users for governance, compliance, and management purposes.

## Features

- **Comprehensive Guest User Data**: Retrieves guest user information including display name, email, creation date, account status, external user state, and last sign-in activity
- **SharePoint Integration**: Automatically creates and updates items in a SharePoint list to maintain current guest user information
- **Stale Entry Management**: Identifies and marks entries for users who are no longer guests in the tenant
- **Throttling Protection**: Implements retry logic with exponential backoff to handle API throttling
- **Comprehensive Logging**: Detailed logging with timestamps and execution metrics
- **WhatIf Support**: Test mode to preview changes without actually updating SharePoint
- **Batch Processing**: Configurable batch processing for improved performance with large user sets

## Prerequisites

### Azure Automation Account
- Azure Automation Account with System-Assigned Managed Identity enabled
- PowerShell 5.1 or later
- Required Azure Automation modules:
  - `Az.Accounts`

### SharePoint Setup
1. **SharePoint Site**: A SharePoint site where the guest directory list will be created
2. **SharePoint List**: A SharePoint list with the following columns (create manually or the script will attempt to use any existing list):
   - `UserId` (Single line of text) - Primary key
   - `DisplayName` (Single line of text)
   - `UPN` (Single line of text) - User Principal Name
   - `Email` (Single line of text)
   - `CreatedDate` (Date and time)
   - `AccountEnabled` (Yes/No)
   - `ExternalState` (Single line of text)
   - `ExternalStateChanged` (Date and time)
   - `LastSignIn` (Date and time)
   - `SyncTimestamp` (Date and time)
   - `Notes` (Multiple lines of text) - Used for stale entry marking

### Microsoft Graph API Permissions
The Automation Account's Managed Identity requires the following Microsoft Graph application permissions:
- `User.Read.All` - Read all users including guest users
- `Sites.ReadWrite.All` - Read and write to SharePoint sites and lists
- `AuditLog.Read.All` - Read sign-in activity (optional but recommended)

## Installation

### Step 1: Import the Runbook
1. Navigate to your Azure Automation Account in the Azure portal
2. Go to **Process Automation** > **Runbooks**
3. Click **+ Create a runbook**
4. Enter a name (e.g., "Sync-GuestUsers")
5. Select **Runbook type**: PowerShell
6. Click **Create**
7. Copy and paste the contents of `Sync-GuestUsers.ps1` into the runbook editor
8. Click **Save** and then **Publish**

### Step 2: Configure Graph API Permissions
Run the included `Add-GraphPermissions.ps1` script to assign the necessary permissions:

```powershell
.\Add-GraphPermissions.ps1 -AutomationAccountName "YourAutomationAccount" -ResourceGroupName "YourResourceGroup" -SubscriptionId "your-subscription-id"
```

**Note**: You must run this script with an account that has sufficient privileges to manage application permissions in Azure AD.

### Step 3: Create SharePoint List
Create a SharePoint list with the columns mentioned in the prerequisites, or use an existing list that matches the expected schema.

## Parameters

### Required Parameters

- **SiteUrl** (string): The full URL of the SharePoint site where the guest directory list is located
  - Example: `"https://contoso.sharepoint.com/sites/GuestGovernance"`

### Optional Parameters

- **ListName** (string): The display name of the SharePoint list where guest user information will be stored
  - Default: `"GuestDirectory"`

- **BatchSize** (int): Number of users to process in each batch
  - Default: `100`
  - Range: `1-1000`

- **MaxRetries** (int): Maximum number of retry attempts for throttled API requests
  - Default: `5`
  - Range: `1-10`

- **InitialBackoffSeconds** (int): Initial backoff period in seconds before retrying a throttled request
  - Default: `5`
  - Range: `1-60`

- **WhatIf** (switch): Preview mode that shows what would be done without actually updating SharePoint
  - Use: `-WhatIf`

## Usage Examples

### Basic Usage
```powershell
# Sync guest users to SharePoint list
$params = @{
    SiteUrl = "https://contoso.sharepoint.com/sites/GuestGovernance"
    ListName = "GuestDirectory"
}
Start-AzAutomationRunbook -AutomationAccountName "MyAutomationAccount" -Name "Sync-GuestUsers" -Parameters $params
```

### Test Mode (WhatIf)
```powershell
# Preview changes without updating SharePoint
$params = @{
    SiteUrl = "https://contoso.sharepoint.com/sites/GuestGovernance"
    ListName = "GuestDirectory"
    WhatIf = $true
}
Start-AzAutomationRunbook -AutomationAccountName "MyAutomationAccount" -Name "Sync-GuestUsers" -Parameters $params
```

### Custom Batch Size
```powershell
# Use smaller batches for better performance monitoring
$params = @{
    SiteUrl = "https://contoso.sharepoint.com/sites/GuestGovernance"
    ListName = "GuestDirectory"
    BatchSize = 50
}
Start-AzAutomationRunbook -AutomationAccountName "MyAutomationAccount" -Name "Sync-GuestUsers" -Parameters $params
```

## Scheduling

To run this runbook on a schedule:

1. In your Azure Automation Account, go to **Shared Resources** > **Schedules**
2. Click **+ Add a schedule**
3. Configure your desired schedule (e.g., daily at 2:00 AM)
4. Go to **Process Automation** > **Runbooks**
5. Select your Guest User Sync runbook
6. Click **Schedules** > **+ Add a schedule**
7. Link the schedule to the runbook and configure parameters

## Output

The runbook returns a PowerShell object with execution results:

```powershell
@{
    Success = $true
    Duration = 2.35                    # Execution time in minutes
    TotalGuests = 150                  # Total guest users found
    CreatedCount = 5                   # New items created in SharePoint
    UpdatedCount = 143                 # Existing items updated
    StaleCount = 2                     # Items marked as stale
    ErrorCount = 0                     # Number of errors encountered
    SiteUrl = "https://..."            # SharePoint site URL
    ListName = "GuestDirectory"        # SharePoint list name
    WhatIf = $false                    # Whether WhatIf mode was used
}
```

## Logging

The runbook provides comprehensive logging with timestamps:

- **INFO**: General information and progress updates
- **WARNING**: Non-critical issues like API throttling
- **ERROR**: Critical errors that prevent processing

Example log output:
```
[2025-09-19 14:30:15] [INFO] === Guest User SharePoint Sync Started ===
[2025-09-19 14:30:15] [INFO] Site URL: https://contoso.sharepoint.com/sites/GuestGovernance
[2025-09-19 14:30:15] [INFO] Authenticating with Managed Identity...
[2025-09-19 14:30:16] [INFO] Token acquired successfully.
[2025-09-19 14:30:16] [INFO] Resolving SharePoint site and list...
[2025-09-19 14:30:17] [INFO] Found SharePoint site: Guest Governance
[2025-09-19 14:30:17] [INFO] Found SharePoint list: GuestDirectory
[2025-09-19 14:30:18] [INFO] Retrieved 150 guest users from Azure AD
[2025-09-19 14:32:45] [INFO] === Guest User SharePoint Sync Completed ===
```

## Troubleshooting

### Common Issues

1. **Authentication Failed**
   - Ensure System-Assigned Managed Identity is enabled
   - Verify Graph API permissions are correctly assigned
   - Check that the `Az.Accounts` module is available

2. **SharePoint Site/List Not Found**
   - Verify the SiteUrl parameter is correct and accessible
   - Ensure the ListName matches the actual SharePoint list display name
   - Check that the Managed Identity has access to the SharePoint site

3. **Permission Denied**
   - Run the `Add-GraphPermissions.ps1` script to assign required permissions
   - Ensure admin consent has been granted for the permissions

4. **API Throttling**
   - The runbook automatically handles throttling with retry logic
   - Consider reducing BatchSize if throttling persists
   - Increase InitialBackoffSeconds for more conservative retry timing

### Getting Support

For support with this runbook:
1. Check the Azure Automation Account logs for detailed error information
2. Verify all prerequisites are met
3. Test with WhatIf mode to identify potential issues
4. Review the troubleshooting section above

## Version History

- **v2.0** (2025-09-19): Complete refactor to match Azure Runbooks toolkit standards
  - Added Managed Identity authentication
  - Implemented comprehensive logging and error handling
  - Added retry logic with exponential backoff
  - Added WhatIf support and execution metrics
  - Improved parameter validation and documentation

- **v1.0** (Initial): Basic guest user sync functionality

## License

This project is licensed under the MIT License - see the main repository LICENSE file for details.