# Sync Intune Device Categories to Action1

Azure Automation runbook that synchronizes Microsoft Intune device categories to Action1 custom attributes using the PSAction1 PowerShell module.

## Overview

This runbook:
1. Uses Managed Identity to authenticate to Microsoft Graph and retrieve Intune device categories
2. Uses the PSAction1 module with credentials from Automation Variables to connect to Action1
3. Matches devices by serial number between Intune and Action1
4. Updates custom attributes in Action1 using PSAction1's native `Update-Action1 Modify CustomAttribute` command

## Prerequisites

### Azure Requirements
- Azure Automation account with **PowerShell 7.2** runtime
- System-Assigned Managed Identity enabled on the Automation account
- Required PowerShell modules in the Automation account:
  - `Az.Accounts` (typically pre-installed)
  - **`PSAction1`** (must be imported - see instructions below)

### Action1 Requirements
- Action1 organization with API access enabled
- Action1 API credentials (Client ID and Client Secret)
- Custom attribute configured in Action1
  - Go to **Configuration** > **Advanced** in Action1
  - You can rename the display label (e.g., "Custom Attribute 1" → "Category")
  - When running the runbook, use the display name you configured (e.g., "Category")

### Permissions
- **Microsoft Graph API**: `DeviceManagementManagedDevices.Read.All` (assigned to Managed Identity)
- **Action1 API**: Token with permissions to update endpoint custom attributes

## Setup Instructions

### 1. Enable System-Assigned Managed Identity

1. Navigate to your Azure Automation account in the Azure Portal
2. Go to **Identity** > **System assigned**
3. Set Status to **On** and click **Save**
4. Copy the **Object ID** - you'll need this for the next step

### 2. Assign Microsoft Graph Permissions

1. Ensure you have the `Microsoft.Graph.Applications` PowerShell module installed locally:
   ```powershell
   Install-Module Microsoft.Graph.Applications -Scope CurrentUser
   ```

2. Run the `Add-GraphPermissions.ps1` script locally with appropriate credentials:
   ```powershell
   .\Add-GraphPermissions.ps1 -AutomationMSI_ID "<YOUR_MSI_OBJECT_ID>"
   ```

3. Authenticate with an account that has:
   - `AppRoleAssignment.ReadWrite.All`
   - `Application.Read.All`

### 3. Import PSAction1 Module to Azure Automation

**Option A: Import from PowerShell Gallery (Recommended)**

1. In your Azure Automation account, go to **Modules**
2. Click **Browse gallery**
3. Search for **PSAction1**
4. Click on **PSAction1**
5. Click **Import**
6. Select **Runtime version**: **7.2**
7. Click **Import**
8. Wait for the import to complete (may take several minutes)

**Option B: Manual Import**

1. Download PSAction1 from PowerShell Gallery: https://www.powershellgallery.com/packages/PSAction1
2. In your Automation account, go to **Modules** > **Add a module**
3. Upload the `.nupkg` file
4. Select **Runtime version**: **7.2**
5. Click **Import**

### 4. Create Action1 API Credentials Automation Variables

1. In Action1, generate API credentials:
   - Navigate to **Settings** > **API Credentials** in the Action1 console
   - Click **Generate API Credentials**
   - Copy both the **Client ID** and **Client Secret**
   - The Client ID will look like: `api-key-9844e782-2506-7488-f599-a5693ce5210737efced4-30ff-4476-839e-044805c3725b@action1.com`
   - The Client Secret will be a hexadecimal string like: `b8513eec230b2af9ae7670fb9c4d0644`

2. In your Azure Automation account, create two encrypted variables:

   **Variable 1 - Client ID:**
   - Go to **Shared Resources** > **Variables**
   - Click **Add a variable**
   - Name: `Action1ClientId`
   - Type: **String**
   - Value: Your Action1 Client ID
   - **Encrypted**: Check this box ✓
   - Click **Create**

   **Variable 2 - Client Secret:**
   - Click **Add a variable** again
   - Name: `Action1ClientSecret`
   - Type: **String**
   - Value: Your Action1 Client Secret
   - **Encrypted**: Check this box ✓
   - Click **Create**

### 5. Import the Runbook

1. In your Azure Automation account, go to **Runbooks**
2. Click **Import a runbook**
3. Select the `Sync-IntuneToAction1Categories-Runbook.ps1` file
4. Set **Runbook type** to **PowerShell**
5. Set **Runtime version** to **7.2**
6. Click **Import**
7. Once imported, click **Publish**

### 6. Test the Runbook

1. Click **Start** on the runbook
2. Configure parameters:
   - **Action1OrgId**: Your Organization ID (see below)
   - **Action1Region**: NorthAmerica, Europe, or Australia (default: NorthAmerica)
   - **Action1CustomAttributeName**: The display name of your custom attribute (e.g., "Category")
   - **WhatIf**: Set to `$true` for a dry-run
3. Click **OK** to start the test
4. Monitor the **Output** tab for progress and results

### 7. Schedule the Runbook

1. In the runbook, go to **Schedules**
2. Click **Add a schedule**
3. Create a new schedule or link to an existing one
   - **Recommended frequency**: Daily or weekly
4. Configure the parameters (especially `Action1OrgId`)
5. Save the schedule

## Parameters

| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `Action1OrgId` | Yes | - | Your Action1 Organization ID (GUID from URL) |
| `Action1Region` | No | `NorthAmerica` | Action1 region: NorthAmerica, Europe, or Australia |
| `Action1ApiClientIdVar` | No | `Action1ClientId` | Name of the Automation Variable storing the Client ID |
| `Action1ApiClientSecretVar` | No | `Action1ClientSecret` | Name of the Automation Variable storing the Client Secret |
| `Action1CustomAttributeName` | No | `Category` | The **display name** of the custom attribute (what you see in UI) |
| `WhatIf` | No | `$false` | If true, logs changes without making them |

## Finding Your Action1 Organization ID

1. Log in to your Action1 console
2. Look at the URL in your browser
3. Find the `?org=` parameter
4. Example: `https://app.action1.com/console/dashboard?org=88c8b425-871e-4ff6-9afc-00df8592c6db`
5. Your Org ID: `88c8b425-871e-4ff6-9afc-00df8592c6db`

## How It Works

### Matching Logic

1. Retrieves all Windows managed devices from Intune via Microsoft Graph
2. Retrieves all endpoints from Action1 via PSAction1 module
3. Builds a lookup table keyed by serial number (case-insensitive)
4. For each Intune device:
   - Matches by serial number with Action1 endpoint
   - Compares current category value
   - Updates if different using PSAction1's `Update-Action1 Modify CustomAttribute`

### Sync Behavior

- **No serial number**: Device skipped
- **No Intune category**: Device skipped
- **Not found in Action1**: Device skipped (logged as warning)
- **Categories match**: Device skipped (no API call)
- **Categories differ**: Custom attribute updated

## Output

The runbook returns a summary object with the following metrics:

```powershell
@{
    Timestamp             = "2025-10-18 19:58:23"
    TotalIntuneDevices    = 117
    MatchedInAction1      = 45
    CategoriesUpdated     = 12
    SkippedNoChange       = 33
    NoCategoryInIntune    = 10
    NotFoundInAction1     = 72
    Errors                = 0
    WhatIfMode            = False
}
```

All operations are logged with timestamps and severity levels (INFO, WARNING, ERROR, SUCCESS) for easy troubleshooting.

## Troubleshooting

### "Failed to acquire Microsoft Graph token"
- Ensure System-Assigned Managed Identity is enabled
- Verify the required Graph API permissions are assigned
- Check that the Az.Accounts module is imported in the Automation account

### "Failed to retrieve Action1 API credentials from variables"
- Verify the Automation Variable names match the parameters (default: `Action1ClientId` and `Action1ClientSecret`)
- Ensure both variables are marked as encrypted
- Check that the values were correctly copied from Action1 console (no extra spaces)

### "Failed to configure PSAction1"
- Verify the PSAction1 module is imported into the Automation account with runtime 7.2
- Check that the Client ID and Client Secret are correct
- Ensure the Action1 region is correct (NorthAmerica/Europe/Australia)
- Verify the Organization ID is correct

### "Device not found in Action1"
- Verify the device has the Action1 agent installed
- Check that serial numbers match between systems
- Ensure the device has synced to Action1 at least once

### "Failed to update category"
- Verify the custom attribute name matches the display name in Action1 UI
- Check that the attribute is defined in Action1 (Configuration > Advanced)
- Ensure your API credentials have permission to modify endpoints

### PSAction1 Module Import Issues
- Ensure you're using PowerShell 7.2 runtime
- Try removing and re-importing the module
- Check the module gallery for the latest version
- Verify the module shows as "Available" in the Modules section

## Security Considerations

- **Managed Identity**: Uses Azure AD-based authentication, no stored credentials for Graph API
- **Encrypted Variables**: Action1 Client ID and Client Secret are stored encrypted in Azure Automation
- **Least Privilege**: Only requests read access to Intune devices
- **Audit Trail**: All operations are logged with timestamps for compliance

## Key Differences from Local Script

**Runbook version:**
- ✅ **Automated**: Runs on schedule without manual intervention
- ✅ **Managed Identity**: No interactive authentication required
- ✅ **Centralized**: Runs in Azure, not on a specific workstation
- ✅ **Logged**: Full audit trail in Azure Automation job history
- ✅ **PSAction1**: Uses the same module that works locally

**Setup differences:**
- Requires PSAction1 module import (one-time)
- Uses Automation Variables instead of interactive prompts
- Uses Managed Identity instead of interactive Graph authentication

## Additional Resources

- [Microsoft Graph API - Intune Devices](https://learn.microsoft.com/en-us/graph/api/resources/intune-devices-manageddevice)
- [PSAction1 Module](https://github.com/Action1Corp/PSAction1)
- [Action1 Custom Attributes](https://www.action1.com/documentation/custom-attributes/)
- [Azure Automation Managed Identity](https://learn.microsoft.com/en-us/azure/automation/enable-managed-identity-for-automation)

## Version History

- **1.0** (October 2025) - Azure Automation runbook version
  - Uses PSAction1 module for custom attribute updates
  - Managed Identity authentication for Microsoft Graph
  - Automation Variables for Action1 credentials
  - Serial number-based matching
  - Comprehensive error handling and logging