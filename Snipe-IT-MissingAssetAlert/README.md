# Snipe-IT Missing Asset Alert

Find Intune managed devices whose serial numbers aren’t present in your Snipe-IT instance. Uses Azure Automation Managed Identity for Microsoft Graph and an Automation Variable for the Snipe-IT API token.

## What it does
- Authenticates to Microsoft Graph with the Automation Account’s System-Assigned Managed Identity.
- Retrieves Intune managed devices (id, serialNumber, deviceName, operatingSystem, lastSyncDateTime, managedDeviceOwnerType, userDisplayName, userPrincipalName).
- De-duplicates devices by serial number before comparison.
- Loads Snipe-IT hardware serials in bulk via API and compares using a case-insensitive HashSet for accuracy and performance.
- Optionally sends an HTML email listing missing devices with columns: Device Name, OS, Serial, Last Sync, Ownership, Assigned User.
- Outputs a structured summary and a list of missing devices.

## Requirements
- Azure Automation account with System-Assigned Managed Identity enabled.
- Grant the Managed Identity Microsoft Graph app role: DeviceManagementManagedDevices.Read.All.
  - Use `Add-GraphPermissions.ps1` in this folder to assign it.
- If using email (`EmailTo`/`EmailSender`), also grant `Mail.Send` to the Managed Identity and ensure the sender has a mailbox that can send.
- One Encrypted Automation Variable:
  - `SnipeItApiToken`: Snipe-IT API Bearer token.

## Parameters

| Parameter | Required | Default | Description |
|---|:---:|---|---|
| `SnipeItBaseUrl` | Yes | — | Base URL of your Snipe-IT instance, e.g. `https://snipeit.contoso.com`. |
| `SnipeItTokenVar` | No | `SnipeItApiToken` | Name of the Encrypted Automation Variable that holds the Snipe-IT API token. |
| `MaxRetries` | No | `5` | Microsoft Graph retry attempts for throttling/5xx responses. |
| `InitialBackoffSeconds` | No | `5` | Initial backoff (seconds) for Graph retries; doubles on each retry. |
| `Limit` | No | `0` (all) | Cap the number of Intune devices for a quick test run. |
| `EmailTo` | No | empty | If set, sends an HTML email report of missing devices to this address. |
| `EmailSender` | No | empty | Sender mailbox UPN used with Graph `sendMail` (requires `Mail.Send`). |

## Setup
1. Import `SnipeIT-MissingAssetAlert.ps1` as a PowerShell 7.2 runbook.
2. Enable System-Assigned Managed Identity on the Automation Account.
3. Run `Add-GraphPermissions.ps1` once (run locally with your global admin creds) to grant:
  - `DeviceManagementManagedDevices.Read.All`
  - `Mail.Send` (only required if using email notifications)
4. Create an Encrypted Automation Variable named `SnipeItApiToken` with your Snipe-IT API token.

## Schedule
Create a schedule (e.g., daily) and link it to the runbook with parameters like:
- `SnipeItBaseUrl`: `https://snipeit.contoso.com`
- `SnipeItTokenVar`: `SnipeItApiToken`
- Optional email alert example:
  - `EmailTo`: `it-notify@contoso.com`
  - `EmailSender`: `automation-notify@contoso.com`

## Output
- Writes a one-line JSON summary to the job output.
- Returns a PSCustomObject with `Summary` and `Missing` properties (visible in job output/Logs & Output).
- Logs progress every ~100 serials.
- The email/console details include Ownership (Corporate/Personal/Unknown) and Assigned User when available.

## Extending
- Add Teams/Email notification when `MissingCount > 0`.
- Persist `Missing` to a storage account or SharePoint for auditing.
- If your Snipe-IT stores serial in a custom field, extend the bulk loader to include that field when building the HashSet.

## Permissions reference
- Microsoft Graph App ID: `00000003-0000-0000-c000-000000000000`
- App role needed: `DeviceManagementManagedDevices.Read.All` (`2f51be20-0bb4-4fed-bf7b-db946066c75e`).

---

