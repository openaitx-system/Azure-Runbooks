# Snipe-IT User Sync

Sync Microsoft 365 users to a Snipe-IT instance using Azure Automation and Managed Identity.

## What it does
- Reads users from Microsoft Graph with Managed Identity (no app secrets).
- Uses email as the anchor to find or create users in Snipe-IT.
- Ensures the Snipe-IT username equals the mail local-part (before the @).
- Updates name fields and profile fields to match Microsoft 365:
  - first name, last name, display name
  - department (by department_id when resolvable)
  - title, phone, mobile phone
- Creates new Snipe-IT users when not found, with a secure password and:
  - login disabled by default (activated=false)
  - invite email disabled by default
  - both can be enabled via parameters or Automation Variables
- Targets a single Snipe-IT instance per run.
- Includes a dry-run mode via `-WhatIf` that logs intended creates/updates (and department creates) without making changes.

## Files
- `SnipeIT-UserSync.ps1` – the runbook script
- `Add-GraphPermissions.ps1` – helper to assign Graph permissions to your Managed Identity

## Requirements
- Azure Automation account (PowerShell 7 Runtime recommended)
- Module: Az.Accounts
- Managed Identity enabled on the Automation account
- Graph permission assigned to the Managed Identity:
  - `User.Read.All`
- Snipe-IT API token(s) stored as Automation variables (encrypted)

## Automation Variables
Configure a single instance via Automation Variables:
- `SnipeItBaseUrl` (String) – e.g. `https://snipe.company.com`
- `SnipeItApiToken` (Encrypted String)

Behavior toggles (optional):
- `SnipeItUserSyncEnableLogin` (Boolean/String) – default for `-EnableLoginForNewUsers`
- `SnipeItUserSyncSendInvite` (Boolean/String) – default for `-SendInviteForNewUsers`

## Parameters
- `-EnableLoginForNewUsers` (bool) – default false. New Snipe-IT users will be created with activated=true when set.
- `-SendInviteForNewUsers` (bool) – default false. Sends an invite/notification if supported by your Snipe-IT version.
- `-CreateMissingDepartments` (bool) – default false. When true, creates the department in Snipe-IT if missing.
- `-OnlyEnabledUsers` (bool) – default true. Limits sync to enabled accounts.
- `-IncludeGuests` (bool) – default false. Include guest accounts.
- `-WhatIf` (bool) – dry run. Set to `true` to log which users would be created/updated and which departments would be created; no writes occur. In the Azure portal, set the WhatIf parameter to `True`. The job will log: `WhatIf mode enabled: no changes will be written to Snipe-IT.`
- `-SnipeRequestDelayMs` (int) – default 100. Milliseconds to sleep between Snipe-IT operations (helps avoid rate limits).
- `-SnipeMaxRetries` (int) – default 5. Max retries for Snipe-IT 429/5xx responses.
- `-SnipeInitialBackoffSeconds` (int) – default 2. Initial backoff delay for Snipe-IT retries (doubles each attempt).

## Property mapping
- Anchor key: `email` (Microsoft 365 `user.mail` falling back to `user.userPrincipalName`)
- `username`: local-part of email (before `@`)
- `first_name`: `givenName`
- `last_name`: `surname`
- `name`: `displayName`
- `jobtitle`: `jobTitle`
- `phone`: first value from `businessPhones`
- `mobile`: `mobilePhone`
- `department_id`: resolved via Snipe-IT departments lookup (optional creation)

## Setup
1) Enable System-Assigned Managed Identity on the Automation account.
2) Grant Graph permissions using the helper:
   - Import and run `Add-GraphPermissions.ps1` locally:
     - requires Microsoft.Graph.Applications
     - assigns `User.Read.All` to the Managed Identity
3) Create the Automation variables as above and add Snipe-IT API token(s).
4) Import `SnipeIT-UserSync.ps1` as a runbook.
5) Perform a dry run first by starting the runbook with `-WhatIf:$true` (optionally include `-CreateMissingDepartments:$true` to preview department creation). In the Azure portal, set the `WhatIf` field to `True` when starting the job.
6) Add a recurring schedule (e.g., nightly).

## Notes
- Passwords for new users are randomly generated with upper, lower, digit, and special characters.
- The script avoids logging password material.
- Invite behavior is subject to your Snipe-IT version; some deployments ignore `send_email` on user creation.
- Rate limits: If you see 429 errors, increase `-SnipeRequestDelayMs` (e.g., 250–500) and/or `-SnipeMaxRetries`. The runbook honors `Retry-After` and uses exponential backoff.

## Troubleshooting
- Ensure the Managed Identity has `User.Read.All` and that consent has been granted.
- Verify the Snipe-IT API token has permission to create and update users, and optionally departments.
- Use the Automation job output to see per-instance summary counts for created/updated/skipped/errors.
