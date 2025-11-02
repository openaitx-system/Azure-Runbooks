# Send Recurring Email Runbook

## Overview
This Azure Automation runbook solution sends scheduled recurring emails using Microsoft Graph API and the Automation Account's System-Assigned Managed Identity for authentication. It provides a reliable alternative to PowerAutomate flows that can experience authentication failures and trigger issues.

## Purpose
The primary purpose of this solution is to:
- Send recurring email messages on a defined schedule (weekly, monthly, etc.)
- Eliminate authentication issues common with PowerAutomate flows
- Use Managed Identity authentication (no client secrets required)
- Support customizable HTML email content
- Send from a specified mailbox to any email address or distribution group
- Provide reliable, automated email delivery with retry logic

This solution is ideal for recurring communications such as maintenance reminders, weekly reports, policy reminders, or any scheduled notification that needs to be sent reliably without manual intervention.

## Solution Components
The solution consists of the following components:

1. **Send-RecurringEmail.ps1** - The main Azure Automation runbook script with customizable email template
2. **Add-GraphPermissions.ps1** - Helper script to configure necessary Graph API permissions
3. **email-template-example.html** - Standalone HTML template file for reference and testing
4. **README.md** - This documentation file

## Prerequisites
- An Azure Automation account with a System-Assigned Managed Identity enabled
- The Managed Identity must have the following Microsoft Graph API permission:
  - `Mail.Send`
- The Az.Accounts PowerShell module must be imported into your Azure Automation account
- A valid mailbox in your Microsoft 365 tenant to use as the sender

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| EmailSender | String | Yes | The email address or mailbox that will send the email. Must be a valid mailbox in your M365 tenant. |
| EmailRecipients | String | Yes | Comma-separated list of email addresses or a single distribution group to receive the email. |
| EmailSubject | String | Yes | The subject line for the email message. |
| MaxRetries | Int | No | Maximum number of retry attempts for throttled API requests. Default is 5. |
| InitialBackoffSeconds | Int | No | Initial backoff period in seconds before retrying a throttled request. Default is 5. |
| WhatIf | Switch | No | If specified, shows what would happen without actually sending the email. |

## Setup Instructions

### 1. Enable System-Assigned Managed Identity
1. Navigate to your Azure Automation Account in the Azure Portal
2. Go to **Identity** under the Account Settings section
3. Under **System assigned**, toggle the status to **On**
4. Click **Save**
5. Copy the **Object ID** for use in the next step

### 2. Assign Microsoft Graph Permissions
Run the `Add-GraphPermissions.ps1` script to grant the necessary permissions:

```powershell
.\Add-GraphPermissions.ps1 -AutomationMSI_ID "your-managed-identity-object-id"
```

This script must be run by a user with Global Administrator privileges or sufficient permissions to grant application permissions in Azure AD.

### 3. Import Required PowerShell Modules
In your Azure Automation Account, ensure the following module is imported:
- **Az.Accounts** (usually pre-installed in newer Automation Accounts)

To import modules:
1. Go to **Modules** in your Automation Account
2. Click **Add a module**
3. Browse to the PowerShell Gallery
4. Search for and import **Az.Accounts** if not already present

### 4. Import the Runbook
1. In the Automation account, go to **Runbooks** > **Import a runbook**
2. Upload the `Send-RecurringEmail.ps1` file
3. Set the runbook type to **PowerShell**
4. Click **Create**

### 5. Customize the Email Content
Before scheduling the runbook, customize the HTML email template:

1. Open `Send-RecurringEmail.ps1` in the Azure Portal editor
2. Find the `Get-EmailHtmlBody` function (around line 200)
3. Replace the HTML content within the `$htmlBody` variable with your desired email template
4. Save the runbook

**Tips for customizing:**
- Use inline CSS for all styling (external stylesheets won't work in email)
- Test your HTML in an email client to ensure proper rendering
- Keep images hosted on publicly accessible URLs (like Azure Blob Storage)
- Consider email client compatibility when designing complex layouts

### 6. Create a Schedule
1. Navigate to the runbook > **Schedules** > **Add a schedule**
2. Create a new schedule or link to an existing one
3. Configure the schedule timing (e.g., weekly on Friday at 2:00 PM)
4. Set the required parameters:
   - **EmailSender**: `noreply@yourdomain.com`
   - **EmailRecipients**: `team@yourdomain.com` or `user1@domain.com,user2@domain.com`
   - **EmailSubject**: `Weekly Maintenance Reminder`
5. Save the schedule

### 7. Test the Runbook
Before the first scheduled run, test the runbook manually:

1. Go to the runbook and click **Start**
2. Use the `-WhatIf` parameter to test without sending: Set `WhatIf` to `True`
3. Review the job output to ensure everything works correctly
4. Run again without `-WhatIf` to send a test email
5. Verify the email is received with correct formatting

## Customizing the Email Template

The email template is defined in the `Get-EmailHtmlBody` function within the runbook script. The provided template is a comprehensive example showing various formatting options that IT professionals can adapt for their needs. Replace with your own custom message.

### Quick Customization Steps
1. Open `Send-RecurringEmail.ps1` in the Azure Portal editor
2. Find the `Get-EmailHtmlBody` function (around line 250)
3. Modify the HTML content within the `$htmlBody` variable
4. Save the runbook


### HTML Email Best Practices
1. **Inline CSS**: Always use inline styles, not external stylesheets
2. **Table-based layouts**: Use tables for layout structure (better email client support)
3. **Absolute URLs**: Use full URLs for all images and links
4. **Test thoroughly**: Test in multiple email clients (Outlook, Gmail, Apple Mail)
5. **Keep it simple**: Avoid complex JavaScript or advanced CSS features
6. **Alt text**: Always include alt text for images
7. **Mobile-friendly**: The template uses responsive design that works on mobile devices
8. **File size**: Keep total HTML under 100KB for best deliverability

### Testing Your Template
Before scheduling the runbook:
1. Use the `-WhatIf` parameter to test without sending
2. Send a test email to yourself first
3. Check how it looks in Outlook, Gmail, and mobile devices
4. Verify all links work correctly
5. Make sure images load properly

### Advanced: Multiple Templates
If you need multiple different emails, you have options:
1. **Multiple Runbooks**: Import the script multiple times with different names and templates
2. **Parameter-based**: Modify the script to accept a template name parameter
3. **Separate Template Files**: Store templates in Azure Blob Storage and load them dynamically

## Execution Flow
1. **Authentication**: The script authenticates to Microsoft Graph API using the Automation Account's Managed Identity
2. **Template Loading**: Loads the HTML email body from the `Get-EmailHtmlBody` function
3. **Email Preparation**: Constructs the email message with recipients, subject, and HTML content
4. **Send Email**: Sends the email via Microsoft Graph API using the specified sender mailbox
5. **Retry Logic**: Implements exponential backoff for throttled requests or transient errors
6. **Output Generation**: Returns execution summary with success/failure status

## Monitoring and Troubleshooting

### Viewing Runbook Output
1. Go to your Automation Account > Runbooks > Send-RecurringEmail
2. Click on the **Jobs** tab to see execution history
3. Click on any job to view detailed output and logs

### Common Issues and Solutions

| Issue | Cause | Solution |
|-------|-------|----------|
| "Failed to acquire access token" | Managed Identity not enabled or configured | Enable System-Assigned Managed Identity in Automation Account |
| "Insufficient privileges" | Missing Mail.Send permission | Run Add-GraphPermissions.ps1 script |
| "Mailbox not found" | Invalid sender email address | Verify sender mailbox exists in your tenant |
| "Module not found: Az.Accounts" | Required module not imported | Import Az.Accounts module in Automation Account |
| Email not received | Recipient spam filter or incorrect address | Check recipient's spam folder, verify email address |
| Formatting issues | CSS not supported by email client | Use inline styles and table-based layouts |

### Debugging Tips
1. **Use WhatIf mode**: Test without actually sending emails
2. **Check job output**: Review detailed logs in the Jobs tab
3. **Test sender mailbox**: Ensure the sender mailbox can send emails manually
4. **Verify recipients**: Test with a single known-good email address first
5. **Review Graph API errors**: Look for specific error messages in the output

## Advantages Over PowerAutomate

| Feature | This Runbook | PowerAutomate |
|---------|--------------|---------------|
| Authentication | Managed Identity (no expiration) | User-based (can expire) |
| Reliability | Runs in Azure Automation with SLA | Subject to connector issues |
| Customization | Full PowerShell scripting | Limited to pre-built actions |
| HTML Control | Complete control over HTML | Limited formatting options |
| Cost | Pay per minute of execution | Pay per run |
| Error Handling | Custom retry logic | Basic retry only |
| Monitoring | Azure Automation job history | Flow run history |
| Source Control | Can be stored in Git | Requires export/import |

## Security Considerations
- **Managed Identity**: Uses Azure AD Managed Identity, eliminating the need for stored credentials
- **No Client Secrets**: No secrets to manage, rotate, or secure
- **Least Privilege**: Only requires Mail.Send permission (no read access)
- **Audit Trail**: All executions are logged in Azure Automation job history
- **Sender Validation**: Can only send from valid mailboxes in your tenant

## Maintenance
- **Module Updates**: Periodically check for Az.Accounts module updates in your Automation Account
- **Permission Review**: Audit Managed Identity permissions quarterly
- **Email Content**: Review and update email templates as needed
- **Schedule Review**: Verify schedule timing aligns with organizational needs

## Limitations
- **Email Size**: Graph API has a 4 MB limit for email message size (including attachments)
- **Recipients**: Large recipient lists may be subject to throttling
- **Send Rate**: Subject to Microsoft Graph API throttling limits
- **Sender Restrictions**: Can only send from mailboxes in your tenant

## Support
If you encounter issues:
1. Review the job output logs in Azure Automation
2. Verify all prerequisites are met
3. Test with the WhatIf parameter to validate configuration
4. Check Microsoft Graph API status for service issues
5. Consult the troubleshooting section above
6. Open an issue on the GitHub repository for further assistance

## Contributing
Improvements and suggestions are welcome! Common enhancements:
- Additional email templates
- Dynamic content generation
- Attachment support
- Calendar integration
- Conditional logic for different audiences
