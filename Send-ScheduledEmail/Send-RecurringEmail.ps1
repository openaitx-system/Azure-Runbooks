<#
.SYNOPSIS
    Sends a scheduled recurring email using Azure Automation and Microsoft Graph API.

.DESCRIPTION
    This Azure Automation runbook sends recurring email messages on a schedule using the Automation Account's
    System-Assigned Managed Identity for authentication. It's designed as a reliable replacement for PowerAutomate
    flows that may fail due to authentication issues.
    
    The script uses Microsoft Graph API to send emails and supports customizable HTML content,
    recipient specification, and sender mailbox selection.

.PARAMETER EmailSender
    The email address or mailbox that will send the email. This must be a valid mailbox in your Microsoft 365 tenant.
    The Managed Identity must have Mail.Send permissions and access to this mailbox.

.PARAMETER EmailRecipients
    A comma-separated list of email addresses or a single distribution group email address that will receive the email.
    Examples: "user@domain.com" or "user1@domain.com,user2@domain.com" or "team@domain.com"

.PARAMETER EmailSubject
    The subject line for the email message.

.PARAMETER MaxRetries
    Optional. Maximum number of retry attempts for throttled API requests. Default is 5.

.PARAMETER InitialBackoffSeconds
    Optional. Initial backoff period in seconds before retrying a throttled request. Default is 5.
    The backoff period doubles with each retry attempt.

.PARAMETER WhatIf
    Optional. If specified, the script will show what would happen without actually sending the email.

.EXAMPLE
    Send-RecurringEmail.ps1 -EmailSender "noreply@contoso.com" -EmailRecipients "team@contoso.com" -EmailSubject "Weekly Maintenance Reminder"

.NOTES
    Author: Ryan Schultz
    Requires: Az.Accounts module
    Graph API Permissions Required: Mail.Send
    
    This script is designed to run as an Azure Automation runbook with a System-Assigned Managed Identity.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$EmailSender,
    
    [Parameter(Mandatory = $true)]
    [string]$EmailRecipients,
    
    [Parameter(Mandatory = $true)]
    [string]$EmailSubject,
    
    [Parameter(Mandatory = $false)]
    [int]$MaxRetries = 5,
    
    [Parameter(Mandatory = $false)]
    [int]$InitialBackoffSeconds = 5,
    
    [Parameter(Mandatory = $false)]
    [switch]$WhatIf
)

# Initialize counters
$emailsSent = 0
$emailsFailed = 0

# ====================================================================================
# LOGGING FUNCTION
# ====================================================================================
function Write-Log {
    param (
        [string]$Message,
        [string]$Type = "INFO"
    )
    
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] [$Type] $Message"
    
    switch ($Type) {
        "ERROR" { 
            Write-Error $Message
            Write-Verbose $LogMessage -Verbose
        }
        "WARNING" { 
            Write-Warning $Message 
            Write-Verbose $LogMessage -Verbose
        }
        "WHATIF" { 
            Write-Verbose "[WHATIF] $Message" -Verbose
        }
        default { 
            Write-Verbose $LogMessage -Verbose
        }
    }
}

# ====================================================================================
# GRAPH API HELPER FUNCTIONS
# ====================================================================================
function Get-MsGraphToken {
    try {
        Write-Log "Authenticating with Managed Identity..."
        Connect-AzAccount -Identity | Out-Null

        $tokenObj = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com"

        if ($tokenObj.Token -is [System.Security.SecureString]) {
            Write-Log "Token is SecureString, converting to plain text..."
            $token = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                [Runtime.InteropServices.Marshal]::SecureStringToBSTR($tokenObj.Token)
            )
        } else {
            Write-Log "Token is plain string, no conversion needed."
            $token = $tokenObj.Token
        }

        if (-not [string]::IsNullOrEmpty($token)) {
            Write-Log "Token acquired successfully."
            return $token
        } else {
            throw "Token was empty."
        }
    }
    catch {
        Write-Error "Failed to acquire Microsoft Graph token using Managed Identity: $_"
        throw
    }
}

function Invoke-MsGraphRequestWithRetry {
    param (
        [string]$Token,
        [string]$Uri,
        [string]$Method = "GET",
        [object]$Body = $null,
        [string]$ContentType = "application/json",
        [int]$MaxRetries = 5,
        [int]$InitialBackoffSeconds = 5
    )
    
    $retryCount = 0
    $backoffSeconds = $InitialBackoffSeconds
    $params = @{
        Uri         = $Uri
        Headers     = @{ Authorization = "Bearer $Token" }
        Method      = $Method
        ContentType = $ContentType
    }
    
    if ($null -ne $Body -and $Method -ne "GET") {
        $params.Add("Body", ($Body | ConvertTo-Json -Depth 10))
    }
    
    while ($true) {
        try {
            return Invoke-RestMethod @params
        }
        catch {
            $statusCode = $null
            if ($_.Exception.Response -ne $null) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }
            
            if (($statusCode -eq 429 -or ($statusCode -ge 500 -and $statusCode -lt 600)) -and $retryCount -lt $MaxRetries) {
                $retryAfter = $backoffSeconds
                if ($_.Exception.Response -ne $null -and $_.Exception.Response.Headers -ne $null) {
                    $retryAfterHeader = $_.Exception.Response.Headers | Where-Object { $_.Key -eq "Retry-After" }
                    if ($retryAfterHeader) {
                        $retryAfter = [int]$retryAfterHeader.Value[0]
                    }
                }
                
                if ($statusCode -eq 429) {
                    Write-Log "Request throttled by Graph API (429). Waiting $retryAfter seconds before retry. Attempt $($retryCount+1) of $MaxRetries" -Type "WARNING"
                }
                else {
                    Write-Log "Server error (5xx). Waiting $retryAfter seconds before retry. Attempt $($retryCount+1) of $MaxRetries" -Type "WARNING"
                }
                
                Start-Sleep -Seconds $retryAfter
                
                $retryCount++
                $backoffSeconds = $backoffSeconds * 2
            }
            else {
                Write-Log "Graph API request failed with status code $statusCode`: $_" -Type "ERROR"
                throw $_
            }
        }
    }
}

# ====================================================================================
# EMAIL SENDING FUNCTION
# ====================================================================================
function Send-GraphMailMessage {
    param (
        [string]$Token,
        [string]$From,
        [string]$To,
        [string]$Subject,
        [string]$HtmlBody,
        [switch]$WhatIf,
        [int]$MaxRetries = 5,
        [int]$InitialBackoffSeconds = 5
    )
    
    try {
        if ($WhatIf) {
            Write-Log "Would send email from $From to $To with subject: $Subject" -Type "WHATIF"
            return $true
        }
        
        Write-Log "Preparing to send email to $To using Microsoft Graph API..."
        
        # Parse recipients (could be comma-separated) and create proper array
        $recipientList = @()
        $To -split ',' | ForEach-Object {
            $recipientList += @{
                emailAddress = @{
                    address = $_.Trim()
                }
            }
        }
        
        $uri = "https://graph.microsoft.com/v1.0/users/$From/sendMail"
        
        # Build the message body with proper structure
        $messageBody = @{
            message = @{
                subject = $Subject
                body = @{
                    contentType = "HTML"
                    content = $HtmlBody
                }
                toRecipients = $recipientList
            }
            saveToSentItems = "true"
        }
        
        Invoke-MsGraphRequestWithRetry -Token $Token -Uri $uri -Method "POST" -Body $messageBody -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        
        Write-Log "Email sent to $To successfully"
        return $true
    }
    catch {
        Write-Log "Failed to send email to $To`: $_" -Type "ERROR"
        return $false
    }
}

# ====================================================================================
# HTML EMAIL TEMPLATE
# ====================================================================================
function Get-EmailHtmlBody {
    # CUSTOMIZE THIS SECTION WITH YOUR HTML CONTENT
    # Replace the HTML below with your desired email template
    # You can use inline CSS for styling
    
    $htmlBody = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Email Notification</title>
</head>
<body style="margin: 0; padding: 0; font-family: Arial, sans-serif; background-color: #f4f4f4;">
    <div style="max-width: 800px; margin: 0 auto; border: 1px solid #ddd; background-color: #ffffff;">
        
        <!-- ============================================================ -->
        <!-- HEADER SECTION - Customize with your logo and colors -->
        <!-- ============================================================ -->
        <div style="background-color: #0078d4; color: white; padding: 20px; text-align: center;">
            <!-- Replace with your organization's logo URL -->
            <img src="https://via.placeholder.com/150x50/0078d4/FFFFFF?text=Your+Logo" alt="Company Logo" width="150" style="display: inline-block;">
            <h2 style="margin-top: 15px; margin-bottom: 0;">Email Subject/Title</h2>
        </div>
        
        <!-- ============================================================ -->
        <!-- CONTENT SECTION - Customize with your message -->
        <!-- ============================================================ -->
        <div style="padding: 20px;">
            
            <!-- Greeting/Introduction -->
            <div style="margin-bottom: 20px;">
                <p>Hello Team,</p>
                <p>This is a sample recurring email template. Replace this text with your message content.</p>
            </div>
            
            <!-- Main Content Section 1 -->
            <div style="margin-bottom: 25px;">
                <h3 style="margin-top: 0; border-bottom: 2px solid #0078d4; padding-bottom: 10px; color: #0078d4;">
                    Section 1: Main Topic
                </h3>
                <p>Add your content here. You can include:</p>
                <ul style="margin-top: 10px; padding-left: 20px;">
                    <li style="margin-bottom: 8px;">Bullet points for easy reading</li>
                    <li style="margin-bottom: 8px;">Important information or instructions</li>
                    <li style="margin-bottom: 8px;">Key dates or deadlines</li>
                    <li style="margin-bottom: 8px;">Action items or next steps</li>
                </ul>
            </div>
            
            <!-- Main Content Section 2 -->
            <div style="margin-bottom: 25px;">
                <h3 style="margin-top: 0; border-bottom: 2px solid #0078d4; padding-bottom: 10px; color: #0078d4;">
                    Section 2: Additional Information
                </h3>
                <p>You can add multiple sections as needed. Here are some examples:</p>
                <ol style="margin-top: 10px; padding-left: 20px;">
                    <li style="margin-bottom: 8px;"><strong>Numbered lists</strong> - For sequential steps or priorities</li>
                    <li style="margin-bottom: 8px;"><strong>Bold text</strong> - To emphasize important points</li>
                    <li style="margin-bottom: 8px;"><strong>Links</strong> - <a href="https://example.com" style="color: #0078d4; text-decoration: underline;">Like this example link</a></li>
                </ol>
            </div>
            
            <!-- Highlighted/Important Notice Box -->
            <div style="margin: 25px 0; padding: 15px; background-color: #fff4ce; border-left: 4px solid #ffb900;">
                <p style="margin: 0;"><strong>⚠️ Important Notice:</strong> Use this styled box to highlight critical information, warnings, or time-sensitive items that need attention.</p>
            </div>
            
            <!-- Alternative Highlight Box (Info Style) -->
            <div style="margin: 25px 0; padding: 15px; background-color: #e6f3ff; border-left: 4px solid #0078d4;">
                <p style="margin: 0;"><strong>💡 Pro Tip:</strong> You can use different colors for different types of notices. This blue box is great for helpful tips or information.</p>
            </div>
            
            <!-- Success/Positive Message Box -->
            <div style="margin: 25px 0; padding: 15px; background-color: #e6f5e6; border-left: 4px solid #107c10;">
                <p style="margin: 0;"><strong>✅ Success Message:</strong> Use green styling for positive messages, confirmations, or success notifications.</p>
            </div>
            
            <!-- Table Example (Optional) -->
            <div style="margin-bottom: 25px;">
                <h3 style="margin-top: 0; border-bottom: 2px solid #0078d4; padding-bottom: 10px; color: #0078d4;">
                    Section 3: Table Example (Optional)
                </h3>
                <p>Tables are useful for structured data:</p>
                <table style="width: 100%; border-collapse: collapse; margin-top: 10px;">
                    <thead>
                        <tr style="background-color: #0078d4; color: white;">
                            <th style="padding: 10px; text-align: left; border: 1px solid #ddd;">Item</th>
                            <th style="padding: 10px; text-align: left; border: 1px solid #ddd;">Description</th>
                            <th style="padding: 10px; text-align: left; border: 1px solid #ddd;">Status</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td style="padding: 10px; border: 1px solid #ddd;">Example 1</td>
                            <td style="padding: 10px; border: 1px solid #ddd;">Sample description</td>
                            <td style="padding: 10px; border: 1px solid #ddd;">✅ Complete</td>
                        </tr>
                        <tr style="background-color: #f9f9f9;">
                            <td style="padding: 10px; border: 1px solid #ddd;">Example 2</td>
                            <td style="padding: 10px; border: 1px solid #ddd;">Another example</td>
                            <td style="padding: 10px; border: 1px solid #ddd;">⏳ In Progress</td>
                        </tr>
                        <tr>
                            <td style="padding: 10px; border: 1px solid #ddd;">Example 3</td>
                            <td style="padding: 10px; border: 1px solid #ddd;">Third example</td>
                            <td style="padding: 10px; border: 1px solid #ddd;">📅 Scheduled</td>
                        </tr>
                    </tbody>
                </table>
            </div>
            
            <!-- Call to Action / Next Steps -->
            <div style="margin-bottom: 25px;">
                <h3 style="margin-top: 0; border-bottom: 2px solid #0078d4; padding-bottom: 10px; color: #0078d4;">
                    What You Need to Do
                </h3>
                <p>Clearly state any actions recipients need to take:</p>
                <ul style="margin-top: 10px; padding-left: 20px;">
                    <li style="margin-bottom: 8px;">Action item 1 with deadline</li>
                    <li style="margin-bottom: 8px;">Action item 2 with instructions</li>
                    <li style="margin-bottom: 8px;">Where to find more information</li>
                </ul>
            </div>
            
            <!-- Contact/Support Information -->
            <div style="margin-top: 20px;">
                <p>If you have any questions or need assistance, please contact:</p>
                <p style="margin-left: 20px;">
                    <strong>IT Support:</strong> <a href="mailto:support@example.com" style="color: #0078d4; text-decoration: underline;">support@example.com</a><br>
                    <strong>Phone:</strong> (555) 123-4567<br>
                    <strong>Help Portal:</strong> <a href="https://support.example.com" style="color: #0078d4; text-decoration: underline;">https://support.example.com</a>
                </p>
                <p>Thank you for your attention!</p>
            </div>
            
        </div>
        
        <!-- ============================================================ -->
        <!-- FOOTER SECTION - Standard disclaimer -->
        <!-- ============================================================ -->
        <div style="padding: 15px; background-color: #f2f2f2; border-top: 1px solid #ddd; text-align: center; font-size: 12px; color: #666;">
            <p style="margin: 0;"><strong>Do not reply to this message.</strong> This email was sent from an automated system.</p>
            <p style="margin: 5px 0 0 0;">© 2025 Your Organization Name. All rights reserved.</p>
        </div>
        
    </div>
</body>
</html>
"@
    
    return $htmlBody
}

# ====================================================================================
# MAIN EXECUTION
# ====================================================================================
try {
    Write-Log "========================================" -Type "INFO"
    Write-Log "Starting Send-RecurringEmail Runbook" -Type "INFO"
    Write-Log "========================================" -Type "INFO"
    
    if ($WhatIf) {
        Write-Log "Running in WhatIf mode - no emails will be sent" -Type "WHATIF"
    }
    
    # Authenticate and get access token
    $token = Get-MsGraphToken
    
    # Get the HTML email body
    Write-Log "Loading email template..."
    $htmlBody = Get-EmailHtmlBody
    
    # Send the email
    Write-Log "Attempting to send email..."
    Write-Log "  From: $EmailSender"
    Write-Log "  To: $EmailRecipients"
    Write-Log "  Subject: $EmailSubject"
    
    try {
        $result = Send-GraphMailMessage -Token $token `
                                         -From $EmailSender `
                                         -To $EmailRecipients `
                                         -Subject $EmailSubject `
                                         -HtmlBody $htmlBody `
                                         -WhatIf:$WhatIf `
                                         -MaxRetries $MaxRetries `
                                         -InitialBackoffSeconds $InitialBackoffSeconds
        
        if ($result) {
            $emailsSent++
        } else {
            $emailsFailed++
        }
    }
    catch {
        $emailsFailed++
        Write-Log "Email sending failed: $_" -Type "ERROR"
    }
    
    # Output summary
    Write-Log "========================================" -Type "INFO"
    Write-Log "Execution Summary:" -Type "INFO"
    Write-Log "  Emails Sent: $emailsSent"
    Write-Log "  Emails Failed: $emailsFailed"
    Write-Log "========================================" -Type "INFO"
    
    # Return structured output
    $output = [PSCustomObject]@{
        EmailsSent = $emailsSent
        EmailsFailed = $emailsFailed
        Sender = $EmailSender
        Recipients = $EmailRecipients
        Subject = $EmailSubject
        ExecutionTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        WhatIfMode = $WhatIf.IsPresent
    }
    
    return $output
}
catch {
    Write-Log "Fatal error in runbook execution: $_" -Type "ERROR"
    Write-Log $_.ScriptStackTrace -Type "ERROR"
    throw
}
finally {
    Write-Log "Runbook execution completed" -Type "INFO"
}