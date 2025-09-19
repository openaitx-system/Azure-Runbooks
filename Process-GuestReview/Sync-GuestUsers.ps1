# Requires -Modules "Az.Accounts"
<#
.SYNOPSIS
    Synchronizes guest users from Azure AD to a SharePoint list for guest governance and management.
    
.DESCRIPTION
    This Azure Runbook script connects to Microsoft Graph API using a System-Assigned Managed Identity,
    retrieves all guest users from Azure AD, and synchronizes their information to a SharePoint list.
    The script updates existing entries and creates new ones as needed, and marks stale entries
    that are no longer present in the tenant.
    
.PARAMETER SiteUrl
    The full URL of the SharePoint site where the guest directory list is located.
    Example: "https://contoso.sharepoint.com/sites/GuestGovernance"
    
.PARAMETER ListName
    The display name of the SharePoint list where guest user information will be stored.
    Default is "GuestDirectory".
    
.PARAMETER BatchSize
    Optional. Number of users to process in each batch. Default is 100.
    
.PARAMETER MaxRetries
    Optional. Maximum number of retry attempts for throttled API requests. Default is 5.
    
.PARAMETER InitialBackoffSeconds
    Optional. Initial backoff period in seconds before retrying a throttled request. Default is 5.
    
.PARAMETER WhatIf
    Optional. If specified, shows what would be done but doesn't actually update SharePoint.
    
.NOTES
    File Name: Sync-GuestUsers.ps1
    Author: Ryan Schultz
    Version: 2.0
    Created: 2025-09-19
    
    Required Graph API Permissions for Managed Identity:
    - User.Read.All
    - Sites.ReadWrite.All
    - AuditLog.Read.All (for sign-in activity)
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory = $false)]
    [string]$ListName = "GuestDirectory",
    
    [Parameter(Mandatory = $false)]
    [int]$BatchSize = 100,
    
    [Parameter(Mandatory = $false)]
    [int]$MaxRetries = 5,
    
    [Parameter(Mandatory = $false)]
    [int]$InitialBackoffSeconds = 5,
    
    [Parameter(Mandatory = $false)]
    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'

function Write-Log {
    param (
        [string]$Message,
        [string]$Type = "INFO"
    )
    
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] [$Type] $Message"
    
    switch ($Type) {
        "ERROR" {
            Write-Error $LogMessage
        }
        "WARNING" {
            Write-Warning $LogMessage
        }
        default {
            Write-Output $LogMessage
        }
    }
}

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
        Write-Log "Failed to acquire Microsoft Graph token using Managed Identity: $_" -Type "ERROR"
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
        if ($Body -is [string]) {
            $params.Add("Body", $Body)
        } else {
            $params.Add("Body", ($Body | ConvertTo-Json -Depth 10))
        }
    }
    
    while ($true) {
        try {
            return Invoke-RestMethod @params
        }
        catch {
            $statusCode = $_.Exception.Response.StatusCode.value__
            $reasonPhrase = $_.Exception.Response.ReasonPhrase
            
            if ($statusCode -eq 429 -or $statusCode -eq 503 -or $statusCode -eq 502) {
                if ($retryCount -lt $MaxRetries) {
                    $retryCount++
                    Write-Log "Request throttled (HTTP $statusCode). Retry $retryCount of $MaxRetries. Waiting $backoffSeconds seconds..." -Type "WARNING"
                    Start-Sleep -Seconds $backoffSeconds
                    $backoffSeconds = [Math]::Min($backoffSeconds * 2, 300)
                    continue
                }
            }
            
            Write-Log "Graph API request failed: HTTP $statusCode - $reasonPhrase" -Type "ERROR"
            Write-Log "Request URI: $Uri" -Type "ERROR"
            throw
        }
    }
}

function Resolve-GraphSiteAndList {
    param(
        [string]$Token,
        [string]$SiteUrl,
        [string]$ListName,
        [int]$MaxRetries = 5,
        [int]$InitialBackoffSeconds = 5
    )
    
    try {
        Write-Log "Resolving SharePoint site and list..."
        $u = [uri]$SiteUrl
        $host = $u.Host                                   # <tenant>.sharepoint.com
        $path = $u.AbsolutePath.TrimStart('/')            # e.g. sites/GuestGovernance
        
        $siteUri = "https://graph.microsoft.com/v1.0/sites/$host`:/$path"
        Write-Log "Getting site information from: $siteUri"
        
        $site = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $siteUri -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        if (-not $site) { 
            throw "Site not found: $SiteUrl" 
        }
        Write-Log "Found SharePoint site: $($site.displayName)"

        $listUri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists?`$filter=displayName eq '$ListName'"
        $listResponse = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $listUri -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        $list = $listResponse.value | Select-Object -First 1
        
        if (-not $list) { 
            throw "List not found: $ListName" 
        }
        Write-Log "Found SharePoint list: $($list.displayName)"
        
        return [pscustomobject]@{ Site = $site; List = $list }
    }
    catch {
        Write-Log "Failed to resolve SharePoint site and list: $_" -Type "ERROR"
        throw
    }
}

function Get-ExistingIndex {
    param(
        [string]$Token,
        [string]$SiteId,
        [string]$ListId,
        [int]$MaxRetries = 5,
        [int]$InitialBackoffSeconds = 5
    )
    
    try {
        Write-Log "Building index of existing items in SharePoint list..."
        $index = @{}
        $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/lists/$ListId/items?`$expand=fields&`$select=id,fields"
        
        do {
            $response = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $uri -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
            
            foreach ($item in $response.value) {
                $uid = $item.fields.UserId
                if ($uid) { 
                    $index[$uid] = $item.id 
                }
            }
            
            $uri = $response.'@odata.nextLink'
        } while ($uri)
        
        Write-Log "Found $($index.Count) existing items in SharePoint list"
        return $index
    }
    catch {
        Write-Log "Failed to build existing items index: $_" -Type "ERROR"
        throw
    }
}

# Main script execution
try {
    Write-Log "=== Guest User SharePoint Sync Started ==="
    Write-Log "Site URL: $SiteUrl"
    Write-Log "List Name: $ListName"
    Write-Log "Batch Size: $BatchSize"
    Write-Log "WhatIf Mode: $($WhatIf.IsPresent)"
    
    $startTime = Get-Date
    
    # ---- Connect with Managed Identity ----
    $token = Get-MsGraphToken

    # ---- Resolve SPO targets ----
    $targets = Resolve-GraphSiteAndList -Token $token -SiteUrl $SiteUrl -ListName $ListName -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
    $siteId = $targets.Site.Id
    $listId = $targets.List.Id

    # ---- Load existing items keyed by UserId ----
    $existing = Get-ExistingIndex -Token $token -SiteId $siteId -ListId $listId -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds

    # ---- Get all guest users (paged) ----
    Write-Log "Retrieving guest users from Azure AD..."
    $select = 'id,displayName,mail,otherMails,userPrincipalName,createdDateTime,accountEnabled,externalUserState,externalUserStateChangeDateTime,signInActivity'
    $uri = "https://graph.microsoft.com/v1.0/users?`$filter=userType eq 'Guest'&`$select=$select&`$top=$BatchSize"
    
    $guests = @()
    do {
        $response = Invoke-MsGraphRequestWithRetry -Token $token -Uri $uri -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
        $guests += $response.value
        $uri = $response.'@odata.nextLink'
    } while ($uri)
    
    Write-Log "Retrieved $($guests.Count) guest users from Azure AD"

    $syncTs = (Get-Date).ToUniversalTime().ToString('o')
    $guestIds = [System.Collections.Generic.HashSet[string]]::new()
    $stats = @{
        TotalGuests = $guests.Count
        UpdatedCount = 0
        CreatedCount = 0
        ErrorCount = 0
    }

    Write-Log "Processing guest users and updating SharePoint list..."
    foreach ($u in $guests) {
        try {
            [void]$guestIds.Add($u.Id)

            $email = if ($u.Mail) { 
                $u.Mail 
            } elseif ($u.OtherMails -and $u.OtherMails.Count) { 
                $u.OtherMails[0] 
            } else { 
                $null 
            }

            # Some properties arrive under AdditionalProperties in SDK
            $extState = $u.AdditionalProperties['externalUserState']
            $extChanged = $u.AdditionalProperties['externalUserStateChangeDateTime']

            $payloadFields = @{
                UserId                  = $u.Id
                DisplayName             = $u.DisplayName
                UPN                     = $u.UserPrincipalName
                Email                   = $email
                CreatedDate             = $u.CreatedDateTime
                AccountEnabled          = $u.AccountEnabled
                ExternalState           = if ($extState) { $extState } else { $u.externalUserState }
                ExternalStateChanged    = if ($extChanged) { $extChanged } else { $u.externalUserStateChangeDateTime }
                LastSignIn              = $u.SignInActivity.LastSignInDateTime
                SyncTimestamp           = $syncTs
            }

            if ($WhatIf) {
                if ($existing.ContainsKey($u.Id)) {
                    Write-Log "WHATIF: Would update existing item for user: $($u.DisplayName) ($($u.UserPrincipalName))"
                } else {
                    Write-Log "WHATIF: Would create new item for user: $($u.DisplayName) ($($u.UserPrincipalName))"
                }
                continue
            }

            if ($existing.ContainsKey($u.Id)) {
                $itemId = $existing[$u.Id]
                $updateUri = "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$listId/items/$itemId/fields"
                Invoke-MsGraphRequestWithRetry -Token $token -Uri $updateUri -Method "PATCH" -Body $payloadFields -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
                $stats.UpdatedCount++
                Write-Log "Updated existing item for user: $($u.DisplayName)"
            } else {
                $body = @{ fields = $payloadFields }
                $createUri = "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$listId/items"
                $new = Invoke-MsGraphRequestWithRetry -Token $token -Uri $createUri -Method "POST" -Body $body -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
                $existing[$u.Id] = $new.id
                $stats.CreatedCount++
                Write-Log "Created new item for user: $($u.DisplayName)"
            }
        }
        catch {
            $stats.ErrorCount++
            Write-Log "Failed to process user $($u.DisplayName): $_" -Type "ERROR"
        }
    }

    # ---- Optional: flag stale list rows (no longer present as guests) ----
    Write-Log "Checking for stale entries in SharePoint list..."
    $staleCount = 0
    foreach ($kv in $existing.GetEnumerator()) {
        if (-not $guestIds.Contains($kv.Key)) {
            try {
                if ($WhatIf) {
                    Write-Log "WHATIF: Would mark item as stale for user ID: $($kv.Key)"
                    $staleCount++
                    continue
                }
                
                $updateUri = "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$listId/items/$($kv.Value)/fields"
                $staleUpdate = @{ 
                    Notes = "Not found in tenant on $syncTs"
                    SyncTimestamp = $syncTs 
                }
                Invoke-MsGraphRequestWithRetry -Token $token -Uri $updateUri -Method "PATCH" -Body $staleUpdate -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
                $staleCount++
                Write-Log "Marked stale entry for user ID: $($kv.Key)"
            }
            catch {
                Write-Log "Failed to mark stale entry for user ID $($kv.Key): $_" -Type "ERROR"
            }
        }
    }

    $endTime = Get-Date
    $duration = $endTime - $startTime

    Write-Log "=== Guest User SharePoint Sync Completed ==="
    Write-Log "Duration: $($duration.TotalMinutes.ToString("0.00")) minutes"
    Write-Log "Total guests processed: $($stats.TotalGuests)"
    Write-Log "Items created: $($stats.CreatedCount)"
    Write-Log "Items updated: $($stats.UpdatedCount)"
    Write-Log "Stale items marked: $staleCount"
    Write-Log "Errors encountered: $($stats.ErrorCount)"
    Write-Log "SharePoint List: $ListName @ $SiteUrl"

    $result = [PSCustomObject]@{
        Success = $true
        Duration = $duration.TotalMinutes
        TotalGuests = $stats.TotalGuests
        CreatedCount = $stats.CreatedCount
        UpdatedCount = $stats.UpdatedCount
        StaleCount = $staleCount
        ErrorCount = $stats.ErrorCount
        SiteUrl = $SiteUrl
        ListName = $ListName
        WhatIf = $WhatIf.IsPresent
    }

    return $result
}
catch {
    Write-Log "Script execution failed: $_" -Type "ERROR"
    throw $_
}
finally {
    Write-Log "Script execution completed"
}
