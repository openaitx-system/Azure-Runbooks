# Requires -Modules "Az.Accounts", "PSAction1"
<#
.SYNOPSIS
    Azure Automation runbook that syncs Intune device categories to Action1 custom attributes using PSAction1 module.

.DESCRIPTION
    This Azure Automation runbook:
    - Uses Managed Identity to authenticate to Microsoft Graph and retrieve Intune device categories
    - Uses PSAction1 module with credentials from Automation Variables to update Action1 custom attributes
    - Matches devices by serial number and syncs categories

.PARAMETER Action1Region
    The Action1 region where your organization is hosted.
    Valid values: NorthAmerica, Europe, Australia
    Default: NorthAmerica

.PARAMETER Action1OrgId
    Your Action1 Organization ID (GUID). Find this in your Action1 console URL.

.PARAMETER Action1ApiClientIdVar
    The name of the Azure Automation Variable (encrypted) that stores the Action1 API Client ID.
    Default: Action1ClientId

.PARAMETER Action1ApiClientSecretVar
    The name of the Azure Automation Variable (encrypted) that stores the Action1 API Client Secret.
    Default: Action1ClientSecret

.PARAMETER Action1CustomAttributeName
    The name of the custom attribute to update in Action1.
    Default: Category

.PARAMETER WhatIf
    If specified, the script will log what changes would be made without actually making them.

.NOTES
    Author:         Ryan Schultz
    Created:        October 2025
    Version:        1.0
    
    Prerequisites:
    - Azure Automation Account with System-Assigned Managed Identity enabled
    - Microsoft Graph API Permission: DeviceManagementManagedDevices.Read.All (assigned to Managed Identity)
    - PSAction1 module imported into Azure Automation account
    - Encrypted Automation Variables: Action1ClientId and Action1ClientSecret
    
.EXAMPLE
    Sync-IntuneToAction1Categories-Runbook.ps1 -Action1OrgId "12345678-1234-1234-1234-123456789012"
    
.EXAMPLE
    Sync-IntuneToAction1Categories-Runbook.ps1 -Action1OrgId "12345678-1234-1234-1234-123456789012" -WhatIf
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet("NorthAmerica", "Europe", "Australia")]
    [string]$Action1Region = "NorthAmerica",
    
    [Parameter(Mandatory = $true)]
    [string]$Action1OrgId,
    
    [Parameter(Mandatory = $false)]
    [string]$Action1ApiClientIdVar = "Action1ClientId",
    
    [Parameter(Mandatory = $false)]
    [string]$Action1ApiClientSecretVar = "Action1ClientSecret",
    
    [Parameter(Mandatory = $false)]
    [string]$Action1CustomAttributeName = "Category",
    
    [Parameter(Mandatory = $false)]
    [switch]$WhatIf
)

#region Helper Functions

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS", "WHATIF")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    switch ($Level) {
        "ERROR"   { Write-Error $Message; Write-Verbose $logMessage -Verbose }
        "WARNING" { Write-Warning $Message; Write-Verbose $logMessage -Verbose }
        "WHATIF"  { Write-Verbose "[WHATIF] $Message" -Verbose }
        default   { Write-Verbose $logMessage -Verbose }
    }
}

function Get-MsGraphToken {
    <#
    .SYNOPSIS
        Acquires a Microsoft Graph API token using Managed Identity.
    #>
    try {
        Write-Log "Authenticating to Microsoft Graph using Managed Identity..."
        Connect-AzAccount -Identity -ErrorAction Stop | Out-Null
        
        $tokenResponse = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com" -ErrorAction Stop
        $token = $tokenResponse.Token
        
        if ($token -is [System.Security.SecureString]) {
            $token = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                [Runtime.InteropServices.Marshal]::SecureStringToBSTR($token)
            )
        }
        
        if ([string]::IsNullOrWhiteSpace($token)) {
            throw "Token acquisition returned an empty token"
        }
        
        Write-Log "Successfully acquired Microsoft Graph token" -Level SUCCESS
        return $token
    }
    catch {
        Write-Log "Failed to acquire Microsoft Graph token: $_" -Level ERROR
        throw
    }
}

function Invoke-GraphRequest {
    <#
    .SYNOPSIS
        Makes a request to Microsoft Graph API with retry logic.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        
        [Parameter(Mandatory = $true)]
        [string]$Token,
        
        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = 5,
        
        [Parameter(Mandatory = $false)]
        [int]$InitialBackoffSeconds = 5
    )
    
    $headers = @{
        "Authorization" = "Bearer $Token"
        "Content-Type"  = "application/json"
    }
    
    $attempt = 0
    $backoff = $InitialBackoffSeconds
    
    while ($attempt -lt $MaxRetries) {
        try {
            $response = Invoke-RestMethod -Uri $Uri -Method GET -Headers $headers -ErrorAction Stop
            return $response
        }
        catch {
            $attempt++
            $statusCode = $_.Exception.Response.StatusCode.value__
            
            if ($statusCode -eq 429 -or $statusCode -ge 500) {
                if ($attempt -lt $MaxRetries) {
                    Write-Log "Request throttled or failed (HTTP $statusCode). Retrying in $backoff seconds... (Attempt $attempt/$MaxRetries)" -Level WARNING
                    Start-Sleep -Seconds $backoff
                    $backoff *= 2
                }
                else {
                    Write-Log "Max retries reached for request to $Uri" -Level ERROR
                    throw
                }
            }
            else {
                Write-Log "Request failed with non-retryable error: $_" -Level ERROR
                throw
            }
        }
    }
}

function Get-IntuneDevices {
    <#
    .SYNOPSIS
        Retrieves all Windows managed devices from Intune with their categories.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Token
    )
    
    Write-Log "Retrieving all Windows managed devices from Intune..."
    
    $devices = @()
    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=operatingSystem eq 'Windows'&`$select=id,deviceName,serialNumber,deviceCategoryDisplayName"
    
    do {
        $response = Invoke-GraphRequest -Uri $uri -Token $Token
        
        if ($response.value) {
            $devices += $response.value
        }
        
        $uri = $response.'@odata.nextLink'
    } while ($uri)
    
    Write-Log "Retrieved $($devices.Count) Windows devices from Intune" -Level SUCCESS
    return $devices
}

#endregion

#region Main Script

try {
    Write-Log "=== Starting Intune to Action1 Category Sync (Runbook) ===" -Level INFO
    Write-Log "Configuration:"
    Write-Log "  Action1 Region: $Action1Region"
    Write-Log "  Action1 Org ID: $Action1OrgId"
    Write-Log "  Custom Attribute: $Action1CustomAttributeName"
    
    if ($WhatIf) {
        Write-Log "WhatIf mode enabled - no changes will be made" -Level WHATIF
    }
    
    # Get Action1 API credentials from Automation Variables
    Write-Log "Retrieving Action1 API credentials from Automation Variables..."
    try {
        $action1ClientId = Get-AutomationVariable -Name $Action1ApiClientIdVar -ErrorAction Stop
        $action1ClientSecret = Get-AutomationVariable -Name $Action1ApiClientSecretVar -ErrorAction Stop
    }
    catch {
        Write-Log "Failed to retrieve Action1 API credentials from variables: $_" -Level ERROR
        throw
    }
    
    if ([string]::IsNullOrWhiteSpace($action1ClientId)) {
        throw "Action1 Client ID is empty. Please ensure the '$Action1ApiClientIdVar' variable is configured."
    }
    
    if ([string]::IsNullOrWhiteSpace($action1ClientSecret)) {
        throw "Action1 Client Secret is empty. Please ensure the '$Action1ApiClientSecretVar' variable is configured."
    }
    
    # Authenticate to Microsoft Graph
    $graphToken = Get-MsGraphToken
    
    # Configure PSAction1
    Write-Log "Configuring PSAction1 module..."
    try {
        Set-Action1Region -Region $Action1Region
        Write-Log "Set Action1 region to: $Action1Region" -Level SUCCESS
        
        Set-Action1Credentials -APIKey $action1ClientId -Secret $action1ClientSecret
        Write-Log "Action1 credentials configured" -Level SUCCESS
        
        Set-Action1DefaultOrg -Org_ID $Action1OrgId
        Write-Log "Set Action1 organization context" -Level SUCCESS
    }
    catch {
        Write-Log "Failed to configure PSAction1: $_" -Level ERROR
        throw
    }
    
    # Get all Windows devices from Intune
    $intuneDevices = Get-IntuneDevices -Token $graphToken
    
    if ($intuneDevices.Count -eq 0) {
        Write-Log "No Windows devices found in Intune. Exiting." -Level WARNING
        return
    }
    
    # Get all endpoints from Action1
    Write-Log "Retrieving endpoints from Action1..."
    try {
        $action1Endpoints = Get-Action1 -Query Endpoints
        Write-Log "Retrieved $($action1Endpoints.Count) endpoints from Action1" -Level SUCCESS
    }
    catch {
        Write-Log "Failed to retrieve Action1 endpoints: $_" -Level ERROR
        throw
    }
    
    if ($action1Endpoints.Count -eq 0) {
        Write-Log "No endpoints found in Action1. Exiting." -Level WARNING
        return
    }
    
    # Build lookup table by serial number
    Write-Log "Building Action1 endpoint lookup table by serial number..."
    $action1Lookup = @{}
    foreach ($endpoint in $action1Endpoints) {
        if (![string]::IsNullOrWhiteSpace($endpoint.serial)) {
            $serialKey = $endpoint.serial.ToLower().Trim()
            $action1Lookup[$serialKey] = $endpoint
        }
    }
    Write-Log "Built lookup table with $($action1Lookup.Count) endpoints"
    
    # Process devices and update Action1
    Write-Log "Processing Intune devices and syncing categories to Action1..."
    
    $stats = @{
        Total              = $intuneDevices.Count
        Matched            = 0
        Updated            = 0
        Skipped            = 0
        NoCategory         = 0
        NotFoundInAction1  = 0
        Errors             = 0
    }
    
    $batchCounter = 0
    $batchSize = 50
    
    foreach ($device in $intuneDevices) {
        $batchCounter++
        
        # Log progress every batch
        if ($batchCounter % $batchSize -eq 0) {
            Write-Log "Processed $batchCounter of $($intuneDevices.Count) devices..."
        }
        
        $deviceName = $device.deviceName
        $serialNumber = $device.serialNumber
        $category = $device.deviceCategoryDisplayName
        
        # Check if device has a serial number
        if ([string]::IsNullOrWhiteSpace($serialNumber)) {
            Write-Log "Device '$deviceName' has no serial number. Skipping." -Level WARNING
            $stats.Skipped++
            continue
        }
        
        # Check if device has a category assigned
        if ([string]::IsNullOrWhiteSpace($category)) {
            Write-Log "Device '$deviceName' (SN: $serialNumber) has no category assigned in Intune. Skipping."
            $stats.NoCategory++
            continue
        }
        
        # Look up device in Action1 by serial number
        $serialKey = $serialNumber.ToLower().Trim()
        $action1Endpoint = $action1Lookup[$serialKey]
        
        if ($null -eq $action1Endpoint) {
            Write-Log "Device '$deviceName' (SN: $serialNumber) not found in Action1. Skipping." -Level WARNING
            $stats.NotFoundInAction1++
            continue
        }
        
        $stats.Matched++
        
        # Check if the category needs to be updated
        $currentCategory = $action1Endpoint.$Action1CustomAttributeName
        if ($currentCategory -eq $category) {
            Write-Log "Device '$deviceName' (SN: $serialNumber) already has category '$category' in Action1. Skipping."
            $stats.Skipped++
            continue
        }
        
        # Update the category in Action1
        try {
            if ($WhatIf) {
                Write-Log "WHATIF: Would update '$deviceName' (SN: $serialNumber) category from '$currentCategory' to '$category'" -Level WHATIF
                $stats.Updated++
            }
            else {
                Write-Log "Updating '$deviceName' (SN: $serialNumber) category from '$currentCategory' to '$category'"
                
                Update-Action1 Modify CustomAttribute -Id $action1Endpoint.id -AttributeName $Action1CustomAttributeName -AttributeValue $category
                
                $stats.Updated++
                Write-Log "Successfully updated category for '$deviceName'" -Level SUCCESS
            }
        }
        catch {
            Write-Log "Failed to update category for '$deviceName' (SN: $serialNumber): $_" -Level ERROR
            $stats.Errors++
        }
    }
    
    # Summary
    Write-Log ""
    Write-Log "=== Sync Complete ===" -Level SUCCESS
    Write-Log "Total Intune Devices: $($stats.Total)"
    Write-Log "Matched in Action1: $($stats.Matched)"
    Write-Log "Categories Updated: $($stats.Updated)"
    Write-Log "Skipped (no change needed): $($stats.Skipped)"
    Write-Log "No Category in Intune: $($stats.NoCategory)"
    Write-Log "Not Found in Action1: $($stats.NotFoundInAction1)"
    Write-Log "Errors: $($stats.Errors)"
    
    return [PSCustomObject]@{
        Timestamp             = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        TotalIntuneDevices    = $stats.Total
        MatchedInAction1      = $stats.Matched
        CategoriesUpdated     = $stats.Updated
        SkippedNoChange       = $stats.Skipped
        NoCategoryInIntune    = $stats.NoCategory
        NotFoundInAction1     = $stats.NotFoundInAction1
        Errors                = $stats.Errors
        WhatIfMode            = $WhatIf.IsPresent
    }
}
catch {
    Write-Log "Critical error in main script execution: $_" -Level ERROR
    throw
}

#endregion