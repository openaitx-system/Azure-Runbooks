<#
.SYNOPSIS
  Sync Microsoft 365 users to a Snipe-IT instance via the Snipe-IT API.

.DESCRIPTION
  Azure Automation runbook that uses the Automation Account's System-Assigned Managed Identity to read users
  from Microsoft Graph, then creates or updates matching users in Snipe-IT.

  Key behaviors
   - Anchor mapping by email address from Microsoft 365 (prefers user.mail, falls back to user.userPrincipalName).
   - Username in Snipe-IT set to the mail local-part (text before the @).
   - Upsert logic: if user exists (email match), update properties; else create new with secure password.
   - New users: login disabled (activated=false) and invite email disabled by default; can be enabled via parameters.
   - Syncs these properties to Snipe-IT: first_name, last_name, name (display), jobtitle, phone, mobile, department (by department_id when resolvable).
  - Targets a single Snipe-IT instance configured via Automation Variables.

.REQUIREMENTS
  PowerShell 7 runbook
  Az.Accounts module (for MSI auth)
  Microsoft Graph application permission: User.Read.All (assign to Managed Identity)

.AUTOMATION VARIABLES
  SnipeItBaseUrl     (String) Snipe-IT base URL (e.g., https://snipe.company.com)
  SnipeItApiToken    (Encrypted String) Snipe-IT API token
  SnipeItUserSyncEnableLogin  (Boolean/String) Optional default for -EnableLoginForNewUsers.
  SnipeItUserSyncSendInvite   (Boolean/String) Optional default for -SendInviteForNewUsers.

.PARAMETER EnableLoginForNewUsers
  When true, new users are created as activated (login enabled). Default: false.

.PARAMETER SendInviteForNewUsers
  When true, attempt to send invite/notification on user creation (if supported by Snipe-IT). Default: false.

.PARAMETER CreateMissingDepartments
  If a department name from Microsoft 365 is not present in Snipe-IT, create it and assign department_id. Default: false.

.PARAMETER OnlyEnabledUsers
  Sync only users where accountEnabled is true. Default: true.

.PARAMETER IncludeGuests
  Include users with userType == Guest. Default: false.

.PARAMETER WhatIf
  Boolean dry-run. Set to true to log intended changes without creating or updating anything.

.OUTPUTS
  A summary object: counts for created, updated, skipped, and errors.

.NOTES
  Author: Ryan Schultz
  Version: 1.0.0
  Date: 2025-09-06
#>

param(
  [Parameter(Mandatory = $false)] [bool] $EnableLoginForNewUsers = $false,
  [Parameter(Mandatory = $false)] [bool] $SendInviteForNewUsers  = $false,
  [Parameter(Mandatory = $false)] [bool] $CreateMissingDepartments = $false,
  [Parameter(Mandatory = $false)] [bool] $OnlyEnabledUsers = $true,
  [Parameter(Mandatory = $false)] [bool] $IncludeGuests = $false,
  [Parameter(Mandatory = $false)] [int]  $SnipeRequestDelayMs = 100,
  [Parameter(Mandatory = $false)] [int]  $SnipeMaxRetries = 5,
  [Parameter(Mandatory = $false)] [int]  $SnipeInitialBackoffSeconds = 2,
  [bool] $WhatIf = $false
)

# Requires -Modules "Az.Accounts"

Set-StrictMode -Version Latest
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
      Write-Error $Message
      Write-Verbose $LogMessage -Verbose
    }
    "WARNING" { 
      Write-Warning $Message 
      Write-Verbose $LogMessage -Verbose
    }
    default { 
      Write-Verbose $LogMessage -Verbose
    }
  }
}

function Get-ObjectPropertySafe {
  param(
    [Parameter(Mandatory=$true)] [object]$Object,
    [Parameter(Mandatory=$true)] [string]$Name
  )
  if ($null -eq $Object) { return $null }
  $prop = $Object.PSObject.Properties[$Name]
  if ($prop) { return $prop.Value }
  return $null
}

function Get-BooleanFromVar([string]$Name, [bool]$Default) {
  try {
    $v = Get-AutomationVariable -Name $Name
    if ($null -eq $v -or "$v" -eq '') { return $Default }
    # Handle string-like bools
    $s = "$v".Trim().ToLower()
    if ($s -in @('true','1','yes','y')) { return $true }
    if ($s -in @('false','0','no','n')) { return $false }
    return [bool]$v
  } catch { return $Default }
}

function Get-SnipeConfig() {
  $baseUrl = $null; $token = $null
  try { $baseUrl = Get-AutomationVariable -Name 'SnipeItBaseUrl' } catch {}
  try { $token   = Get-AutomationVariable -Name 'SnipeItApiToken' } catch {}
  if (-not $baseUrl) { throw "Missing Automation Variable 'SnipeItBaseUrl'" }
  if (-not $token)   { throw "Missing Automation Variable 'SnipeItApiToken'" }
  return [pscustomobject]@{ BaseUrl = ("$baseUrl").TrimEnd('/'); ApiToken = $token }
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

function Get-TenantUsers([string]$Token, [bool]$OnlyEnabled, [bool]$IncludeGuests) {
  $select = 'id,displayName,givenName,surname,mail,userPrincipalName,department,jobTitle,businessPhones,mobilePhone,userType,accountEnabled'
  $url = "https://graph.microsoft.com/v1.0/users?$`select=$select&$`top=999"
  $all = @()
  do {
    $resp = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $url
    if ($resp -and $resp.value) { $all += $resp.value }
    $nextUrl = $null
    if ($resp) {
      try { $nextUrl = $resp.'@odata.nextLink' } catch { $nextUrl = $null }
      if (-not $nextUrl -and ($resp.PSObject -and $resp.PSObject.Properties.Name -contains '@odata.nextLink')) {
        $nextUrl = ($resp.PSObject.Properties | Where-Object Name -eq '@odata.nextLink').Value
      }
    }
    $url = $nextUrl
  } while ($url)

  $out = $all
  if ($OnlyEnabled) { $out = $out | Where-Object { $_.accountEnabled -eq $true } }
  if (-not $IncludeGuests) { $out = $out | Where-Object { $_.userType -ne 'Guest' } }
  return $out
}

function New-SecurePassword([int]$Length = 20) {
  if ($Length -lt 12) { $Length = 12 }
  $upper = 'ABCDEFGHJKLMNPQRSTUVWXYZ'
  $lower = 'abcdefghijkmnpqrstuvwxyz'
  $digits = '23456789'
  $special = '!@#$%^&*()-_=+[]{}:,.?'
  $all = ($upper + $lower + $digits + $special)
  $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
  function Pick([string]$chars) {
    $b = New-Object byte[] 4; $rng.GetBytes($b); $idx = [BitConverter]::ToUInt32($b,0) % $chars.Length; return $chars[$idx]
  }
  $req = @((Pick $upper),(Pick $lower),(Pick $digits),(Pick $special))
  $remain = for ($i=0; $i -lt ($Length - $req.Count); $i++) { Pick $all }
  return -join ($req + $remain)
}

function Invoke-SnipeRequestWithRetry {
  param(
    [string]$BaseUrl,
    [hashtable]$Headers,
    [string]$Path,
    [ValidateSet('GET','POST','PATCH','PUT','DELETE')] [string]$Method = 'GET',
    [object]$Body,
    [int]$MaxRetries = 5,
    [int]$InitialBackoffSeconds = 2
  )
  $uri = "$BaseUrl$Path"
  $attempt = 0
  $backoff = $InitialBackoffSeconds
  while ($true) {
    $attempt++
    $params = @{ Uri = $uri; Method = $Method; Headers = $Headers }
    if ($PSBoundParameters.ContainsKey('Body')) {
      $params.ContentType = 'application/json'
      $params.Body = ($Body | ConvertTo-Json -Depth 8)
    }
    try {
      $resp = Invoke-RestMethod @params
      return $resp
    } catch {
      $code = $null; $retryAfter = $backoff
      if ($_.Exception.Response) {
        try { $code = [int]$_.Exception.Response.StatusCode } catch {}
        try {
          $hdrs = $_.Exception.Response.Headers
          if ($hdrs) {
            $ra = $hdrs | Where-Object { $_.Key -match 'Retry-After' }
            if ($ra) { $retryAfter = [int]($ra.Value[0]) }
          }
        } catch {}
      }
      if (($code -eq 429 -or ($code -ge 500 -and $code -lt 600)) -and $attempt -le $MaxRetries) {
        Write-Log "Snipe-IT $Method $Path failed with $code. Retrying in $retryAfter s (attempt $attempt/$MaxRetries)" -Type WARNING
        Start-Sleep -Seconds $retryAfter
        $backoff = [Math]::Min($backoff * 2, 60)
        # On retry, also attempt form-encoded for write methods
        if ($Method -in @('POST','PATCH','PUT') -and $Body) {
          try {
            $params.ContentType = 'application/x-www-form-urlencoded'
            $params.Body = $Body
            $resp2 = Invoke-RestMethod @params
            return $resp2
          } catch {}
        }
        continue
      }
      throw
    }
  }
}

function Find-SnipeUserByEmail([string]$BaseUrl, [hashtable]$Headers, [string]$Email) {
  if ([string]::IsNullOrWhiteSpace($Email)) { return $null }
  $q = [Uri]::EscapeDataString($Email)
  $resp = Invoke-SnipeRequestWithRetry -BaseUrl $BaseUrl -Headers $Headers -Path "/api/v1/users?search=$q" -Method GET -MaxRetries $SnipeMaxRetries -InitialBackoffSeconds $SnipeInitialBackoffSeconds
  $rows = @(); if ($resp -and $resp.rows) { $rows = @($resp.rows) }
  if ($rows.Count -eq 0) { return $null }
  $exact = $rows | Where-Object { $_.email -and ($_.email.ToString().Trim().ToLower() -eq $Email.Trim().ToLower()) } | Select-Object -First 1
  if ($exact) { return $exact }
  return ($rows | Select-Object -First 1)
}

function Get-SnipeDepartmentId([string]$BaseUrl, [hashtable]$Headers, [string]$DepartmentName, [bool]$CreateIfMissing, [bool]$WhatIf) {
  if ([string]::IsNullOrWhiteSpace($DepartmentName)) { return $null }
  $q = [Uri]::EscapeDataString($DepartmentName)
  $resp = Invoke-SnipeRequestWithRetry -BaseUrl $BaseUrl -Headers $Headers -Path "/api/v1/departments?search=$q" -Method GET -MaxRetries $SnipeMaxRetries -InitialBackoffSeconds $SnipeInitialBackoffSeconds
  $rows = @(); if ($resp -and $resp.rows) { $rows = @($resp.rows) }
  $match = $rows | Where-Object { $_.name -and ($_.name.ToString().Trim().ToLower() -eq $DepartmentName.Trim().ToLower()) } | Select-Object -First 1
  if ($match) { return $match.id }
  if (-not $CreateIfMissing) { return $null }
  if ($WhatIf) { Write-Log "WhatIf: Would create Snipe-IT department '$DepartmentName'"; return $null }
  try {
    $created = Invoke-SnipeRequestWithRetry -BaseUrl $BaseUrl -Headers $Headers -Path '/api/v1/departments' -Method POST -Body @{ name = $DepartmentName } -MaxRetries $SnipeMaxRetries -InitialBackoffSeconds $SnipeInitialBackoffSeconds
    if ($created -and $created.status -eq 'success' -and $created.payload -and $created.payload.id) { return $created.payload.id }
  } catch {}
  return $null
}

function Build-DesiredSnipeUserBody([object]$MsUser, [int]$DepartmentId) {
  $email = $MsUser.mail; if (-not $email) { $email = $MsUser.userPrincipalName }
  $local = $email.Split('@')[0]
  $body = [ordered]@{
    first_name = $MsUser.givenName
    last_name  = $MsUser.surname
    name       = $MsUser.displayName
    username   = $local
    email      = $email
    jobtitle   = $MsUser.jobTitle
    phone      = ($MsUser.businessPhones | Select-Object -First 1)
    mobile     = $MsUser.mobilePhone
  }
  if ($DepartmentId) { $body.department_id = $DepartmentId }
  return $body
}

function Build-NewSnipeUserBody([object]$MsUser, [bool]$Activate, [bool]$SendInvite, [int]$PasswordLength, [int]$DepartmentId) {
  $desired = Build-DesiredSnipeUserBody -MsUser $MsUser -DepartmentId $DepartmentId
  $generatedPassword = New-SecurePassword -Length $PasswordLength
  $desired.activated = [bool]$Activate
  $desired.password = $generatedPassword
  $desired.password_confirmation = $generatedPassword
  if ($SendInvite) { $desired.send_email = $true }
  return $desired
}

function Build-UpdateSnipeUserBody([object]$MsUser, [object]$Existing, [int]$DepartmentId) {
  $desired = Build-DesiredSnipeUserBody -MsUser $MsUser -DepartmentId $DepartmentId

  $delta = [ordered]@{}
  foreach ($k in $desired.Keys) {
    $new = $desired[$k]
    $cur = $null
    if ($k -eq 'department_id') {
      if ($Existing.department -and $Existing.department.id) { $cur = $Existing.department.id } else { $cur = $null }
    } else {
      $cur = Get-ObjectPropertySafe -Object $Existing -Name $k
      # Some Snipe-IT instances may expose snake_case; try alternate mapping for job title
      if ($null -eq $cur -and $k -eq 'jobtitle') { $cur = Get-ObjectPropertySafe -Object $Existing -Name 'job_title' }
    }
    # Only send non-null/non-empty values to avoid wiping existing data with blanks
    if ($null -ne $new -and "$new" -ne '') {
      if (("$cur") -ne ("$new")) { $delta[$k] = $new }
    }
  }
  return $delta
}

# Main
try {
  Write-Log '=== Snipe-IT User Sync start ==='
  if ($WhatIf) { Write-Log 'WhatIf mode enabled: no changes will be written to Snipe-IT.' -Type WARNING }

  # Allow Automation Variables to override the two booleans if params omitted
  if (-not $PSBoundParameters.ContainsKey('EnableLoginForNewUsers')) {
    $EnableLoginForNewUsers = Get-BooleanFromVar -Name 'SnipeItUserSyncEnableLogin' -Default:$EnableLoginForNewUsers
  }
  if (-not $PSBoundParameters.ContainsKey('SendInviteForNewUsers')) {
    $SendInviteForNewUsers = Get-BooleanFromVar -Name 'SnipeItUserSyncSendInvite' -Default:$SendInviteForNewUsers
  }

  $cfg = Get-SnipeConfig
  $graphToken = Get-MsGraphToken
  $msUsers = Get-TenantUsers -Token $graphToken -OnlyEnabled:$OnlyEnabledUsers -IncludeGuests:$IncludeGuests
  Write-Log ("Loaded {0} Microsoft 365 users after filters" -f $msUsers.Count)

  Write-Log "Syncing to Snipe-IT at $($cfg.BaseUrl)"
  $headers = @{ Authorization = "Bearer $($cfg.ApiToken)"; Accept = 'application/json' }

  $stats = [ordered]@{ Created=0; Updated=0; Skipped=0; Errors=0 }

  foreach ($u in $msUsers) {
      try {
        $email = if ($u.mail) { $u.mail } else { $u.userPrincipalName }
        if (-not $email -or ($email -notlike '*@*')) { $stats.Skipped++; continue }

        $deptId = $null
        if ($u.department) { $deptId = Get-SnipeDepartmentId -BaseUrl $cfg.BaseUrl -Headers $headers -DepartmentName $u.department -CreateIfMissing:$CreateMissingDepartments -WhatIf:$WhatIf }

    $existing = Find-SnipeUserByEmail -BaseUrl $cfg.BaseUrl -Headers $headers -Email $email
        if ($existing) {
          $body = Build-UpdateSnipeUserBody -MsUser $u -Existing $existing -DepartmentId $deptId
          if ($body.Keys.Count -gt 0) {
            if ($WhatIf) {
              $fields = ($body.Keys -join ', ')
              Write-Log "WhatIf: Would update Snipe-IT user id=$($existing.id) email=$email fields=[$fields]"
              $stats.Updated++
            } else {
      [void](Invoke-SnipeRequestWithRetry -BaseUrl $cfg.BaseUrl -Headers $headers -Path "/api/v1/users/$($existing.id)" -Method PATCH -Body $body -MaxRetries $SnipeMaxRetries -InitialBackoffSeconds $SnipeInitialBackoffSeconds)
              $stats.Updated++
            }
          } else { $stats.Skipped++ }
        } else {
          if ($WhatIf) {
            $local = ($email.Split('@')[0])
            $deptMsg = if ($u.department) { " department='$($u.department)'." } else { '.' }
            Write-Log "WhatIf: Would create Snipe-IT user email=$email username=$local activated=$EnableLoginForNewUsers sendInvite=$SendInviteForNewUsers$deptMsg"
            $stats.Created++
          } else {
            $newBody = Build-NewSnipeUserBody -MsUser $u -Activate:$EnableLoginForNewUsers -SendInvite:$SendInviteForNewUsers -PasswordLength 20 -DepartmentId $deptId
            # Never log password fields
    [void](Invoke-SnipeRequestWithRetry -BaseUrl $cfg.BaseUrl -Headers $headers -Path '/api/v1/users' -Method POST -Body $newBody -MaxRetries $SnipeMaxRetries -InitialBackoffSeconds $SnipeInitialBackoffSeconds)
            $stats.Created++
          }
        }
      } catch {
        $stats.Errors++
        Write-Log "Error syncing user '$($u.displayName)' ($($u.userPrincipalName)): $_" -Type ERROR
      }
  if ($SnipeRequestDelayMs -gt 0) { Start-Sleep -Milliseconds $SnipeRequestDelayMs }
  }

  Write-Log '=== Snipe-IT User Sync complete ==='
  $summary = [pscustomobject]$stats
  Write-Output $summary
  return $summary
}
catch {
  Write-Log "Runbook failed: $_" -Type ERROR
  throw
}
