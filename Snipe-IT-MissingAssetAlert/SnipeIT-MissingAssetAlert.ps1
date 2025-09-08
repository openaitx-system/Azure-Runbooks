# Requires -Modules "Az.Accounts"
<#
.SYNOPSIS
  Azure Automation runbook: find Intune managed devices whose serial numbers are NOT present in Snipe-IT.

.DESCRIPTION
  Authenticates to Microsoft Graph using the Automation Account's System-Assigned Managed Identity,
  retrieves all Intune managed devices (with serial numbers), and checks each serial against your
  Snipe-IT instance via the Snipe-IT API. Outputs a list of missing assets and a compact summary.

.PARAMETER SnipeItBaseUrl
  Base URL of your Snipe-IT instance, e.g. https://snipeit.contoso.com

.PARAMETER SnipeItTokenVar
  Name of the Azure Automation Variable holding the Snipe-IT API token (Encrypted) [default: SnipeItApiToken]

.PARAMETER MaxRetries
  Max retries for Graph API requests when throttled/5xx. [default: 5]

.PARAMETER InitialBackoffSeconds
  Initial backoff for Graph API retries. Doubles on each retry. [default: 5]

.PARAMETER Limit
  Optional cap on number of Intune devices to examine (for testing). Omit or set 0 for all.

.NOTES
  Author: Ryan Schultz
  Created: 2025-09-07
  Permissions needed on Managed Identity: DeviceManagementManagedDevices.Read.All, Mail.Send (if using EmailTo/EmailSender)
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$SnipeItBaseUrl,
  [string]$SnipeItTokenVar = "SnipeItApiToken",
  [int]$MaxRetries = 5,
  [int]$InitialBackoffSeconds = 5,
  [int]$Limit = 0,
  [string]$EmailTo = "",
  [string]$EmailSender = ""
)

function Write-Log {
  param(
    [string]$Message,
    [ValidateSet("INFO","WARNING","ERROR","WHATIF")]
    [string]$Type = "INFO"
  )
  $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  $line = "[$ts] [$Type] $Message"
  switch ($Type) {
    "ERROR"   { Write-Error $Message; Write-Verbose $line -Verbose }
    "WARNING" { Write-Warning $Message; Write-Verbose $line -Verbose }
    "WHATIF"  { Write-Verbose "[WHATIF] $Message" -Verbose }
    default   { Write-Verbose $line -Verbose }
  }
}

function Get-MsGraphToken {
  try {
    Write-Log "Authenticating with Managed Identity…"
    Connect-AzAccount -Identity | Out-Null
    $tok = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com"
    $token = $tok.Token
    if ($token -is [System.Security.SecureString]) {
      $token = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [Runtime.InteropServices.Marshal]::SecureStringToBSTR($token)
      )
    }
    if ([string]::IsNullOrWhiteSpace($token)) { throw "Empty Graph token." }
    Write-Log "Graph token acquired."
    return $token
  } catch { Write-Error "Failed to acquire Graph token via Managed Identity: $_"; throw }
}

function Invoke-MsGraphRequestWithRetry {
  param(
    [Parameter(Mandatory)][string]$Token,
    [Parameter(Mandatory)][string]$Uri,
    [ValidateSet("GET","POST","PATCH","PUT","DELETE")]
    [string]$Method = "GET",
    [object]$Body = $null,
    [string]$ContentType = "application/json",
    [int]$MaxRetries = 5,
    [int]$InitialBackoffSeconds = 5
  )

  $retry = 0
  $backoff = $InitialBackoffSeconds
  $params = @{ Uri = $Uri; Headers = @{ Authorization = "Bearer $Token" }; Method = $Method; ContentType = $ContentType }
  if ($null -ne $Body -and $Method -ne "GET") {
    $params.Body = ($Body -is [string]) ? $Body : ($Body | ConvertTo-Json -Depth 10)
  }
  while ($true) {
    try { return Invoke-RestMethod @params }
    catch {
      $status = $null; if ($_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode }
      if (($status -eq 429 -or ($status -ge 500 -and $status -lt 600)) -and $retry -lt $MaxRetries) {
        $retryAfter = $backoff
        try {
          if ($_.Exception.Response -and $_.Exception.Response.Headers) {
            $hdr = $_.Exception.Response.Headers | Where-Object { $_.Key -eq "Retry-After" }
            if ($hdr) { $retryAfter = [int]$hdr.Value[0] }
          }
        } catch {}
        $msg = ($status -eq 429) ? "Graph throttled (429)" : "Graph 5xx"
        Write-Log "$msg. Sleeping $retryAfter sec. Attempt $($retry+1)/$MaxRetries" -Type "WARNING"
        Start-Sleep -Seconds $retryAfter
        $retry++; $backoff = [Math]::Max($backoff * 2, 1)
      } else { Write-Log "Graph request failed ($status): $_" -Type "ERROR"; throw }
    }
  }
}

function Get-IntuneManagedDevices {
  param(
    [Parameter(Mandatory)][string]$Token,
    [int]$MaxRetries = 5,
    [int]$InitialBackoffSeconds = 5,
    [int]$Limit = 0
  )
  try {
    Write-Log "Retrieving Intune managed devices…"
  $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$select=id,serialNumber,deviceName,operatingSystem,lastSyncDateTime,managedDeviceOwnerType,userDisplayName,userPrincipalName&`$orderby=deviceName"
    $col = @()
    $resp = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $uri -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
    $col += $resp.value
    while ($resp.'@odata.nextLink') {
      if ($Limit -gt 0 -and $col.Count -ge $Limit) { break }
      $resp = Invoke-MsGraphRequestWithRetry -Token $Token -Uri $resp.'@odata.nextLink' -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds
      $col += $resp.value
    }
  if ($Limit -gt 0 -and $col.Count -gt $Limit) { $col = $col | Select-Object -First $Limit }
  $withSerial = $col | Where-Object { $_.serialNumber -and $_.serialNumber.Trim().Length -gt 0 }
    Write-Log "Total devices: $($col.Count); with serial: $($withSerial.Count)"
  return $withSerial
  } catch { Write-Log "Failed to retrieve Intune devices: $_" -Type "ERROR"; throw }
}

function Get-SnipeItHeaders {
  param(
    [Parameter(Mandatory)][string]$TokenVar
  )
  try { $apiToken = Get-AutomationVariable -Name $TokenVar } catch { throw "Failed to read Automation Variable '$TokenVar'. $_" }
  if ([string]::IsNullOrWhiteSpace($apiToken)) { throw "Missing Snipe-IT API token in variable '$TokenVar'" }
  return @{ "Authorization" = "Bearer $apiToken"; "Accept" = "application/json" }
}

function Test-SnipeItAssetBySerial {
  param(
    [Parameter(Mandatory)][string]$BaseUrl,
    [Parameter(Mandatory)][hashtable]$Headers,
    [Parameter(Mandatory)][string]$Serial
  )
  # Try byserial endpoint first
  $root = ($BaseUrl.TrimEnd('/')) + "/api/v1"
  $encoded = [Uri]::EscapeDataString($Serial)
  $bySerial = "$root/hardware/byserial/$encoded"
  try {
    $resp = Invoke-RestMethod -Uri $bySerial -Headers $Headers -Method GET -TimeoutSec 60
    if ($resp) { return $true }
  } catch {
    $code = $null; if ($_.Exception.Response) { $code = [int]$_.Exception.Response.StatusCode }
    if ($code -eq 404) { return $false }
    # Fallback to search if endpoint unsupported or other error
    try {
      $searchUrl = "$root/hardware?search=$encoded"
      $s = Invoke-RestMethod -Uri $searchUrl -Headers $Headers -Method GET -TimeoutSec 60
      $items = $null
      if ($s -and $s.PSObject.Properties.Name -contains 'rows')      { $items = @($s.rows) }
      elseif ($s.PSObject.Properties.Name -contains 'data')          { $items = @($s.data) }
      elseif ($s -is [System.Collections.IEnumerable])               { $items = @($s) }
      if ($items) {
        foreach ($it in $items) {
          $ser = $it.serial
          if (-not $ser) { $ser = $it.serial_number }
          if (-not $ser) { $ser = $it.asset_tag }
          if ($ser -and ($ser.ToString().Trim().ToUpper() -eq $Serial.Trim().ToUpper())) { return $true }
        }
      }
      return $false
    } catch { return $false }
  }
}

function Get-SnipeItSerialHashSet {
  param(
    [Parameter(Mandatory)][string]$BaseUrl,
    [Parameter(Mandatory)][hashtable]$Headers,
    [int]$PageSize = 100
  )
  $root = ($BaseUrl.TrimEnd('/')) + "/api/v1"
  $offset = 0
  $serials = New-Object 'System.Collections.Generic.HashSet[string]'
  $total = $null
  do {
    $url = "$root/hardware?limit=$PageSize&offset=$offset"
    $resp = $null
    try {
      $resp = Invoke-RestMethod -Uri $url -Headers $Headers -Method GET -TimeoutSec 120
    } catch {
      $code = $null; if ($_.Exception.Response) { $code = [int]$_.Exception.Response.StatusCode }
      throw "Snipe-IT hardware list failed (HTTP $code). Verify API base URL and token. $_"
    }
    if ($null -eq $resp) { break }
    if ($null -eq $total) {
      if ($resp.PSObject.Properties.Name -contains 'total') { $total = [int]$resp.total } else { $total = 0 }
    }
    $items = $null
    if ($resp.PSObject.Properties.Name -contains 'rows') { $items = @($resp.rows) }
    elseif ($resp.PSObject.Properties.Name -contains 'data') { $items = @($resp.data) }
    elseif ($resp -is [System.Collections.IEnumerable]) { $items = @($resp) }
    if ($items) {
      foreach ($it in $items) {
        $ser = $it.serial
        if (-not $ser) { $ser = $it.serial_number }
        if ($ser -and $ser.ToString().Trim().Length -gt 0) {
          [void]$serials.Add($ser.ToString().Trim().ToUpperInvariant())
        }
      }
    }
    $offset += $PageSize
    if ($total -eq 0 -and $items) { $total = $items.Count }
  } while ($items -and ($offset -lt $total))

  Write-Log "Loaded $($serials.Count) serials from Snipe-IT hardware."
  return $serials
}

function Build-MissingHtml {
  param(
    [Parameter(Mandatory)][System.Collections.IEnumerable]$Missing,
    [Parameter(Mandatory)][object]$Summary
  )
  $rows = ""
  foreach ($m in ($Missing | Sort-Object deviceName)) {
    $name = [System.Web.HttpUtility]::HtmlEncode($m.deviceName)
    $os   = [System.Web.HttpUtility]::HtmlEncode($m.operatingSystem)
    $sn   = [System.Web.HttpUtility]::HtmlEncode($m.serialNumber)
    $ls   = [System.Web.HttpUtility]::HtmlEncode([string]$m.lastSyncDateTime)
    $own  = [System.Web.HttpUtility]::HtmlEncode($m.ownership)
    $usr  = $null
    if ($m.assignedUserDisplayName -or $m.assignedUserPrincipalName) {
      $disp = [System.Web.HttpUtility]::HtmlEncode([string]$m.assignedUserDisplayName)
      $upn  = [System.Web.HttpUtility]::HtmlEncode([string]$m.assignedUserPrincipalName)
      if ([string]::IsNullOrWhiteSpace($disp)) { $usr = $upn }
      elseif ([string]::IsNullOrWhiteSpace($upn)) { $usr = $disp }
      else { $usr = "$disp ($upn)" }
    } else { $usr = "" }
    $rows += "<tr>"+
             "<td style='border:1px solid #e5e5e5;padding:8px 10px;text-align:left;'>$name</td>"+
             "<td style='border:1px solid #e5e5e5;padding:8px 10px;text-align:left;'>$os</td>"+
             "<td style='border:1px solid #e5e5e5;padding:8px 10px;text-align:left;'><span style=""font-family:Consolas,'Courier New',monospace;background:#f2f2f2;padding:1px 4px;border-radius:3px;"">$sn</span></td>"+
             "<td style='border:1px solid #e5e5e5;padding:8px 10px;text-align:left;'>$ls</td>"+
             "<td style='border:1px solid #e5e5e5;padding:8px 10px;text-align:left;'>$own</td>"+
             "<td style='border:1px solid #e5e5e5;padding:8px 10px;text-align:left;'>$usr</td>"+
             "</tr>"
  }
  $title = "Intune devices not found in Snipe-IT ($($Summary.MissingCount))"
  $meta  = "Generated: $([DateTime]::UtcNow.ToString('u')) | Checked Serials: $($Summary.UniqueSerialsChecked) | Intune Devices: $($Summary.TotalIntuneDevices)"
  return @"
<!DOCTYPE html>
<html>
  <head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'></head>
  <body style="font-family:Segoe UI,Arial,sans-serif;color:#222;margin:0;padding:0;">
    <div style="max-width:900px;margin:0 auto;border:1px solid #e0e0e0;border-radius:6px;overflow:hidden;">
      <div style="background:#0078d4;color:#ffffff;padding:14px 18px;font-size:16px;">$title</div>
      <div style="padding:16px;">
        <p style="margin:0 0 12px 0;">The following Intune devices were not found in Snipe-IT by serial number. Add corporate owned devices to Snipe-IT.</p>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%;border:1px solid #e5e5e5;mso-table-lspace:0pt;mso-table-rspace:0pt;">
          <thead>
            <tr>
              <th align="left" style="background:#f5f5f5;border:1px solid #e5e5e5;padding:8px 10px;">Device Name</th>
              <th align="left" style="background:#f5f5f5;border:1px solid #e5e5e5;padding:8px 10px;">OS</th>
              <th align="left" style="background:#f5f5f5;border:1px solid #e5e5e5;padding:8px 10px;">Serial</th>
              <th align="left" style="background:#f5f5f5;border:1px solid #e5e5e5;padding:8px 10px;">Last Sync</th>
              <th align="left" style="background:#f5f5f5;border:1px solid #e5e5e5;padding:8px 10px;">Ownership</th>
              <th align="left" style="background:#f5f5f5;border:1px solid #e5e5e5;padding:8px 10px;">Assigned User</th>
            </tr>
          </thead>
          <tbody>
            $rows
          </tbody>
        </table>
        <p style="font-size:12px;color:#555;margin:12px 0 0 0;">$meta</p>
      </div>
    </div>
  </body>
</html>
"@
}

function Send-GraphMailMessage {
  param (
    [Parameter(Mandatory)][string]$Token,
    [Parameter(Mandatory)][string]$From,
    [Parameter(Mandatory)][string]$To,
    [Parameter(Mandatory)][string]$Subject,
    [Parameter(Mandatory)][string]$HtmlBody,
    [int]$MaxRetries = 5,
    [int]$InitialBackoffSeconds = 5
  )
  try {
    $uri = "https://graph.microsoft.com/v1.0/users/$From/sendMail"
    $body = @{
      message = @{
        subject = $Subject
        body = @{ contentType = "HTML"; content = $HtmlBody }
        toRecipients = @(@{ emailAddress = @{ address = $To } })
      }
      saveToSentItems = $true
    }
    Invoke-MsGraphRequestWithRetry -Token $Token -Uri $uri -Method POST -Body $body -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds | Out-Null
    Write-Log "Email sent to $To successfully"
    return $true
  } catch { Write-Log ("Failed to send email to {0}: {1}" -f $To, $_) -Type "ERROR"; return $false }
}

try {
  Write-Log "=== Snipe-IT Missing Asset Alert run starting ==="
  $graphToken = Get-MsGraphToken
  $headers = Get-SnipeItHeaders -TokenVar $SnipeItTokenVar

  $snipeSerials = Get-SnipeItSerialHashSet -BaseUrl $SnipeItBaseUrl -Headers $headers -PageSize 200

  $devices = Get-IntuneManagedDevices -Token $graphToken -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds -Limit $Limit
  # Deduplicate by serial (some devices may share or have duplicates across platforms)
  $serialSeen = New-Object 'System.Collections.Generic.HashSet[string]'
  $unique = @()
  foreach ($d in $devices) {
    $sn = $d.serialNumber.Trim()
    if (-not [string]::IsNullOrWhiteSpace($sn)) {
      if ($serialSeen.Add($sn)) { $unique += $d }
    }
  }
  Write-Log "Unique serials to check: $($unique.Count)"

  $missing = New-Object System.Collections.ArrayList
  $checked = 0
  foreach ($d in $unique) {
    $sn = $d.serialNumber
    $norm = $sn.Trim().ToUpperInvariant()
    $exists = $snipeSerials.Contains($norm)
    if (-not $exists) {
      $owner = switch -Regex ($d.managedDeviceOwnerType) {
        '^company$' { 'Corporate' }
        '^personal$' { 'Personal' }
        default { 'Unknown' }
      }
      [void]$missing.Add([PSCustomObject]@{
        serialNumber      = $sn
        deviceName        = $d.deviceName
        operatingSystem   = $d.operatingSystem
        lastSyncDateTime  = $d.lastSyncDateTime
        ownership         = $owner
        assignedUserDisplayName    = $d.userDisplayName
        assignedUserPrincipalName  = $d.userPrincipalName
      })
    }
    $checked++
    if (($checked % 100) -eq 0) { Write-Log "Checked $checked / $($unique.Count) serials…" }
  }

  $summary = [PSCustomObject][ordered]@{
    TotalIntuneDevices   = $devices.Count
    UniqueSerialsChecked = $unique.Count
    MissingCount         = $missing.Count
    TimestampUtc         = (Get-Date).ToUniversalTime().ToString('s') + 'Z'
  }

  Write-Output "Summary: $(($summary | ConvertTo-Json -Compress))"
  if ($missing.Count -gt 0) {
    Write-Output "Devices missing in Snipe-IT: $($missing.Count)"
    $missing | Sort-Object deviceName | ForEach-Object {
      $own = $_.ownership
      $usr = if ($_.assignedUserDisplayName -or $_.assignedUserPrincipalName) {
        if ($_.assignedUserDisplayName -and $_.assignedUserPrincipalName) {
          "{0} ({1})" -f $_.assignedUserDisplayName, $_.assignedUserPrincipalName
        } elseif ($_.assignedUserDisplayName) { $_.assignedUserDisplayName }
        else { $_.assignedUserPrincipalName }
      } else { '' }
      Write-Output ("Missing: {0} [{1}] SN={2} Owner={3}{4}" -f $_.deviceName, $_.operatingSystem, $_.serialNumber, $own, $(if ($usr) { " User=$usr" } else { '' }))
    }
    # Optional email alert
    if (-not [string]::IsNullOrWhiteSpace($EmailTo)) {
      if ([string]::IsNullOrWhiteSpace($EmailSender)) {
        Write-Log "EmailTo specified but EmailSender is empty. Skipping email." -Type "WARNING"
      } else {
        $subject = "Alert: $($missing.Count) Intune devices not found in Snipe-IT"
        $html = Build-MissingHtml -Missing $missing -Summary $summary
        [void](Send-GraphMailMessage -Token $graphToken -From $EmailSender -To $EmailTo -Subject $subject -HtmlBody $html -MaxRetries $MaxRetries -InitialBackoffSeconds $InitialBackoffSeconds)
      }
    }
  } else {
    Write-Output "All checked serials exist in Snipe-IT."
  }

  # Return structured object for output
  [PSCustomObject]@{ Summary = $summary; Missing = $missing }
}
catch {
  Write-Log "Run failed: $_" -Type "ERROR"
  throw
}
finally { Write-Log "=== Run complete ===" }
