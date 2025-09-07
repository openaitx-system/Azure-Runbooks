# Requires -Modules "Microsoft.Graph.Applications"
<#
.SYNOPSIS
  Assign required Microsoft Graph application permissions to an Automation Account's Managed Identity for Snipe-IT user sync.

.DESCRIPTION
  Grants the Managed Identity the minimal Graph application permission needed to read users:
   - User.Read.All

  Run locally with sufficient privileges. You must be a Directory Administrator or have the ability to assign app roles.

.PARAMETER AutomationMSI_ID
  The Object ID of the Automation Account's System-Assigned Managed Identity.
#>

[CmdletBinding()] param(
  [Parameter(Mandatory = $true)] [string] $AutomationMSI_ID = '<REPLACE_WITH_YOUR_AUTOMATION_ACCOUNT_MSI_OBJECT_ID>'
)

$GRAPH_APP_ID = '00000003-0000-0000-c000-000000000000'

Write-Host 'Connecting to Microsoft Graph...' -ForegroundColor Cyan
try {
  Connect-MgGraph -Scopes 'AppRoleAssignment.ReadWrite.All','Application.Read.All' -NoWelcome
} catch {
  Write-Host "Failed to connect to Graph: $_" -ForegroundColor Red; exit 1
}

try {
  $graphSp = Get-MgServicePrincipal -Filter "appId eq '$GRAPH_APP_ID'"
  if (-not $graphSp) { throw 'Microsoft Graph service principal not found.' }
} catch {
  Write-Host "Error locating Graph service principal: $_" -ForegroundColor Red; exit 1
}

$required = @(
  @{ Name = 'User.Read.All'; Id = 'df021288-bdef-4463-88db-98f22de89214' }
)

Write-Host "Assigning permissions to Managed Identity: $AutomationMSI_ID" -ForegroundColor Cyan
foreach ($perm in $required) {
  $existing = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $AutomationMSI_ID |
              Where-Object { $_.AppRoleId -eq $perm.Id }
  if (-not $existing) {
    try {
      New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $AutomationMSI_ID `
        -PrincipalId $AutomationMSI_ID -ResourceId $graphSp.Id -AppRoleId $perm.Id | Out-Null
      Write-Host "Assigned $($perm.Name)" -ForegroundColor Green
    } catch {
      Write-Host "Failed to assign $($perm.Name): $_" -ForegroundColor Red
    }
  } else {
    Write-Host "$($perm.Name) already assigned" -ForegroundColor Yellow
  }
}

Write-Host 'Done.' -ForegroundColor Green
