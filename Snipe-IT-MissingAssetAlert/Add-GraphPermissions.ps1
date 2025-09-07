# Requires -Modules "Microsoft.Graph.Applications"
<#
.SYNOPSIS
  Assign Microsoft Graph App Roles required by Snipe-IT Missing Asset Alert runbook to an Automation Managed Identity.

.DESCRIPTION
  Grants DeviceManagementManagedDevices.Read.All to the specified System-Assigned Managed Identity so the runbook
  can read Intune managed devices via Microsoft Graph using Managed Identity authentication.

.PARAMETER AutomationMSI_ID
  Object ID of the Automation Account's System-Assigned Managed Identity.

.NOTES
  Requires the operator to have: AppRoleAssignment.ReadWrite.All and Application.Read.All.
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)][string]$AutomationMSI_ID = "<REPLACE_WITH_YOUR_AUTOMATION_ACCOUNT_MSI_OBJECT_ID>"
)

$GRAPH_APP_ID = "00000003-0000-0000-c000-000000000000"

Write-Host "Connecting to Microsoft Graph…"
Connect-MgGraph -Scopes "AppRoleAssignment.ReadWrite.All","Application.Read.All" -NoWelcome

Write-Host "Locating Microsoft Graph service principal…"
$graphSp = Get-MgServicePrincipal -Filter "appId eq '$GRAPH_APP_ID'"
if (-not $graphSp) { throw "Microsoft Graph service principal not found." }

# Required roles for this runbook
$roles = @(
  @{ Name = 'DeviceManagementManagedDevices.Read.All'; Id = '2f51be20-0bb4-4fed-bf7b-db946066c75e' }
  @{ Name = 'Mail.Send'; Id = 'b633e1c5-b582-4048-a93e-9f11b44c7e96' } # required if using EmailTo/EmailSender
)

foreach ($r in $roles) {
  $exists = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $AutomationMSI_ID | Where-Object { $_.AppRoleId -eq $r.Id }
  if (-not $exists) {
    Write-Host "Granting $($r.Name)…"
    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $AutomationMSI_ID -PrincipalId $AutomationMSI_ID -ResourceId $graphSp.Id -AppRoleId $r.Id | Out-Null
    Write-Host "Granted $($r.Name)."
  } else {
    Write-Host "$($r.Name) already assigned."
  }
}

Write-Host "Done."
