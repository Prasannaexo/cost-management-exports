<#
Provisions an Azure Automation Account that runs Get-AzureAccessReview.ps1
on a monthly schedule, uploading the resulting workbook to the same
storage account the Cost Management exports use.

Deliberately a separate, dedicated Automation Account rather than reusing
the AI-guardrails one in c:\my_work -- this one's managed identity only
ever needs read access (RBAC + group membership) across three
subscriptions plus write access to one storage container, and keeping it
isolated means a compromise or misconfiguration here can't touch the
guardrails automation (budget actions, deployment scaling) or vice versa.

What this script automates:
  1. Resource group + Automation Account, with a system-assigned managed
     identity.
  2. Reader role for that identity on the three source subscriptions.
  3. Storage Blob Data Contributor for that identity on the destination
     storage account (so the runbook can upload the workbook without
     needing shared-key access, consistent with that account's firewall
     setup in New-CostExportStorage.ps1).
  4. Imports Get-AzureAccessReview.ps1 as a runbook and publishes it.
  5. Creates a monthly schedule (1st of each month, 06:00 UTC) and links it
     to the runbook with the right parameters.

What this script does NOT automate (both require tenant-level admin
consent this script shouldn't silently assume it has -- same pattern as
the manual follow-ups at the end of c:\my_work\configure-ai-guardrails.ps1):
  - Importing the PowerShell modules the runbook depends on
    (Az.Accounts, Az.Resources, Az.Storage, ImportExcel) into the
    Automation Account. Do this from the portal: Automation Account >
    Modules > Browse gallery > search and import each one. Az.* modules
    may already be preloaded depending on the account's runtime version;
    ImportExcel will not be.
  - Granting the managed identity's service principal the Microsoft Graph
    application permissions it needs to expand group membership
    (GroupMember.Read.All, User.Read.All). This needs a Global
    Administrator or Privileged Role Administrator. See the printed
    instructions at the end of this script for the exact commands to hand
    to whoever holds that role.

Usage:
  .\New-AccessReviewAutomation.ps1 `
      -ProductionSubscriptionId 865fa361-a42e-4d3f-8566-cc34114cb8be `
      -TestSubscriptionId c0ae4ee0-ebef-4f37-a8e8-d590fc7417ba `
      -DevSubscriptionId 86a78b30-a351-425e-9692-6a3938a559cf `
      -StorageAccountResourceId "/subscriptions/865fa361-.../storageAccounts/stldmcostexports"
#>
param(
  [Parameter(Mandatory)][string]$ProductionSubscriptionId,
  [Parameter(Mandatory)][string]$TestSubscriptionId,
  [Parameter(Mandatory)][string]$DevSubscriptionId,
  [Parameter(Mandatory)][string]$StorageAccountResourceId,
  [string]$ResourceGroupName = "rg-access-review",
  [string]$AutomationAccountName = "aa-access-review",
  [string]$Location = "eastus2",
  [string]$RunbookName = "Get-AzureAccessReview"
)

$ErrorActionPreference = "Stop"
$hostSubscriptionId = $ProductionSubscriptionId  # Automation Account lives alongside the storage account, in Production

Write-Host "Setting context to $hostSubscriptionId..." -ForegroundColor Cyan
Set-AzContext -SubscriptionId $hostSubscriptionId | Out-Null

$rg = Get-AzResourceGroup -Name $ResourceGroupName -ErrorAction SilentlyContinue
if (-not $rg) {
  Write-Host "Creating resource group $ResourceGroupName..." -ForegroundColor Cyan
  $rg = New-AzResourceGroup -Name $ResourceGroupName -Location $Location
}

$aa = Get-AzAutomationAccount -ResourceGroupName $ResourceGroupName -Name $AutomationAccountName -ErrorAction SilentlyContinue
if (-not $aa) {
  Write-Host "Creating Automation Account $AutomationAccountName..." -ForegroundColor Cyan
  $aa = New-AzAutomationAccount -ResourceGroupName $ResourceGroupName -Name $AutomationAccountName -Location $Location -AssignSystemIdentity
} else {
  Write-Host "Automation Account $AutomationAccountName already exists -- reusing." -ForegroundColor Yellow
}

$principalId = $aa.Identity.PrincipalId
Write-Host "Managed identity principal ID: $principalId" -ForegroundColor Cyan

Write-Host "Granting Reader on the three source subscriptions..." -ForegroundColor Cyan
foreach ($subId in @($ProductionSubscriptionId, $TestSubscriptionId, $DevSubscriptionId)) {
  $scope = "/subscriptions/$subId"
  $existing = Get-AzRoleAssignment -ObjectId $principalId -Scope $scope -RoleDefinitionName "Reader" -ErrorAction SilentlyContinue
  if (-not $existing) {
    New-AzRoleAssignment -ObjectId $principalId -Scope $scope -RoleDefinitionName "Reader" | Out-Null
    Write-Host "  Granted Reader on $subId"
  } else {
    Write-Host "  Reader already present on $subId"
  }
}

Write-Host "Granting Storage Blob Data Contributor on the destination storage account..." -ForegroundColor Cyan
$existingStorageRole = Get-AzRoleAssignment -ObjectId $principalId -Scope $StorageAccountResourceId `
  -RoleDefinitionName "Storage Blob Data Contributor" -ErrorAction SilentlyContinue
if (-not $existingStorageRole) {
  New-AzRoleAssignment -ObjectId $principalId -Scope $StorageAccountResourceId `
    -RoleDefinitionName "Storage Blob Data Contributor" | Out-Null
}

Write-Host "Importing runbook $RunbookName..." -ForegroundColor Cyan
$scriptPath = Join-Path $PSScriptRoot "Get-AzureAccessReview.ps1"
Import-AzAutomationRunbook -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName `
  -Name $RunbookName -Path $scriptPath -Type PowerShell72 -Force | Out-Null
Publish-AzAutomationRunbook -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName `
  -Name $RunbookName | Out-Null

Write-Host "Creating monthly schedule..." -ForegroundColor Cyan
$scheduleName = "monthly-1st-06utc"
$nextMonthStart = (Get-Date -Day 1).AddMonths(1).Date.AddHours(6)  # 1st of next month, 06:00 UTC
$schedule = Get-AzAutomationSchedule -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName `
  -Name $scheduleName -ErrorAction SilentlyContinue
if (-not $schedule) {
  $schedule = New-AzAutomationSchedule -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName `
    -Name $scheduleName -StartTime $nextMonthStart -MonthInterval 1 -TimeZone "UTC"
}

Write-Host "Linking schedule to runbook with parameters..." -ForegroundColor Cyan
$params = @{
  ProductionSubscriptionId = $ProductionSubscriptionId
  TestSubscriptionId       = $TestSubscriptionId
  DevSubscriptionId        = $DevSubscriptionId
  StorageAccountResourceId = $StorageAccountResourceId
  UseManagedIdentity       = $true
}
Register-AzAutomationScheduledRunbook -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName `
  -RunbookName $RunbookName -ScheduleName $scheduleName -Parameters $params -ErrorAction SilentlyContinue | Out-Null

Write-Host "`nDone. Automation Account: $AutomationAccountName in $ResourceGroupName." -ForegroundColor Green
Write-Host "First scheduled run: $nextMonthStart UTC. Run it manually now via Start-AzAutomationRunbook to test before then:" -ForegroundColor Green
Write-Host "  Start-AzAutomationRunbook -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName -Name $RunbookName -Parameters `$paramsHashtable"

Write-Host "`n=== MANUAL FOLLOW-UP REQUIRED (needs Global Administrator / Privileged Role Administrator) ===" -ForegroundColor Yellow
Write-Host @"
The runbook cannot expand group membership until the managed identity
($principalId) is granted these Microsoft Graph application permissions.
Run this as a Global Admin (PowerShell, Microsoft Graph SDK):

  Connect-MgGraph -Scopes "AppRoleAssignment.ReadWrite.All"
  `$graphSpId = (Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'").Id
  `$miSpId = "$principalId"
  foreach (`$perm in "GroupMember.Read.All", "User.Read.All", "PrivilegedAccess.Read.AzureADGroup") {
    `$appRole = (Get-MgServicePrincipal -ServicePrincipalId `$graphSpId).AppRoles |
      Where-Object { `$_.Value -eq `$perm }
    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId `$miSpId -PrincipalId `$miSpId ``
      -ResourceId `$graphSpId -AppRoleId `$appRole.Id
  }

  # PrivilegedAccess.Read.AzureADGroup is what lets the report see PIM
  # "eligible" (just-in-time) membership in admin groups like
  # BDECompute-Platform-SuperAdmins-Prod -- without it those groups show
  # zero members even though people are eligible to activate access into
  # them. The other two are for ordinary group/user lookups. The report
  # runs fine without PrivilegedAccess.Read.AzureADGroup, it just won't
  # show "Eligible" rows for PIM-managed groups (its "READ ME -
  # Limitations" sheet says so explicitly when this is missing).

Also import these modules into the Automation Account from the portal
(Automation Account > Modules > Browse gallery) before the first run:
  Az.Accounts, Az.Resources, Az.Storage, ImportExcel
"@ -ForegroundColor Yellow
