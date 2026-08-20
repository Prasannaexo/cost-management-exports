# Azure Access Review (monthly)

Automates a recurring Azure RBAC access review workbook, delivered the same
way as the Cost Management exports: a scheduled Azure-native job writing
into the same firewalled storage account (`stldmcostexports`, container
`access-reviews`), so both feeds live in one place.

There was no prior script for this -- it recreates a manually-built
workbook (`260804 AR PK azure access review.xlsx`) that had sheets for
Production, Test, Dev, an "All Subscriptions" rollup, and a "Role
Definitions" reference. See the methodology notes at the top of
[Get-AzureAccessReview.ps1](Get-AzureAccessReview.ps1) for exactly how role
and group membership are attributed to each user.

**Group membership note, confirmed against this tenant:** groups like
`BDECompute-Platform-SuperAdmins-Prod` initially appeared to have zero
members at all, which looked like a permissions bug. It isn't one --
confirmed via a direct Graph query that the group is NOT a
role-assignable group (`isAssignableToRole: false`). The real cause is
PIM: these admin groups use "eligible" (just-in-time) membership rather
than standing membership, so nobody is a permanent member and a plain
membership query genuinely finds nobody active most of the time. The
script now separately queries PIM eligibility schedules and marks those
users **"Eligible"** in the report (vs **"Active"** for real, current
access) -- see the methodology notes in
[Get-AzureAccessReview.ps1](Get-AzureAccessReview.ps1). This needs one
additional Graph permission the reporting identity may not have yet; see
"Closing the PIM-eligibility gap" below. Check the runbook job's verbose
logs to see whether that permission was actually available on a given run.

## Why an Azure Automation runbook instead of Cost Management's native scheduling

Cost Management exports are scheduled by Azure itself -- there's no
equivalent first-party "recurring RBAC report" feature. An Azure Automation
Account + Runbook + Schedule is the standard way to get the same
"set it up once, runs monthly forever" behavior for a custom script. This
mirrors the pattern already used for the AI usage guardrails automation in
`c:\my_work` (see its "Bonus -- Daily Usage Report Email" section).

## Setup

```powershell
.\New-AccessReviewAutomation.ps1 `
    -ProductionSubscriptionId "865fa361-a42e-4d3f-8566-cc34114cb8be" `
    -TestSubscriptionId "c0ae4ee0-ebef-4f37-a8e8-d590fc7417ba" `
    -DevSubscriptionId "86a78b30-a351-425e-9692-6a3938a559cf" `
    -StorageAccountResourceId "/subscriptions/865fa361-a42e-4d3f-8566-cc34114cb8be/resourceGroups/rg-cost-exports/providers/Microsoft.Storage/storageAccounts/stldmcostexports"
```

This creates a dedicated `rg-access-review` resource group and
`aa-access-review` Automation Account (separate from the guardrails
automation on purpose -- see the comment header in
[New-AccessReviewAutomation.ps1](New-AccessReviewAutomation.ps1)), grants
its managed identity Reader on the three subscriptions and Storage Blob
Data Contributor on the storage account, imports the runbook, and schedules
it for the 1st of every month at 06:00 UTC.

## Manual follow-up required before the first run

The provisioning script prints these at the end too -- repeated here since
they need a different person (a Global Administrator or Privileged Role
Administrator) to actually carry out:

1. **Import PowerShell modules into the Automation Account** (portal:
   Automation Account > Modules > Browse gallery): `Az.Accounts`,
   `Az.Resources`, `Az.Storage`, `ImportExcel`. The runbook will fail
   without these.
2. **Grant the managed identity Microsoft Graph application permissions**
   (`GroupMember.Read.All`, `User.Read.All`) so it can expand group
   membership -- see the exact `New-MgServicePrincipalAppRoleAssignment`
   commands printed at the end of `New-AccessReviewAutomation.ps1`.
3. **Grant one more Graph permission for PIM-eligible group data**:
   `PrivilegedAccess.Read.AzureADGroup` -- see "Closing the PIM-eligibility
   gap" below. Without it the report still runs fine, just without
   "Eligible" rows for PIM-managed admin groups.

## Closing the PIM-eligibility gap (not yet done)

To see PIM-eligible (not-yet-activated) membership in admin groups like
`BDECompute-Platform-SuperAdmins-Prod`, the reporting identity's Graph
permissions need `PrivilegedAccess.Read.AzureADGroup` in addition to
`GroupMember.Read.All`/`User.Read.All`. This requires a Global
Administrator or Privileged Role Administrator to grant, same as the
other Graph permissions:

```powershell
Connect-MgGraph -Scopes "AppRoleAssignment.ReadWrite.All"
$graphSpId = (Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'").Id
$targetSpId = "<the reporting identity's service principal object ID>"
$appRole = (Get-MgServicePrincipal -ServicePrincipalId $graphSpId).AppRoles |
  Where-Object { $_.Value -eq "PrivilegedAccess.Read.AzureADGroup" }
New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $targetSpId -PrincipalId $targetSpId `
  -ResourceId $graphSpId -AppRoleId $appRole.Id
```

Verify it worked with the same check used to originally confirm the gap:
```powershell
$tokenObj = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com"
$token = if ($tokenObj.Token -is [securestring]) {
  [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($tokenObj.Token))
} else { $tokenObj.Token }
Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/identityGovernance/privilegedAccess/group/eligibilitySchedules" -Headers @{ Authorization = "Bearer $token" }
```
A 403 means the permission hasn't propagated yet or wasn't granted to the
right identity; a 200 with data (or an empty `value` array, which is a
valid "nobody eligible right now" answer) means it worked.

## Storage firewall exposure window

`stldmcostexports` sits at `defaultAction: Deny` by default, same as the
Cost Management exports storage account -- Cost Management writes still
work because the account's `bypass: AzureServices` setting already covers
it. Azure Automation isn't on that trusted list, and (confirmed against
Microsoft docs) Automation Accounts can't use storage resource instance
rules, and Automation cloud jobs can't reach private-endpoint-secured
resources at all -- so there's no way to give the runbook a standing,
narrowly-scoped exception the way an IP rule does for a person's machine.

Instead, `Get-AzureAccessReview.ps1` flips `defaultAction` to `Allow`
immediately before uploading the workbook and back to `Deny` in a
`finally` block right after (see `Set-StorageFirewallDefaultAction`) --
open only for the seconds it takes to upload one file, not for the whole
run. This needs the "Storage Account Contributor" role (control-plane,
scoped to just this storage account) on top of the existing "Storage Blob
Data Contributor" (data-plane) role -- `New-AccessReviewAutomation.ps1`
grants both.

If the runbook is killed mid-run, the firewall could be left open; check
`az storage account show --name stldmcostexports --resource-group
rg-cost-exports --query networkRuleSet.defaultAction` after an
unexpectedly-terminated job and reset to `Deny` manually if needed.

## Testing before the schedule fires

Run it on demand instead of waiting for the 1st of the month:

```powershell
Start-AzAutomationRunbook -ResourceGroupName "rg-access-review" -AutomationAccountName "aa-access-review" `
    -Name "Get-AzureAccessReview" -Parameters @{
      ProductionSubscriptionId = "865fa361-a42e-4d3f-8566-cc34114cb8be"
      TestSubscriptionId       = "c0ae4ee0-ebef-4f37-a8e8-d590fc7417ba"
      DevSubscriptionId        = "86a78b30-a351-425e-9692-6a3938a559cf"
      StorageAccountResourceId = "/subscriptions/865fa361-a42e-4d3f-8566-cc34114cb8be/resourceGroups/rg-cost-exports/providers/Microsoft.Storage/storageAccounts/stldmcostexports"
      UseManagedIdentity       = $true
    }
```

Or run `Get-AzureAccessReview.ps1` directly on your own machine after
`Connect-AzAccount` (omit `-UseManagedIdentity`) to sanity-check the output
before relying on the scheduled runbook.

## Output

Each run uploads to `access-reviews/<yyyyMM>/azure-access-review-<yyyyMMdd>.xlsx`
in the storage account -- one file per month, dated, so history
accumulates rather than being overwritten. Download via the portal or
Azure Storage Explorer using an Entra account with Storage Blob Data
Reader on the account (shared keys are disabled, same as the cost exports).

## Sensitivity

This workbook contains real user identities, role assignments, and group
memberships across three subscriptions. Treat it like the source access
review file -- keep it in the authorized storage location, don't paste its
contents into chat tools or other unmanaged channels, and grant read access
only to people who actually need it for the review.
