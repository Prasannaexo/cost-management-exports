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

**Known gap, confirmed against this tenant (not just a theoretical risk):**
roles granted via Entra groups marked "assignable to Azure roles" --
every `BDECompute-Platform-*`/`BDECompute-*-SuperAdmins` group, which
carry most of the Owner/Contributor/AcrPush/Key Vault/Storage role grants
-- do not show up. `Get-AzADGroupMember` returns zero members for these
groups regardless of the caller's own admin rights, because Microsoft
Graph restricts membership visibility for role-assignable groups at the
calling-application permission-scope level, separate from the user's own
directory role. Ordinary groups and directly-assigned user roles resolve
correctly. The generated workbook includes a "READ ME - Limitations"
sheet stating this so it isn't mistaken for a complete picture. See
"Closing the group-membership gap" below for the actual fix, which hasn't
been applied yet.

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
   commands printed at the end of `New-AccessReviewAutomation.ps1`. Note:
   as of this writing these two are NOT sufficient to read role-assignable
   group membership -- see "Closing the group-membership gap" below.

## Closing the group-membership gap (not yet done)

To actually see roles granted via `BDECompute-*` groups, the reporting
identity's Graph permissions need `RoleManagement.Read.Directory` in
addition to `GroupMember.Read.All`/`User.Read.All` -- that scope is
specifically for reading role-assignable group membership, which is
otherwise restricted regardless of the caller's own directory role. This
requires a Global Administrator or Privileged Role Administrator to grant,
same as the other Graph permissions. Whether this alone fully resolves it
hasn't been confirmed -- test against one known role-assignable group
after granting it, the same way the current gap was confirmed:
```powershell
Get-AzADGroupMember -GroupObjectId "<a BDECompute-* group's object ID>" | Select-Object DisplayName, Id
```
If members show up now, re-run the report and the "READ ME - Limitations"
sheet's group count should drop.

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
