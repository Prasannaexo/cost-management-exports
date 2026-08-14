# Cost Management Exports

Standalone PowerShell automation for Azure Cost Management scheduled
exports -- provisioning a firewalled destination storage account and
creating the exports against it.

Reference: [Tutorial - Create and manage Cost Management exports](https://learn.microsoft.com/en-us/azure/cost-management-billing/costs/tutorial-improved-exports)

Scripts: [scripts/](scripts/)

## What gets created

| Export | Scope | Dataset type | Recurrence |
| --- | --- | --- | --- |
| `finops-actualcost-mg` | Management group | `Usage` (actual cost) | Daily |
| `finops-pricesheet` | Billing account | `PriceSheet` | Monthly |
| `finops-reservationrecs` | Billing account | `ReservationRecommendations` | Daily |

All three write CSV files to one dedicated, firewalled storage account.

## Prerequisites

- **Confirm your agreement type before running this.** Management group
  scope exports are only supported for **Enterprise Agreement (EA)**
  billing, not MCA. If you're on MCA, skip `-ManagementGroupId` in
  `New-CostManagementExports.ps1` and use a subscription scope export
  instead (see the tutorial's supported-scopes table).
- Owner (or Contributor + `Microsoft.Authorization/roleAssignments/write` +
  `Microsoft.Authorization/permissions/read`) on the storage account's
  subscription.
- Owner or Reader with export permission on the target management group /
  billing account -- see [Understand and work with scopes](https://learn.microsoft.com/en-us/azure/cost-management-billing/costs/understand-work-scopes).
- PowerShell `Az` module: `Az.Accounts`, `Az.Storage`, `Az.Resources`.
- New subscriptions can take up to 48 hours before Cost Management features
  are usable.

## Setup

```powershell
# 1. Provision the firewalled storage account + container
.\scripts\New-CostExportStorage.ps1 `
    -SubscriptionId <subscription-id> `
    -StorageAccountName stcostexportsprod `
    -Location eastus2

# 2. Create the exports against it
.\scripts\New-CostManagementExports.ps1 `
    -ManagementGroupId <mg-id> `
    -BillingAccountId <billing-account-id> `
    -StorageAccountResourceId "<resource ID printed by step 1>"
```

Use `-DryRun` on `New-CostManagementExports.ps1` to print the request
bodies without creating anything, e.g. to review before a compliance
sign-off.

Re-running `New-CostManagementExports.ps1` with the same export names
updates the existing exports in place (the underlying API call is a PUT).

## Firewall mechanics (why this works)

Cost Management's export service is a first-party control-plane service,
not a VNet client. Exporting to a storage account with a firewall requires:

- Public network access **enabled** on the storage account, with the
  network rules' default action set to **Deny** and **Allow trusted Azure
  service access** (`Bypass = AzureServices`) turned on. `New-CostExportStorage.ps1`
  configures this.
- The subscription registered against the `Microsoft.CostManagementExports`
  resource provider, so the storage account can be referenced by resource
  ID. `New-CostExportStorage.ps1` handles this too.
- A system-assigned managed identity on each export (`New-CostManagementExports.ps1`
  requests one via `"identity": {"type": "SystemAssigned"}`). Azure
  auto-grants it *Storage Blob Data Contributor*, scoped to the container,
  the first time the export is created or updated -- provided the
  principal running the PUT has `Microsoft.Authorization/roleAssignments/write`
  on the storage account. No role assignment for the export's identity is
  ever pre-created by the scripts; it doesn't exist until the export does.
- If you later change the storage account's network configuration, you
  must re-save (re-PUT) each affected export for the change to take
  effect.
- Firewalled destinations are only supported within the same Entra tenant
  as the export -- not for cross-tenant exports.

## Verifying data landed

1. In the Azure portal, open the storage account (or use Azure Storage
   Explorer, "Open in Explorer" from the export's details pane).
2. Browse `<container>/<rootFolderPath>/<exportName>/<YYYYMMDD-YYYYMMDD>/<runId>/`.
3. Each run produces one or more partitioned CSVs (files are split at
   ~1 GB regardless of dataset size -- always driven by size, not row
   count) plus a `manifest.json` listing every partition in order. Don't
   hardcode partition file names; read the manifest.
4. First-run data can take up to 24 hours to appear. Subsequent runs land
   within ~4 hours of the scheduled run starting.

## Managing exports

Via the portal (Cost Management > Exports at the relevant scope) or
`Az.CostManagement` cmdlets:

- `Get-AzCostManagementExport -Scope <scope>` -- list/inspect, including run history.
- `Invoke-AzCostManagementExportExecute` -- "Run now", or backfill a specific
  historical month (`-Selected dates`, up to 13 months via the API path
  used here; up to 7 years for cost/usage datasets via the Exports Execute
  REST API directly -- see the tutorial's "How much historical data can I
  retrieve" FAQ).
- Disable/delete an export from the portal or by PATCHing/DELETEing the
  same resource path the scripts PUT to. Deleting or changing an export's
  destination automatically removes its managed identity's role assignment,
  provided the caller has `Microsoft.Authorization/roleAssignments/delete`
  -- otherwise it must be removed manually.

## Known limitations (from the tutorial)

- Management group scope: EA only, usage/actual-cost data only -- no
  amortized cost, no purchases/reservations, no FOCUS format, single
  currency only.
- Price sheet and reservation recommendations: billing account or billing
  profile scope only -- never subscription or management group.
- Reservation recommendations reflect a current snapshot only; no
  historical backfill.
- File partitioning cannot be disabled.
- Management groups are capped at 3,000 subscriptions for cost purposes;
  split larger orgs into multiple management groups with one export each.
