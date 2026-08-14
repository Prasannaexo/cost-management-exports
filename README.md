# Cost Management Exports

Standalone PowerShell automation for Azure Cost Management scheduled
exports -- provisioning a firewalled destination storage account and
creating the exports against it.

Reference: [Tutorial - Create and manage Cost Management exports](https://learn.microsoft.com/en-us/azure/cost-management-billing/costs/tutorial-improved-exports)

Scripts: [scripts/](scripts/)

## What gets created

| Export | Scope | Dataset type | Recurrence | When it's created |
| --- | --- | --- | --- | --- |
| `finops-actualcost-<label>` | Subscription (one per entry in `-SubscriptionIds`) | `ActualCost` | Daily | `-SubscriptionIds` passed |
| `finops-actualcost-mg` | Management group | `Usage` (actual cost) | Daily | `-ManagementGroupId` passed (EA only) |
| `finops-pricesheet` | Billing account | `PriceSheet` | Monthly | `-BillingAccountId` passed |
| `finops-reservationrecs` | Billing account | `ReservationRecommendations` | Daily | `-BillingAccountId` passed |

Pass whichever combination of `-SubscriptionIds` / `-ManagementGroupId` /
`-BillingAccountId` matches what you actually have access to -- only the
matching exports are created. Every export, from every subscription, writes
to the **same** dedicated, firewalled storage account -- that's what makes
it possible to build one Power BI report across all of them.

**If your billing is managed by a reseller/vendor (CSP / Microsoft Partner
Agreement)**, you typically won't have a billing account ID or management
group to point at -- only `-SubscriptionIds` is available to you. That's
enough to get actual daily cost/usage data flowing to storage for Power BI;
price sheet and reservation recommendations require direct billing account
access you likely don't have as a CSP customer, so just omit
`-BillingAccountId` and `-ManagementGroupId`.

## Prerequisites

- Owner (or Contributor + `Microsoft.Authorization/roleAssignments/write` +
  `Microsoft.Authorization/permissions/read`) on the storage account's
  subscription.
- Owner or Reader with export permission on **every** source subscription
  you list in `-SubscriptionIds` -- see [Understand and work with scopes](https://learn.microsoft.com/en-us/azure/cost-management-billing/costs/understand-work-scopes).
  Cross-subscription data transfer into the storage account itself doesn't
  need per-subscription storage RBAC; see "Firewall mechanics" below for
  why.
- PowerShell `Az` module: `Az.Accounts`, `Az.Storage`, `Az.Resources`.
- New subscriptions can take up to 48 hours before Cost Management features
  are usable.

## Setup

Example using this org's subscriptions (three from the `scaling-infra`
Terraform repo's runbooks, plus the Dev subscription referenced but not
recorded with an ID in that repo) -- one central storage account in the
**LDM internal prod** subscription (the actual primary subscription), fed
by exports from all four:

| Label | Subscription ID | Source |
| --- | --- | --- |
| `ldm-prod` | `865fa361-a42e-4d3f-8566-cc34114cb8be` | LDM internal prod (primary -- hosts the storage account) |
| `sandbox` | `eb9a9f59-a1df-475f-b951-bfd41f253982` | Sandbox |
| `dev` | `86a78b30-a351-425e-9692-6a3938a559cf` | Azure Subscription- Dev |
| `test` | `c0ae4ee0-ebef-4f37-a8e8-d590fc7417ba` | Test Subscription |

```powershell
# 1. Provision the firewalled storage account + container, in the LDM
#    internal prod subscription (the primary subscription)
.\scripts\New-CostExportStorage.ps1 `
    -SubscriptionId "865fa361-a42e-4d3f-8566-cc34114cb8be" `
    -StorageAccountName "stcostexportsprod" `
    -Location "eastus2"

# 2. Create one export per subscription against it, all landing in the
#    same container
.\scripts\New-CostManagementExports.ps1 `
    -SubscriptionIds "865fa361-a42e-4d3f-8566-cc34114cb8be","eb9a9f59-a1df-475f-b951-bfd41f253982","86a78b30-a351-425e-9692-6a3938a559cf","c0ae4ee0-ebef-4f37-a8e8-d590fc7417ba" `
    -SubscriptionLabels "ldm-prod","sandbox","dev","test" `
    -StorageAccountResourceId "<resource ID printed by step 1>"
```

Add more subscriptions later by re-running step 2 with the full updated
list -- it's a PUT per export, so existing ones are left as-is and only new
labels create new exports.

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
- **Cross-subscription data transfer specifically** (multiple subscriptions'
  exports writing into one storage account that lives in a different
  subscription) also depends on the storage account's **"Permitted scope
  for copy operations"** setting being **"From any storage account"**
  (unrestricted) rather than scoped to "AAD" or "Private Link".
  `New-CostExportStorage.ps1` achieves this by simply never setting
  `-AllowedCopyScope` on the account (the ARM default is unrestricted) and
  prints a warning at the end if it ever finds that setting restricted.

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

## Connecting Power BI

The exported CSVs land at
`<container>/<rootFolderPath>/<exportName>/<YYYYMMDD-YYYYMMDD>/<runId>/`,
partitioned into multiple files with a `manifest.json` per run (see above).
Two ways to bring that into Power BI:

**Option A -- Power Query against the container (simplest, good starting point)**

1. In Power BI Desktop: **Get Data > Azure > Azure Blob Storage**.
2. Enter the storage account name and authenticate (Organizational
   account / Entra ID works if your account has the Storage Blob Data
   Reader role on the account -- ask whoever ran `New-CostExportStorage.ps1`
   to grant it, since shared keys are disabled on this account).
3. Navigate to the `cost-management-exports` container. Each subscription's
   data lands under its own label, e.g. `actualcost/ldm-prod/finops-actualcost-ldm-prod/`,
   `actualcost/sandbox/finops-actualcost-sandbox/`, `actualcost/dev/finops-actualcost-dev/`,
   `actualcost/test/finops-actualcost-test/`.
4. Use **Combine & Transform** on the `actualcost/` folder (not just one
   subscription's subfolder) so Power Query unions every subscription and
   every partition automatically -- this also picks up new daily runs
   without changing the query. Filter out `manifest.json` (it isn't a CSV)
   in the query, e.g. `Table.SelectRows(Source, each [Extension] = ".csv")`.
   Add a column derived from the folder path if you want to slice the
   report by subscription/label.
5. Set the dataset to refresh on a schedule that trails the export's own
   daily run (see run-history timing below) so Power BI always has the
   prior day's file.

**Option B -- Power BI dataflow / Fabric ingestion (more scale, less manual query wrangling)**

For larger, ongoing FinOps reporting, point a Power BI dataflow or a
Fabric pipeline at the same container instead of a Desktop query -- same
underlying files, but with proper incremental refresh and a shared,
governed dataset other reports can reuse. Worth moving to once the
Option A query is proven out.

Either way, budget for the up-to-24-hour delay on the first file and the
~4-hour delay on subsequent daily runs when you set the refresh schedule.

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
  profile scope only -- never subscription or management group. Under a
  CSP/vendor-managed agreement you typically don't have billing account
  access at all, so these two are simply unavailable to you -- subscription
  scope actual cost is what you get.
- Reservation recommendations reflect a current snapshot only; no
  historical backfill.
- File partitioning cannot be disabled.
- Management groups are capped at 3,000 subscriptions for cost purposes;
  split larger orgs into multiple management groups with one export each.
