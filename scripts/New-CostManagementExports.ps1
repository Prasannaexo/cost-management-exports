<#
Creates scheduled Azure Cost Management exports against a firewalled storage
account, per:
https://learn.microsoft.com/en-us/azure/cost-management-billing/costs/tutorial-improved-exports

Run New-CostExportStorage.ps1 first -- this script needs the storage account
resource ID it provisions.

Datasets and scope, per tutorial limitations:
  - Subscription scope: actual cost/usage export ("ActualCost"), one per
    subscription in -SubscriptionIds. This is the scope to use when your
    billing is managed by a reseller/vendor (CSP / Microsoft Partner
    Agreement) and you don't have billing account access -- MPA customers
    are supported at customer, subscription, and resource group scope for
    actual cost, but NOT at billing account scope. Every subscription's
    export writes to its own rootFolderPath (actualcost/<label>) under the
    same storage account, so multiple subscriptions can feed one Power BI
    report from a single container. This only requires: (a) Owner/Contributor
    on each source subscription to create its export, and (b)
    roleAssignments/write on the destination storage account (wherever it
    lives) so Azure can auto-grant each export's managed identity access --
    the storage account's own "Permitted scope for copy operations" needs
    to stay unrestricted ("From any storage account") for this
    cross-subscription write to work; New-CostExportStorage.ps1 leaves it
    that way by not setting -AllowedCopyScope.
  - Management group scope only supports the "Usage" (actual cost) export
    type, and only for Enterprise Agreement (EA) billing -- not MCA/MPA.
    Price sheet, amortized cost, and reservation datasets are not supported
    at management group scope.
  - Price sheet and reservation recommendations are billing-account-wide
    datasets -- they require direct billing account access (EA or MCA that
    you bought directly), which a CSP/vendor-managed subscription normally
    doesn't have. Skip -BillingAccountId if that's your situation; those
    two exports simply won't be created.

This calls the Cost Management REST API directly via Invoke-AzRestMethod
(api-version 2023-08-01, the documented minimum for firewalled-storage
support) rather than the Az.CostManagement module cmdlets, which are still
in preview and don't expose the reservation-recommendation dataset filters
(ReservationScope / ResourceType / LookBackPeriod) used below.

PUT is idempotent here: re-running with the same export names updates the
existing exports in place.

Usage (vendor-managed billing -- one export per subscription, all landing in
the same central storage account so they can all feed one Power BI report):
  .\New-CostManagementExports.ps1 `
      -SubscriptionIds "eb9a9f59-...","c0ae4ee0-...","865fa361-..." `
      -SubscriptionLabels "sandbox","test","ldm-prod" `
      -StorageAccountResourceId "/subscriptions/.../storageAccounts/stcostexportsprod" `
      -Location eastus2

Usage (EA with management group + billing account access):
  .\New-CostManagementExports.ps1 `
      -ManagementGroupId mg-acme `
      -BillingAccountId 1234567 `
      -StorageAccountResourceId "/subscriptions/.../storageAccounts/stcostexportsprod" `
      -Location eastus2

  Add -DryRun to print the request bodies without calling the API.
#>
param(
  [string[]]$SubscriptionIds = @(),
  [string[]]$SubscriptionLabels = @(),
  [string]$ManagementGroupId = "",
  [string]$BillingAccountId = "",
  [Parameter(Mandatory)][string]$StorageAccountResourceId,
  [string]$ContainerName = "cost-management-exports",
  [string]$Location = "eastus2",
  [string]$ExportNamePrefix = "finops",
  [string]$ApiVersion = "2023-08-01",
  [string]$ReservationScope = "Single",
  [string]$ReservationResourceType = "VirtualMachines",
  [string]$ReservationLookBackPeriod = "Last7Days",
  [switch]$DryRun
)

$ErrorActionPreference = "Stop"

if ($SubscriptionIds.Count -eq 0 -and -not $ManagementGroupId -and -not $BillingAccountId) {
  Write-Error "Pass at least one of -SubscriptionIds, -ManagementGroupId, or -BillingAccountId."
}
if ($SubscriptionLabels.Count -gt 0 -and $SubscriptionLabels.Count -ne $SubscriptionIds.Count) {
  Write-Error "-SubscriptionLabels must have the same number of entries as -SubscriptionIds (or be omitted)."
}

$recurrenceFrom = (Get-Date).ToUniversalTime().Date
$recurrenceTo   = $recurrenceFrom.AddYears(2)

function New-ExportBody {
  param(
    [string]$Type,
    [string]$Timeframe,
    [string]$Recurrence,
    [string]$RootFolderPath,
    [string]$Granularity,
    [object[]]$Filters
  )

  $dataSet = @{ configuration = @{} }
  if ($Granularity) { $dataSet["granularity"] = $Granularity }
  if ($Filters) { $dataSet.configuration["filters"] = $Filters }
  if ($dataSet.configuration.Count -eq 0) { $dataSet.Remove("configuration") }

  return @{
    location = $Location
    identity = @{ type = "SystemAssigned" }
    properties = @{
      format                = "Csv"
      dataOverwriteBehavior = "OverwritePreviousReport"
      partitionData         = $true
      definition            = @{
        type      = $Type
        timeframe = $Timeframe
        dataSet   = $dataSet
      }
      deliveryInfo = @{
        destination = @{
          type           = "AzureBlob"
          resourceId     = $StorageAccountResourceId
          container      = $ContainerName
          rootFolderPath = $RootFolderPath
        }
      }
      schedule = @{
        recurrence       = $Recurrence
        recurrencePeriod = @{
          from = $recurrenceFrom.ToString("yyyy-MM-ddTHH:mm:ssZ")
          to   = $recurrenceTo.ToString("yyyy-MM-ddTHH:mm:ssZ")
        }
        status = "Active"
      }
    }
  }
}

function Publish-Export {
  param(
    [string]$ScopePath,
    [string]$ExportName,
    [hashtable]$Body
  )

  $json = $Body | ConvertTo-Json -Depth 12
  $uri  = "$ScopePath/providers/Microsoft.CostManagement/exports/${ExportName}?api-version=$ApiVersion"

  if ($DryRun) {
    Write-Host "`n--- DRY RUN: PUT $uri ---" -ForegroundColor Yellow
    Write-Host $json
    return
  }

  Write-Host "Creating/updating export $ExportName at $ScopePath..." -ForegroundColor Cyan
  $resp = Invoke-AzRestMethod -Path $uri -Method PUT -Payload $json
  if ($resp.StatusCode -notin 200, 201) {
    Write-Error "Failed to create $ExportName ($($resp.StatusCode)): $($resp.Content)"
  } else {
    Write-Host "  OK" -ForegroundColor Green
  }
}

for ($i = 0; $i -lt $SubscriptionIds.Count; $i++) {
  $subId = $SubscriptionIds[$i]
  $label = if ($SubscriptionLabels.Count -gt $i) { $SubscriptionLabels[$i] } else { $subId }

  $subScope = "/subscriptions/$subId"
  $body = New-ExportBody -Type "ActualCost" -Timeframe "MonthToDate" -Recurrence "Daily" `
    -RootFolderPath "actualcost/$label" -Granularity "Daily"
  Publish-Export -ScopePath $subScope -ExportName "$ExportNamePrefix-actualcost-$label" -Body $body
}

if ($ManagementGroupId) {
  Write-Warning "Management group scope exports require an Enterprise Agreement (EA) -- not MCA. Confirm your agreement type before relying on this export; MCA will fail here with an unsupported-scope error."

  $mgScope = "/providers/Microsoft.Management/managementGroups/$ManagementGroupId"
  $body = New-ExportBody -Type "Usage" -Timeframe "MonthToDate" -Recurrence "Daily" `
    -RootFolderPath "actualcost" -Granularity "Daily"
  Publish-Export -ScopePath $mgScope -ExportName "$ExportNamePrefix-actualcost-mg" -Body $body
}

if ($BillingAccountId) {
  $billingScope = "/providers/Microsoft.Billing/billingAccounts/$BillingAccountId"

  $priceSheetBody = New-ExportBody -Type "PriceSheet" -Timeframe "TheCurrentMonth" -Recurrence "Monthly" `
    -RootFolderPath "pricesheet"
  Publish-Export -ScopePath $billingScope -ExportName "$ExportNamePrefix-pricesheet" -Body $priceSheetBody

  $reservationFilters = @(
    @{ name = "ReservationScope"; value = $ReservationScope }
    @{ name = "ResourceType"; value = $ReservationResourceType }
    @{ name = "LookBackPeriod"; value = $ReservationLookBackPeriod }
  )
  $reservationBody = New-ExportBody -Type "ReservationRecommendations" -Timeframe "MonthToDate" -Recurrence "Daily" `
    -RootFolderPath "reservationrecommendations" -Filters $reservationFilters
  Publish-Export -ScopePath $billingScope -ExportName "$ExportNamePrefix-reservationrecs" -Body $reservationBody
}

Write-Host "`nExports can take up to 24 hours to produce their first file. Verify with Azure Storage Explorer or Get-AzCostManagementExport." -ForegroundColor Green
