<#
Provisions a dedicated, firewalled Azure Storage account for Cost Management
scheduled exports:
https://learn.microsoft.com/en-us/azure/cost-management-billing/costs/tutorial-improved-exports

What this does:
  1. Registers the Microsoft.CostManagementExports resource provider on the
     target subscription. Required to reference the destination storage
     account by resource ID, which is what lets Cost Management reach an
     account that has a firewall.
  2. Creates (or reuses) a resource group and storage account with:
       - public network access enabled, default-deny network rules, and the
         "AzureServices" trusted-service bypass -- Cost Management's export
         service is a first-party control-plane service, not a VNet client,
         so it can only reach a firewalled account through this bypass.
       - TLS 1.2 minimum, HTTPS-only, no shared-key access (Entra ID / RBAC
         only), blob versioning + 14-day soft delete.
  3. Creates the destination blob container.
  4. Grants Storage Blob Data Contributor on the account to the caller (or
     -GrantAccessToPrincipalId) -- container creation is a data-plane
     operation, which management-plane roles like Owner do not cover once
     shared-key access is disabled.

Run New-CostManagementExports.ps1 afterwards to create the exports against
the storage account this script provisions.

Usage:
  .\New-CostExportStorage.ps1 -SubscriptionId <sub-id> `
      -StorageAccountName stcostexportsprod -Location eastus2

  .\New-CostExportStorage.ps1 -SubscriptionId <sub-id> `
      -StorageAccountName stcostexportsprod -AllowedIpRanges "203.0.113.4" `
      -GrantAccessToPrincipalId <object-id>
#>
param(
  [Parameter(Mandatory)][string]$SubscriptionId,
  [string]$ResourceGroupName = "rg-cost-exports",
  [Parameter(Mandatory)][ValidatePattern('^[a-z0-9]{3,24}$')][string]$StorageAccountName,
  [string]$Location = "eastus2",
  [string]$ContainerName = "cost-management-exports",
  [string[]]$AllowedIpRanges = @(),
  [string]$GrantAccessToPrincipalId = ""
)

$ErrorActionPreference = "Stop"

Write-Host "Setting context to subscription $SubscriptionId..." -ForegroundColor Cyan
Set-AzContext -SubscriptionId $SubscriptionId | Out-Null

Write-Host "Registering Microsoft.CostManagementExports resource provider..." -ForegroundColor Cyan
Register-AzResourceProvider -ProviderNamespace "Microsoft.CostManagementExports" | Out-Null
do {
  Start-Sleep -Seconds 5
  $rpState = (Get-AzResourceProvider -ProviderNamespace "Microsoft.CostManagementExports" |
    Select-Object -First 1).RegistrationState
  Write-Host "  Registration state: $rpState"
} while ($rpState -ne "Registered")

$rg = Get-AzResourceGroup -Name $ResourceGroupName -ErrorAction SilentlyContinue
if (-not $rg) {
  Write-Host "Creating resource group $ResourceGroupName in $Location..." -ForegroundColor Cyan
  $rg = New-AzResourceGroup -Name $ResourceGroupName -Location $Location
}

$sa = Get-AzStorageAccount -ResourceGroupName $ResourceGroupName -Name $StorageAccountName -ErrorAction SilentlyContinue
if (-not $sa) {
  Write-Host "Creating storage account $StorageAccountName..." -ForegroundColor Cyan
  $sa = New-AzStorageAccount `
    -ResourceGroupName $ResourceGroupName `
    -Name $StorageAccountName `
    -Location $Location `
    -SkuName "Standard_LRS" `
    -Kind "StorageV2" `
    -MinimumTlsVersion "TLS1_2" `
    -EnableHttpsTrafficOnly $true `
    -AllowBlobPublicAccess $false `
    -AllowSharedKeyAccess $false `
    -PublicNetworkAccess "Enabled"
} else {
  Write-Host "Storage account $StorageAccountName already exists -- reusing." -ForegroundColor Yellow
}

Write-Host "Configuring firewall: default deny + AzureServices trusted-service bypass..." -ForegroundColor Cyan
$ruleSetParams = @{
  ResourceGroupName = $ResourceGroupName
  Name              = $StorageAccountName
  DefaultAction     = "Deny"
  Bypass            = "AzureServices"
}
if ($AllowedIpRanges.Count -gt 0) {
  $ruleSetParams["IPRule"] = @($AllowedIpRanges | ForEach-Object { @{ IPAddressOrRange = $_; Action = "allow" } })
}
Update-AzStorageAccountNetworkRuleSet @ruleSetParams | Out-Null

Write-Host "Enabling blob versioning and soft delete..." -ForegroundColor Cyan
Update-AzStorageBlobServiceProperty -ResourceGroupName $ResourceGroupName -StorageAccountName $StorageAccountName `
  -IsVersioningEnabled $true | Out-Null
Enable-AzStorageBlobDeleteRetentionPolicy -ResourceGroupName $ResourceGroupName -StorageAccountName $StorageAccountName `
  -RetentionDays 14 | Out-Null
Enable-AzStorageContainerDeleteRetentionPolicy -ResourceGroupName $ResourceGroupName -StorageAccountName $StorageAccountName `
  -RetentionDays 14 | Out-Null

if (-not $GrantAccessToPrincipalId) {
  Write-Host "No -GrantAccessToPrincipalId supplied -- attempting to resolve the signed-in user..." -ForegroundColor Cyan
  try {
    $GrantAccessToPrincipalId = (Get-AzADUser -SignedIn -ErrorAction Stop).Id
  } catch {
    Write-Error "Could not auto-resolve the signed-in user's object ID (often blocked by Conditional Access). Re-run with -GrantAccessToPrincipalId <objectId>."
  }
}

Write-Host "Granting Storage Blob Data Contributor to $GrantAccessToPrincipalId..." -ForegroundColor Cyan
$existingRole = Get-AzRoleAssignment -ObjectId $GrantAccessToPrincipalId -Scope $sa.Id `
  -RoleDefinitionName "Storage Blob Data Contributor" -ErrorAction SilentlyContinue
if (-not $existingRole) {
  New-AzRoleAssignment -ObjectId $GrantAccessToPrincipalId -Scope $sa.Id `
    -RoleDefinitionName "Storage Blob Data Contributor" | Out-Null
  Write-Host "Waiting for role assignment to propagate..." -ForegroundColor Cyan
  Start-Sleep -Seconds 30
}

Write-Host "Creating container $ContainerName..." -ForegroundColor Cyan
$ctx = New-AzStorageContext -StorageAccountName $StorageAccountName -UseConnectedAccount
$container = Get-AzStorageContainer -Name $ContainerName -Context $ctx -ErrorAction SilentlyContinue
if (-not $container) {
  New-AzStorageContainer -Name $ContainerName -Context $ctx -Permission Off | Out-Null
}

Write-Host "`nDone. Pass this resource ID as -StorageAccountResourceId to New-CostManagementExports.ps1:" -ForegroundColor Green
Write-Host $sa.Id
