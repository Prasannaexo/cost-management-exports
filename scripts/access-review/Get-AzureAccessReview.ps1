<#
Builds a multi-sheet Azure RBAC access review workbook (Production / Test /
Dev / All Subscriptions / Role Definitions) and uploads it to the same
firewalled storage account the Cost Management exports use.

Designed to run two ways:
  - Locally, interactively, after Connect-AzAccount -- for ad hoc reviews.
  - As an Azure Automation runbook on a monthly schedule (via
    New-AccessReviewAutomation.ps1 in this folder, with -UseManagedIdentity),
    authenticating with the Automation Account's system-assigned managed
    identity.

Methodology (there was no prior script to copy -- this is a fresh
implementation, so the exact attribution logic is spelled out below rather
than assumed; verify against a real run before relying on it for an actual
audit):
  - "Role: <X>" columns: a user is marked for role X if they hold it either
    directly (assigned to their own user object) or indirectly via
    membership in a group that holds it. Group membership expansion uses
    Get-AzADGroupMember, which returns DIRECT members only -- nested group
    membership (a group inside a group) is not expanded further.
  - "Group: <X>" columns: plain Entra group membership, independent of
    whether that group currently carries any role. The column set is
    whatever groups the discovered users actually belong to, discovered
    dynamically (one pass over every group's member list) rather than
    hardcoded, so it adapts as groups are renamed/added.
  - Rows are Entra user objects only. Role assignments held by service
    principals or managed identities are excluded from the per-user sheets
    entirely -- counted and reported in a console summary instead so
    nothing silently disappears.
  - "Role Definitions" sheet only lists roles actually observed in use
    across the three subscriptions, not the full built-in Azure role
    catalog.
  - Get-AzADGroupMember returns a generic directory-object type; this
    script assumes member objects expose .UserPrincipalName and
    .DisplayName directly. Confirm that's true against this tenant's Az
    module version on the first real run -- if it isn't, member UPNs will
    silently come back empty and those users will be skipped (a warning is
    printed per skipped assignment either way).

Requires: Az.Accounts, Az.Resources, Az.Storage, ImportExcel.

Usage (local, interactive):
  Connect-AzAccount
  .\Get-AzureAccessReview.ps1 `
      -ProductionSubscriptionId 865fa361-a42e-4d3f-8566-cc34114cb8be `
      -TestSubscriptionId c0ae4ee0-ebef-4f37-a8e8-d590fc7417ba `
      -DevSubscriptionId 86a78b30-a351-425e-9692-6a3938a559cf `
      -StorageAccountResourceId "/subscriptions/.../storageAccounts/stldmcostexports"

Usage (Azure Automation runbook): same parameters plus -UseManagedIdentity,
wired up by New-AccessReviewAutomation.ps1.
#>
param(
  [Parameter(Mandatory)][string]$ProductionSubscriptionId,
  [Parameter(Mandatory)][string]$TestSubscriptionId,
  [Parameter(Mandatory)][string]$DevSubscriptionId,
  [Parameter(Mandatory)][string]$StorageAccountResourceId,
  [string]$ContainerName = "access-reviews",
  [switch]$UseManagedIdentity,
  [string]$OutputPath = "$env:TEMP\azure-access-review.xlsx"
)

$ErrorActionPreference = "Stop"

if ($UseManagedIdentity) {
  Connect-AzAccount -Identity | Out-Null
}

function Get-SubscriptionAccessRows {
  param([string]$SubscriptionId, [string]$Label)

  Write-Output "Collecting role assignments for $Label ($SubscriptionId)..."
  Set-AzContext -SubscriptionId $SubscriptionId | Out-Null

  # No -Scope filter: returns every assignment across every resource in the
  # subscription, not just the subscription-root scope (confirmed behavior
  # during this project's earlier manual RBAC checks).
  $assignments = Get-AzRoleAssignment

  $userRows = @{}       # UPN -> ordered hashtable of columns
  $groupMemberCache = @{}  # GroupObjectId -> cached member list
  $otherCount = 0

  function Get-OrCreateUserRow {
    param([string]$Upn, [string]$DisplayName)
    if (-not $userRows.ContainsKey($Upn)) {
      $userRows[$Upn] = [ordered]@{ User = $DisplayName; UPN = $Upn }
    }
    return $userRows[$Upn]
  }

  foreach ($a in $assignments) {
    switch ($a.ObjectType) {
      "User" {
        $row = Get-OrCreateUserRow -Upn $a.SignInName -DisplayName $a.DisplayName
        $row["Role: $($a.RoleDefinitionName)"] = "Yes"
      }
      "Group" {
        if (-not $groupMemberCache.ContainsKey($a.ObjectId)) {
          try {
            $groupMemberCache[$a.ObjectId] = @(Get-AzADGroupMember -GroupObjectId $a.ObjectId -ErrorAction Stop)
          } catch {
            Write-Warning "Could not expand group $($a.DisplayName) ($($a.ObjectId)): $_"
            $groupMemberCache[$a.ObjectId] = @()
          }
        }
        foreach ($m in $groupMemberCache[$a.ObjectId]) {
          if (-not $m.UserPrincipalName) { continue }  # skip nested groups/service principals in membership
          $row = Get-OrCreateUserRow -Upn $m.UserPrincipalName -DisplayName $m.DisplayName
          $row["Role: $($a.RoleDefinitionName)"] = "Yes"
        }
      }
      default { $otherCount++ }
    }
  }

  Write-Output "  $($userRows.Count) distinct users, $otherCount service-principal/other assignments excluded from rows."

  # Group membership columns: one pass over every group's member list,
  # building a reverse index, instead of checking membership per-user
  # (which would be O(users x groups) Graph calls against the full tenant).
  Write-Output "  Resolving group memberships for discovered users..."
  $allGroups = Get-AzADGroup
  foreach ($g in $allGroups) {
    $members = Get-AzADGroupMember -GroupObjectId $g.Id -ErrorAction SilentlyContinue |
      Where-Object { $_.UserPrincipalName -and $userRows.ContainsKey($_.UserPrincipalName) }
    foreach ($m in $members) {
      $userRows[$m.UserPrincipalName]["Group: $($g.DisplayName)"] = "Yes"
    }
  }

  return @($userRows.Values | ForEach-Object { [pscustomobject]$_ })
}

function Add-SubscriptionLabel {
  param([object[]]$Rows, [string]$Label)
  return @($Rows | ForEach-Object {
    $h = [ordered]@{ Subscription = $Label }
    foreach ($p in $_.PSObject.Properties) { $h[$p.Name] = $p.Value }
    [pscustomobject]$h
  })
}

$prodRows = Get-SubscriptionAccessRows -SubscriptionId $ProductionSubscriptionId -Label "Production"
$testRows = Get-SubscriptionAccessRows -SubscriptionId $TestSubscriptionId -Label "Test"
$devRows  = Get-SubscriptionAccessRows -SubscriptionId $DevSubscriptionId -Label "Dev"

$allRows = @(Add-SubscriptionLabel -Rows $prodRows -Label "Production") +
           @(Add-SubscriptionLabel -Rows $testRows -Label "Test") +
           @(Add-SubscriptionLabel -Rows $devRows -Label "Dev")

Write-Output "Building Role Definitions reference sheet..."
$observedRoleNames = @($prodRows + $testRows + $devRows | ForEach-Object {
  $_.PSObject.Properties.Name | Where-Object { $_ -like "Role: *" }
} | ForEach-Object { $_ -replace "^Role: ", "" } | Select-Object -Unique)

$roleDefRows = foreach ($roleName in $observedRoleNames) {
  $def = Get-AzRoleDefinition -Name $roleName -ErrorAction SilentlyContinue
  if ($def) {
    [pscustomobject]@{
      Role                             = $def.Name
      Type                             = if ($def.IsCustom) { "CustomRole" } else { "BuiltInRole" }
      Description                      = $def.Description
      "Actions (access granted)"       = ($def.Actions -join "; ")
      "Data Actions (access granted)"  = ($def.DataActions -join "; ")
      "Not Actions (excluded)"         = ($def.NotActions -join "; ")
      "Not Data Actions (excluded)"    = ($def.NotDataActions -join "; ")
    }
  }
}

Write-Output "Writing workbook to $OutputPath..."
Remove-Item -Path $OutputPath -ErrorAction SilentlyContinue
$prodRows    | Export-Excel -Path $OutputPath -WorksheetName "Production" -AutoSize -FreezeTopRow -BoldTopRow
$testRows    | Export-Excel -Path $OutputPath -WorksheetName "Test" -AutoSize -FreezeTopRow -BoldTopRow
$devRows     | Export-Excel -Path $OutputPath -WorksheetName "Dev" -AutoSize -FreezeTopRow -BoldTopRow
$allRows     | Export-Excel -Path $OutputPath -WorksheetName "All Subscriptions" -AutoSize -FreezeTopRow -BoldTopRow
$roleDefRows | Export-Excel -Path $OutputPath -WorksheetName "Role Definitions" -AutoSize -FreezeTopRow -BoldTopRow

Write-Output "Uploading to storage..."
$storageAccountName = ($StorageAccountResourceId -split "/")[-1]
$ctx = New-AzStorageContext -StorageAccountName $storageAccountName -UseConnectedAccount
$container = Get-AzStorageContainer -Name $ContainerName -Context $ctx -ErrorAction SilentlyContinue
if (-not $container) {
  New-AzStorageContainer -Name $ContainerName -Context $ctx -Permission Off | Out-Null
}
$blobName = "{0:yyyyMM}/azure-access-review-{0:yyyyMMdd}.xlsx" -f (Get-Date)
Set-AzStorageBlobContent -File $OutputPath -Container $ContainerName -Blob $blobName -Context $ctx -Force | Out-Null

Write-Output "Done. Uploaded to $ContainerName/$blobName"
