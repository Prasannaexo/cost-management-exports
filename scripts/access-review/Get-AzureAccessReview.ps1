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
    dynamically rather than hardcoded, so it adapts as groups are
    renamed/added.
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
    silently come back empty and those users will be skipped.

Performance: group membership is tenant-wide, not per-subscription, so it's
resolved ONCE (a single pass over every Entra group) and reused for both
(a) expanding group-based role assignments to member users and (b) the
"Group: X" membership columns, across all three subscriptions. An earlier
version of this script re-resolved it separately per subscription (3x the
Graph calls) -- don't reintroduce that.

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

foreach ($requiredModule in "Az.Accounts", "Az.Resources", "Az.Storage", "ImportExcel") {
  if (-not (Get-Module -ListAvailable -Name $requiredModule)) {
    throw "Required module '$requiredModule' is not installed in this PowerShell session ($($PSVersionTable.PSVersion), $($PSVersionTable.PSEdition)). Install it with: Install-Module -Name $requiredModule -Scope CurrentUser"
  }
  Import-Module -Name $requiredModule -ErrorAction Stop
}

if ($UseManagedIdentity) {
  Connect-AzAccount -Identity | Out-Null
}

# --- Tenant-wide group membership index (built once, reused everywhere) ---

function Get-MemberUpnAndName {
  # Get-AzADGroupMember's return type varies by Az.Resources version -- some
  # expose .UserPrincipalName/.DisplayName directly, others only populate
  # .AdditionalProperties['userPrincipalName']/['displayName']. Check both so
  # membership resolution doesn't silently come back empty on a version
  # where the flat properties aren't populated.
  param($Member)
  $upn = $Member.UserPrincipalName
  if (-not $upn -and $Member.AdditionalProperties -and $Member.AdditionalProperties.ContainsKey('userPrincipalName')) {
    $upn = $Member.AdditionalProperties['userPrincipalName']
  }
  $name = $Member.DisplayName
  if (-not $name -and $Member.AdditionalProperties -and $Member.AdditionalProperties.ContainsKey('displayName')) {
    $name = $Member.AdditionalProperties['displayName']
  }
  if (-not $upn) { return $null }
  return [pscustomobject]@{ UserPrincipalName = $upn; DisplayName = $name }
}

Write-Host "Building tenant-wide group membership index (one pass over every Entra group)..."
$allGroups = Get-AzADGroup
Write-Host "  $($allGroups.Count) groups found. Resolving members..."

$groupMembersById = @{}  # GroupObjectId -> array of normalized {UserPrincipalName, DisplayName} objects
$groupNameById    = @{}  # GroupObjectId -> DisplayName
$i = 0
foreach ($g in $allGroups) {
  $i++
  if ($i % 25 -eq 0 -or $i -eq $allGroups.Count) {
    Write-Host "  ...$i / $($allGroups.Count) groups resolved"
  }
  $groupNameById[$g.Id] = $g.DisplayName
  $rawMembers = Get-AzADGroupMember -GroupObjectId $g.Id -ErrorAction SilentlyContinue
  $groupMembersById[$g.Id] = @($rawMembers | ForEach-Object { Get-MemberUpnAndName $_ } | Where-Object { $_ })
}

$totalResolvedMembers = ($groupMembersById.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum
if ($totalResolvedMembers -eq 0 -and $allGroups.Count -gt 0) {
  Write-Warning "Zero group members resolved across $($allGroups.Count) groups -- Get-AzADGroupMember's return shape may not match what this script expects in this Az.Resources version, or these groups are genuinely all empty. Group-based role attribution and 'Group:' columns will be empty in the report. Verify with: Get-AzADGroupMember -GroupObjectId <id> | Format-List *"
}

# --- Per-subscription role assignment collection ---

function Get-SubscriptionAccessRows {
  param([string]$SubscriptionId, [string]$Label)

  # Write-Host, not Write-Output: this function's result is captured into a
  # variable by the caller ($prodRows = Get-SubscriptionAccessRows ...), and
  # PowerShell functions return everything written to the success stream --
  # Write-Output here would silently prepend these status strings onto the
  # actual row data, corrupting the Excel output. Write-Host bypasses that.
  Write-Host "Collecting role assignments for $Label ($SubscriptionId)..."
  Set-AzContext -SubscriptionId $SubscriptionId | Out-Null

  # No -Scope filter: returns every assignment across every resource in the
  # subscription, not just the subscription-root scope (confirmed behavior
  # during this project's earlier manual RBAC checks).
  $assignments = Get-AzRoleAssignment

  $userRows = @{}       # UPN -> ordered hashtable of columns
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
        if (-not $groupMembersById.ContainsKey($a.ObjectId)) {
          Write-Warning "Group $($a.DisplayName) ($($a.ObjectId)) holds role $($a.RoleDefinitionName) but wasn't in the tenant group index -- skipped (may have been deleted since the index was built)."
          continue
        }
        $members = $groupMembersById[$a.ObjectId]
        if ($members.Count -eq 0) {
          Write-Warning "Group $($a.DisplayName) ($($a.ObjectId)) holds role $($a.RoleDefinitionName) but resolved to zero members -- either genuinely empty, or Get-AzADGroupMember's return shape didn't match (see the warning printed during index building, if any)."
        }
        foreach ($m in $members) {
          $row = Get-OrCreateUserRow -Upn $m.UserPrincipalName -DisplayName $m.DisplayName
          $row["Role: $($a.RoleDefinitionName)"] = "Yes"
        }
      }
      default { $otherCount++ }
    }
  }

  Write-Host "  $($userRows.Count) distinct users, $otherCount service-principal/other assignments excluded from rows."

  # Group membership columns: pure in-memory lookup against the index built
  # above -- no additional Graph calls.
  foreach ($groupId in $groupMembersById.Keys) {
    foreach ($m in $groupMembersById[$groupId]) {
      if ($userRows.ContainsKey($m.UserPrincipalName)) {
        $userRows[$m.UserPrincipalName]["Group: $($groupNameById[$groupId])"] = "Yes"
      }
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
