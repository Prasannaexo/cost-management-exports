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
  - Get-AzADGroupMember returns a generic directory-object type; member
    objects may expose .UserPrincipalName/.DisplayName directly or only
    under .AdditionalProperties depending on Az.Resources version --
    Get-MemberUpnAndName below checks both.

PIM-ELIGIBLE GROUP MEMBERSHIP (confirmed against this tenant -- read this
before treating an empty admin-group result as "nobody has access"):
Get-AzADGroupMember only ever sees ACTIVE group membership. Groups like
BDECompute-Platform-SuperAdmins-Prod initially appeared to have zero
members entirely -- that was NOT a role-assignable-group permission
restriction (confirmed via direct Graph query: isAssignableToRole is
false for that group) and NOT a bug in Get-AzADGroupMember. It's PIM
"eligible" (just-in-time) group membership by design: nobody is a
standing member; people activate temporary membership when they need
elevated access, so a snapshot query genuinely finds nobody active most
of the time. This script now separately queries PIM eligibility
schedules (Microsoft Graph identityGovernance/privilegedAccess/group/
eligibilitySchedules) and merges eligible-but-inactive assignments into
the report as "Eligible" (vs "Active" for real, current access) --
see Get-GraphBearerToken and the PIM section below. This requires the
PrivilegedAccess.Read.AzureADGroup Graph permission; without it, this
degrades gracefully to active-membership-only (previous behavior); check
the runbook job's verbose logs for whether PIM data was actually available
on a given run.

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

Write-Verbose "Building tenant-wide group membership index (one pass over every Entra group)..." -Verbose
$allGroups = Get-AzADGroup
Write-Verbose "  $($allGroups.Count) groups found. Resolving members..." -Verbose

$groupMembersById = @{}  # GroupObjectId -> array of normalized {UserPrincipalName, DisplayName} objects, ACTIVE members
$groupNameById    = @{}  # GroupObjectId -> DisplayName
$i = 0
foreach ($g in $allGroups) {
  $i++
  if ($i % 25 -eq 0 -or $i -eq $allGroups.Count) {
    Write-Verbose "  ...$i / $($allGroups.Count) groups resolved" -Verbose
  }
  $groupNameById[$g.Id] = $g.DisplayName
  $rawMembers = Get-AzADGroupMember -GroupObjectId $g.Id -ErrorAction SilentlyContinue
  $groupMembersById[$g.Id] = @($rawMembers | ForEach-Object { Get-MemberUpnAndName $_ } | Where-Object { $_ })
}

$totalResolvedMembers = ($groupMembersById.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum
if ($totalResolvedMembers -eq 0 -and $allGroups.Count -gt 0) {
  Write-Warning "Zero group members resolved across $($allGroups.Count) groups -- Get-AzADGroupMember's return shape may not match what this script expects in this Az.Resources version, or these groups are genuinely all empty. Group-based role attribution and 'Group:' columns will be empty in the report. Verify with: Get-AzADGroupMember -GroupObjectId <id> | Format-List *"
}

# --- PIM-eligible group membership (separate from active membership above) ---
# Some groups -- confirmed on this tenant for BDECompute-Platform-SuperAdmins-Prod,
# which is NOT a role-assignable group (isAssignableToRole: false) -- have zero
# ACTIVE members by design: access is granted via PIM "eligible" assignments that
# only become real ("active") when a user activates them. Get-AzADGroupMember only
# ever sees active membership. This queries eligibility schedules separately so
# eligible-but-not-activated access still shows up in the report, clearly marked
# "Eligible" rather than "Active" so the two aren't confused.
# Requires the Microsoft Graph PrivilegedAccess.Read.AzureADGroup permission on
# whichever identity runs this. Without it, this 403s -- caught below, reported
# once, and the script continues using active-membership data only (previous
# behavior), so this is safe to leave in even before that permission is granted.

$groupEligibleMembersById = @{}  # GroupObjectId -> array of normalized {UserPrincipalName, DisplayName} objects
$pimAvailable = $true
$pimErrorMessage = $null

function Get-GraphBearerToken {
  $tokenObj = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com"
  if ($tokenObj.Token -is [securestring]) {
    return [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($tokenObj.Token))
  }
  return $tokenObj.Token
}

Write-Verbose "Checking PIM-eligible group membership (requires PrivilegedAccess.Read.AzureADGroup)..." -Verbose
try {
  $graphHeaders = @{ Authorization = "Bearer $(Get-GraphBearerToken)" }
  # No $expand=principal here -- that combination 400'd against this endpoint.
  # Resolve principals separately via Get-AzADUser instead, reusing the same
  # cmdlet already used elsewhere in this script rather than guessing at Graph
  # $expand support further.
  $uri = "https://graph.microsoft.com/v1.0/identityGovernance/privilegedAccess/group/eligibilitySchedules"
  $eligiblePrincipalIdsByGroup = @{}
  do {
    $resp = Invoke-RestMethod -Uri $uri -Headers $graphHeaders -ErrorAction Stop
    foreach ($sched in $resp.value) {
      if (-not $eligiblePrincipalIdsByGroup.ContainsKey($sched.groupId)) {
        $eligiblePrincipalIdsByGroup[$sched.groupId] = @()
      }
      $eligiblePrincipalIdsByGroup[$sched.groupId] += $sched.principalId
    }
    $uri = $resp.'@odata.nextLink'
  } while ($uri)

  # Resolve each unique principal once, not once per group it's eligible for.
  $allPrincipalIds = @($eligiblePrincipalIdsByGroup.Values | ForEach-Object { $_ } | Select-Object -Unique)
  $resolvedPrincipals = @{}  # principalId -> {UserPrincipalName, DisplayName} or $null if not a user (e.g. a nested group)
  foreach ($pid in $allPrincipalIds) {
    try {
      $u = Get-AzADUser -ObjectId $pid -ErrorAction Stop
      $resolvedPrincipals[$pid] = [pscustomobject]@{ UserPrincipalName = $u.UserPrincipalName; DisplayName = $u.DisplayName }
    } catch {
      $resolvedPrincipals[$pid] = $null  # not a user (group/service principal) -- not expanded further, same as active-membership limitation
    }
  }

  foreach ($groupId in $eligiblePrincipalIdsByGroup.Keys) {
    $resolved = @($eligiblePrincipalIdsByGroup[$groupId] | ForEach-Object { $resolvedPrincipals[$_] } | Where-Object { $_ })
    $groupEligibleMembersById[$groupId] = $resolved
  }
  $eligibleCount = ($groupEligibleMembersById.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum
  Write-Verbose "  PIM eligibility data available -- $eligibleCount eligible user assignments found across $($groupEligibleMembersById.Keys.Count) groups." -Verbose
} catch {
  $pimAvailable = $false
  $pimErrorMessage = $_.Exception.Message
  Write-Warning "PIM eligibility data not available ($pimErrorMessage) -- continuing with active-membership data only. Grant PrivilegedAccess.Read.AzureADGroup to close this gap; see the README in this folder."
}

# --- Per-subscription role assignment collection ---

function Get-SubscriptionAccessRows {
  param([string]$SubscriptionId, [string]$Label)

  # Write-Verbose (with -Verbose to force it on regardless of caller
  # preference), not Write-Output/Write-Host: this function's result is
  # captured into a variable by the caller ($prodRows = Get-SubscriptionAccessRows
  # ...), and PowerShell functions return everything written to the success
  # stream -- Write-Output here would silently prepend these status strings
  # onto the actual row data, corrupting the Excel output. Write-Host avoids
  # that but Azure Automation doesn't capture the Information stream it uses
  # at all (confirmed: none of these messages appeared in the job's logs) --
  # Verbose is captured, provided the runbook has logVerbose enabled (see
  # New-AccessReviewAutomation.ps1 / Set-AzAutomationRunbook -LogVerbose $true).
  Write-Verbose "Collecting role assignments for $Label ($SubscriptionId)..." -Verbose
  $context = Set-AzContext -SubscriptionId $SubscriptionId
  if (-not $context -or $context.Subscription.Id -ne $SubscriptionId) {
    throw "Set-AzContext did not switch to subscription $SubscriptionId for $Label -- got '$($context.Subscription.Id)' instead. Aborting rather than silently reporting the wrong (or no) subscription's data."
  }

  # No -Scope filter: returns every assignment across every resource in the
  # subscription, not just the subscription-root scope (confirmed behavior
  # during this project's earlier manual RBAC checks).
  try {
    $assignments = @(Get-AzRoleAssignment -ErrorAction Stop)
  } catch {
    throw "Get-AzRoleAssignment failed for $Label ($SubscriptionId): $_"
  }
  Write-Verbose "  Get-AzRoleAssignment returned $($assignments.Count) assignment(s) for $Label." -Verbose
  if ($assignments.Count -eq 0) {
    Write-Warning "Zero role assignments returned for $Label ($SubscriptionId). If this subscription has any RBAC assignments at all (nearly all do -- Owner/Contributor/etc. at minimum), this points to a permission or context problem for the identity running this script, not a genuinely empty subscription."
  }
  Write-Verbose "  Distinct role names in raw assignments for $Label`: $(($assignments | Select-Object -ExpandProperty RoleDefinitionName -Unique) -join ', ')" -Verbose
  $knownUserAssignments = @($assignments | Where-Object { $_.ObjectId -eq "10365227-109f-4439-953c-41c6d58bfeaa" })
  if ($knownUserAssignments.Count -gt 0) {
    Write-Verbose "  Raw assignments for known-user ObjectId 10365227... in $Label`: $(($knownUserAssignments | ForEach-Object { $_.RoleDefinitionName }) -join ' | ')" -Verbose
  }

  $userRows = @{}       # UPN -> ordered hashtable of columns
  $otherCount = 0
  $otherSamples = @()   # up to 5 full property dumps of unmatched assignments, for diagnosing classification misses

  function Get-OrCreateUserRow {
    param([string]$Upn, [string]$DisplayName)
    if (-not $userRows.ContainsKey($Upn)) {
      $userRows[$Upn] = [ordered]@{ User = $DisplayName; UPN = $Upn }
    }
    return $userRows[$Upn]
  }

  # Neither ObjectType nor SignInName can be trusted here: confirmed via a
  # real run that both come back empty/"Unknown" for every assignment when
  # this script runs as the Automation Account's managed identity --
  # including for a role assignment we independently confirmed belongs to
  # a real user (SignInName was blank for it too, not just for genuine
  # service principals). Classify by explicit lookup instead: an ObjectId
  # matching the tenant group index built earlier is a group; anything else
  # gets a one-time (cached) Get-AzADUser lookup to determine if it's a
  # real user or a genuine service principal/managed identity to exclude.
  $userLookupCache = @{}  # ObjectId -> resolved {UserPrincipalName, DisplayName} or $null if not a user

  foreach ($a in $assignments) {
    if (-not $userLookupCache.ContainsKey($a.ObjectId) -and -not $groupNameById.ContainsKey($a.ObjectId)) {
      try {
        $u = Get-AzADUser -ObjectId $a.ObjectId -ErrorAction Stop
        $userLookupCache[$a.ObjectId] = [pscustomobject]@{ UserPrincipalName = $u.UserPrincipalName; DisplayName = $u.DisplayName }
      } catch {
        $userLookupCache[$a.ObjectId] = $null  # genuinely not a user -- service principal/managed identity
      }
    }

    if ($userLookupCache.ContainsKey($a.ObjectId) -and $userLookupCache[$a.ObjectId]) {
      # Get-AzRoleAssignment only ever returns ACTIVE assignments -- a
      # directly-assigned role that's PIM-eligible-but-not-activated
      # wouldn't appear here at all (a separate, narrower gap than the
      # group-eligibility one this script now covers; not yet handled).
      $u = $userLookupCache[$a.ObjectId]
      $row = Get-OrCreateUserRow -Upn $u.UserPrincipalName -DisplayName $u.DisplayName
      $row["Role: $($a.RoleDefinitionName)"] = "Active"
    } elseif ($groupNameById.ContainsKey($a.ObjectId)) {
      $activeMembers = $groupMembersById[$a.ObjectId]
      $eligibleMembers = if ($groupEligibleMembersById.ContainsKey($a.ObjectId)) { $groupEligibleMembersById[$a.ObjectId] } else { @() }
      if ($activeMembers.Count -eq 0 -and $eligibleMembers.Count -eq 0) {
        Write-Warning "Group $($groupNameById[$a.ObjectId]) ($($a.ObjectId)) holds role $($a.RoleDefinitionName) but resolved to zero active or eligible members."
      }
      foreach ($m in $activeMembers) {
        $row = Get-OrCreateUserRow -Upn $m.UserPrincipalName -DisplayName $m.DisplayName
        $row["Role: $($a.RoleDefinitionName)"] = "Active"
      }
      foreach ($m in $eligibleMembers) {
        $row = Get-OrCreateUserRow -Upn $m.UserPrincipalName -DisplayName $m.DisplayName
        # Don't downgrade a cell that's already "Active" via some other path
        # (e.g. direct assignment, or active in a different group with the same role).
        if ($row["Role: $($a.RoleDefinitionName)"] -ne "Active") {
          $row["Role: $($a.RoleDefinitionName)"] = "Eligible"
        }
      }
    } else {
      # Neither a resolvable user nor a group in our tenant-wide index --
      # genuinely a service principal/managed identity, or a group that was
      # deleted since the index was built at the top of this script. Sample
      # the first few so we can see exactly what's actually in these
      # objects if this count looks too high -- ObjectType/ObjectId have
      # both turned out to be unreliable in this identity context before.
      $otherCount++
      if ($otherSamples.Count -lt 5) {
        $otherSamples += ($a.PSObject.Properties | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join "; "
      }
    }
  }

  Write-Verbose "  $($userRows.Count) distinct users, $otherCount service-principal/other assignments excluded from rows." -Verbose
  if ($otherCount -gt 0) {
    Write-Verbose "  Sample of excluded assignments (up to 5), full property dump:" -Verbose
    $otherSamples | ForEach-Object { Write-Verbose "    $_" -Verbose }
  }
  $knownUserUpn = ($userLookupCache.Values | Where-Object { $_ -and $_.UserPrincipalName -like "pkbiswal@*" } | Select-Object -First 1).UserPrincipalName
  if ($knownUserUpn -and $userRows.ContainsKey($knownUserUpn)) {
    $rowDump = ($userRows[$knownUserUpn].Keys | ForEach-Object { "$_=$($userRows[$knownUserUpn][$_])" }) -join "; "
    Write-Verbose "  Final row contents for $knownUserUpn`: $rowDump" -Verbose
  }

  # Group membership columns: pure in-memory lookup against the indexes built
  # above -- no additional Graph calls. Active membership takes precedence
  # over eligible if a user is somehow both (shouldn't normally happen).
  foreach ($groupId in $groupMembersById.Keys) {
    foreach ($m in $groupMembersById[$groupId]) {
      if ($userRows.ContainsKey($m.UserPrincipalName)) {
        $userRows[$m.UserPrincipalName]["Group: $($groupNameById[$groupId])"] = "Active"
      }
    }
  }
  foreach ($groupId in $groupEligibleMembersById.Keys) {
    $groupLabel = if ($groupNameById.ContainsKey($groupId)) { $groupNameById[$groupId] } else { $groupId }
    foreach ($m in $groupEligibleMembersById[$groupId]) {
      if ($userRows.ContainsKey($m.UserPrincipalName) -and $userRows[$m.UserPrincipalName]["Group: $groupLabel"] -ne "Active") {
        $userRows[$m.UserPrincipalName]["Group: $groupLabel"] = "Eligible"
      }
    }
  }

  # Unary comma is load-bearing: without it, "return @(emptyOrOneItemCollection)"
  # gets unrolled by PowerShell's pipeline output semantics, so a genuinely
  # empty result becomes $null (not @()) once captured by the caller, and a
  # single-item result becomes that bare item instead of a 1-element array.
  # Confirmed to matter here: a run that found zero assignments produced
  # $null instead of an empty array, which downstream became a 1-element
  # array containing $null when re-wrapped with @(...), which Export-Excel
  # then rendered as garbage placeholder rows instead of an empty sheet.
  return ,@($userRows.Values | ForEach-Object { [pscustomobject]$_ })
}

function Add-SubscriptionLabel {
  param([object[]]$Rows, [string]$Label)
  return ,@($Rows | ForEach-Object {
    $h = [ordered]@{ Subscription = $Label }
    foreach ($p in $_.PSObject.Properties) { $h[$p.Name] = $p.Value }
    [pscustomobject]$h
  })
}

$prodRows = Get-SubscriptionAccessRows -SubscriptionId $ProductionSubscriptionId -Label "Production"
$testRows = Get-SubscriptionAccessRows -SubscriptionId $TestSubscriptionId -Label "Test"
$devRows  = Get-SubscriptionAccessRows -SubscriptionId $DevSubscriptionId -Label "Dev"

$allRows = (Add-SubscriptionLabel -Rows $prodRows -Label "Production") +
           (Add-SubscriptionLabel -Rows $testRows -Label "Test") +
           (Add-SubscriptionLabel -Rows $devRows -Label "Dev")

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

function ConvertTo-UniformRows {
  # Export-Excel derives its column headers from the FIRST object in the
  # piped collection. Since each user's set of Role:/Group: properties
  # depends on what they individually hold, later objects routinely have
  # properties the first one doesn't -- and Export-Excel silently drops
  # those from the output. Confirmed the hard way: a user's in-memory row
  # had 28 roles, only 3 made it into the exported sheet, because whichever
  # user happened to be first in hashtable enumeration order only held 3.
  # Normalize every object to the same full property set (the union across
  # all of them) before anything gets exported.
  param([object[]]$Rows)
  if ($Rows.Count -eq 0) { return ,@() }
  $allKeys = [ordered]@{}
  foreach ($r in $Rows) {
    foreach ($p in $r.PSObject.Properties) { $allKeys[$p.Name] = $true }
  }
  return ,@($Rows | ForEach-Object {
    $src = $_
    $h = [ordered]@{}
    foreach ($k in $allKeys.Keys) {
      $h[$k] = if ($src.PSObject.Properties.Name -contains $k) { $src.$k } else { $null }
    }
    [pscustomobject]$h
  })
}

Write-Output "Writing workbook to $OutputPath..."
Remove-Item -Path $OutputPath -ErrorAction SilentlyContinue
(ConvertTo-UniformRows $prodRows) | Export-Excel -Path $OutputPath -WorksheetName "Production" -AutoSize -FreezeTopRow -BoldTopRow
(ConvertTo-UniformRows $testRows) | Export-Excel -Path $OutputPath -WorksheetName "Test" -AutoSize -FreezeTopRow -BoldTopRow
(ConvertTo-UniformRows $devRows)  | Export-Excel -Path $OutputPath -WorksheetName "Dev" -AutoSize -FreezeTopRow -BoldTopRow
(ConvertTo-UniformRows $allRows)  | Export-Excel -Path $OutputPath -WorksheetName "All Subscriptions" -AutoSize -FreezeTopRow -BoldTopRow
(ConvertTo-UniformRows $roleDefRows) | Export-Excel -Path $OutputPath -WorksheetName "Role Definitions" -AutoSize -FreezeTopRow -BoldTopRow

function Set-StorageFirewallDefaultAction {
  # Automation Accounts aren't on Microsoft's supported list for storage
  # resource instance rules, and Automation cloud jobs architecturally
  # cannot reach private-endpoint-secured resources at all -- both confirmed
  # against Microsoft docs. Toggling defaultAction is the only lever left,
  # so the account sits at Deny by default (see New-CostExportStorage.ps1)
  # and this narrows the exposure window to just the upload below instead
  # of leaving the firewall open indefinitely.
  param([string]$StorageAccountResourceId, [ValidateSet("Allow", "Deny")][string]$DefaultAction)
  $path = "$($StorageAccountResourceId)?api-version=2023-01-01"
  $current = Invoke-AzRestMethod -Method GET -Path $path
  if ($current.StatusCode -ge 300) { throw "Failed to read storage account network config: $($current.Content)" }
  $account = $current.Content | ConvertFrom-Json
  $account.properties.networkAcls.defaultAction = $DefaultAction
  $body = @{ properties = @{ networkAcls = $account.properties.networkAcls } } | ConvertTo-Json -Depth 10
  $update = Invoke-AzRestMethod -Method PATCH -Path $path -Payload $body
  if ($update.StatusCode -ge 300) { throw "Failed to set storage firewall defaultAction=$DefaultAction : $($update.Content)" }
}

Write-Output "Uploading to storage..."
$storageAccountName = ($StorageAccountResourceId -split "/")[-1]
$blobName = "{0:yyyyMM}/azure-access-review-{0:yyyyMMdd}.xlsx" -f (Get-Date)

Write-Verbose "Opening storage firewall for upload..." -Verbose
Set-StorageFirewallDefaultAction -StorageAccountResourceId $StorageAccountResourceId -DefaultAction "Allow"
try {
  $maxAttempts = 6
  for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
    try {
      $ctx = New-AzStorageContext -StorageAccountName $storageAccountName -UseConnectedAccount
      $container = Get-AzStorageContainer -Name $ContainerName -Context $ctx -ErrorAction SilentlyContinue
      if (-not $container) {
        New-AzStorageContainer -Name $ContainerName -Context $ctx -Permission Off | Out-Null
      }
      Set-AzStorageBlobContent -File $OutputPath -Container $ContainerName -Blob $blobName -Context $ctx -Force | Out-Null
      break
    } catch {
      if ($attempt -eq $maxAttempts) { throw }
      Write-Verbose "Upload attempt $attempt failed (firewall rule likely not yet propagated): $($_.Exception.Message). Retrying in 15s..." -Verbose
      Start-Sleep -Seconds 15
    }
  }
} finally {
  Write-Verbose "Closing storage firewall..." -Verbose
  Set-StorageFirewallDefaultAction -StorageAccountResourceId $StorageAccountResourceId -DefaultAction "Deny"
}

Write-Output "Done. Uploaded to $ContainerName/$blobName"
