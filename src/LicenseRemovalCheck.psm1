#Requires -Version 7.0
<#
  LicenseRemovalCheck
  A read-only check that runs before an offboarding workflow removes groups and licenses
  from a Microsoft 365 user whose mailbox has been converted to a shared mailbox.

  It reads one JSON snapshot (Graph user + Exchange mailbox data) and the full planned
  removal batch, and returns ALLOW, BLOCK or REVIEW with reasons and sources.
  It connects to nothing, writes nothing, holds no secret and calls no AI model.
  When a piece of data it needs is missing or unreadable, it does not decide: REVIEW.
#>
Set-StrictMode -Version Latest

# 50 GB, the shared mailbox size limit without a license, and 100 GB with Exchange Online (Plan 2).
# Exchange sizes are binary: Set-Mailbox documents "1.999999999 terabytes (2199023254528 bytes)",
# so 1 GB = 1,073,741,824 bytes, 50 GB = 53,687,091,200 bytes, 100 GB = 107,374,182,400 bytes.
$script:SharedMailboxLimitBytes = [int64]53687091200
$script:Plan2CeilingBytes = [int64]107374182400
$script:Inv = [System.Globalization.CultureInfo]::InvariantCulture

$script:Sources = @{
    SharedMailboxLimit = 'https://learn.microsoft.com/en-us/microsoft-365/admin/email/about-shared-mailboxes'
    ConvertMailbox     = 'https://learn.microsoft.com/en-us/microsoft-365/admin/email/convert-user-mailbox-to-shared-mailbox'
    GroupLicensing     = 'https://learn.microsoft.com/en-us/microsoft-365/admin/manage/manage-group-licenses'
    LicenseState       = 'https://learn.microsoft.com/en-us/graph/api/resources/licenseassignmentstate?view=graph-rest-1.0'
    UserResource       = 'https://learn.microsoft.com/en-us/graph/api/resources/user?view=graph-rest-1.0'
    HoldTypes          = 'https://learn.microsoft.com/en-us/purview/edisc-hold-types-mailboxes'
    GetExoMailbox      = 'https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/get-exomailbox'
    MailboxStats       = 'https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/get-exomailboxstatistics'
    ArchiveProperties  = 'https://learn.microsoft.com/en-us/troubleshoot/microsoft-365/purview/archive-mailboxes/archivestatus-set-none'
    OverQuotaEffect    = 'https://techcommunity.microsoft.com/blog/exchange/correcting-shared-mailbox-provisioning-and-sizing/607991'
}

$script:Reference = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot '..' 'data' 'licensing-reference.json') | ConvertFrom-Json -AsHashtable
$script:Plan2ServicePlanId = $script:Reference.service_plan.id
$script:Plan2SkuIds = @($script:Reference.skus_with_plan | ForEach-Object { $_.skuId.ToLowerInvariant() })

function Get-SnapshotValue {
    # Returns $null when any segment of the path is missing, instead of throwing under StrictMode.
    param($Object, [string[]]$Path)
    $cur = $Object
    foreach ($p in $Path) {
        if ($null -eq $cur) { return $null }
        if ($cur -is [System.Collections.IDictionary]) {
            if (-not $cur.Contains($p)) { return $null }
            $cur = $cur[$p]
        } else { return $null }
    }
    return , $cur
}

function Test-HasKey {
    param($Object, [string]$Key)
    return ($Object -is [System.Collections.IDictionary]) -and $Object.Contains($Key)
}

function Test-IsList {
    param($Value)
    return ($null -ne $Value) -and ($Value -is [System.Collections.IList])
}

function Get-ProductName {
    param([string]$SkuId)
    $n = $script:Reference.product_names[$SkuId.ToLowerInvariant()]
    if ($n) { return $n } else { return "SKU $SkuId" }
}

function Get-ByteCount {
    # The snapshot contract carries the primary mailbox size as a whole number of bytes.
    # Anything else (text, negative, fractional) is treated as unreadable.
    param([AllowNull()]$Value)
    if ($Value -is [int] -or $Value -is [long]) {
        if ($Value -ge 0) { return [int64]$Value }
    }
    return $null
}

function Format-Size {
    # Plain GB far from the 50 GB line; exact bytes within 1 GB of it, so a reason never
    # reads "over its 50 GB quota at 50 GB".
    param([int64]$Bytes)
    if ([math]::Abs($Bytes - $script:SharedMailboxLimitBytes) -lt 1GB) {
        return [string]::Format($script:Inv, '{0:N0} bytes', $Bytes)
    }
    return ([math]::Round($Bytes / 1GB, 1)).ToString($script:Inv) + ' GB'
}

function Get-HoldKind {
    # Meaning of an InPlaceHolds entry from its prefix, as written on the Microsoft Purview page
    # "Identify Exchange mailbox hold types in eDiscovery" (read 2026-11-09).
    # Returns $null when that page gives no meaning for the prefix.
    param([string]$Entry)
    switch -Regex ($Entry) {
        '^-mbx' { return 'exclusion from an organization-wide retention policy' }
        '^UniH' { return 'eDiscovery hold' }
        '^mbx'  { return 'retention policy' }
        '^skp'  { return 'retention policy applied to Skype for Business conversations' }
        '^grp'  { return 'organization-wide retention policy applied to Microsoft 365 groups and Teams channel messages' }
        default { return $null }
    }
}

function New-Reason {
    param([string]$Rule, [string]$Decision, [string]$Text, [string]$Source, [string]$MissingData, [System.Collections.IDictionary]$Extra)
    $r = [ordered]@{ rule = $Rule; decision = $Decision; text = $Text }
    if ($MissingData) { $r.missing_data = $MissingData }
    if ($Extra) { foreach ($k in $Extra.Keys) { $r[$k] = $Extra[$k] } }
    $r.source = $Source
    return $r
}

function Test-OffboardingRemovalSafety {
    <#
    .SYNOPSIS
      Decides whether a planned batch of group and license removals is safe for a user
      whose mailbox was converted to a shared mailbox. Read-only.
    .PARAMETER AadUserId
      Entra object ID of the departing user (Rewst-style input name).
    .PARAMETER TenantId
      Tenant ID the workflow is running against.
    .PARAMETER SnapshotPath
      Path to a JSON snapshot: graph, exchange, planned_removals, conversion_requested.
    .PARAMETER Snapshot
      The same snapshot, already loaded as a hashtable.
    .OUTPUTS
      Ordered dictionary: inputs, success, status, decision, reasons, capacity_kept_by,
      checked_batch, automation_log, writes.
    #>
    [CmdletBinding(DefaultParameterSetName = 'Path')]
    param(
        [Parameter(Mandatory)][string]$AadUserId,
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory, ParameterSetName = 'Path')][string]$SnapshotPath,
        [Parameter(Mandatory, ParameterSetName = 'Object')][hashtable]$Snapshot
    )

    $log = [System.Collections.Generic.List[string]]::new()
    $reasons = [System.Collections.Generic.List[object]]::new()
    $readWarnings = 0

    $out = [ordered]@{
        inputs           = [ordered]@{ aad_user_id = $AadUserId; tenant_id = $TenantId }
        success          = $false
        status           = 2000
        decision         = 'REVIEW'
        reasons          = $null
        capacity_kept_by = @()
        checked_batch    = $null
        automation_log   = $null
        writes           = 0
        note             = 'success and status describe the data collection. Only decision says whether the batch may run.'
    }

    function Complete-Output {
        $out.reasons = @($reasons)
        $out.automation_log = ($log -join "`n")
        return $out
    }
    function Stop-Collection([string]$Rule, [string]$Text, [string]$Missing, [string]$Source) {
        $log.Add("Collection failed: $Text")
        $reasons.Add((New-Reason -Rule $Rule -Decision 'REVIEW' -Text "Human decision required: $Text" -MissingData $Missing -Source $Source))
        return Complete-Output
    }

    # 1. Load the snapshot.
    if ($PSCmdlet.ParameterSetName -eq 'Path') {
        if (-not (Test-Path -LiteralPath $SnapshotPath)) {
            return Stop-Collection 'snapshot-unreadable' 'the snapshot file was not found.' 'snapshot file' $script:Sources.GetExoMailbox
        }
        $parsed = $null
        try { $parsed = Get-Content -Raw -LiteralPath $SnapshotPath | ConvertFrom-Json -AsHashtable -ErrorAction Stop }
        catch { return Stop-Collection 'snapshot-unreadable' 'the snapshot is not valid JSON.' 'snapshot file' $script:Sources.GetExoMailbox }
        if ($parsed -isnot [System.Collections.IDictionary]) {
            return Stop-Collection 'snapshot-unreadable' 'the snapshot is not a JSON object.' 'snapshot file' $script:Sources.GetExoMailbox
        }
        $Snapshot = $parsed
    }
    $log.Add("Read-only check started for aad_user_id $AadUserId in tenant $TenantId.")

    $user    = Get-SnapshotValue $Snapshot 'graph', 'user'
    $mailbox = Get-SnapshotValue $Snapshot 'exchange', 'mailbox'
    if ($user -isnot [System.Collections.IDictionary]) {
        return Stop-Collection 'collection-failed' 'graph.user is missing from the snapshot.' 'graph.user' $script:Sources.UserResource
    }
    if ($mailbox -isnot [System.Collections.IDictionary]) {
        return Stop-Collection 'collection-failed' 'exchange.mailbox is missing from the snapshot.' 'exchange.mailbox' $script:Sources.GetExoMailbox
    }
    $snapUserId = [string](Get-SnapshotValue $user 'id')
    if ($snapUserId -ne $AadUserId) {
        return Stop-Collection 'user-mismatch' "the Graph data describes user $snapUserId, not $AadUserId." 'graph.user.id' $script:Sources.UserResource
    }

    # 2. The mailbox must be this user's mailbox.
    $owner = [string](Get-SnapshotValue $mailbox 'ExternalDirectoryObjectId')
    if (-not $owner) {
        return Stop-Collection 'mailbox-owner-unknown' 'the mailbox ExternalDirectoryObjectId was not read, so the mailbox cannot be tied to the user.' 'mailbox ExternalDirectoryObjectId' $script:Sources.GetExoMailbox
    }
    if ($owner -ne $AadUserId) {
        return Stop-Collection 'mailbox-owner-mismatch' "the mailbox belongs to $owner, not $AadUserId." 'mailbox of this user' $script:Sources.GetExoMailbox
    }
    $exGuid = [string](Get-SnapshotValue $mailbox 'ExchangeGuid')
    $statGuid = [string](Get-SnapshotValue $Snapshot 'exchange', 'mailbox_statistics', 'MailboxGuid')
    if ($exGuid -and $statGuid -and ($exGuid -ne $statGuid)) {
        return Stop-Collection 'statistics-mismatch' "the size statistics describe mailbox $statGuid, not $exGuid." 'statistics of this mailbox' $script:Sources.MailboxStats
    }

    # 3. Tenant identity on each block of data.
    $graphTenant = [string](Get-SnapshotValue $Snapshot 'graph', 'tenant_id')
    $exoTenant   = [string](Get-SnapshotValue $Snapshot 'exchange', 'tenant_id')
    if ($graphTenant -ne $TenantId -or $exoTenant -ne $TenantId) {
        $reasons.Add((New-Reason -Rule 'tenant-mismatch' -Decision 'REVIEW' -Text "Human decision required: tenant_id differs between the request ($TenantId), Graph data ($graphTenant) and Exchange data ($exoTenant)." -MissingData 'one tenant_id for all data' -Source $script:Sources.UserResource))
        $log.Add('Tenant IDs differ between request, Graph and Exchange data.')
    }

    # 4. Hybrid identities are out of scope. Absent is not the same as false.
    if (-not (Test-HasKey $user 'onPremisesSyncEnabled')) {
        $readWarnings++
        $reasons.Add((New-Reason -Rule 'identity-unknown' -Decision 'REVIEW' -Text 'Human decision required: onPremisesSyncEnabled was not read, so a synced identity cannot be ruled out.' -MissingData 'Graph user onPremisesSyncEnabled' -Source $script:Sources.UserResource))
    } else {
        $sync = $user['onPremisesSyncEnabled']
        if ($sync -is [bool] -and $sync) {
            $reasons.Add((New-Reason -Rule 'identity-synced' -Decision 'REVIEW' -Text 'Human decision required: identity synchronized from on-premises Active Directory, out of scope for this check.' -MissingData 'cloud-only identity' -Source $script:Sources.UserResource))
            $log.Add('onPremisesSyncEnabled is true: hybrid identity, out of scope.')
        } elseif ($null -ne $sync -and $sync -isnot [bool]) {
            $readWarnings++
            $reasons.Add((New-Reason -Rule 'identity-unknown' -Decision 'REVIEW' -Text "Human decision required: onPremisesSyncEnabled is '$sync', not true, false or null." -MissingData 'Graph user onPremisesSyncEnabled' -Source $script:Sources.UserResource))
        }
    }

    # 5. Mailbox type read back from Exchange. Only a SharedMailbox can be allowed.
    $type = [string](Get-SnapshotValue $mailbox 'RecipientTypeDetails')
    $conversionRequested = (Get-SnapshotValue $Snapshot 'conversion_requested') -eq $true
    if (-not $type) {
        $readWarnings++
        $reasons.Add((New-Reason -Rule 'type-unknown' -Decision 'REVIEW' -Text 'Human decision required: the mailbox type was not read.' -MissingData 'Get-Mailbox RecipientTypeDetails' -Source $script:Sources.ConvertMailbox))
    } elseif ($type -ne 'SharedMailbox') {
        if ($conversionRequested -and $type -eq 'UserMailbox') {
            $reasons.Add((New-Reason -Rule 'not-yet-shared' -Decision 'REVIEW' -Text "Human decision required: conversion to shared was requested but the mailbox reads back as $type." -MissingData 'RecipientTypeDetails = SharedMailbox' -Source $script:Sources.ConvertMailbox))
        } else {
            $reasons.Add((New-Reason -Rule 'not-shared' -Decision 'REVIEW' -Text "Human decision required: the mailbox reads back as $type, not SharedMailbox. This check only approves batches for converted shared mailboxes." -MissingData 'RecipientTypeDetails = SharedMailbox' -Source $script:Sources.ConvertMailbox))
        }
    }
    $log.Add("Mailbox type read back: $(if ($type) { $type } else { 'not read' }).")

    # 6. Mailbox size, in bytes.
    $rawSize = Get-SnapshotValue $Snapshot 'exchange', 'mailbox_statistics', 'TotalItemSize'
    $sizeBytes = Get-ByteCount $rawSize
    if ($null -eq $sizeBytes) {
        $readWarnings++
        $why = if ($null -eq $rawSize) { 'missing' } else { 'unreadable (the snapshot contract expects a whole number of bytes)' }
        $reasons.Add((New-Reason -Rule 'size-unreadable' -Decision 'REVIEW' -Text "Human decision required: mailbox size $why." -MissingData 'Get-EXOMailboxStatistics TotalItemSize, in bytes' -Source $script:Sources.MailboxStats))
        $log.Add("Mailbox size $why.")
    } else {
        $log.Add([string]::Format($script:Inv, 'Primary mailbox size: {0:N0} bytes (limit without license {1:N0}).', $sizeBytes, $script:SharedMailboxLimitBytes))
    }

    # 7. Archive. Microsoft says to read ArchiveGuid and ArchiveDatabase, not ArchiveStatus.
    $zeroGuid = '00000000-0000-0000-0000-000000000000'
    if (-not (Test-HasKey $mailbox 'ArchiveGuid') -or -not (Test-HasKey $mailbox 'ArchiveDatabase') -or $null -eq $mailbox['ArchiveGuid']) {
        $readWarnings++
        $reasons.Add((New-Reason -Rule 'archive-unknown' -Decision 'REVIEW' -Text 'Human decision required: the archive fields were not read.' -MissingData 'Get-Mailbox ArchiveGuid and ArchiveDatabase' -Source $script:Sources.ArchiveProperties))
    } else {
        $hasGuid = ([string]$mailbox['ArchiveGuid']) -and ([string]$mailbox['ArchiveGuid'] -ne $zeroGuid)
        $hasDb = [bool]([string]$mailbox['ArchiveDatabase'])
        if ($hasGuid -and $hasDb) {
            $reasons.Add((New-Reason -Rule 'archive-present' -Decision 'REVIEW' -Text 'Human decision required: an online archive is present. Microsoft: "To increase the size of the archive mailbox, an Exchange Online Plan 2 license or an Exchange Online Plan 1 license with an Exchange Online Archiving add-on license is required."' -MissingData 'license decision for the archive' -Source $script:Sources.SharedMailboxLimit))
        } elseif ($hasGuid -or $hasDb) {
            $readWarnings++
            $reasons.Add((New-Reason -Rule 'archive-inconsistent' -Decision 'REVIEW' -Text 'Human decision required: archive state inconsistent (only one of ArchiveGuid and ArchiveDatabase is set).' -MissingData 'consistent ArchiveGuid and ArchiveDatabase' -Source $script:Sources.ArchiveProperties))
        }
    }

    # 8. Holds on the mailbox. A flag that is missing, null or not a boolean counts as not read.
    foreach ($flag in 'LitigationHoldEnabled', 'ComplianceTagHoldApplied', 'DelayHoldApplied', 'DelayReleaseHoldApplied') {
        $v = if (Test-HasKey $mailbox $flag) { $mailbox[$flag] } else { $null }
        if ($v -isnot [bool]) {
            $readWarnings++
            $reasons.Add((New-Reason -Rule 'hold-unknown' -Decision 'REVIEW' -Text "Human decision required: $flag was not read." -MissingData "Get-Mailbox $flag" -Source $script:Sources.HoldTypes))
        } elseif ($v) {
            $reasons.Add((New-Reason -Rule 'hold-present' -Decision 'REVIEW' -Text "Human decision required: $flag is true on the mailbox, so the mailbox is on hold." -MissingData 'hold decision by whoever owns the hold' -Source $script:Sources.HoldTypes -Extra ([ordered]@{ hold_scope = 'mailbox'; hold_property = $flag })))
        }
    }
    $exclusions = @()
    if (-not (Test-HasKey $mailbox 'InPlaceHolds') -or $null -eq $mailbox['InPlaceHolds']) {
        $readWarnings++
        $reasons.Add((New-Reason -Rule 'hold-unknown' -Decision 'REVIEW' -Text 'Human decision required: mailbox InPlaceHolds was not read.' -MissingData 'Get-Mailbox InPlaceHolds' -Source $script:Sources.HoldTypes))
    } else {
        foreach ($h in @($mailbox['InPlaceHolds'] | Where-Object { $_ })) {
            if ($h -like '-mbx*') { $exclusions += [string]$h; $log.Add("Mailbox InPlaceHolds lists exclusion $h."); continue }
            $kind = Get-HoldKind $h
            $extra = [ordered]@{ hold_scope = 'mailbox'; hold_entry = $h }
            if ($kind) { $extra.hold_type = $kind }
            $reasons.Add((New-Reason -Rule 'hold-present' -Decision 'REVIEW' -Text "Human decision required: the mailbox InPlaceHolds lists $h$(if ($kind) { ", a $kind" })." -MissingData 'hold decision by whoever owns the hold' -Source $script:Sources.HoldTypes -Extra $extra))
        }
    }

    # 9. Organization-wide retention: an empty mailbox InPlaceHolds does not rule it out.
    $org = Get-SnapshotValue $Snapshot 'exchange', 'organization_config'
    if ($null -eq $org -or -not (Test-HasKey $org 'InPlaceHolds') -or $null -eq $org['InPlaceHolds']) {
        $readWarnings++
        $reasons.Add((New-Reason -Rule 'org-retention-unknown' -Decision 'REVIEW' -Text 'Human decision required: the organization''s retention settings could not be read, so the check does not approve the batch.' -MissingData 'Get-OrganizationConfig InPlaceHolds' -Source $script:Sources.HoldTypes))
        $log.Add('Organization-wide retention not readable.')
    } else {
        foreach ($h in @($org['InPlaceHolds'] | Where-Object { $_ })) {
            $kind = Get-HoldKind $h
            $extra = [ordered]@{ hold_scope = 'organization'; hold_entry = $h }
            if ($kind) { $extra.hold_type = $kind }
            $text = "Human decision required: the organization InPlaceHolds lists $h$(if ($kind) { ", a $kind" })."
            $guid = ([string]$h -replace '^(mbx|grp|skp)', '') -replace ':\d+$', ''
            $match = @($exclusions | Where-Object { $_ -eq "-mbx$guid" })
            if ($match.Count -gt 0) {
                $text += " The mailbox InPlaceHolds also lists $($match[0])."
                $extra.mailbox_exclusion = $match[0]
            }
            $reasons.Add((New-Reason -Rule 'hold-present' -Decision 'REVIEW' -Text $text -MissingData 'retention policy scope and exclusions' -Source $script:Sources.HoldTypes -Extra $extra))
        }
    }
    $holdFound = [bool]@($reasons | Where-Object { $_.rule -eq 'hold-present' }).Count

    # 10. The removal batch. Missing lists are not empty lists.
    $groupNames = @{}
    $groupList = Get-SnapshotValue $Snapshot 'graph', 'groups'
    foreach ($g in @($groupList)) {
        if ($g -is [System.Collections.IDictionary] -and $g.Contains('id') -and $g['id']) {
            $gid = ([string]$g['id']).ToLowerInvariant()
            $groupNames[$gid] = if ($g.Contains('displayName') -and $g['displayName']) { [string]$g['displayName'] } else { $null }
        }
    }
    function Get-GroupLabel([string]$Id) { if ($groupNames[$Id]) { $groupNames[$Id] } else { $Id } }

    $batchKnown = $true
    $plannedGroups = @(); $plannedSkus = @()
    $pr = Get-SnapshotValue $Snapshot 'planned_removals'
    $prGroups = Get-SnapshotValue $pr 'groups'
    $prSkus = Get-SnapshotValue $pr 'license_sku_ids'
    if (-not (Test-IsList $prGroups) -or -not (Test-IsList $prSkus)) {
        $batchKnown = $false
        $readWarnings++
        $reasons.Add((New-Reason -Rule 'batch-unreadable' -Decision 'REVIEW' -Text 'Human decision required: the planned removal batch was not read (planned_removals.groups and planned_removals.license_sku_ids must both be lists; an empty list is fine).' -MissingData 'planned_removals.groups and planned_removals.license_sku_ids' -Source $script:Sources.GroupLicensing))
        $log.Add('Removal batch not readable.')
    } else {
        $plannedGroups = @($prGroups | Where-Object { $_ } | ForEach-Object { ([string]$_).ToLowerInvariant() } | Select-Object -Unique)
        $plannedSkus   = @($prSkus | Where-Object { $_ } | ForEach-Object { ([string]$_).ToLowerInvariant() } | Select-Object -Unique)
    }
    $out.checked_batch = [ordered]@{
        groups   = @($plannedGroups | ForEach-Object { [ordered]@{ id = $_; displayName = (Get-GroupLabel $_) } })
        licenses = @($plannedSkus | ForEach-Object { [ordered]@{ skuId = $_; product = (Get-ProductName $_) } })
    }

    # 11. License assignments: which Plan 2 assignments survive the whole batch.
    $licensesKnown = $true
    $assignments = Get-SnapshotValue $user 'licenseAssignmentStates'
    if (-not (Test-IsList $assignments)) {
        $licensesKnown = $false
        $readWarnings++
        $reasons.Add((New-Reason -Rule 'licenses-unknown' -Decision 'REVIEW' -Text 'Human decision required: license assignments were not read.' -MissingData 'Graph user licenseAssignmentStates' -Source $script:Sources.LicenseState))
        $assignments = @()
    }

    $removedActive = [System.Collections.Generic.List[object]]::new()
    $keptActive    = [System.Collections.Generic.List[object]]::new()
    $keptInError   = [System.Collections.Generic.List[object]]::new()
    $removedOther  = 0
    foreach ($a in @($assignments)) {
        $complete = ($a -is [System.Collections.IDictionary]) -and $a.Contains('assignedByGroup') -and $a.Contains('disabledPlans') -and
                    $a.Contains('skuId') -and $a['skuId'] -and $a.Contains('state') -and (Test-IsList $a['disabledPlans'])
        if (-not $complete) {
            $licensesKnown = $false
            $readWarnings++
            $reasons.Add((New-Reason -Rule 'assignment-incomplete' -Decision 'REVIEW' -Text 'Human decision required: a license assignment is missing skuId, state, disabledPlans or assignedByGroup (only an explicit null assignedByGroup means a direct assignment).' -MissingData 'complete licenseAssignmentState' -Source $script:Sources.LicenseState))
            continue
        }
        $sku = ([string]$a['skuId']).ToLowerInvariant()
        if ($script:Plan2SkuIds -notcontains $sku) { continue }
        $disabled = @($a['disabledPlans'] | ForEach-Object { ([string]$_).ToLowerInvariant() })
        if ($disabled -contains $script:Plan2ServicePlanId) { continue }
        $groupId = if ($a['assignedByGroup']) { ([string]$a['assignedByGroup']).ToLowerInvariant() } else { $null }
        $removed = if ($groupId) { $plannedGroups -contains $groupId } else { $plannedSkus -contains $sku }
        $state = [string]$a['state']
        $err = if ($a.Contains('error')) { [string]$a['error'] } else { '' }
        $healthy = ($state -eq 'Active') -and (-not $err -or $err -eq 'None')
        $desc = [ordered]@{
            skuId           = $sku
            product         = (Get-ProductName $sku)
            assigned        = if ($groupId) { 'by group' } else { 'directly' }
            assignedByGroup = $groupId
            group_name      = if ($groupId) { Get-GroupLabel $groupId } else { $null }
            state           = $state
            error           = $err
        }
        if ($removed) { if ($healthy) { $removedActive.Add($desc) } else { $removedOther++ } }
        elseif ($healthy) { $keptActive.Add($desc) }
        else { $keptInError.Add($desc) }
    }
    $log.Add("Plan 2 assignments: $($removedActive.Count) healthy removed by the batch, $($keptActive.Count) healthy kept, $($keptInError.Count) kept but not healthy.")

    # 12. The 50 GB rule. Only decided when size, batch and every assignment were read.
    if ($null -ne $sizeBytes -and $batchKnown -and $licensesKnown -and $sizeBytes -ge $script:SharedMailboxLimitBytes -and ($type -eq 'SharedMailbox' -or $type -eq 'UserMailbox')) {
        $size = Format-Size $sizeBytes
        if ($keptActive.Count -gt 0) {
            $out.capacity_kept_by = @($keptActive)
            $log.Add('Mailbox capacity stays covered by a healthy Plan 2 assignment outside the batch.')
            if ($type -eq 'SharedMailbox' -and $sizeBytes -gt $script:Plan2CeilingBytes) {
                $reasons.Add((New-Reason -Rule 'over-plan2-ceiling' -Decision 'REVIEW' -Text "Human decision required: the mailbox is $size, above the 100 GB that Exchange Online Plan 2 gives a shared mailbox (a nominal limit; an admin can set a lower quota)." -MissingData 'storage decision for the mailbox' -Source $script:Sources.SharedMailboxLimit))
            }
        } elseif ($removedActive.Count -gt 0 -and $keptInError.Count -gt 0) {
            $e = $keptInError[0]
            $reasons.Add((New-Reason -Rule 'replacement-in-error' -Decision 'REVIEW' -Text "Human decision required: the replacement $($e.product) assignment outside the batch is $($e.state) (error $($e.error)), so it does not yet cover the $size mailbox." -MissingData 'replacement assignment Active with no error' -Source $script:Sources.LicenseState))
        } elseif ($removedActive.Count -gt 0) {
            $byGroup = @($removedActive | Where-Object { $_.assignedByGroup })
            $carriers = @($removedActive | ForEach-Object {
                if ($_.assignedByGroup) { "group $($_.group_name)" } else { "a direct assignment of $($_.product)" }
            } | Select-Object -Unique)
            $groupsCarrying = @($byGroup | ForEach-Object { [ordered]@{ id = $_.assignedByGroup; displayName = $_.group_name } })
            $quota = if ($sizeBytes -eq $script:SharedMailboxLimitBytes) { "reaches the 50 GB quota at exactly $size" } else { "goes over its 50 GB quota at $size" }
            $what = if ($type -eq 'SharedMailbox') { "The mailbox stays shared but $quota." } else { "The mailbox still reads back as UserMailbox; as a shared mailbox without that license it $quota." }
            if ($byGroup.Count -gt 0) {
                $first = if ($holdFound) { 'Keep Plan 2 assigned independently of the group' } else { 'Keep Exchange Online (Plan 2) assigned independently of the group' }
                $fix = "$($first): assign it directly, confirm the assignment is Active, then check again."
            } else {
                $first = if ($holdFound) { 'Keep Plan 2 assigned' } else { 'Keep Exchange Online (Plan 2) assigned' }
                $fix = "$($first): take the direct Exchange Online (Plan 2) license out of the batch, then check again."
            }
            $fix += if ($holdFound) { ' Bringing the mailbox under 50 GB is an option only after whoever owns the hold approves.' } else { ' Or bring the mailbox under 50 GB, then check again.' }
            $reasons.Insert(0, (New-Reason -Rule 'last-plan2-removed' -Decision 'BLOCK' `
                -Text "This batch removes the last healthy Exchange Online (Plan 2) assignment, carried by $($carriers -join ' and '). $what" `
                -Source $script:Sources.SharedMailboxLimit `
                -Extra ([ordered]@{
                    carried_by         = $groupsCarrying
                    consequence        = 'Exchange team blog: when the license is removed from a converted shared mailbox over 50 GB, the 50 GB quota is enforced, sending and receiving stop, and incoming mail fails with "Mailbox is full", code 5.2.2.'
                    consequence_source = $script:Sources.OverQuotaEffect
                    remediation        = $fix
                    remediation_source = $script:Sources.ConvertMailbox
                })))
            $log.Add('BLOCK: the batch removes the last healthy Plan 2 for a mailbox of 50 GB or more.')
        } elseif ($type -eq 'SharedMailbox') {
            $reasons.Add((New-Reason -Rule 'over-quota-without-plan2' -Decision 'REVIEW' -Text "Human decision required: the shared mailbox is $size and no healthy (Active, no error) Exchange Online (Plan 2) assignment covers it now." -MissingData 'license decision for the mailbox size' -Source $script:Sources.SharedMailboxLimit))
        }
    }

    # 13. Decision. BLOCK wins over REVIEW, REVIEW over ALLOW. ALLOW needs a SharedMailbox.
    $decision = 'ALLOW'
    if ($reasons | Where-Object { $_.decision -eq 'REVIEW' }) { $decision = 'REVIEW' }
    if ($reasons | Where-Object { $_.decision -eq 'BLOCK' }) { $decision = 'BLOCK' }
    if ($decision -eq 'ALLOW' -and $type -ne 'SharedMailbox') { $decision = 'REVIEW' }
    $out.success = $true
    $out.status = if ($readWarnings -gt 0) { 1001 } else { 1000 }
    $out.decision = $decision
    $log.Add("Decision: $decision. Nothing was changed in the tenant.")
    return Complete-Output
}

Export-ModuleMember -Function Test-OffboardingRemovalSafety
