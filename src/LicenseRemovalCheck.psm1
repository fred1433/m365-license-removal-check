#Requires -Version 7.0
<#
  LicenseRemovalCheck
  A read-only check that runs before an offboarding workflow removes groups and licenses
  from a Microsoft 365 user whose mailbox has been converted to a shared mailbox.

  It reads one JSON snapshot (Graph user + Exchange mailbox data) and the full planned
  removal batch, and returns ALLOW, BLOCK or REVIEW with reasons and sources.
  It connects to nothing, writes nothing, holds no secret and calls no AI model.
#>
Set-StrictMode -Version Latest

# 50 GB, the shared mailbox size limit without a license.
# Exchange sizes are binary: Set-Mailbox documents "1.999999999 terabytes (2199023254528 bytes)",
# so 1 GB = 1,073,741,824 bytes and 50 GB = 53,687,091,200 bytes.
$script:SharedMailboxLimitBytes = [int64]53687091200
$script:Inv = [System.Globalization.CultureInfo]::InvariantCulture

$script:Sources = @{
    SharedMailboxLimit = 'https://learn.microsoft.com/en-us/microsoft-365/admin/email/about-shared-mailboxes'
    ConvertMailbox     = 'https://learn.microsoft.com/en-us/microsoft-365/admin/email/convert-user-mailbox-to-shared-mailbox'
    ByteUnits          = 'https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/set-mailbox'
    GroupLicensing     = 'https://learn.microsoft.com/en-us/microsoft-365/admin/manage/manage-group-licenses'
    LicenseState       = 'https://learn.microsoft.com/en-us/graph/api/resources/licenseassignmentstate'
    UserResource       = 'https://learn.microsoft.com/en-us/graph/api/resources/user'
    IdentifyHold       = 'https://learn.microsoft.com/en-us/purview/ediscovery-identify-a-hold-on-an-exchange-online-mailbox'
    GetMailbox         = 'https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/get-mailbox'
    OrgConfig          = 'https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/get-organizationconfig'
    MailboxStats       = 'https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/get-exomailboxstatistics'
    Licensing          = 'https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference'
    ArchiveProperties  = 'https://learn.microsoft.com/en-us/troubleshoot/microsoft-365/purview/archive-mailboxes/archivestatus-set-none'
    OverQuotaEffect    = 'https://techcommunity.microsoft.com/blog/exchange/correcting-shared-mailbox-provisioning-and-sizing/607991'
    RewstRemoveLicense = 'https://docs.rewst.help/documentation/automations/subworkflows/rewst-task-m365-remove-licenses'
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

function Get-ProductName {
    param([string]$SkuId)
    $n = $script:Reference.product_names[$SkuId.ToLowerInvariant()]
    if ($n) { return $n } else { return "SKU $SkuId" }
}

function Get-ByteCount {
    # The snapshot carries the primary mailbox size as a whole number of bytes.
    # Anything else (text, negative, fractional) is treated as unreadable.
    param([AllowNull()]$Value)
    if ($Value -is [int] -or $Value -is [long]) {
        if ($Value -ge 0) { return [int64]$Value }
    }
    return $null
}

function Get-HoldKind {
    # Meaning of an InPlaceHolds entry from its prefix, as written on the Purview page
    # "Identify the type of hold placed on an Exchange Online mailbox" (read 2026-11-09).
    # Returns $null when the page gives no meaning for the prefix.
    param([string]$Entry)
    switch -Regex ($Entry) {
        '^-mbx' { return 'exclusion from a retention policy' }
        '^UniH' { return 'eDiscovery hold' }
        '^cld'  { return 'In-Place Hold' }
        '^mbx'  { return 'retention policy' }
        '^skp'  { return 'retention policy on Skype for Business conversations' }
        '^grp'  { return 'retention policy on Microsoft 365 groups and Teams channel messages' }
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
        note             = 'success means the data was collected. It does not mean the removal is allowed: read decision.'
    }

    function Complete-Output {
        $out.reasons = @($reasons)
        $out.automation_log = ($log -join "`n")
        return $out
    }

    # 1. Load the snapshot.
    if ($PSCmdlet.ParameterSetName -eq 'Path') {
        if (-not (Test-Path -LiteralPath $SnapshotPath)) {
            $log.Add("Snapshot not found: $SnapshotPath")
            $reasons.Add((New-Reason -Rule 'snapshot-unreadable' -Decision 'REVIEW' -Text 'Human decision required: the snapshot could not be read.' -MissingData 'snapshot file' -Source $script:Sources.GetMailbox))
            return Complete-Output
        }
        try { $Snapshot = Get-Content -Raw -LiteralPath $SnapshotPath | ConvertFrom-Json -AsHashtable -ErrorAction Stop }
        catch {
            $log.Add("Snapshot is not valid JSON: $SnapshotPath")
            $reasons.Add((New-Reason -Rule 'snapshot-unreadable' -Decision 'REVIEW' -Text 'Human decision required: the snapshot is not valid JSON.' -MissingData 'snapshot file' -Source $script:Sources.GetMailbox))
            return Complete-Output
        }
    }
    $log.Add("Read-only check started for aad_user_id $AadUserId in tenant $TenantId.")

    $user    = Get-SnapshotValue $Snapshot 'graph', 'user'
    $mailbox = Get-SnapshotValue $Snapshot 'exchange', 'mailbox'
    if ($null -eq $user -or $null -eq $mailbox) {
        $missing = if ($null -eq $user) { 'graph.user' } else { 'exchange.mailbox' }
        $log.Add("Collection failed: $missing is missing.")
        $reasons.Add((New-Reason -Rule 'collection-failed' -Decision 'REVIEW' -Text "Human decision required: $missing is missing from the snapshot." -MissingData $missing -Source $script:Sources.GetMailbox))
        return Complete-Output
    }
    $snapUserId = Get-SnapshotValue $user 'id'
    if ([string]$snapUserId -ne $AadUserId) {
        $log.Add("Collection failed: snapshot user $snapUserId is not $AadUserId.")
        $reasons.Add((New-Reason -Rule 'user-mismatch' -Decision 'REVIEW' -Text "Human decision required: the snapshot describes user $snapUserId, not $AadUserId." -MissingData 'graph.user.id' -Source $script:Sources.UserResource))
        return Complete-Output
    }

    # 2. Tenant identity on each block of data.
    $graphTenant = [string](Get-SnapshotValue $Snapshot 'graph', 'tenant_id')
    $exoTenant   = [string](Get-SnapshotValue $Snapshot 'exchange', 'tenant_id')
    if ($graphTenant -ne $TenantId -or $exoTenant -ne $TenantId) {
        $reasons.Add((New-Reason -Rule 'tenant-mismatch' -Decision 'REVIEW' -Text "Human decision required: tenant_id differs between the request ($TenantId), Graph data ($graphTenant) and Exchange data ($exoTenant)." -MissingData 'one tenant_id for all data' -Source $script:Sources.UserResource))
        $log.Add('Tenant IDs differ between request, Graph and Exchange data.')
    }

    # 3. Hybrid identities are out of scope.
    if ((Get-SnapshotValue $user 'onPremisesSyncEnabled') -eq $true) {
        $reasons.Add((New-Reason -Rule 'identity-synced' -Decision 'REVIEW' -Text 'Human decision required: identity synchronized from on-premises Active Directory, out of scope for this check.' -MissingData 'cloud-only identity' -Source $script:Sources.UserResource))
        $log.Add('onPremisesSyncEnabled is true: hybrid identity, out of scope.')
    }

    # 4. Mailbox type read back from Exchange.
    $type = [string](Get-SnapshotValue $mailbox 'RecipientTypeDetails')
    $conversionRequested = (Get-SnapshotValue $Snapshot 'conversion_requested') -eq $true
    if ($conversionRequested -and $type -ne 'SharedMailbox') {
        $shown = if ($type) { $type } else { 'unknown' }
        $reasons.Add((New-Reason -Rule 'not-yet-shared' -Decision 'REVIEW' -Text "Human decision required: conversion to shared was requested but the mailbox reads back as $shown." -MissingData 'RecipientTypeDetails = SharedMailbox' -Source $script:Sources.ConvertMailbox))
    }
    $log.Add("Mailbox type read back: $(if ($type) { $type } else { 'unknown' }).")

    # 5. Mailbox size, in bytes.
    $rawSize = Get-SnapshotValue $Snapshot 'exchange', 'mailbox_statistics', 'TotalItemSize'
    $sizeBytes = Get-ByteCount $rawSize
    if ($null -eq $sizeBytes) {
        $readWarnings++
        $why = if ($null -eq $rawSize) { 'missing' } else { 'unreadable (expected a whole number of bytes)' }
        $reasons.Add((New-Reason -Rule 'size-unreadable' -Decision 'REVIEW' -Text "Human decision required: mailbox size $why." -MissingData 'Get-EXOMailboxStatistics TotalItemSize' -Source $script:Sources.MailboxStats))
        $log.Add("Mailbox size $why.")
    } else {
        $log.Add([string]::Format($script:Inv, 'Primary mailbox size: {0:N0} bytes (limit without license {1:N0}).', $sizeBytes, $script:SharedMailboxLimitBytes))
    }

    # 6. Archive. Microsoft says to read ArchiveGuid and ArchiveDatabase, not ArchiveStatus.
    if (-not (Test-HasKey $mailbox 'ArchiveGuid') -or -not (Test-HasKey $mailbox 'ArchiveDatabase')) {
        $readWarnings++
        $reasons.Add((New-Reason -Rule 'archive-unknown' -Decision 'REVIEW' -Text 'Human decision required: archive status was not read.' -MissingData 'Get-Mailbox ArchiveGuid and ArchiveDatabase' -Source $script:Sources.ArchiveProperties))
    } else {
        $aGuid = [string]$mailbox['ArchiveGuid']
        $aDb = [string]$mailbox['ArchiveDatabase']
        $hasGuid = $aGuid -and ($aGuid -ne '00000000-0000-0000-0000-000000000000')
        if ($hasGuid -or $aDb) {
            $reasons.Add((New-Reason -Rule 'archive-present' -Decision 'REVIEW' -Text 'Human decision required: an online archive is present (ArchiveGuid or ArchiveDatabase is set). A shared mailbox with an archive needs Exchange Online Plan 2, or Plan 1 with Exchange Online Archiving.' -MissingData 'license decision for the archive' -Source $script:Sources.SharedMailboxLimit))
        }
    }

    # 7. Holds on the mailbox.
    foreach ($flag in 'LitigationHoldEnabled', 'ComplianceTagHoldApplied', 'DelayHoldApplied', 'DelayReleaseHoldApplied') {
        if (-not (Test-HasKey $mailbox $flag)) {
            $readWarnings++
            $reasons.Add((New-Reason -Rule 'hold-unknown' -Decision 'REVIEW' -Text "Human decision required: $flag was not read." -MissingData "Get-Mailbox $flag" -Source $script:Sources.IdentifyHold))
        } elseif ($mailbox[$flag] -eq $true) {
            $reasons.Add((New-Reason -Rule 'hold-present' -Decision 'REVIEW' -Text "Human decision required: $flag is true on the mailbox." -MissingData 'hold decision by the client' -Source $script:Sources.IdentifyHold -Extra ([ordered]@{ hold_scope = 'mailbox'; hold_property = $flag })))
        }
    }
    if (-not (Test-HasKey $mailbox 'InPlaceHolds')) {
        $readWarnings++
        $reasons.Add((New-Reason -Rule 'hold-unknown' -Decision 'REVIEW' -Text 'Human decision required: mailbox InPlaceHolds was not read.' -MissingData 'Get-Mailbox InPlaceHolds' -Source $script:Sources.IdentifyHold))
    } else {
        foreach ($h in @($mailbox['InPlaceHolds'] | Where-Object { $_ })) {
            if ($h -like '-mbx*') { $log.Add("Mailbox is excluded from retention policy $h."); continue }
            $kind = Get-HoldKind $h
            $extra = [ordered]@{ hold_scope = 'mailbox'; hold_entry = $h }
            if ($kind) { $extra.hold_type = $kind }
            $reasons.Add((New-Reason -Rule 'hold-present' -Decision 'REVIEW' -Text "Human decision required: InPlaceHolds on the mailbox lists $h$(if ($kind) { " ($kind)" })." -MissingData 'hold decision by the client' -Source $script:Sources.IdentifyHold -Extra $extra))
        }
    }

    # 8. Organization-wide retention: an empty mailbox InPlaceHolds does not rule it out.
    $org = Get-SnapshotValue $Snapshot 'exchange', 'organization_config'
    if ($null -eq $org -or -not (Test-HasKey $org 'InPlaceHolds') -or $null -eq $org['InPlaceHolds']) {
        $readWarnings++
        $reasons.Add((New-Reason -Rule 'org-retention-unknown' -Decision 'REVIEW' -Text 'Human decision required: organization-wide retention could not be read, so its scope is unknown.' -MissingData 'Get-OrganizationConfig InPlaceHolds' -Source $script:Sources.IdentifyHold))
        $log.Add('Organization-wide retention not readable.')
    } else {
        foreach ($h in @($org['InPlaceHolds'] | Where-Object { $_ })) {
            $kind = Get-HoldKind $h
            $extra = [ordered]@{ hold_scope = 'organization'; hold_entry = $h }
            if ($kind) { $extra.hold_type = $kind }
            $reasons.Add((New-Reason -Rule 'hold-present' -Decision 'REVIEW' -Text "Human decision required: organization InPlaceHolds lists $h$(if ($kind) { " ($kind)" }), which can cover this mailbox." -MissingData 'retention policy scope and exclusions' -Source $script:Sources.IdentifyHold -Extra $extra))
        }
    }

    # 9. License assignments: which ones survive the whole batch.
    $plannedGroups = @((Get-SnapshotValue $Snapshot 'planned_removals', 'groups') | Where-Object { $_ } | ForEach-Object { ([string]$_).ToLowerInvariant() })
    $plannedSkus   = @((Get-SnapshotValue $Snapshot 'planned_removals', 'license_sku_ids') | Where-Object { $_ } | ForEach-Object { ([string]$_).ToLowerInvariant() })
    $groupNames = @{}
    $groupList = Get-SnapshotValue $Snapshot 'graph', 'groups'
    foreach ($g in @($groupList)) {
        if ($g -is [System.Collections.IDictionary] -and $g.Contains('id')) { $groupNames[([string]$g.id).ToLowerInvariant()] = [string]$g.displayName }
    }
    $out.checked_batch = [ordered]@{
        groups   = @($plannedGroups | ForEach-Object { [ordered]@{ id = $_; displayName = $groupNames[$_] } })
        licenses = @($plannedSkus | ForEach-Object { [ordered]@{ skuId = $_; product = (Get-ProductName $_) } })
    }

    $assignments = Get-SnapshotValue $user 'licenseAssignmentStates'
    if ($null -eq $assignments) {
        $readWarnings++
        $reasons.Add((New-Reason -Rule 'licenses-unknown' -Decision 'REVIEW' -Text 'Human decision required: license assignments were not read.' -MissingData 'Graph user licenseAssignmentStates' -Source $script:Sources.LicenseState))
        $assignments = @()
    }

    $removedActive = [System.Collections.Generic.List[object]]::new()
    $keptActive    = [System.Collections.Generic.List[object]]::new()
    $keptInError   = [System.Collections.Generic.List[object]]::new()
    foreach ($a in @($assignments)) {
        $sku = ([string](Get-SnapshotValue $a 'skuId')).ToLowerInvariant()
        if ($script:Plan2SkuIds -notcontains $sku) { continue }
        $disabled = @((Get-SnapshotValue $a 'disabledPlans') | ForEach-Object { ([string]$_).ToLowerInvariant() })
        if ($disabled -contains $script:Plan2ServicePlanId) { continue }
        $byGroup = Get-SnapshotValue $a 'assignedByGroup'
        $groupId = if ($byGroup) { ([string]$byGroup).ToLowerInvariant() } else { $null }
        $removed = if ($groupId) { $plannedGroups -contains $groupId } else { $plannedSkus -contains $sku }
        $state = [string](Get-SnapshotValue $a 'state')
        $err = [string](Get-SnapshotValue $a 'error')
        $healthy = ($state -eq 'Active') -and (-not $err -or $err -eq 'None')
        $desc = [ordered]@{
            skuId           = $sku
            product         = (Get-ProductName $sku)
            assigned        = if ($groupId) { 'by group' } else { 'directly' }
            assignedByGroup = $groupId
            group_name      = if ($groupId) { $groupNames[$groupId] } else { $null }
            state           = $state
            error           = $err
        }
        if ($removed) { if ($healthy) { $removedActive.Add($desc) } }
        elseif ($healthy) { $keptActive.Add($desc) }
        else { $keptInError.Add($desc) }
    }
    $log.Add("Plan 2 assignments: $($removedActive.Count) active removed by the batch, $($keptActive.Count) active kept, $($keptInError.Count) kept but not active.")

    # 10. The 50 GB rule for a shared mailbox.
    if ($null -ne $sizeBytes -and $sizeBytes -ge $script:SharedMailboxLimitBytes -and $type -eq 'SharedMailbox') {
        $gb = ([math]::Round($sizeBytes / 1GB, 1)).ToString($script:Inv)
        if ($keptActive.Count -gt 0) {
            $out.capacity_kept_by = @($keptActive)
            $log.Add('Mailbox capacity stays covered by an active Plan 2 assignment outside the batch.')
        } elseif ($removedActive.Count -gt 0 -and $keptInError.Count -gt 0) {
            $e = $keptInError[0]
            $reasons.Add((New-Reason -Rule 'replacement-in-error' -Decision 'REVIEW' -Text "Human decision required: the replacement $($e.product) assignment is $($e.state) (error $($e.error)), so it does not yet cover the $gb GB mailbox." -MissingData 'replacement license state = Active' -Source $script:Sources.LicenseState))
        } elseif ($removedActive.Count -gt 0) {
            $carriers = @($removedActive | ForEach-Object {
                if ($_.assignedByGroup) { "group $($_.group_name) ($($_.assignedByGroup))" } else { "direct assignment of $($_.product)" }
            })
            $groupsCarrying = @($removedActive | Where-Object { $_.assignedByGroup } | ForEach-Object { [ordered]@{ id = $_.assignedByGroup; displayName = $_.group_name } })
            $reasons.Add((New-Reason -Rule 'last-plan2-removed' -Decision 'BLOCK' `
                -Text "This batch removes the last active Exchange Online (Plan 2) assignment, carried by $($carriers -join ' and '). The mailbox stays shared but goes over its 50 GB quota at $gb GB." `
                -Source $script:Sources.SharedMailboxLimit `
                -Extra ([ordered]@{
                    carried_by         = $groupsCarrying
                    consequence        = 'Exchange team blog: when the license is removed from a converted shared mailbox over 50 GB, the 50 GB quota is enforced, sending and receiving stop, and incoming mail fails with "Mailbox is full", code 5.2.2.'
                    consequence_source = $script:Sources.OverQuotaEffect
                    remediation        = 'Assign Exchange Online (Plan 2) independently of the group, confirm the assignment is Active, then check again. Or bring the mailbox under 50 GB, then check again.'
                    remediation_source = $script:Sources.ConvertMailbox
                })))
            $log.Add('BLOCK: the batch removes the last active Plan 2 for a shared mailbox over 50 GB.')
        } else {
            $reasons.Add((New-Reason -Rule 'over-quota-without-plan2' -Decision 'REVIEW' -Text "Human decision required: the shared mailbox is $gb GB and no active Exchange Online (Plan 2) assignment covers it now." -MissingData 'license decision for the mailbox size' -Source $script:Sources.SharedMailboxLimit))
        }
    }

    # 11. Decision.
    $decision = 'ALLOW'
    if ($reasons | Where-Object { $_.decision -eq 'REVIEW' }) { $decision = 'REVIEW' }
    if ($reasons | Where-Object { $_.decision -eq 'BLOCK' }) { $decision = 'BLOCK' }
    $out.success = $true
    $out.status = if ($readWarnings -gt 0) { 1001 } else { 1000 }
    $out.decision = $decision
    $log.Add("Decision: $decision. Nothing was changed in the tenant.")
    return Complete-Output
}

Export-ModuleMember -Function Test-OffboardingRemovalSafety
