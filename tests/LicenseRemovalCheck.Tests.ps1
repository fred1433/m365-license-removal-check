BeforeAll {
    $script:Root = Split-Path -Parent $PSScriptRoot
    Import-Module (Join-Path $script:Root 'src' 'LicenseRemovalCheck.psm1') -Force

    $script:Tenant = '7d3e5c1a-2b4f-4e8a-9c6d-0f1e2a3b4c5d'
    $script:User = 'a41f6b2e-8c3d-4f7a-b1e9-52d06c8e7f13'
    $script:FirmStaff = 'e2b7c9d4-1a5f-4d3e-8b6c-9f0a1e2d3c4b'
    $script:Plan2 = '19ec0d23-8335-4cbd-94ac-6050e30712fa'
    $script:Limit = [int64]53687091200

    function script:Get-Fixture([string]$Name) {
        Get-Content -Raw -LiteralPath (Join-Path $script:Root 'fixtures' "$Name.json") | ConvertFrom-Json -AsHashtable
    }
    function script:Invoke-Check([hashtable]$Snapshot) {
        Test-OffboardingRemovalSafety -AadUserId $script:User -TenantId $script:Tenant -Snapshot $Snapshot
    }
    function script:Save-Output($Result, [string]$Name) {
        $dir = Join-Path $script:Root 'examples' 'output'
        $json = ($Result | ConvertTo-Json -Depth 10) -replace "`r`n", "`n"
        [System.IO.File]::WriteAllText((Join-Path $dir "$Name.json"), $json + "`n")
    }
    function script:New-Plan2([string]$GroupId, [string]$State = 'Active', [string]$ErrorValue = 'None') {
        [ordered]@{ assignedByGroup = $GroupId; disabledPlans = @(); error = $ErrorValue; lastUpdatedDateTime = '2026-11-09T08:57:40Z'; skuId = $script:Plan2; state = $State }
    }
}

Describe 'The three tickets shown on the page' {
    It 'Scenario 1: Firm Staff carries the only active Plan 2, so the group removal is blocked' {
        $r = Invoke-Check (Get-Fixture 'scenario-1-firm-staff-only-plan2')
        Save-Output $r 'scenario-1-firm-staff-only-plan2'
        $r.decision | Should -Be 'BLOCK'
        $r.success | Should -BeTrue
        $r.status | Should -Be 1000
        $block = @($r.reasons | Where-Object { $_.decision -eq 'BLOCK' })
        $block.Count | Should -Be 1
        $block[0].carried_by[0].id | Should -Be $script:FirmStaff
        $block[0].carried_by[0].displayName | Should -Be 'Firm Staff'
        $block[0].text | Should -Match 'Firm Staff'
        $block[0].remediation | Should -Match 'independently of the group'
        $block[0].remediation | Should -Match 'under 50 GB'
    }

    It 'Scenario 2: a second, direct Plan 2 stays assigned, so the removal is allowed and the page can name it' {
        $r = Invoke-Check (Get-Fixture 'scenario-2-second-plan2-kept')
        Save-Output $r 'scenario-2-second-plan2-kept'
        $r.decision | Should -Be 'ALLOW'
        $r.status | Should -Be 1000
        @($r.reasons).Count | Should -Be 0
        @($r.capacity_kept_by).Count | Should -Be 1
        $r.capacity_kept_by[0].assigned | Should -Be 'directly'
        $r.capacity_kept_by[0].state | Should -Be 'Active'
    }

    It 'Scenario 3: organization retention is unreadable, so a human decides' {
        $r = Invoke-Check (Get-Fixture 'scenario-3-retention-unreadable')
        Save-Output $r 'scenario-3-retention-unreadable'
        $r.decision | Should -Be 'REVIEW'
        $r.success | Should -BeTrue
        $r.status | Should -Be 1001
        $r.reasons[0].rule | Should -Be 'org-retention-unknown'
        $r.reasons[0].missing_data | Should -Be 'Get-OrganizationConfig InPlaceHolds'
    }
}

Describe 'Batches that look fine but are not' {
    It 'Two groups carry Plan 2 and the batch removes both: BLOCK, naming both groups' {
        $s = Get-Fixture 'scenario-1-firm-staff-only-plan2'
        $other = 'b9c1d2e3-4f5a-4b6c-8d7e-0a1b2c3d4e5f'
        $s.graph.groups += [ordered]@{ id = $other; displayName = 'Litigation Team' }
        $s.graph.user.licenseAssignmentStates += (New-Plan2 $other)
        $s.planned_removals.groups = @($script:FirmStaff, $other)
        $r = Invoke-Check $s
        $r.decision | Should -Be 'BLOCK'
        $ids = @($r.reasons | Where-Object { $_.decision -eq 'BLOCK' } | ForEach-Object { $_.carried_by } | ForEach-Object { $_.id })
        $ids | Should -Contain $script:FirmStaff
        $ids | Should -Contain $other
    }

    It 'Two groups carry Plan 2 and the batch removes only one: ALLOW, kept by the other group' {
        $s = Get-Fixture 'scenario-1-firm-staff-only-plan2'
        $other = 'b9c1d2e3-4f5a-4b6c-8d7e-0a1b2c3d4e5f'
        $s.graph.groups += [ordered]@{ id = $other; displayName = 'Litigation Team' }
        $s.graph.user.licenseAssignmentStates += (New-Plan2 $other)
        $r = Invoke-Check $s
        $r.decision | Should -Be 'ALLOW'
        $r.capacity_kept_by[0].group_name | Should -Be 'Litigation Team'
    }

    It 'Replacement Plan 2 requested but in error: REVIEW' -ForEach @(
        @{ State = 'Error'; Err = 'CountViolation' }
        @{ State = 'ActiveWithError'; Err = 'DependencyViolation' }
    ) {
        $s = Get-Fixture 'scenario-1-firm-staff-only-plan2'
        $s.graph.user.licenseAssignmentStates += (New-Plan2 $null $State $Err)
        $r = Invoke-Check $s
        $r.decision | Should -Be 'REVIEW'
        @($r.reasons.rule) | Should -Contain 'replacement-in-error'
    }

    It 'Conversion requested but the mailbox still reads back as UserMailbox: REVIEW' {
        $s = Get-Fixture 'scenario-2-second-plan2-kept'
        $s.exchange.mailbox.RecipientTypeDetails = 'UserMailbox'
        $r = Invoke-Check $s
        $r.decision | Should -Be 'REVIEW'
        @($r.reasons.rule) | Should -Contain 'not-yet-shared'
    }

    It 'Graph and Exchange data come from different tenants: REVIEW' {
        $s = Get-Fixture 'scenario-2-second-plan2-kept'
        $s.exchange.tenant_id = '00000000-1111-4222-8333-444444444444'
        $r = Invoke-Check $s
        $r.decision | Should -Be 'REVIEW'
        @($r.reasons.rule) | Should -Contain 'tenant-mismatch'
    }

    It 'Identity synchronized from on-premises: REVIEW, out of scope' {
        $s = Get-Fixture 'scenario-2-second-plan2-kept'
        $s.graph.user.onPremisesSyncEnabled = $true
        $r = Invoke-Check $s
        $r.decision | Should -Be 'REVIEW'
        @($r.reasons.rule) | Should -Contain 'identity-synced'
    }

    It 'Mailbox under 50 GB, no hold, but an online archive is present: REVIEW' {
        $s = Get-Fixture 'scenario-1-firm-staff-only-plan2'
        $s.exchange.mailbox_statistics.TotalItemSize = [int64]20GB
        $s.exchange.mailbox.ArchiveGuid = '3c4d5e6f-7a8b-4c9d-8e0f-1a2b3c4d5e6f'
        $s.exchange.mailbox.ArchiveDatabase = 'NAMPR01DG123-db045'
        $r = Invoke-Check $s
        $r.decision | Should -Be 'REVIEW'
        @($r.reasons.rule) | Should -Be @('archive-present')
    }

    It 'Size given as text instead of bytes: REVIEW, size unreadable' {
        $s = Get-Fixture 'scenario-2-second-plan2-kept'
        $s.exchange.mailbox_statistics.TotalItemSize = '62 GB'
        $r = Invoke-Check $s
        $r.decision | Should -Be 'REVIEW'
        @($r.reasons.rule) | Should -Contain 'size-unreadable'
    }
}

Describe 'The 50 GB line, in bytes' {
    It '<Label>: <Decision>' -ForEach @(
        @{ Label = 'one byte under 50 GB'; Bytes = [int64]53687091199; Decision = 'ALLOW' }
        @{ Label = 'exactly 50 GB'; Bytes = [int64]53687091200; Decision = 'BLOCK' }
        @{ Label = 'one byte over 50 GB'; Bytes = [int64]53687091201; Decision = 'BLOCK' }
    ) {
        $s = Get-Fixture 'scenario-1-firm-staff-only-plan2'
        $s.exchange.mailbox_statistics.TotalItemSize = $Bytes
        (Invoke-Check $s).decision | Should -Be $Decision
    }
}

Describe 'Holds always go to a human' {
    It '<Flag> on the mailbox: REVIEW' -ForEach @(
        @{ Flag = 'LitigationHoldEnabled' }
        @{ Flag = 'ComplianceTagHoldApplied' }
        @{ Flag = 'DelayHoldApplied' }
        @{ Flag = 'DelayReleaseHoldApplied' }
    ) {
        $s = Get-Fixture 'scenario-2-second-plan2-kept'
        $s.exchange.mailbox[$Flag] = $true
        $r = Invoke-Check $s
        $r.decision | Should -Be 'REVIEW'
        $hold = @($r.reasons | Where-Object { $_.rule -eq 'hold-present' })
        $hold[0].hold_scope | Should -Be 'mailbox'
        $hold[0].hold_property | Should -Be $Flag
    }

    It 'InPlaceHolds on the mailbox lists an eDiscovery hold: REVIEW, typed by its prefix' {
        $s = Get-Fixture 'scenario-2-second-plan2-kept'
        $s.exchange.mailbox.InPlaceHolds = @('UniH7d3e5c1a2b4f4e8a9c6d0f1e2a3b4c5d')
        $r = Invoke-Check $s
        $r.decision | Should -Be 'REVIEW'
        $r.reasons[0].hold_scope | Should -Be 'mailbox'
        $r.reasons[0].hold_type | Should -Be 'eDiscovery hold'
    }

    It 'Mailbox InPlaceHolds empty but the organization has a retention policy: REVIEW, scope organization' {
        $s = Get-Fixture 'scenario-2-second-plan2-kept'
        $s.exchange.mailbox.InPlaceHolds = @()
        $s.exchange.organization_config.InPlaceHolds = @('mbx5c2a8f1e9d3b4a7c8e6f0a1b2c3d4e5f:2')
        $r = Invoke-Check $s
        $r.decision | Should -Be 'REVIEW'
        $r.reasons[0].hold_scope | Should -Be 'organization'
        $r.reasons[0].hold_type | Should -Be 'retention policy'
    }

    It 'A mailbox excluded from a retention policy (-mbx) is not a hold by itself' {
        $s = Get-Fixture 'scenario-2-second-plan2-kept'
        $s.exchange.mailbox.InPlaceHolds = @('-mbx5c2a8f1e9d3b4a7c8e6f0a1b2c3d4e5f')
        (Invoke-Check $s).decision | Should -Be 'ALLOW'
    }
}

Describe 'Read-only and replayable' {
    BeforeAll {
        # Stand-ins for the cmdlets a removal would use. The check must never call them.
        function global:Set-Mailbox { }
        function global:Set-MgUserLicense { }
        function global:Remove-MgGroupMemberDirectoryObjectByRef { }
        Mock -ModuleName LicenseRemovalCheck Set-Mailbox { }
        Mock -ModuleName LicenseRemovalCheck Set-MgUserLicense { }
        Mock -ModuleName LicenseRemovalCheck Remove-MgGroupMemberDirectoryObjectByRef { }
    }
    AfterAll {
        Remove-Item Function:\Set-Mailbox, Function:\Set-MgUserLicense, Function:\Remove-MgGroupMemberDirectoryObjectByRef -ErrorAction SilentlyContinue
    }

    It 'Run twice on the same snapshot file: same output, file untouched, zero writes' {
        $path = Join-Path $script:Root 'fixtures' 'scenario-1-firm-staff-only-plan2.json'
        $before = (Get-FileHash -LiteralPath $path).Hash
        $a = Test-OffboardingRemovalSafety -AadUserId $script:User -TenantId $script:Tenant -SnapshotPath $path
        $b = Test-OffboardingRemovalSafety -AadUserId $script:User -TenantId $script:Tenant -SnapshotPath $path
        ($a | ConvertTo-Json -Depth 10) | Should -BeExactly ($b | ConvertTo-Json -Depth 10)
        $a.writes | Should -Be 0
        (Get-FileHash -LiteralPath $path).Hash | Should -Be $before
        Should -Invoke -ModuleName LicenseRemovalCheck Set-Mailbox -Times 0 -Exactly
        Should -Invoke -ModuleName LicenseRemovalCheck Set-MgUserLicense -Times 0 -Exactly
        Should -Invoke -ModuleName LicenseRemovalCheck Remove-MgGroupMemberDirectoryObjectByRef -Times 0 -Exactly
    }
}

Describe 'Collection status is not permission' {
    It 'A BLOCK still reports success true: the data was collected' {
        $r = Invoke-Check (Get-Fixture 'scenario-1-firm-staff-only-plan2')
        $r.success | Should -BeTrue
        $r.decision | Should -Not -Be 'ALLOW'
    }

    It 'Snapshot for another user: success false, status 2000, REVIEW' {
        $r = Test-OffboardingRemovalSafety -AadUserId '11111111-2222-4333-8444-555555555555' -TenantId $script:Tenant -Snapshot (Get-Fixture 'scenario-2-second-plan2-kept')
        $r.success | Should -BeFalse
        $r.status | Should -Be 2000
        $r.decision | Should -Be 'REVIEW'
    }

    It 'Missing snapshot file: success false, REVIEW' {
        $r = Test-OffboardingRemovalSafety -AadUserId $script:User -TenantId $script:Tenant -SnapshotPath (Join-Path $TestDrive 'nope.json')
        $r.success | Should -BeFalse
        $r.decision | Should -Be 'REVIEW'
    }

    It 'Snapshot without organization_config key at all: REVIEW, status 1001' {
        $s = Get-Fixture 'scenario-2-second-plan2-kept'
        $s.exchange.Remove('organization_config')
        $r = Invoke-Check $s
        $r.decision | Should -Be 'REVIEW'
        $r.status | Should -Be 1001
    }
}

Describe 'Fix 1: archive' {
    It 'Archive present: REVIEW, quoting the full Microsoft sentence' {
        $s = Get-Fixture 'scenario-2-second-plan2-kept'
        $s.exchange.mailbox.ArchiveGuid = '3c4d5e6f-7a8b-4c9d-8e0f-1a2b3c4d5e6f'
        $s.exchange.mailbox.ArchiveDatabase = 'NAMPR01DG123-db045'
        $r = Invoke-Check $s
        $r.decision | Should -Be 'REVIEW'
        $r.reasons[0].text | Should -Match ([regex]::Escape('"To increase the size of the archive mailbox, an Exchange Online Plan 2 license or an Exchange Online Plan 1 license with an Exchange Online Archiving add-on license is required."'))
        $r.reasons[0].text | Should -Not -Match 'its own license'
    }
    It 'Only <Set> is set: REVIEW, archive state inconsistent' -ForEach @(
        @{ Set = 'ArchiveGuid'; Guid = '3c4d5e6f-7a8b-4c9d-8e0f-1a2b3c4d5e6f'; Db = $null }
        @{ Set = 'ArchiveDatabase'; Guid = '00000000-0000-0000-0000-000000000000'; Db = 'NAMPR01DG123-db045' }
    ) {
        $s = Get-Fixture 'scenario-2-second-plan2-kept'
        $s.exchange.mailbox.ArchiveGuid = $Guid
        $s.exchange.mailbox.ArchiveDatabase = $Db
        $r = Invoke-Check $s
        $r.decision | Should -Be 'REVIEW'
        @($r.reasons.rule) | Should -Be @('archive-inconsistent')
        $r.reasons[0].text | Should -Match 'archive state inconsistent'
    }
}

Describe 'Fix 2: no decision without the data' {
    It '<Label>: REVIEW <Rule>, status 1001' -ForEach @(
        @{ Label = 'planned_removals missing'; Rule = 'batch-unreadable'; Change = { param($s) $s.Remove('planned_removals') } }
        @{ Label = 'planned_removals.groups missing'; Rule = 'batch-unreadable'; Change = { param($s) $s.planned_removals.Remove('groups') } }
        @{ Label = 'planned_removals.license_sku_ids missing'; Rule = 'batch-unreadable'; Change = { param($s) $s.planned_removals.Remove('license_sku_ids') } }
        @{ Label = 'planned_removals.groups misnamed'; Rule = 'batch-unreadable'; Change = { param($s) $s.planned_removals = [ordered]@{ group_ids = @($script:FirmStaff); license_sku_ids = @() } } }
        @{ Label = 'planned_removals.groups a single string'; Rule = 'batch-unreadable'; Change = { param($s) $s.planned_removals.groups = $script:FirmStaff } }
        @{ Label = 'assignedByGroup key missing'; Rule = 'assignment-incomplete'; Change = { param($s) $s.graph.user.licenseAssignmentStates[2].Remove('assignedByGroup') } }
        @{ Label = 'disabledPlans missing'; Rule = 'assignment-incomplete'; Change = { param($s) $s.graph.user.licenseAssignmentStates[2].Remove('disabledPlans') } }
        @{ Label = 'LitigationHoldEnabled null'; Rule = 'hold-unknown'; Change = { param($s) $s.exchange.mailbox.LitigationHoldEnabled = $null } }
        @{ Label = 'LitigationHoldEnabled missing'; Rule = 'hold-unknown'; Change = { param($s) $s.exchange.mailbox.Remove('LitigationHoldEnabled') } }
        @{ Label = 'LitigationHoldEnabled the string True'; Rule = 'hold-unknown'; Change = { param($s) $s.exchange.mailbox.LitigationHoldEnabled = 'True' } }
        @{ Label = 'mailbox InPlaceHolds null'; Rule = 'hold-unknown'; Change = { param($s) $s.exchange.mailbox.InPlaceHolds = $null } }
        @{ Label = 'mailbox InPlaceHolds missing'; Rule = 'hold-unknown'; Change = { param($s) $s.exchange.mailbox.Remove('InPlaceHolds') } }
        @{ Label = 'RecipientTypeDetails missing'; Rule = 'type-unknown'; Change = { param($s) $s.exchange.mailbox.Remove('RecipientTypeDetails') } }
        @{ Label = 'RecipientTypeDetails empty'; Rule = 'type-unknown'; Change = { param($s) $s.exchange.mailbox.RecipientTypeDetails = '' } }
        @{ Label = 'onPremisesSyncEnabled missing'; Rule = 'identity-unknown'; Change = { param($s) $s.graph.user.Remove('onPremisesSyncEnabled') } }
        @{ Label = 'ArchiveGuid null'; Rule = 'archive-unknown'; Change = { param($s) $s.exchange.mailbox.ArchiveGuid = $null } }
    ) {
        $s = Get-Fixture 'scenario-2-second-plan2-kept'
        & $Change $s
        $r = Invoke-Check $s
        $r.decision | Should -Be 'REVIEW'
        $r.status | Should -Be 1001
        @($r.reasons.rule) | Should -Contain $Rule
    }

    It 'licenseAssignmentStates missing on a 62 GB shared mailbox: REVIEW licenses-unknown only' {
        $s = Get-Fixture 'scenario-1-firm-staff-only-plan2'
        $s.graph.user.Remove('licenseAssignmentStates')
        $r = Invoke-Check $s
        $r.decision | Should -Be 'REVIEW'
        @($r.reasons.rule) | Should -Be @('licenses-unknown')
    }

    It 'An explicit empty license list is a valid batch' {
        $s = Get-Fixture 'scenario-2-second-plan2-kept'
        $s.planned_removals.license_sku_ids = @()
        (Invoke-Check $s).decision | Should -Be 'ALLOW'
    }

    It 'onPremisesSyncEnabled false is a cloud identity: ALLOW' {
        $s = Get-Fixture 'scenario-2-second-plan2-kept'
        $s.graph.user.onPremisesSyncEnabled = $false
        (Invoke-Check $s).decision | Should -Be 'ALLOW'
    }
}

Describe 'Fix 3: the mailbox must be hers' {
    It 'Mailbox ExternalDirectoryObjectId is another user: REVIEW, success false' {
        $s = Get-Fixture 'scenario-2-second-plan2-kept'
        $s.exchange.mailbox.ExternalDirectoryObjectId = '11111111-2222-4333-8444-555555555555'
        $r = Invoke-Check $s
        $r.decision | Should -Be 'REVIEW'
        $r.success | Should -BeFalse
        $r.reasons[0].rule | Should -Be 'mailbox-owner-mismatch'
    }
    It 'Mailbox ExternalDirectoryObjectId missing: REVIEW' {
        $s = Get-Fixture 'scenario-2-second-plan2-kept'
        $s.exchange.mailbox.Remove('ExternalDirectoryObjectId')
        (Invoke-Check $s).reasons[0].rule | Should -Be 'mailbox-owner-unknown'
    }
    It 'Statistics MailboxGuid differs from mailbox ExchangeGuid: REVIEW' {
        $s = Get-Fixture 'scenario-2-second-plan2-kept'
        $s.exchange.mailbox.ExchangeGuid = '0d1e2f3a-4b5c-4d6e-8f7a-9b0c1d2e3f4a'
        $s.exchange.mailbox_statistics.MailboxGuid = 'ffffffff-4b5c-4d6e-8f7a-9b0c1d2e3f4a'
        $r = Invoke-Check $s
        $r.decision | Should -Be 'REVIEW'
        $r.reasons[0].rule | Should -Be 'statistics-mismatch'
    }
    It 'Statistics MailboxGuid equals mailbox ExchangeGuid: ALLOW' {
        $s = Get-Fixture 'scenario-2-second-plan2-kept'
        $s.exchange.mailbox.ExchangeGuid = '0d1e2f3a-4b5c-4d6e-8f7a-9b0c1d2e3f4a'
        $s.exchange.mailbox_statistics.MailboxGuid = '0d1e2f3a-4b5c-4d6e-8f7a-9b0c1d2e3f4a'
        (Invoke-Check $s).decision | Should -Be 'ALLOW'
    }
}

Describe 'Fix 4: no ALLOW unless the mailbox reads back as shared' {
    It 'UserMailbox, conversion not requested, 5 GB: REVIEW not-shared' {
        $s = Get-Fixture 'scenario-1-firm-staff-only-plan2'
        $s.conversion_requested = $false
        $s.exchange.mailbox.RecipientTypeDetails = 'UserMailbox'
        $s.exchange.mailbox_statistics.TotalItemSize = [int64]5GB
        $r = Invoke-Check $s
        $r.decision | Should -Be 'REVIEW'
        @($r.reasons.rule) | Should -Contain 'not-shared'
    }
    It 'UserMailbox, conversion_requested missing, 62 GB, last Plan 2 removed: BLOCK, both reasons listed' {
        $s = Get-Fixture 'scenario-1-firm-staff-only-plan2'
        $s.Remove('conversion_requested')
        $s.exchange.mailbox.RecipientTypeDetails = 'UserMailbox'
        $r = Invoke-Check $s
        $r.decision | Should -Be 'BLOCK'
        @($r.reasons.rule) | Should -Contain 'last-plan2-removed'
        @($r.reasons.rule) | Should -Contain 'not-shared'
    }
}

Describe 'Fix 5: the green light names the exact batch' {
    It 'Group only, direct Plan 2 kept: ALLOW' {
        (Invoke-Check (Get-Fixture 'scenario-2-second-plan2-kept')).decision | Should -Be 'ALLOW'
    }
    It 'Same group plus the direct Plan 2 SKU: BLOCK' {
        $s = Get-Fixture 'scenario-2-second-plan2-kept'
        $s.planned_removals.license_sku_ids = @($script:Plan2)
        (Invoke-Check $s).decision | Should -Be 'BLOCK'
    }
    It 'Direct Plan 2 SKU removed, no group in the batch: BLOCK, remediation does not talk about a group' {
        $s = Get-Fixture 'scenario-1-firm-staff-only-plan2'
        $s.graph.user.licenseAssignmentStates = @((New-Plan2 $null))
        $s.planned_removals.groups = @()
        $s.planned_removals.license_sku_ids = @($script:Plan2)
        $r = Invoke-Check $s
        $r.decision | Should -Be 'BLOCK'
        $r.reasons[0].remediation | Should -Not -Match 'group'
        $r.reasons[0].text | Should -Match 'direct assignment'
    }
}

Describe 'Fix 6: BLOCK and a hold together' {
    It 'Both reasons listed, remediation starts with keeping Plan 2, size reduction only with the hold owner' {
        $s = Get-Fixture 'scenario-1-firm-staff-only-plan2'
        $s.exchange.mailbox.LitigationHoldEnabled = $true
        $r = Invoke-Check $s
        $r.decision | Should -Be 'BLOCK'
        @($r.reasons.rule) | Should -Be @('last-plan2-removed', 'hold-present')
        $r.reasons[0].remediation | Should -Match '^Keep Plan 2 assigned independently'
        $r.reasons[0].remediation | Should -Match 'only after whoever owns the hold approves'
        $r.reasons[0].remediation | Should -Not -Match 'Or bring'
    }
}

Describe 'Fix 7 and 8: limits and wording at the line' {
    It 'Shared mailbox above 100 GB with Plan 2 kept: REVIEW over-plan2-ceiling' {
        $s = Get-Fixture 'scenario-2-second-plan2-kept'
        $s.exchange.mailbox_statistics.TotalItemSize = [int64]107374182401
        $r = Invoke-Check $s
        $r.decision | Should -Be 'REVIEW'
        @($r.reasons.rule) | Should -Be @('over-plan2-ceiling')
        $r.reasons[0].text | Should -Match 'nominal limit'
    }
    It 'Exactly 50 GB: the reason says it reaches the quota, in bytes' {
        $s = Get-Fixture 'scenario-1-firm-staff-only-plan2'
        $s.exchange.mailbox_statistics.TotalItemSize = [int64]53687091200
        $t = (Invoke-Check $s).reasons[0].text
        $t | Should -Match 'reaches the 50 GB quota at exactly 53,687,091,200 bytes'
        $t | Should -Not -Match 'at 50 GB'
    }
    It 'One byte over: the reason gives bytes, never "at 50 GB"' {
        $s = Get-Fixture 'scenario-1-firm-staff-only-plan2'
        $s.exchange.mailbox_statistics.TotalItemSize = [int64]53687091201
        $t = (Invoke-Check $s).reasons[0].text
        $t | Should -Match 'goes over its 50 GB quota at 53,687,091,201 bytes'
        $t | Should -Not -Match 'at 50 GB'
    }
}

Describe 'Fix 9: holds, from the worldwide Purview page' {
    It 'Every hold reason cites the worldwide page, not the 21Vianet one' {
        $s = Get-Fixture 'scenario-2-second-plan2-kept'
        $s.exchange.mailbox.DelayHoldApplied = $true
        $r = Invoke-Check $s
        $r.reasons[0].source | Should -Be 'https://learn.microsoft.com/en-us/purview/edisc-hold-types-mailboxes'
    }
    It 'Organization grp entry: REVIEW, described with the page''s words' {
        $s = Get-Fixture 'scenario-2-second-plan2-kept'
        $s.exchange.organization_config.InPlaceHolds = @('grp1a0a132ee8944501a4bb6a452ec31171:3')
        $r = Invoke-Check $s
        $r.decision | Should -Be 'REVIEW'
        $r.reasons[0].hold_type | Should -Be 'organization-wide retention policy applied to Microsoft 365 groups and Teams channel messages'
        $r.reasons[0].text | Should -Not -Match 'can cover'
    }
    It 'Organization mbx policy with a matching -mbx exclusion on the mailbox: still REVIEW, the exclusion is quoted' {
        $s = Get-Fixture 'scenario-2-second-plan2-kept'
        $s.exchange.organization_config.InPlaceHolds = @('mbxe9b52bf7ab3b46a286308ecb29624696:2')
        $s.exchange.mailbox.InPlaceHolds = @('-mbxe9b52bf7ab3b46a286308ecb29624696')
        $r = Invoke-Check $s
        $r.decision | Should -Be 'REVIEW'
        $r.reasons[0].mailbox_exclusion | Should -Be '-mbxe9b52bf7ab3b46a286308ecb29624696'
    }
}

Describe 'Fix 11 and 13: labels' {
    It 'Group without displayName: label is the GUID, no double space' {
        $s = Get-Fixture 'scenario-1-firm-staff-only-plan2'
        $s.graph.Remove('groups')
        $r = Invoke-Check $s
        $r.decision | Should -Be 'BLOCK'
        $r.reasons[0].text | Should -Match "group $($script:FirmStaff)"
        $r.reasons[0].text | Should -Not -Match '  '
    }
    It 'Only Plan 2 is ActiveWithError and removed: the reason says healthy (Active, no error)' {
        $s = Get-Fixture 'scenario-1-firm-staff-only-plan2'
        $s.graph.user.licenseAssignmentStates[1].state = 'ActiveWithError'
        $s.graph.user.licenseAssignmentStates[1].error = 'CountViolation'
        $r = Invoke-Check $s
        $r.decision | Should -Be 'REVIEW'
        $r.reasons[0].text | Should -Match 'no healthy \(Active, no error\)'
    }
    It 'A JSON file holding a bare string: "not a JSON object"' {
        $p = Join-Path $TestDrive 'x.json'
        Set-Content -LiteralPath $p -Value '"x"'
        $r = Test-OffboardingRemovalSafety -AadUserId $script:User -TenantId $script:Tenant -SnapshotPath $p
        $r.success | Should -BeFalse
        $r.reasons[0].text | Should -Match 'not a JSON object'
    }
    It 'A truncated file: "not valid JSON"' {
        $p = Join-Path $TestDrive 't.json'
        Set-Content -LiteralPath $p -Value '{"graph": {'
        (Test-OffboardingRemovalSafety -AadUserId $script:User -TenantId $script:Tenant -SnapshotPath $p).reasons[0].text | Should -Match 'not valid JSON'
    }
}

Describe 'No retired modules' {
    It 'No MSOnline or AzureAD cmdlet name appears in the repository' {
        $pattern = '\b[A-Za-z]+-(' + 'Ms' + 'ol|Azure' + 'AD)[A-Za-z]*'
        $files = Get-ChildItem -LiteralPath $script:Root -Recurse -File -Include *.ps1, *.psm1, *.md, *.json |
            Where-Object { $_.FullName -notmatch '[\\/]\.git[\\/]' }
        $hits = $files | Select-String -Pattern $pattern
        $hits | Should -BeNullOrEmpty
    }
}
