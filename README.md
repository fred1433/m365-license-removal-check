# m365-license-removal-check

A read-only PowerShell check that runs **before** an offboarding workflow removes a departing user's groups and licenses, once their mailbox has been converted to a shared mailbox.

It reads one JSON snapshot of the user (Microsoft Graph) and of the mailbox (Exchange Online), plus the **whole batch** of planned group and license removals, and answers `ALLOW`, `BLOCK` or `REVIEW`, with each reason and its source.

Everything here runs on a **fictional tenant**: the firm, the person, the group and every ID are invented. The check connects to nothing, writes nothing, holds no secret and calls no AI model. Model API cost: $0.

Live walkthrough of the three tickets: https://theaipipe.com/demos/m365-license-removal-check/

## The case it is built for

A partner leaves a firm. Sign-in is blocked, sessions are revoked, and her 62 GB mailbox is converted and reads back as `SharedMailbox`. Next step on the ticket: remove her from the group **Firm Staff**.

If Firm Staff is the group that carries her only active Exchange Online (Plan 2) license, removing her from it removes the license too ([group licensing](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/manage-group-licenses)). A shared mailbox without a license is limited to 50 GB ([about shared mailboxes](https://learn.microsoft.com/en-us/microsoft-365/admin/email/about-shared-mailboxes)). The mailbox stays shared, but goes over quota. The check flags that removal before it runs.

The check only holds back the group and license cleanup. Blocking sign-in and revoking sessions go ahead as planned.

This is not a new idea: the Exchange team has written about shared mailbox sizing and quotas since 2018 ([Exchange team blog](https://techcommunity.microsoft.com/blog/exchange/correcting-shared-mailbox-provisioning-and-sizing/607991)). What this repo does is put the question in front of the group removal, on the whole batch.

## Reproduce

Requires PowerShell 7 and Pester 5.5 or later (tested with PowerShell 7.6 and Pester 6.2 on macOS).

```powershell
Install-Module Pester -MinimumVersion 5.5 -Scope CurrentUser   # once
pwsh ./run-tests.ps1
```

The test run writes the three page scenarios to `examples/output/`. Those JSON files are committed exactly as the run produced them; the demo page displays them as they are.

Run the check on one snapshot:

```powershell
Import-Module ./src/LicenseRemovalCheck.psm1
Test-OffboardingRemovalSafety `
  -AadUserId 'a41f6b2e-8c3d-4f7a-b1e9-52d06c8e7f13' `
  -TenantId  '7d3e5c1a-2b4f-4e8a-9c6d-0f1e2a3b4c5d' `
  -SnapshotPath ./fixtures/scenario-1-firm-staff-only-plan2.json | ConvertTo-Json -Depth 10
```

## Contract

Inspired by Rewst's public documentation, not an export.

| | Field | Meaning |
|---|---|---|
| in | `aad_user_id` | Entra object ID of the user, the input name used by the public [Remove Licenses](https://docs.rewst.help/documentation/automations/subworkflows/rewst-task-m365-remove-licenses) subworkflow doc |
| in | `tenant_id` | Tenant the workflow runs against. Every block of the snapshot carries its own `tenant_id` and they must match |
| out | `success` | `true` when the data was collected. **It does not mean the removal is allowed.** Read `decision` |
| out | `status` | `1000` all data read; `1001` collected, but something was unreadable (decision is then `REVIEW`); `2000` collection failed (`success` is `false`) |
| out | `automation_log` | One line per step, as a string |
| out | `decision` | `ALLOW`, `BLOCK` or `REVIEW`. Our field, not a Rewst one |
| out | `reasons` | Each with `rule`, `text`, `source`, and `missing_data` when a human has to supply something |
| out | `capacity_kept_by` | For a shared mailbox over 50 GB: the Plan 2 assignments that survive the batch |
| out | `writes` | Always `0` |

About `status`: the public Rewst docs do not define these codes. The [User Offboarding crate](https://docs.rewst.help/documentation/crates/existing-crate-documentation/microsoft-user-offboarding-crate) example log shows `Status: 1000` on completed steps and `Status: 1001` on a step "completed with warning", and the Remove Licenses page uses `status < 2000` when it publishes `automation_log`. The meanings above are this repo's reading of those examples.

## Rules and sources

`BLOCK` wins over `REVIEW`, which wins over `ALLOW`. `ALLOW` only when no rule below fires. All sources read on 2026-11-09.

| Rule | Decision | Source |
|---|---|---|
| Batch removes the last **active** Exchange Online (Plan 2) assignment (direct, or inherited through a group in the batch) of a `SharedMailbox` of 50 GB or more. The reason names the group that carries it | BLOCK | [About shared mailboxes](https://learn.microsoft.com/en-us/microsoft-365/admin/email/about-shared-mailboxes): "A shared mailbox can store up to 50 GB of data without assigning a license to it." "To increase the size limit to 100 GB, the shared mailbox must be assigned an Exchange Online Plan 2 license." |
| Remediation: assign Plan 2 independently of the group, confirm it is `Active`, check again; or bring the mailbox under 50 GB, check again | | [Convert a user mailbox](https://learn.microsoft.com/en-us/microsoft-365/admin/email/convert-user-mailbox-to-shared-mailbox): "If the mailbox is smaller than 50 GB, you can remove the license from the user". [Group licensing](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/manage-group-licenses): "If you remove the user from the source group first, the user is unlicensed until group-based licensing finishes processing the new group's assignment." |
| What happens over quota | | [Exchange team blog](https://techcommunity.microsoft.com/blog/exchange/correcting-shared-mailbox-provisioning-and-sizing/607991): "sending and receiving of email will not work (incoming email will error with "Mailbox is full" error, code 5.2.2)" |
| 50 GB is `53,687,091,200` bytes | | Exchange sizes are binary: [Set-Mailbox](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/set-mailbox) gives "1.999999999 terabytes (2199023254528 bytes)". Tested one byte under, at, and one byte over |
| Which licenses count as Plan 2: any SKU that contains service plan `EXCHANGE_S_ENTERPRISE` (`efb87545-963c-4e0d-99df-69c6916d9eb0`) not listed in `disabledPlans`. Business Premium does not count | | [Licensing service plan reference](https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference), CSV copied to `data/licensing-reference.json` (49 SKUs contain the plan) |
| License assignment fields `assignedByGroup` (null when direct), `skuId`, `state` (`Active`, `ActiveWithError`, `Disabled`, `Error`), `error`, `disabledPlans` | | [licenseAssignmentState](https://learn.microsoft.com/en-us/graph/api/resources/licenseassignmentstate?view=graph-rest-1.0) |
| Replacement Plan 2 outside the batch exists but is not `Active` with no error | REVIEW | same page |
| Conversion requested but `RecipientTypeDetails` is still `UserMailbox` | REVIEW | [Convert a user mailbox](https://learn.microsoft.com/en-us/microsoft-365/admin/email/convert-user-mailbox-to-shared-mailbox) |
| Online archive present (`ArchiveGuid` not empty, or `ArchiveDatabase` set) | REVIEW | [About shared mailboxes](https://learn.microsoft.com/en-us/microsoft-365/admin/email/about-shared-mailboxes): an archive needs "an Exchange Online Plan 2 license or an Exchange Online Plan 1 license with an Exchange Online Archiving add-on license". Read `ArchiveGuid` and `ArchiveDatabase`, not `ArchiveStatus`: [ArchiveStatus set to None](https://learn.microsoft.com/en-us/troubleshoot/microsoft-365/purview/archive-mailboxes/archivestatus-set-none) |
| `LitigationHoldEnabled`, `ComplianceTagHoldApplied`, `DelayHoldApplied` or `DelayReleaseHoldApplied` is true | REVIEW | [Identify a hold on a mailbox](https://learn.microsoft.com/en-us/purview/ediscovery-identify-a-hold-on-an-exchange-online-mailbox) |
| `InPlaceHolds` on the mailbox, or on the organization (`Get-OrganizationConfig`), lists an entry. The reason says `mailbox` or `organization`, and the hold type when the prefix is one the page lists (`UniH`, `cld`, `mbx`, `skp`, `grp`; `-mbx` is an exclusion and is only logged) | REVIEW | same page: "If the InPlaceHolds property is empty when you run the Get-Mailbox cmdlet, there still may be one or more organization-wide Microsoft Purview retention policies applied to the mailbox." |
| Organization `InPlaceHolds` not readable: retention scope unknown | REVIEW | same page |
| `tenant_id` differs between the request, Graph data and Exchange data | REVIEW | |
| `onPremisesSyncEnabled` is true | REVIEW | [user resource](https://learn.microsoft.com/en-us/graph/api/resources/user?view=graph-rest-1.0) |
| Size missing or not a whole number of bytes; any hold or archive field not read | REVIEW | |

Note on the hold page: it now opens with "The guidance in this article only applies to organizations hosted in Microsoft 365 operated by 21Vianet (China)." It is still the page that documents these mailbox properties and prefixes; check them against your own tenant before relying on them.

## Scope

Cloud identities only; hybrid is out of scope. When an identity is synchronized from on-premises Active Directory, a mailbox converted in the cloud can be converted back at the next directory synchronization, and Microsoft sends hybrid setups to a separate procedure (see "Convert a user's mailbox in a hybrid environment" on the [conversion page](https://learn.microsoft.com/en-us/microsoft-365/admin/email/convert-user-mailbox-to-shared-mailbox)). The check returns `REVIEW` for a synced user rather than guess.

## Collecting a real snapshot

This repo does not ship a collector. The snapshot fields map to these reads (all read-only):

| Snapshot | Read with |
|---|---|
| `graph.user.licenseAssignmentStates`, `onPremisesSyncEnabled` | Graph `GET /users/{id}?$select=id,onPremisesSyncEnabled,licenseAssignmentStates` (both need `$select`) |
| `exchange.mailbox.*` | `Get-Mailbox` |
| `exchange.mailbox_statistics.TotalItemSize` | `Get-EXOMailboxStatistics`, written as a whole number of **bytes**. The cmdlet's reference page does not document the format of `TotalItemSize`, so this repo does not parse text sizes: anything other than bytes gives `REVIEW` |
| `exchange.organization_config.InPlaceHolds` | `Get-OrganizationConfig` |

## Wiring into a Rewst workflow (not validated)

- The check sits **before** group removal and license removal.
- Only `ALLOW` on the batch that was checked lets the workflow continue to those two steps. `BLOCK` and `REVIEW` stop there and hand the reasons to a person.
- The public doc of the [Remove Group Memberships](https://docs.rewst.help/documentation/automations/subworkflows/rewst-task-remove-group-memberships) task does not describe a per-group exclusion. So one protected group holds back the whole group removal task, which is why the check judges the batch as a whole.
- Importing and running this inside Rewst has not been validated.

## Files

```
src/LicenseRemovalCheck.psm1      the check (Test-OffboardingRemovalSafety)
tests/LicenseRemovalCheck.Tests.ps1
fixtures/                         three fictional snapshots, one per page scenario
examples/output/                  JSON written by the test run, committed as produced
data/licensing-reference.json     Plan 2 SKUs and product names, from Microsoft's licensing CSV
run-tests.ps1
```

## Author

Frederic de Lavenne de Choulot, The AI Pipe. MIT license.
