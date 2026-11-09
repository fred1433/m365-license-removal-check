# m365-license-removal-check

A read-only PowerShell check that runs **before** an offboarding workflow removes a departing user's groups and licenses, once their mailbox has been converted to a shared mailbox.

It reads one JSON snapshot of the user (Microsoft Graph) and of the mailbox (Exchange Online), plus the **whole batch** of planned group and license removals, and answers `ALLOW`, `BLOCK` or `REVIEW`, with each reason and its source. When a piece of data it needs is missing or unreadable, it does not decide: `REVIEW`.

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
| out | `success`, `status` | The **quality of the data collection**, nothing else. `1000` everything read; `1001` collected, but something was missing or unreadable; `2000` collection failed (`success` is `false`: no snapshot, no user, or a mailbox or size statistics that are not this user's) |
| out | `decision` | **The only field that says whether the batch may run.** `ALLOW`, `BLOCK` or `REVIEW`. Our field, not a Rewst one |
| out | `reasons` | Each with `rule`, `text`, `source`, and `missing_data` when a human has to supply something |
| out | `capacity_kept_by` | For a mailbox of 50 GB or more: the healthy Plan 2 assignments that survive the batch |
| out | `checked_batch` | The groups and licenses that were judged, as read from `planned_removals` |
| out | `automation_log` | One line per step, as a string |
| out | `writes` | Always `0` |

`BLOCK` wins over `REVIEW`, which wins over `ALLOW`, and the two fields are independent: a `BLOCK` can come out with `status` `1001` when something else was also unreadable, and a `REVIEW` can come with `1000` when everything was read and a person still has to decide (a hold, for example).

About the codes: the public Rewst docs do not define them. The [User Offboarding crate](https://docs.rewst.help/documentation/crates/existing-crate-documentation/microsoft-user-offboarding-crate) example log shows `Status: 1000` on completed steps and `Status: 1001` on a step "completed with warning", and the Remove Licenses page uses `status < 2000` when it publishes `automation_log`. The meanings above are this repo's reading of those examples.

**The `ALLOW` covers the exact batch that was checked**, as listed in `checked_batch`. In the demo's safe ticket that is "remove Firm Staff only, keep the direct Exchange Online (Plan 2)". If the workflow cannot guarantee it will run that exact batch (another group, another license, a later step that removes "all remaining licenses"), it runs the check again on the new batch.

## Rules and sources

`ALLOW` only when no rule below fires and the mailbox reads back as `SharedMailbox`. All sources read on 2026-11-09.

| Rule | Decision | Source |
|---|---|---|
| Batch removes the last **healthy** (state `Active`, no error) Exchange Online (Plan 2) assignment, direct or inherited through a group in the batch, of a `SharedMailbox` or `UserMailbox` of 50 GB or more. The reason names the group that carries it, or says it is a direct assignment. Exception: if a replacement Plan 2 outside the batch exists but is not yet healthy, the decision is `REVIEW` instead (next rows) | BLOCK | [About shared mailboxes](https://learn.microsoft.com/en-us/microsoft-365/admin/email/about-shared-mailboxes): "A shared mailbox can store up to 50 GB of data without assigning a license to it." "To increase the size limit to 100 GB, the shared mailbox must be assigned an Exchange Online Plan 2 license." |
| Remediation, group carrier: keep Plan 2 assigned independently of the group, confirm it is `Active`, check again; or bring the mailbox under 50 GB, check again. Direct carrier: take the direct license out of the batch. When a hold is also found, bringing the mailbox under 50 GB is offered only after whoever owns the hold approves | | [Convert a user mailbox](https://learn.microsoft.com/en-us/microsoft-365/admin/email/convert-user-mailbox-to-shared-mailbox): "If the mailbox is smaller than 50 GB, you can remove the license from the user, and stop paying for it." [Group licensing](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/manage-group-licenses): "If you remove the user from the source group first, the user is unlicensed until group-based licensing finishes processing the new group's assignment." |
| What happens over quota | | [Exchange team blog](https://techcommunity.microsoft.com/blog/exchange/correcting-shared-mailbox-provisioning-and-sizing/607991): "sending and receiving of email will not work (incoming email will error with "Mailbox is full" error, code 5.2.2)" |
| 50 GB is `53,687,091,200` bytes, 100 GB is `107,374,182,400`. At exactly 50 GB the reason says the mailbox "reaches" the quota; within 1 GB of the line it gives the size in bytes | | Exchange sizes are binary: [Set-Mailbox](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/set-mailbox) gives "1.999999999 terabytes (2199023254528 bytes)". Tested one byte under, at, and one byte over |
| Which licenses count as Plan 2: any SKU that contains service plan `EXCHANGE_S_ENTERPRISE` (`efb87545-963c-4e0d-99df-69c6916d9eb0`) not listed in `disabledPlans`. Business Premium does not count: it carries Exchange Online (Plan 1) | | [Licensing service plan reference](https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference), CSV copied to `data/licensing-reference.json` (49 SKUs contain the plan) |
| In the fixtures, Business Premium has its Exchange Online (Plan 1) plan (`9aaf7827-d63c-4b61-89c3-182f06f82e5c`) in `disabledPlans`, because Plan 2 is also assigned | | [Group licensing](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/manage-group-licenses), section on conflicting service plans: "Some plans overlap or conflict, such as:" followed by "Exchange Online Plan 1" and "Exchange Online Plan 2", then "Disable one conflicting plan, remove the redundant direct license, or adjust the group license assignment." |
| Replacement Plan 2 outside the batch exists but is not healthy (`Error`, `ActiveWithError`, or an error value) while the batch removes the last healthy one | REVIEW | [licenseAssignmentState](https://learn.microsoft.com/en-us/graph/api/resources/licenseassignmentstate?view=graph-rest-1.0) |
| Shared mailbox of 50 GB or more with no healthy Plan 2 at all | REVIEW | About shared mailboxes |
| Shared mailbox above 100 GB with Plan 2 kept (100 GB is the nominal limit; an admin can set a lower quota) | REVIEW | About shared mailboxes |
| Mailbox reads back as `UserMailbox` (conversion requested or not), another type, or the type was not read | REVIEW | [Convert a user mailbox](https://learn.microsoft.com/en-us/microsoft-365/admin/email/convert-user-mailbox-to-shared-mailbox) |
| Online archive present (`ArchiveGuid` not empty and `ArchiveDatabase` set); only one of the two set: "archive state inconsistent" | REVIEW | [About shared mailboxes](https://learn.microsoft.com/en-us/microsoft-365/admin/email/about-shared-mailboxes): "To increase the size of the archive mailbox, an Exchange Online Plan 2 license or an Exchange Online Plan 1 license with an Exchange Online Archiving add-on license is required." Read `ArchiveGuid` and `ArchiveDatabase`, not `ArchiveStatus`: [ArchiveStatus set to None](https://learn.microsoft.com/en-us/troubleshoot/microsoft-365/purview/archive-mailboxes/archivestatus-set-none) |
| `LitigationHoldEnabled`, `ComplianceTagHoldApplied`, `DelayHoldApplied` or `DelayReleaseHoldApplied` is true | REVIEW | [Identify Exchange mailbox hold types](https://learn.microsoft.com/en-us/purview/edisc-hold-types-mailboxes) |
| `InPlaceHolds` on the mailbox, or on the organization (`Get-OrganizationConfig`), lists an entry. The reason says `mailbox` or `organization`, and gives the type only for the prefixes that page describes (`UniH`, `mbx`, `skp`, `grp`). A `-mbx` entry on the mailbox is an exclusion: it is logged, and quoted in the organization reason when it matches, but the organization entry still goes to a person | REVIEW | same page: "If the InPlaceHolds property is empty when you run the Get-Mailbox cmdlet, one or more organization-wide Microsoft Purview retention policies might still apply to the mailbox." |
| Organization `InPlaceHolds` not readable: the check does not approve the batch | REVIEW | same page |
| `tenant_id` differs between the request, Graph data and Exchange data | REVIEW | |
| `onPremisesSyncEnabled` is true, missing, or not a boolean (`null` and `false` mean cloud identity) | REVIEW | [user resource](https://learn.microsoft.com/en-us/graph/api/resources/user?view=graph-rest-1.0) |
| Data not read: `planned_removals.groups` or `planned_removals.license_sku_ids` missing or not a list (an empty list is fine); `licenseAssignmentStates` missing; an assignment without `skuId`, `state`, `disabledPlans` or the `assignedByGroup` key (only an explicit `null` means direct); a hold flag missing, `null` or not a boolean; mailbox `InPlaceHolds` missing or `null`; archive fields missing; size missing or not a whole number of bytes | REVIEW, `status` 1001 | |
| The mailbox `ExternalDirectoryObjectId` is missing or is not `aad_user_id`; or the statistics `MailboxGuid` and the mailbox `ExchangeGuid` are both present and differ | REVIEW, `success` false, `status` 2000 | [Get-EXOMailbox](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/get-exomailbox) |

## Scope

Cloud identities only; hybrid is out of scope. When an identity is synchronized from on-premises Active Directory, a mailbox converted in the cloud can be converted back at the next directory synchronization, and Microsoft sends hybrid setups to a separate procedure (see "Convert a user's mailbox in a hybrid environment" on the [conversion page](https://learn.microsoft.com/en-us/microsoft-365/admin/email/convert-user-mailbox-to-shared-mailbox)). The check returns `REVIEW` for a synced user rather than guess.

## Collecting a real snapshot

This repo does not ship a collector. The snapshot fields map to these reads (all read-only):

| Snapshot | Read with |
|---|---|
| `graph.user.licenseAssignmentStates`, `onPremisesSyncEnabled` | Graph `GET /users/{id}?$select=id,onPremisesSyncEnabled,licenseAssignmentStates` (both need `$select`) |
| `exchange.mailbox.*` | `Get-Mailbox` / `Get-EXOMailbox` |
| `exchange.mailbox_statistics.TotalItemSize` | `Get-EXOMailboxStatistics`. A choice of this contract: the collector normalizes the size to a whole number of **bytes**, and any other form (text, decimal, negative) gives `REVIEW`. The cmdlet's reference page does not document the format of `TotalItemSize`, so this repo does not parse text sizes |
| `exchange.organization_config.InPlaceHolds` | `Get-OrganizationConfig` |

## Wiring into a Rewst workflow (not validated)

- The check sits **before** group removal and license removal.
- Only `ALLOW` on the batch that was checked lets the workflow continue to those two steps, with that exact batch. `BLOCK` and `REVIEW` stop there and hand the reasons to a person.
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
