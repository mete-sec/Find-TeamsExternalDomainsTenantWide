# Find Teams External Domains Tenant-Wide — Exchange Online + Unified Audit Log

A PowerShell script that generates tenant-wide reports for Microsoft Teams external/federated domain communication by using Exchange Online recipients and Microsoft 365 Unified Audit Log.

The script answers a practical investigation question:

> Which internal users communicated with which external Microsoft Teams domains?

It does **not** use Microsoft Graph. It relies on:

- `Get-EXORecipient` to enumerate mail-enabled internal recipients
- `Search-UnifiedAuditLog` to retrieve Microsoft Teams audit events
- `ParticipantInfo.ParticipatingDomains` to identify domains participating in Teams conversations

---

## Why this script exists

Microsoft Teams external access and federation can create visibility challenges in large tenants. Security, compliance, and messaging teams may need to understand which users have communicated with external Teams domains, especially when reviewing external collaboration exposure, investigating suspicious domains, or preparing customer-facing reports.

This script collects Teams audit events, extracts participating domains, removes the internal domains you define, and exports several CSV reports that summarize the findings by user, external domain, chat thread, actor, and raw message-level audit event.

---

## Key features

- Tenant-wide Teams external/federated domain reporting
- No Microsoft Graph dependency
- Uses Exchange Online and Unified Audit Log only
- Maps Teams audit records to internal users through multiple correlation methods:
  - Actor `UserId`
  - `UserKey` / `ExternalDirectoryObjectId`
  - GUIDs found in `ChatThreadId`, `ItemName`, and raw `AuditData`
  - Email addresses found in raw `AuditData`
- Removes internal domains from `ParticipantInfo.ParticipatingDomains`
- Produces investigation-friendly CSV outputs
- Supports optional raw audit data export
- Supports optional related Teams audit operations beyond `MessageSent`
- Includes unattributed thread reporting for manual review

---

## Requirements

The account and PowerShell session running the script must be able to use:

- Exchange Online PowerShell
- `Connect-ExchangeOnline`
- `Get-EXORecipient`
- `Search-UnifiedAuditLog`

Required PowerShell module:

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
```

Required Microsoft 365 permissions typically include access to:

- Exchange Online recipient data
- Microsoft 365 Unified Audit Log

A practical role combination is usually:

- Exchange Admin
- View-Only Audit Logs

Unified Audit Log must be enabled in the tenant.

You can check this with:

```powershell
Get-AdminAuditLogConfig | Format-List UnifiedAuditLogIngestionEnabled
```

Expected result:

```text
UnifiedAuditLogIngestionEnabled : True
```

---

## Quick start

```powershell
.\Find-TeamsExternalDomainsTenantWide-EXO2.ps1 `
  -StartUtc "2026-06-03T00:00:00" `
  -EndUtc "2026-06-03T21:00:00" `
  -InternalDomains @("example.com") `
  -ExchangeUserPrincipalName "admin@example.com" `
  -OutputFolder ".\TeamsTenantAuditReport"
```

Example with multiple internal domains:

```powershell
.\Find-TeamsExternalDomainsTenantWide-EXO2.ps1 `
  -StartUtc "2026-06-01T00:00:00" `
  -EndUtc "2026-06-07T23:59:59" `
  -InternalDomains @("example.com", "example.com.tr", "subsidiary.com") `
  -ExchangeUserPrincipalName "admin@example.com" `
  -OutputFolder ".\TeamsTenantAuditReport"
```

Example including raw `AuditData`:

```powershell
.\Find-TeamsExternalDomainsTenantWide-EXO2.ps1 `
  -StartUtc "2026-06-01T00:00:00" `
  -EndUtc "2026-06-07T23:59:59" `
  -InternalDomains @("example.com") `
  -ExchangeUserPrincipalName "admin@example.com" `
  -OutputFolder ".\TeamsTenantAuditReport" `
  -IncludeRawAuditData
```

Example using an already connected Exchange Online session:

```powershell
Connect-ExchangeOnline -UserPrincipalName admin@example.com

.\Find-TeamsExternalDomainsTenantWide-EXO2.ps1 `
  -StartUtc "2026-06-01T00:00:00" `
  -EndUtc "2026-06-07T23:59:59" `
  -InternalDomains @("example.com") `
  -SkipExchangeConnect
```

---

## Parameters

| Parameter | Required | Default | Description |
|---|---:|---|---|
| `StartUtc` | Yes | None | Start date/time for the Unified Audit Log search. Use UTC. |
| `EndUtc` | Yes | None | End date/time for the Unified Audit Log search. Use UTC. Must be later than `StartUtc`. |
| `InternalDomains` | Yes | None | One or more internal domains. These domains are removed from participating domains. Remaining domains are treated as external. |
| `ExchangeUserPrincipalName` | No | None | UPN used by `Connect-ExchangeOnline`. If omitted, interactive/default connection behavior is used. |
| `OutputFolder` | No | `.\TeamsTenantAuditReport` | Folder where CSV reports and a generated `README.txt` are written. |
| `RecipientTypeDetails` | No | `UserMailbox`, `SharedMailbox`, `MailUser` | Exchange recipient types used for internal user mapping. |
| `IncludeRelatedEvents` | No | Disabled | Also searches selected related Teams operations in addition to `MessageSent`. |
| `IncludeRawAuditData` | No | Disabled | Includes raw `AuditData` JSON in the detailed message CSV. Review before sharing externally. |
| `SkipExchangeConnect` | No | Disabled | Skips `Connect-ExchangeOnline`. Use this if you already connected manually. |
| `MaxBatches` | No | `50` | Maximum number of `Search-UnifiedAuditLog` batches to retrieve. Each batch requests up to 5000 records. |

---

## How it works

At a high level, the script follows this flow:

1. Validates the date range and normalizes internal domains.
2. Imports the `ExchangeOnlineManagement` module.
3. Connects to Exchange Online unless `-SkipExchangeConnect` is used.
4. Pulls mail-enabled internal recipients using `Get-EXORecipient`.
5. Builds internal user lookup tables by object ID, UPN, primary SMTP address, and proxy SMTP addresses.
6. Searches Unified Audit Log for Microsoft Teams audit events.
7. Parses `AuditData` JSON from each audit record.
8. Extracts `ParticipantInfo.ParticipatingDomains`.
9. Removes internal domains passed through `-InternalDomains`.
10. Maps Teams events to internal users through actor, object ID, GUID, and email correlation.
11. Aggregates results by chat thread, user, external domain, and actor.
12. Exports CSV reports.

---

## Generated files

The script creates timestamped CSV files in the selected output folder.

### `TeamsTenant_UserExternalDomainSummary_*.csv`

The most important report.

Shows which internal user communicated with which external Teams domain.

Useful columns:

- `UserPrincipalName`
- `ExternalDomain`
- `FirstSeen`
- `LastSeen`
- `ThreadCount`
- `EventCount`
- `ChatThreadIds`
- `Actors`
- `ClientIPs`
- `Directions`
- `ParticipatingDomains`
- `ParticipatingTenantIds`

Use this first for most investigations and customer-facing summaries.

---

### `TeamsTenant_UserSummary_*.csv`

User-based summary.

Shows how many external domains each internal user communicated with.

Useful columns:

- `UserPrincipalName`
- `ExternalDomainCount`
- `ExternalDomains`
- `ThreadCount`
- `EventCount`
- `FirstSeen`
- `LastSeen`

Useful for identifying users with broad external Teams communication.

---

### `TeamsTenant_ExternalDomainSummary_*.csv`

External domain-based summary.

Shows which external domains appeared and how many internal users were involved.

Useful columns:

- `ExternalDomain`
- `UserCount`
- `Users`
- `ThreadCount`
- `EventCount`
- `FirstSeen`
- `LastSeen`

Useful for suspicious domain review, external collaboration assessment, and allow/block decisions.

---

### `TeamsTenant_ThreadSummary_*.csv`

Chat thread-based summary.

Shows the internal users and external domains detected in each Teams chat thread.

Useful columns:

- `ChatThreadId`
- `InternalUsers`
- `ExternalDomains`
- `Actors`
- `ClientIPs`
- `ParticipatingDomains`
- `ParticipatingTenantIds`
- `FirstSeen`
- `LastSeen`
- `EventCount`

Useful for evidence validation and thread-level correlation.

---

### `TeamsTenant_MessageDetails_*.csv`

Most detailed audit event report.

Contains one row per parsed Teams audit event.

Useful columns:

- `CreationTime`
- `Operation`
- `Direction`
- `ActorUserId`
- `ClientIP`
- `ChatThreadId`
- `MessageId`
- `ExternalDomainsInEvent`
- `InternalUsersInEvent`
- `ParticipatingDomains`
- `ParticipatingTenantIds`
- `RawAuditData`

If `-IncludeRawAuditData` is used, this file contains raw audit JSON.

Use this file for deep technical analysis, troubleshooting, and evidence preservation.

---

### `TeamsTenant_ActorSummary_*.csv`

Actor-based summary.

Shows which users appeared as actors in audit records.

Useful columns:

- `ActorUserId`
- `ActorDomain`
- `IsActorInternal`
- `FirstSeen`
- `LastSeen`
- `EventCount`
- `Directions`
- `InternalUsersInEvent`
- `ExternalDomainsInEvent`
- `ChatThreadIds`
- `ClientIPs`

Useful for actor-level investigation and understanding whether external users appeared directly as actors.

---

### `TeamsTenant_EXORecipients_*.csv`

Internal recipient inventory pulled from Exchange Online.

This file is used for internal user mapping.

Useful columns:

- `Id`
- `DisplayName`
- `UserPrincipalName`
- `PrimarySmtpAddress`
- `ProxySmtpAddresses`
- `RecipientTypeDetails`
- `Domains`

Useful for debugging why a user was or was not mapped.

---

### `TeamsTenant_UnattributedExternalThreads_*.csv`

Contains threads where external domains were detected but the script could not reliably map the thread to an internal user.

Ideally, this file should be empty.

Possible reasons for unattributed rows:

- Actor `UserId` did not map to an internal recipient
- `UserKey` did not match `ExternalDirectoryObjectId`
- `ChatThreadId` or raw `AuditData` did not contain a mappable internal GUID
- Raw `AuditData` did not contain an internal email address
- The internal user was not returned by `Get-EXORecipient`
- The account is not mail-enabled
- The account uses a recipient type not included in `RecipientTypeDetails`

---

## Important notes

### Audit log delay

Microsoft Teams audit events may not appear immediately in Unified Audit Log. Very recent messages may require additional time before they are searchable.

### Internal domain accuracy matters

The `-InternalDomains` parameter is critical.

If you forget to include one of your organization’s internal domains, that domain may be incorrectly reported as external.

Example:

```powershell
-InternalDomains @("company.com", "company.com.tr", "subsidiary.com")
```

### ClientIP interpretation

`ClientIP` is audit metadata for the actor in that audit record.

Do not directly treat it as the user’s physical location without validation. VPNs, proxies, mobile networks, Microsoft service behavior, and federation behavior may affect interpretation.

### Raw audit data sensitivity

When `-IncludeRawAuditData` is used, the detailed message report may contain sensitive metadata, including:

- IP addresses
- Tenant IDs
- Message IDs
- Chat thread IDs
- Raw audit JSON
- User and domain identifiers

Review the generated files carefully before sharing them externally.

---

## Troubleshooting

### `ExchangeOnlineManagement module is not installed`

Install the module:

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
```

Then re-run the script.

---

### `Search-UnifiedAuditLog is not available`

Try connecting manually:

```powershell
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -UserPrincipalName admin@example.com
Get-Command Search-UnifiedAuditLog
```

If the command is still unavailable, verify that the account has the required audit log permissions.

---

### `Get-EXORecipient is not available`

Try connecting manually:

```powershell
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -UserPrincipalName admin@example.com
Get-Command Get-EXORecipient
```

If the command is unavailable, verify Exchange Online PowerShell connectivity and permissions.

---

### No external domains found

Possible reasons:

- No Teams external/federated communication occurred in the selected time range
- Audit events have not appeared yet
- The selected time range is too narrow
- `InternalDomains` contains incorrect or overly broad values
- The relevant Teams audit events are not available in Unified Audit Log

Try widening the time range and re-running the script.

---

### External domains detected but no internal user mapped

Check:

```text
TeamsTenant_UnattributedExternalThreads_*.csv
```

Then validate:

- Whether the user is mail-enabled
- Whether their recipient type is included in `RecipientTypeDetails`
- Whether the user exists in `TeamsTenant_EXORecipients_*.csv`
- Whether raw `AuditData` should be included for deeper troubleshooting

---

## Security and privacy considerations

The generated reports can contain sensitive tenant metadata. Store them securely and share them only with authorized people.

Recommended handling:

- Keep raw audit data disabled unless needed
- Sanitize CSV files before sending them to customers or third parties
- Avoid publishing generated reports publicly
- Treat IP addresses, tenant IDs, object IDs, message IDs, and chat thread IDs as sensitive investigation data

---

## Limitations

- The script depends on what is available in Microsoft 365 Unified Audit Log.
- It does not retrieve Teams message content.
- It does not call Microsoft Graph.
- It does not modify Teams, Exchange, or tenant configuration.
- Mapping accuracy depends on Exchange recipient data, audit event structure, and the internal domains provided by the user.
- Some external-domain threads may require manual review if the audit data cannot be reliably mapped to an internal user.

---

## Suggested investigation workflow

1. Start with `TeamsTenant_UserExternalDomainSummary_*.csv`.
2. Review `TeamsTenant_ExternalDomainSummary_*.csv` for suspicious or unexpected external domains.
3. Use `TeamsTenant_UserSummary_*.csv` to identify users with broad external collaboration.
4. Validate suspicious rows with `TeamsTenant_ThreadSummary_*.csv`.
5. Use `TeamsTenant_MessageDetails_*.csv` only when deeper technical evidence is needed.
6. Check `TeamsTenant_UnattributedExternalThreads_*.csv` for rows that need manual correlation.

---

## Disclaimer

This script is intended for security, compliance, audit, and administrative review of Microsoft Teams external/federated communication metadata in tenants where you are authorized to perform such analysis.

Review the output carefully before making enforcement, disciplinary, or legal decisions based on the generated reports.
