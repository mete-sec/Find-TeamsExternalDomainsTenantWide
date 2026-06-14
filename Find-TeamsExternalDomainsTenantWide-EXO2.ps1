<#
.SYNOPSIS
    Tenant-wide Microsoft Teams external/federated domain report using Exchange Online recipients + Unified Audit Log.

.DESCRIPTION
    - Does NOT use Microsoft Graph.
    - Uses Get-EXORecipient to pull mail-enabled internal recipients.
    - Uses Search-UnifiedAuditLog to pull Microsoft Teams audit events.
    - Parses Teams MessageSent audit records.
    - Extracts ParticipantInfo.ParticipatingDomains.
    - Removes internal domains.
    - Maps Teams audit events to internal users by:
        * Actor UserId
        * UserKey / ExternalDirectoryObjectId
        * GUIDs inside ChatThreadId / ItemName / raw AuditData
        * Email addresses inside raw AuditData
    - Exports tenant-wide CSV reports.

.REQUIREMENTS
    - ExchangeOnlineManagement module
    - Connect-ExchangeOnline works
    - Search-UnifiedAuditLog is available
    - Get-EXORecipient is available
    - Audit Logs / View-Only Audit Logs permission
    - Unified Audit Log enabled

.EXAMPLE
    .\Find-TeamsExternalDomainsTenantWide-EXO2.ps1 `
      -StartUtc "2026-06-03T00:00:00" `
      -EndUtc "2026-06-03T21:00:00" `
      -InternalDomains @("example.com") `
      -ExchangeUserPrincipalName "admin@example.com" `
      -OutputFolder ".\TeamsTenantAuditReport" `
      -IncludeRawAuditData
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [datetime]$StartUtc,

    [Parameter(Mandatory = $true)]
    [datetime]$EndUtc,

    [Parameter(Mandatory = $true)]
    [string[]]$InternalDomains,

    [Parameter(Mandatory = $false)]
    [string]$ExchangeUserPrincipalName,

    [Parameter(Mandatory = $false)]
    [string]$OutputFolder = ".\TeamsTenantAuditReport",

    [Parameter(Mandatory = $false)]
    [string[]]$RecipientTypeDetails = @("UserMailbox", "SharedMailbox", "MailUser"),

    [Parameter(Mandatory = $false)]
    [switch]$IncludeRelatedEvents,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeRawAuditData,

    [Parameter(Mandatory = $false)]
    [switch]$SkipExchangeConnect,

    [Parameter(Mandatory = $false)]
    [int]$MaxBatches = 50
)

$ErrorActionPreference = "Stop"

function Normalize-String {
    param($Value)

    if ($null -eq $Value) {
        return ""
    }

    $text = [string]$Value

    if ([string]::IsNullOrWhiteSpace($text)) {
        return ""
    }

    return $text.Trim().ToLowerInvariant()
}

function Convert-ToArraySafe {
    param($Value)

    if ($null -eq $Value) {
        return @()
    }

    if ($Value -is [System.Array]) {
        return @($Value)
    }

    return @($Value)
}

function Get-SafeProperty {
    param(
        $Object,
        [string]$PropertyName
    )

    if ($null -eq $Object) {
        return $null
    }

    $prop = $Object.PSObject.Properties[$PropertyName]

    if ($null -eq $prop) {
        return $null
    }

    return $prop.Value
}

function Get-EmailDomain {
    param($Email)

    $normalizedEmail = Normalize-String $Email

    if ($normalizedEmail -notmatch "@") {
        return ""
    }

    return (($normalizedEmail -split "@")[-1]).Trim().ToLowerInvariant()
}

function Test-IsInternalDomain {
    param(
        $Domain,
        $InternalDomains
    )

    $normalizedDomain = Normalize-String $Domain

    foreach ($internalDomain in (Convert-ToArraySafe $InternalDomains)) {
        if ($normalizedDomain -eq (Normalize-String $internalDomain)) {
            return $true
        }
    }

    return $false
}

function Test-IsInternalEmail {
    param(
        $Email,
        $InternalDomains
    )

    $domain = Get-EmailDomain $Email

    if ([string]::IsNullOrWhiteSpace($domain)) {
        return $false
    }

    return Test-IsInternalDomain -Domain $domain -InternalDomains $InternalDomains
}

function Join-UniqueValues {
    param($Values)

    $cleanValues = @()

    foreach ($value in (Convert-ToArraySafe $Values)) {
        if ($null -ne $value) {
            $text = ([string]$value).Trim()

            if ($text -ne "") {
                $cleanValues += $text
            }
        }
    }

    return ($cleanValues | Sort-Object -Unique) -join "; "
}

function ConvertTo-CompactJson {
    param($Object)

    if ($null -eq $Object) {
        return ""
    }

    try {
        return ($Object | ConvertTo-Json -Compress -Depth 50)
    }
    catch {
        return ""
    }
}

function Get-ExtraPropertyValue {
    param(
        $AuditJson,
        [string]$Key
    )

    $extraProperties = Get-SafeProperty -Object $AuditJson -PropertyName "ExtraProperties"

    foreach ($item in (Convert-ToArraySafe $extraProperties)) {
        $itemKey = Get-SafeProperty -Object $item -PropertyName "Key"
        $itemValue = Get-SafeProperty -Object $item -PropertyName "Value"

        if ((Normalize-String $itemKey) -eq (Normalize-String $Key)) {
            return $itemValue
        }
    }

    return ""
}

function Get-EmailMatchesFromText {
    param($Text)

    if ([string]::IsNullOrWhiteSpace([string]$Text)) {
        return @()
    }

    $emailRegex = '[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}'

    return @(
        [regex]::Matches([string]$Text, $emailRegex) |
            ForEach-Object { Normalize-String $_.Value } |
            Sort-Object -Unique
    )
}

function Get-GuidMatchesFromText {
    param($Text)

    if ([string]::IsNullOrWhiteSpace([string]$Text)) {
        return @()
    }

    $guidRegex = '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}'

    return @(
        [regex]::Matches([string]$Text, $guidRegex) |
            ForEach-Object { Normalize-String $_.Value } |
            Sort-Object -Unique
    )
}

function Export-ObjectsToCsvSafe {
    param(
        $Objects,
        [string]$Path
    )

    $items = @(Convert-ToArraySafe $Objects)

    if ($items.Count -gt 0) {
        $items | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
    }
    else {
        [pscustomobject]@{
            Status = "No data found"
        } | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
    }
}

function Add-UniqueToArray {
    param(
        [object[]]$Array,
        $Value
    )

    $text = ([string]$Value).Trim()

    if ($text -eq "") {
        return @($Array)
    }

    if (@($Array) -notcontains $text) {
        return @($Array + $text)
    }

    return @($Array)
}

function Get-ProxySmtpAddresses {
    param($EmailAddresses)

    $addresses = @()

    foreach ($addr in (Convert-ToArraySafe $EmailAddresses)) {
        $text = [string]$addr

        if ([string]::IsNullOrWhiteSpace($text)) {
            continue
        }

        if ($text -match "^(smtp|SMTP):(.+)$") {
            $smtp = Normalize-String $Matches[2]

            if ($smtp -ne "") {
                $addresses += $smtp
            }
        }
        elseif ($text -match "^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$") {
            $addresses += (Normalize-String $text)
        }
    }

    return @($addresses | Sort-Object -Unique)
}

if ($EndUtc -le $StartUtc) {
    throw "EndUtc must be later than StartUtc."
}

$InternalDomainsNormalized = @()

foreach ($domain in (Convert-ToArraySafe $InternalDomains)) {
    $normalized = Normalize-String $domain

    if ($normalized -ne "") {
        $InternalDomainsNormalized += $normalized
    }
}

$InternalDomainsNormalized = @($InternalDomainsNormalized | Sort-Object -Unique)

Write-Host ""
Write-Host "Start UTC:        $StartUtc"
Write-Host "End UTC:          $EndUtc"
Write-Host "Internal domains: $($InternalDomainsNormalized -join ', ')"
Write-Host "Recipient types:  $($RecipientTypeDetails -join ', ')"
Write-Host ""

if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    throw "ExchangeOnlineManagement module is not installed. Install it with: Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber"
}

Import-Module ExchangeOnlineManagement -Force

if (-not $SkipExchangeConnect) {
    Write-Host "Connecting to Exchange Online..."

    if ([string]::IsNullOrWhiteSpace($ExchangeUserPrincipalName)) {
        Connect-ExchangeOnline -ShowBanner:$false
    }
    else {
        Connect-ExchangeOnline -UserPrincipalName $ExchangeUserPrincipalName -ShowBanner:$false
    }
}

if ($null -eq (Get-Command Search-UnifiedAuditLog -ErrorAction SilentlyContinue)) {
    throw @"
Search-UnifiedAuditLog is not available in this PowerShell session.

Try manually:
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -UserPrincipalName your-admin@domain.com
Get-Command Search-UnifiedAuditLog
"@
}

if ($null -eq (Get-Command Get-EXORecipient -ErrorAction SilentlyContinue)) {
    throw @"
Get-EXORecipient is not available in this PowerShell session.

Try manually:
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -UserPrincipalName your-admin@domain.com
Get-Command Get-EXORecipient
"@
}

# ----------------------------
# Fetch Exchange recipients
# ----------------------------

Write-Host "Fetching mail-enabled internal recipients from Exchange Online..."

$RecipientProperties = @(
    "ExternalDirectoryObjectId",
    "PrimarySmtpAddress",
    "RecipientTypeDetails",
    "DisplayName",
    "EmailAddresses"
)

try {
    $RecipientsRaw = @(
        Get-EXORecipient `
            -ResultSize Unlimited `
            -RecipientTypeDetails $RecipientTypeDetails `
            -Properties $RecipientProperties
    )
}
catch {
    Write-Warning "Get-EXORecipient with -RecipientTypeDetails failed. Falling back to all recipients and local filtering. Error: $($_.Exception.Message)"

    $RecipientsRaw = @(
        Get-EXORecipient `
            -ResultSize Unlimited `
            -Properties $RecipientProperties
    )
}

$InternalRecipients = @()

foreach ($r in $RecipientsRaw) {
    $recipientType = [string](Get-SafeProperty -Object $r -PropertyName "RecipientTypeDetails")

    if ($RecipientTypeDetails.Count -gt 0) {
        if ($RecipientTypeDetails -notcontains $recipientType) {
            continue
        }
    }

    $primarySmtp = Normalize-String (Get-SafeProperty -Object $r -PropertyName "PrimarySmtpAddress")
    $externalDirectoryObjectId = Normalize-String (Get-SafeProperty -Object $r -PropertyName "ExternalDirectoryObjectId")
    $displayName = [string](Get-SafeProperty -Object $r -PropertyName "DisplayName")
    $emailAddresses = Get-SafeProperty -Object $r -PropertyName "EmailAddresses"
    $proxySmtpAddresses = @(Get-ProxySmtpAddresses -EmailAddresses $emailAddresses)

    if ([string]::IsNullOrWhiteSpace($primarySmtp) -and $proxySmtpAddresses.Count -eq 0) {
        continue
    }

    $domainCandidates = @()

    if ($primarySmtp -ne "") {
        $domainCandidates += (Get-EmailDomain $primarySmtp)
    }

    foreach ($proxy in $proxySmtpAddresses) {
        $domainCandidates += (Get-EmailDomain $proxy)
    }

    $domainCandidates = @($domainCandidates | Where-Object { $_ -ne "" } | Sort-Object -Unique)

    $isInternal = $false

    foreach ($candidateDomain in $domainCandidates) {
        if (Test-IsInternalDomain -Domain $candidateDomain -InternalDomains $InternalDomainsNormalized) {
            $isInternal = $true
            break
        }
    }

    if (-not $isInternal) {
        continue
    }

    $primaryIdentity = $primarySmtp

    if ([string]::IsNullOrWhiteSpace($primaryIdentity) -and $proxySmtpAddresses.Count -gt 0) {
        $primaryIdentity = $proxySmtpAddresses[0]
    }

    $InternalRecipients += [pscustomobject]@{
        Id                  = $externalDirectoryObjectId
        DisplayName          = $displayName
        UserPrincipalName    = $primaryIdentity
        PrimarySmtpAddress   = $primarySmtp
        ProxySmtpAddresses   = Join-UniqueValues $proxySmtpAddresses
        RecipientTypeDetails = $recipientType
        Domains              = Join-UniqueValues $domainCandidates
    }
}

Write-Host "Internal mail-enabled recipients loaded: $($InternalRecipients.Count)"
Write-Host ""

$UserById = @{}
$UserByUpn = @{}
$UserByMail = @{}

foreach ($u in $InternalRecipients) {
    if (-not [string]::IsNullOrWhiteSpace($u.Id)) {
        $UserById[$u.Id] = $u
    }

    if (-not [string]::IsNullOrWhiteSpace($u.UserPrincipalName)) {
        $UserByUpn[$u.UserPrincipalName] = $u
        $UserByMail[$u.UserPrincipalName] = $u
    }

    if (-not [string]::IsNullOrWhiteSpace($u.PrimarySmtpAddress)) {
        $UserByMail[$u.PrimarySmtpAddress] = $u
    }

    foreach ($proxy in @(([string]$u.ProxySmtpAddresses) -split ";")) {
        $cleanProxy = Normalize-String $proxy

        if ($cleanProxy -ne "") {
            $UserByMail[$cleanProxy] = $u
        }
    }
}

# ----------------------------
# Search Unified Audit Log
# ----------------------------

$Operations = @("MessageSent")

if ($IncludeRelatedEvents) {
    $Operations += @(
        "MessageCreatedHasLink",
        "MessageEditedHasLink",
        "GraphMessageUpdated",
        "CreateThreadProbe",
        "ChatCreated",
        "MemberAdded"
    )
}

$Operations = @($Operations | Sort-Object -Unique)

Write-Host "Operations: $($Operations -join ', ')"
Write-Host "Searching Unified Audit Log..."
Write-Host ""

$SessionId = "TeamsTenantExternalDomainsEXO-$([guid]::NewGuid())"
$AllLogs = @()
$SeenIds = @{}
$BatchNumber = 0

do {
    $BatchNumber++

    $RawBatch = Search-UnifiedAuditLog `
        -StartDate $StartUtc `
        -EndDate $EndUtc `
        -Operations $Operations `
        -SessionId $SessionId `
        -SessionCommand ReturnLargeSet `
        -ResultSize 5000

    $Batch = @(Convert-ToArraySafe $RawBatch)
    $NewInBatch = 0

    foreach ($item in $Batch) {
        if ($null -eq $item) {
            continue
        }

        $identity = [string]$item.Identity

        if ([string]::IsNullOrWhiteSpace($identity)) {
            $identity = "$($item.CreationDate)-$($item.UserIds)-$($item.Operations)-$($item.ResultIndex)"
        }

        if (-not $SeenIds.ContainsKey($identity)) {
            $SeenIds[$identity] = $true
            $AllLogs += $item
            $NewInBatch++
        }
    }

    Write-Host "Batch $BatchNumber fetched: $($Batch.Count) records. New unique: $NewInBatch. Total unique: $($AllLogs.Count)"

    if ($BatchNumber -ge $MaxBatches) {
        Write-Warning "Max batch count reached. Stopping."
        break
    }

} while ($Batch.Count -gt 0 -and $NewInBatch -gt 0)

Write-Host ""
Write-Host "Total unique audit records fetched: $($AllLogs.Count)"
Write-Host "Parsing AuditData..."
Write-Host ""

# ----------------------------
# Parse audit records
# ----------------------------

$ParsedEvents = @()

foreach ($record in $AllLogs) {
    $rawAuditData = [string]$record.AuditData

    if ([string]::IsNullOrWhiteSpace($rawAuditData)) {
        continue
    }

    try {
        $auditJson = $rawAuditData | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        Write-Warning "Could not parse AuditData JSON for record: $($record.Identity)"
        continue
    }

    $workload = Normalize-String (Get-SafeProperty -Object $auditJson -PropertyName "Workload")

    if ($workload -ne "microsoftteams") {
        continue
    }

    $operation = Get-SafeProperty -Object $auditJson -PropertyName "Operation"
    $creationTime = Get-SafeProperty -Object $auditJson -PropertyName "CreationTime"
    $actorUserId = Normalize-String (Get-SafeProperty -Object $auditJson -PropertyName "UserId")
    $actorDomain = Get-EmailDomain $actorUserId
    $chatThreadId = [string](Get-SafeProperty -Object $auditJson -PropertyName "ChatThreadId")
    $messageId = Get-SafeProperty -Object $auditJson -PropertyName "MessageId"
    $messageVersion = Get-SafeProperty -Object $auditJson -PropertyName "MessageVersion"
    $clientIp = Get-SafeProperty -Object $auditJson -PropertyName "ClientIP"
    $userAgent = Get-SafeProperty -Object $auditJson -PropertyName "UserAgent"
    $communicationType = Get-SafeProperty -Object $auditJson -PropertyName "CommunicationType"
    $resourceTenantId = Get-SafeProperty -Object $auditJson -PropertyName "ResourceTenantId"
    $organizationId = Get-SafeProperty -Object $auditJson -PropertyName "OrganizationId"
    $userTenantId = Get-SafeProperty -Object $auditJson -PropertyName "UserTenantId"
    $userType = Get-SafeProperty -Object $auditJson -PropertyName "UserType"
    $userKey = Normalize-String (Get-SafeProperty -Object $auditJson -PropertyName "UserKey")
    $itemName = [string](Get-SafeProperty -Object $auditJson -PropertyName "ItemName")

    $participantInfo = Get-SafeProperty -Object $auditJson -PropertyName "ParticipantInfo"

    $participatingDomains = @()
    $participatingTenantIds = @()
    $participatingSipDomainsJson = ""

    if ($null -ne $participantInfo) {
        $participatingDomainsRaw = Get-SafeProperty -Object $participantInfo -PropertyName "ParticipatingDomains"
        $participatingTenantIdsRaw = Get-SafeProperty -Object $participantInfo -PropertyName "ParticipatingTenantIds"
        $participatingSipDomainsRaw = Get-SafeProperty -Object $participantInfo -PropertyName "ParticipatingSIPDomains"

        foreach ($domain in (Convert-ToArraySafe $participatingDomainsRaw)) {
            $normalizedDomain = Normalize-String $domain

            if ($normalizedDomain -ne "") {
                $participatingDomains += $normalizedDomain
            }
        }

        foreach ($tenantId in (Convert-ToArraySafe $participatingTenantIdsRaw)) {
            $tenantText = ([string]$tenantId).Trim()

            if ($tenantText -ne "") {
                $participatingTenantIds += $tenantText
            }
        }

        $participatingSipDomainsJson = ConvertTo-CompactJson $participatingSipDomainsRaw
    }

    $participatingDomains = @($participatingDomains | Sort-Object -Unique)
    $participatingTenantIds = @($participatingTenantIds | Sort-Object -Unique)

    $externalDomainsInEvent = @()

    foreach ($domain in $participatingDomains) {
        if (-not (Test-IsInternalDomain -Domain $domain -InternalDomains $InternalDomainsNormalized)) {
            $externalDomainsInEvent += $domain
        }
    }

    $externalDomainsInEvent = @($externalDomainsInEvent | Sort-Object -Unique)

    $hasForeignTenantUsers = Get-SafeProperty -Object $participantInfo -PropertyName "HasForeignTenantUsers"
    $hasGuestUsers = Get-SafeProperty -Object $participantInfo -PropertyName "HasGuestUsers"
    $hasOtherGuestUsers = Get-SafeProperty -Object $participantInfo -PropertyName "HasOtherGuestUsers"
    $hasUnauthenticatedUsers = Get-SafeProperty -Object $participantInfo -PropertyName "HasUnauthenticatedUsers"

    $allEmailsInRaw = @(Get-EmailMatchesFromText -Text $rawAuditData)
    $allGuidsInRaw = @(Get-GuidMatchesFromText -Text ($rawAuditData + " " + $chatThreadId + " " + $itemName + " " + $userKey))

    $internalUsersInEvent = @()
    $internalRecipientIdsInEvent = @()

    # 1) Actor UPN/mail mapping
    if ($UserByUpn.ContainsKey($actorUserId)) {
        $mapped = $UserByUpn[$actorUserId]
        $internalUsersInEvent = Add-UniqueToArray -Array $internalUsersInEvent -Value $mapped.UserPrincipalName
        $internalRecipientIdsInEvent = Add-UniqueToArray -Array $internalRecipientIdsInEvent -Value $mapped.Id
    }
    elseif ($UserByMail.ContainsKey($actorUserId)) {
        $mapped = $UserByMail[$actorUserId]
        $internalUsersInEvent = Add-UniqueToArray -Array $internalUsersInEvent -Value $mapped.UserPrincipalName
        $internalRecipientIdsInEvent = Add-UniqueToArray -Array $internalRecipientIdsInEvent -Value $mapped.Id
    }
    elseif (Test-IsInternalEmail -Email $actorUserId -InternalDomains $InternalDomainsNormalized) {
        $internalUsersInEvent = Add-UniqueToArray -Array $internalUsersInEvent -Value $actorUserId
    }

    # 2) UserKey object ID mapping
    if ($userKey -ne "" -and $UserById.ContainsKey($userKey)) {
        $mapped = $UserById[$userKey]
        $internalUsersInEvent = Add-UniqueToArray -Array $internalUsersInEvent -Value $mapped.UserPrincipalName
        $internalRecipientIdsInEvent = Add-UniqueToArray -Array $internalRecipientIdsInEvent -Value $mapped.Id
    }

    # 3) GUIDs from ChatThreadId / ItemName / RawAuditData
    foreach ($guid in $allGuidsInRaw) {
        if ($UserById.ContainsKey($guid)) {
            $mapped = $UserById[$guid]
            $internalUsersInEvent = Add-UniqueToArray -Array $internalUsersInEvent -Value $mapped.UserPrincipalName
            $internalRecipientIdsInEvent = Add-UniqueToArray -Array $internalRecipientIdsInEvent -Value $mapped.Id
        }
    }

    # 4) Emails from raw AuditData
    foreach ($email in $allEmailsInRaw) {
        if ($UserByUpn.ContainsKey($email)) {
            $mapped = $UserByUpn[$email]
            $internalUsersInEvent = Add-UniqueToArray -Array $internalUsersInEvent -Value $mapped.UserPrincipalName
            $internalRecipientIdsInEvent = Add-UniqueToArray -Array $internalRecipientIdsInEvent -Value $mapped.Id
        }
        elseif ($UserByMail.ContainsKey($email)) {
            $mapped = $UserByMail[$email]
            $internalUsersInEvent = Add-UniqueToArray -Array $internalUsersInEvent -Value $mapped.UserPrincipalName
            $internalRecipientIdsInEvent = Add-UniqueToArray -Array $internalRecipientIdsInEvent -Value $mapped.Id
        }
        elseif (Test-IsInternalEmail -Email $email -InternalDomains $InternalDomainsNormalized) {
            $internalUsersInEvent = Add-UniqueToArray -Array $internalUsersInEvent -Value $email
        }
    }

    $isActorInternal = Test-IsInternalEmail -Email $actorUserId -InternalDomains $InternalDomainsNormalized

    $direction = "Unknown"

    if ($isActorInternal) {
        $direction = "OutboundFromInternalActor"
    }
    elseif (-not $isActorInternal -and $actorUserId -like "*@*") {
        $direction = "InboundFromExternalActor"
    }

    $eventHasExternalDomain = $externalDomainsInEvent.Count -gt 0

    $ParsedEvents += [pscustomobject]@{
        CreationTime                = $creationTime
        Operation                   = $operation
        Direction                   = $direction
        ActorUserId                 = $actorUserId
        ActorDomain                 = $actorDomain
        IsActorInternal             = $isActorInternal
        InternalUsersInEvent        = Join-UniqueValues $internalUsersInEvent
        InternalRecipientIdsInEvent = Join-UniqueValues $internalRecipientIdsInEvent
        CommunicationType           = $communicationType
        ClientIP                    = $clientIp
        UserAgent                   = $userAgent
        ChatThreadId                = $chatThreadId
        MessageId                   = $messageId
        MessageVersion              = $messageVersion
        UserType                    = $userType
        UserKey                     = $userKey
        UserTenantId                = $userTenantId
        ResourceTenantId            = $resourceTenantId
        OrganizationId              = $organizationId
        ItemName                    = $itemName
        HasForeignTenantUsers       = $hasForeignTenantUsers
        HasGuestUsers               = $hasGuestUsers
        HasOtherGuestUsers          = $hasOtherGuestUsers
        HasUnauthenticatedUsers     = $hasUnauthenticatedUsers
        EventHasExternalDomain      = $eventHasExternalDomain
        ParticipatingDomains        = Join-UniqueValues $participatingDomains
        ExternalDomainsInEvent      = Join-UniqueValues $externalDomainsInEvent
        ParticipatingTenantIds      = Join-UniqueValues $participatingTenantIds
        ParticipatingSIPDomains     = $participatingSipDomainsJson
        Extra_TimeZone              = Get-ExtraPropertyValue -AuditJson $auditJson -Key "TimeZone"
        Extra_OsName                = Get-ExtraPropertyValue -AuditJson $auditJson -Key "OsName"
        Extra_OsVersion             = Get-ExtraPropertyValue -AuditJson $auditJson -Key "OsVersion"
        Extra_Country               = Get-ExtraPropertyValue -AuditJson $auditJson -Key "Country"
        Extra_ClientName            = Get-ExtraPropertyValue -AuditJson $auditJson -Key "ClientName"
        Extra_ClientVersion         = Get-ExtraPropertyValue -AuditJson $auditJson -Key "ClientVersion"
        Extra_ClientUtcOffsetSecs   = Get-ExtraPropertyValue -AuditJson $auditJson -Key "ClientUtcOffsetSeconds"
        AllEmailsInRaw              = Join-UniqueValues $allEmailsInRaw
        AllGuidsInRaw               = Join-UniqueValues $allGuidsInRaw
        RawAuditData                = $(if ($IncludeRawAuditData) { $rawAuditData } else { "" })
        RecordIdentity              = [string]$record.Identity
    }
}

Write-Host "Parsed Teams audit events: $($ParsedEvents.Count)"
Write-Host ""

# ----------------------------
# Thread-level enrichment
# ----------------------------

$ThreadMap = @{}

foreach ($event in $ParsedEvents) {
    $threadId = [string]$event.ChatThreadId

    if ([string]::IsNullOrWhiteSpace($threadId)) {
        continue
    }

    if (-not $ThreadMap.ContainsKey($threadId)) {
        $ThreadMap[$threadId] = [pscustomobject]@{
            ChatThreadId           = $threadId
            InternalUsers          = @()
            InternalRecipientIds   = @()
            ExternalDomains        = @()
            Actors                 = @()
            ClientIPs              = @()
            Operations             = @()
            Directions             = @()
            ParticipatingDomains   = @()
            ParticipatingTenantIds = @()
            FirstSeen              = $event.CreationTime
            LastSeen               = $event.CreationTime
            EventCount             = 0
        }
    }

    $thread = $ThreadMap[$threadId]
    $thread.EventCount++

    foreach ($u in @(([string]$event.InternalUsersInEvent) -split ";")) {
        $cleanUser = Normalize-String $u
        if ($cleanUser -ne "") {
            $thread.InternalUsers = Add-UniqueToArray -Array $thread.InternalUsers -Value $cleanUser
        }
    }

    foreach ($id in @(([string]$event.InternalRecipientIdsInEvent) -split ";")) {
        $cleanId = Normalize-String $id
        if ($cleanId -ne "") {
            $thread.InternalRecipientIds = Add-UniqueToArray -Array $thread.InternalRecipientIds -Value $cleanId
        }
    }

    foreach ($d in @(([string]$event.ExternalDomainsInEvent) -split ";")) {
        $cleanDomain = Normalize-String $d
        if ($cleanDomain -ne "") {
            $thread.ExternalDomains = Add-UniqueToArray -Array $thread.ExternalDomains -Value $cleanDomain
        }
    }

    foreach ($d in @(([string]$event.ParticipatingDomains) -split ";")) {
        $cleanDomain = Normalize-String $d
        if ($cleanDomain -ne "") {
            $thread.ParticipatingDomains = Add-UniqueToArray -Array $thread.ParticipatingDomains -Value $cleanDomain
        }
    }

    foreach ($t in @(([string]$event.ParticipatingTenantIds) -split ";")) {
        $cleanTenant = ([string]$t).Trim()
        if ($cleanTenant -ne "") {
            $thread.ParticipatingTenantIds = Add-UniqueToArray -Array $thread.ParticipatingTenantIds -Value $cleanTenant
        }
    }

    $thread.Actors = Add-UniqueToArray -Array $thread.Actors -Value $event.ActorUserId
    $thread.ClientIPs = Add-UniqueToArray -Array $thread.ClientIPs -Value $event.ClientIP
    $thread.Operations = Add-UniqueToArray -Array $thread.Operations -Value $event.Operation
    $thread.Directions = Add-UniqueToArray -Array $thread.Directions -Value $event.Direction

    try {
        if ([datetime]$event.CreationTime -lt [datetime]$thread.FirstSeen) {
            $thread.FirstSeen = $event.CreationTime
        }

        if ([datetime]$event.CreationTime -gt [datetime]$thread.LastSeen) {
            $thread.LastSeen = $event.CreationTime
        }
    }
    catch {
        # Ignore date conversion issues.
    }
}

$ThreadSummary = @()

foreach ($key in $ThreadMap.Keys) {
    $thread = $ThreadMap[$key]

    $ThreadSummary += [pscustomobject]@{
        ChatThreadId           = $thread.ChatThreadId
        FirstSeen              = $thread.FirstSeen
        LastSeen               = $thread.LastSeen
        EventCount             = $thread.EventCount
        InternalUsers          = Join-UniqueValues $thread.InternalUsers
        InternalRecipientIds   = Join-UniqueValues $thread.InternalRecipientIds
        ExternalDomains        = Join-UniqueValues $thread.ExternalDomains
        Actors                 = Join-UniqueValues $thread.Actors
        ClientIPs              = Join-UniqueValues $thread.ClientIPs
        Operations             = Join-UniqueValues $thread.Operations
        Directions             = Join-UniqueValues $thread.Directions
        ParticipatingDomains   = Join-UniqueValues $thread.ParticipatingDomains
        ParticipatingTenantIds = Join-UniqueValues $thread.ParticipatingTenantIds
    }
}

$ThreadSummary = @($ThreadSummary | Sort-Object LastSeen -Descending)

# ----------------------------
# User-domain rows
# ----------------------------

$UserDomainRows = @()
$UnattributedExternalThreads = @()

foreach ($thread in $ThreadSummary) {
    $internalUsers = @(([string]$thread.InternalUsers) -split ";") |
        ForEach-Object { Normalize-String $_ } |
        Where-Object { $_ -ne "" } |
        Sort-Object -Unique

    $externalDomains = @(([string]$thread.ExternalDomains) -split ";") |
        ForEach-Object { Normalize-String $_ } |
        Where-Object { $_ -ne "" } |
        Sort-Object -Unique

    if ($externalDomains.Count -eq 0) {
        continue
    }

    if ($internalUsers.Count -eq 0) {
        $UnattributedExternalThreads += [pscustomobject]@{
            ChatThreadId            = $thread.ChatThreadId
            FirstSeen               = $thread.FirstSeen
            LastSeen                = $thread.LastSeen
            EventCount              = $thread.EventCount
            ExternalDomains         = Join-UniqueValues $externalDomains
            Actors                  = $thread.Actors
            ClientIPs               = $thread.ClientIPs
            ParticipatingDomains    = $thread.ParticipatingDomains
            ParticipatingTenantIds  = $thread.ParticipatingTenantIds
            Note                    = "External domain detected but no internal user could be mapped from actor, UserKey, GUIDs, or raw emails."
        }

        continue
    }

    foreach ($internalUser in $internalUsers) {
        foreach ($externalDomain in $externalDomains) {
            $UserDomainRows += [pscustomobject]@{
                UserPrincipalName      = $internalUser
                ExternalDomain         = $externalDomain
                ChatThreadId           = $thread.ChatThreadId
                FirstSeen              = $thread.FirstSeen
                LastSeen               = $thread.LastSeen
                EventCount             = $thread.EventCount
                Actors                 = $thread.Actors
                ClientIPs              = $thread.ClientIPs
                Directions             = $thread.Directions
                ParticipatingDomains   = $thread.ParticipatingDomains
                ParticipatingTenantIds = $thread.ParticipatingTenantIds
            }
        }
    }
}

$UserExternalDomainSummary = @(
    $UserDomainRows |
        Group-Object UserPrincipalName, ExternalDomain |
        ForEach-Object {
            $group = @($_.Group)

            [pscustomobject]@{
                UserPrincipalName      = $group[0].UserPrincipalName
                ExternalDomain         = $group[0].ExternalDomain
                FirstSeen              = ($group | Sort-Object FirstSeen | Select-Object -First 1).FirstSeen
                LastSeen               = ($group | Sort-Object LastSeen | Select-Object -Last 1).LastSeen
                ThreadCount            = @($group.ChatThreadId | Sort-Object -Unique).Count
                EventCount             = ($group | Measure-Object -Property EventCount -Sum).Sum
                ChatThreadIds          = Join-UniqueValues $group.ChatThreadId
                Actors                 = Join-UniqueValues $group.Actors
                ClientIPs              = Join-UniqueValues $group.ClientIPs
                Directions             = Join-UniqueValues $group.Directions
                ParticipatingDomains   = Join-UniqueValues $group.ParticipatingDomains
                ParticipatingTenantIds = Join-UniqueValues $group.ParticipatingTenantIds
            }
        } |
        Sort-Object UserPrincipalName, ExternalDomain
)

$UserSummary = @(
    $UserDomainRows |
        Group-Object UserPrincipalName |
        ForEach-Object {
            $group = @($_.Group)

            [pscustomobject]@{
                UserPrincipalName    = $_.Name
                FirstSeen            = ($group | Sort-Object FirstSeen | Select-Object -First 1).FirstSeen
                LastSeen             = ($group | Sort-Object LastSeen | Select-Object -Last 1).LastSeen
                ExternalDomainCount  = @($group.ExternalDomain | Sort-Object -Unique).Count
                ExternalDomains      = Join-UniqueValues $group.ExternalDomain
                ThreadCount          = @($group.ChatThreadId | Sort-Object -Unique).Count
                EventCount           = ($group | Measure-Object -Property EventCount -Sum).Sum
                ChatThreadIds        = Join-UniqueValues $group.ChatThreadId
                ClientIPs            = Join-UniqueValues $group.ClientIPs
            }
        } |
        Sort-Object LastSeen -Descending
)

$ExternalDomainSummary = @(
    $UserDomainRows |
        Group-Object ExternalDomain |
        ForEach-Object {
            $group = @($_.Group)

            [pscustomobject]@{
                ExternalDomain = $_.Name
                FirstSeen      = ($group | Sort-Object FirstSeen | Select-Object -First 1).FirstSeen
                LastSeen       = ($group | Sort-Object LastSeen | Select-Object -Last 1).LastSeen
                UserCount      = @($group.UserPrincipalName | Sort-Object -Unique).Count
                Users          = Join-UniqueValues $group.UserPrincipalName
                ThreadCount    = @($group.ChatThreadId | Sort-Object -Unique).Count
                EventCount     = ($group | Measure-Object -Property EventCount -Sum).Sum
                ChatThreadIds  = Join-UniqueValues $group.ChatThreadId
                ClientIPs      = Join-UniqueValues $group.ClientIPs
            }
        } |
        Sort-Object LastSeen -Descending
)

$ActorSummary = @(
    $ParsedEvents |
        Group-Object ActorUserId |
        ForEach-Object {
            $group = @($_.Group)

            [pscustomobject]@{
                ActorUserId              = $_.Name
                ActorDomain              = Join-UniqueValues $group.ActorDomain
                IsActorInternal          = Join-UniqueValues $group.IsActorInternal
                FirstSeen                = ($group | Sort-Object CreationTime | Select-Object -First 1).CreationTime
                LastSeen                 = ($group | Sort-Object CreationTime | Select-Object -Last 1).CreationTime
                EventCount               = $group.Count
                Directions               = Join-UniqueValues $group.Direction
                InternalUsersInEvent     = Join-UniqueValues $group.InternalUsersInEvent
                ExternalDomainsInEvent   = Join-UniqueValues $group.ExternalDomainsInEvent
                ChatThreadIds            = Join-UniqueValues $group.ChatThreadId
                ClientIPs                = Join-UniqueValues $group.ClientIP
            }
        } |
        Sort-Object LastSeen -Descending
)

# ----------------------------
# Export reports
# ----------------------------

if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}

$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"

$recipientsPath = Join-Path $OutputFolder "TeamsTenant_EXORecipients_$timestamp.csv"
$messageDetailsPath = Join-Path $OutputFolder "TeamsTenant_MessageDetails_$timestamp.csv"
$threadSummaryPath = Join-Path $OutputFolder "TeamsTenant_ThreadSummary_$timestamp.csv"
$userExternalDomainSummaryPath = Join-Path $OutputFolder "TeamsTenant_UserExternalDomainSummary_$timestamp.csv"
$userSummaryPath = Join-Path $OutputFolder "TeamsTenant_UserSummary_$timestamp.csv"
$externalDomainSummaryPath = Join-Path $OutputFolder "TeamsTenant_ExternalDomainSummary_$timestamp.csv"
$actorSummaryPath = Join-Path $OutputFolder "TeamsTenant_ActorSummary_$timestamp.csv"
$unattributedPath = Join-Path $OutputFolder "TeamsTenant_UnattributedExternalThreads_$timestamp.csv"
$readmePath = Join-Path $OutputFolder "README.txt"

Export-ObjectsToCsvSafe -Objects $InternalRecipients -Path $recipientsPath
Export-ObjectsToCsvSafe -Objects $ParsedEvents -Path $messageDetailsPath
Export-ObjectsToCsvSafe -Objects $ThreadSummary -Path $threadSummaryPath
Export-ObjectsToCsvSafe -Objects $UserExternalDomainSummary -Path $userExternalDomainSummaryPath
Export-ObjectsToCsvSafe -Objects $UserSummary -Path $userSummaryPath
Export-ObjectsToCsvSafe -Objects $ExternalDomainSummary -Path $externalDomainSummaryPath
Export-ObjectsToCsvSafe -Objects $ActorSummary -Path $actorSummaryPath
Export-ObjectsToCsvSafe -Objects $UnattributedExternalThreads -Path $unattributedPath

$readmeContent = @"
Microsoft Teams External/Federated Domain Audit Report
======================================================

Generated At
------------
$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")

Search Parameters
-----------------
Start UTC:        $StartUtc
End UTC:          $EndUtc
Internal domains: $($InternalDomainsNormalized -join ", ")
Recipient types:  $($RecipientTypeDetails -join ", ")
Operations:       $($Operations -join ", ")
Raw AuditData included: $IncludeRawAuditData
Related events included: $IncludeRelatedEvents

Purpose
-------
This report analyzes Microsoft Teams audit events from Microsoft 365 Unified Audit Log.

The script identifies external/federated Teams domains by parsing:

    ParticipantInfo.ParticipatingDomains

Then it removes the internal domains provided with the -InternalDomains parameter.
The remaining domains are treated as external Teams domains.

Main question answered by this report:

    Which internal users communicated with which external Teams domains?

Generated CSV Files
===================

1) TeamsTenant_UserExternalDomainSummary_*.csv
----------------------------------------------
Most important report.

Answers:

    Which internal user communicated with which external Teams domain?

Important columns:

    UserPrincipalName
    ExternalDomain
    FirstSeen
    LastSeen
    ThreadCount
    EventCount
    ChatThreadIds
    Actors
    ClientIPs
    Directions
    ParticipatingDomains
    ParticipatingTenantIds

Recommended for:

    Management/customer-facing summary
    User-to-external-domain investigation
    Quick external collaboration review


2) TeamsTenant_UserSummary_*.csv
--------------------------------
User-based summary.

Answers:

    How many external domains did each internal user communicate with?

Important columns:

    UserPrincipalName
    ExternalDomainCount
    ExternalDomains
    ThreadCount
    EventCount
    FirstSeen
    LastSeen

Recommended for:

    Quick user risk review
    Finding users with broad external Teams communication


3) TeamsTenant_ExternalDomainSummary_*.csv
------------------------------------------
External domain-based summary.

Answers:

    Which external domains were contacted and how many internal users were involved?

Important columns:

    ExternalDomain
    UserCount
    Users
    ThreadCount
    EventCount
    FirstSeen
    LastSeen

Recommended for:

    Suspicious external domain analysis
    External collaboration review
    Domain-level blocking/allowlisting decisions


4) TeamsTenant_ThreadSummary_*.csv
----------------------------------
Teams chat thread-based summary.

Answers:

    Which internal users and external domains existed in each Teams chat thread?

Important columns:

    ChatThreadId
    InternalUsers
    ExternalDomains
    Actors
    ClientIPs
    ParticipatingDomains
    ParticipatingTenantIds
    FirstSeen
    LastSeen
    EventCount

Recommended for:

    Evidence validation
    ChatThreadId-based correlation
    Verifying both directions of a conversation


5) TeamsTenant_MessageDetails_*.csv
-----------------------------------
Most detailed audit event report.

Contains one row per parsed audit event.

Important columns:

    CreationTime
    Operation
    ActorUserId
    ClientIP
    ChatThreadId
    MessageId
    ExternalDomainsInEvent
    InternalUsersInEvent
    ParticipatingDomains
    ParticipatingTenantIds
    RawAuditData

If -IncludeRawAuditData was used, this file contains raw AuditData JSON.

Recommended for:

    Deep technical analysis
    Evidence preservation
    Troubleshooting mapping/correlation problems

Warning:

    This file may contain sensitive metadata such as IP addresses, tenant IDs,
    session IDs, message IDs, and raw audit JSON. Review before sharing externally.


6) TeamsTenant_ActorSummary_*.csv
---------------------------------
Actor-based summary.

Answers:

    Which users appeared as actors in the audit events?

Important columns:

    ActorUserId
    ActorDomain
    IsActorInternal
    FirstSeen
    LastSeen
    EventCount
    Directions
    InternalUsersInEvent
    ExternalDomainsInEvent
    ChatThreadIds
    ClientIPs

Recommended for:

    Actor-level investigation
    Checking whether external users appeared directly as actors


7) TeamsTenant_EXORecipients_*.csv
----------------------------------
Internal recipient list pulled from Exchange Online.

This file contains the internal mail-enabled recipients used for mapping audit events.

Important columns:

    Id
    DisplayName
    UserPrincipalName
    PrimarySmtpAddress
    ProxySmtpAddresses
    RecipientTypeDetails
    Domains

Recommended for:

    Debugging
    Verifying why a user was or was not mapped
    Checking object ID / SMTP address mapping


8) TeamsTenant_UnattributedExternalThreads_*.csv
------------------------------------------------
External domains were detected, but the script could not reliably map the thread to an internal user.

Ideally this file should be empty.

If it contains rows, manual review is recommended.

Possible reasons:

    Actor UserId did not map to an internal recipient
    UserKey did not match ExternalDirectoryObjectId
    ChatThreadId or raw AuditData did not contain a mappable internal GUID
    Raw AuditData did not contain an internal email address
    The internal user was not returned by Get-EXORecipient
    The account is not mail-enabled or uses an unexpected recipient type

Important Notes
===============

1) Audit Log Delay
------------------
Microsoft Teams audit logs may appear with delay.
Recent messages may require 60-90 minutes before they are searchable.

2) Required Backend
-------------------
This script uses:

    Search-UnifiedAuditLog
    Get-EXORecipient

It does not use:

    Microsoft Graph
    Microsoft Sentinel
    Microsoft Defender
    Intune
    Defender for Cloud Apps

3) Required Permissions
-----------------------
The account running the script needs permissions for:

    Search-UnifiedAuditLog
    Get-EXORecipient

Typical practical role combination:

    Exchange Admin
    +
    View-Only Audit Logs

4) Unified Audit Log
--------------------
Unified Audit Log must be enabled.

Check:

    Get-AdminAuditLogConfig | Format-List UnifiedAuditLogIngestionEnabled

Expected:

    UnifiedAuditLogIngestionEnabled : True

5) ClientIP Interpretation
--------------------------
ClientIP is audit metadata for the actor in that audit record.

Do not directly report it as the user's physical location without validation.
VPN, proxy, mobile ISP, Microsoft service behavior, and federation behavior may affect interpretation.

6) Internal Domain Accuracy
---------------------------
The -InternalDomains parameter is critical.

If internal domains are missing, they may be incorrectly reported as external domains.

Example:

    -InternalDomains @("company.com", "company.com.tr", "subsidiary.com")

7) Most Useful Files
--------------------
For normal reporting, start with:

    TeamsTenant_UserExternalDomainSummary_*.csv
    TeamsTenant_UserSummary_*.csv
    TeamsTenant_ExternalDomainSummary_*.csv

For technical validation, use:

    TeamsTenant_ThreadSummary_*.csv
    TeamsTenant_MessageDetails_*.csv
    TeamsTenant_UnattributedExternalThreads_*.csv

"@

Set-Content -Path $readmePath -Value $readmeContent -Encoding UTF8

Write-Host ""
Write-Host "Done."
Write-Host "EXO recipients CSV:                $recipientsPath"
Write-Host "Message details CSV:               $messageDetailsPath"
Write-Host "Thread summary CSV:                $threadSummaryPath"
Write-Host "User-external domain summary CSV:  $userExternalDomainSummaryPath"
Write-Host "User summary CSV:                  $userSummaryPath"
Write-Host "External domain summary CSV:       $externalDomainSummaryPath"
Write-Host "Actor summary CSV:                 $actorSummaryPath"
Write-Host "Unattributed external threads CSV: $unattributedPath"
Write-Host "README file:                       $readmePath"
Write-Host ""

Write-Host "External domain summary preview:"
if ($ExternalDomainSummary.Count -gt 0) {
    $ExternalDomainSummary | Format-Table ExternalDomain, UserCount, FirstSeen, LastSeen, ThreadCount, EventCount -AutoSize
}
else {
    Write-Host "No external domains found in the selected time range."
}

Write-Host ""
Write-Host "User summary preview:"
if ($UserSummary.Count -gt 0) {
    $UserSummary | Format-Table UserPrincipalName, ExternalDomainCount, ExternalDomains, FirstSeen, LastSeen, ThreadCount -AutoSize
}
else {
    Write-Host "No user-to-external-domain mapping found."
}

Write-Host ""
Write-Host "Notes:"
Write-Host "- It uses Get-EXORecipient to map mail-enabled internal recipients."
Write-Host "- It does not request WindowsEmailAddress or UserPrincipalName from Get-EXORecipient."
Write-Host "- The most important report is TeamsTenant_UserExternalDomainSummary."
Write-Host "- If a thread has external domains but no internal user can be mapped, check TeamsTenant_UnattributedExternalThreads."
Write-Host "- Teams audit logs can appear with delay. Recent messages may require 60-90 minutes."
Write-Host "- ClientIP is audit metadata for the actor in that record. Be careful before interpreting it as physical location."