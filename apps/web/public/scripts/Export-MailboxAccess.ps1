#Requires -Version 5.1
<#
.SYNOPSIS
Exports who holds Full Access, Send As and Send on Behalf on shared mailboxes, for Prism365.

.DESCRIPTION
Microsoft offers no read-only API for these rights, so Prism365 reads them from this file
instead. The script runs in your own Exchange Online session and only reads: every Exchange
cmdlet it calls is a Get- cmdlet. It writes one UTF-8 JSON file, which you import on Prism365's
Microsoft access page. Prism365 never sees your Exchange session or any credential.

Before running it:
  1. Install ExchangeOnlineManagement 3.0.0 or later.
  2. Connect once, to the tenant Prism365 reads: Connect-ExchangeOnline
  3. Check that your role may run the permission cmdlets:
     Get-ManagementRole -Cmdlet Get-MailboxPermission

Full Access is read with Get-EXOMailboxPermission, without inherited or deny entries and without
the mailbox's own NT AUTHORITY\SELF entry. Send As is read with Get-EXORecipientPermission,
without inherited or deny entries. Send on Behalf is the mailbox's GrantSendOnBehalfTo list.
Every trustee is looked up once with Get-EXORecipient. A trustee that does not resolve to
exactly one recipient (a deleted account, an ambiguous name) is written as unresolved, with the
text Exchange returned, and is never guessed.

.PARAMETER OutputPath
Where to write the file. The script refuses to replace an existing file unless -Force is given.

.PARAMETER IncludeUserMailboxes
Also export rights on user mailboxes, not only shared mailboxes. This takes longer and makes the
file larger; Prism365 reads files up to 25 MB.

.PARAMETER Force
Replace the file at OutputPath if it exists.

.EXAMPLE
.\Export-MailboxAccess.ps1 -OutputPath .\mailbox-access.json
#>
[CmdletBinding()]
param(
  [string]$OutputPath = '.\prism365-mailbox-access.json',
  [switch]$IncludeUserMailboxes,
  [switch]$Force
)

$ErrorActionPreference = 'Stop'
$roleCheck = 'Get-ManagementRole -Cmdlet Get-MailboxPermission'

# Text cut to the length Prism365 accepts, so one long value never makes the whole file refused.
function Get-BoundedText([object]$Value, [int]$Max) {
  if ($null -eq $Value) { return '' }
  $text = ([string]$Value).Trim()
  if ($text.Length -gt $Max) { $text = $text.Substring(0, $Max).Trim() }
  return $text
}

# The value of a field if the object has it, or $null. Microsoft does not document every output
# field of these cmdlets, so a missing one is read as missing rather than failing the export.
function Get-FieldValue([object]$Item, [string]$Name) {
  if ($null -eq $Item) { return $null }
  $property = $Item.PSObject.Properties[$Name]
  if ($null -eq $property) { return $null }
  return $property.Value
}

# True only for an actual true value. A string such as 'False' is not true here, whatever
# PowerShell's own conversion would make of it.
function Get-IsTrue([object]$Value) {
  return ([string]$Value) -eq 'True'
}

# A directory object id as Prism365 stores it, or $null for anything that is not a GUID.
function Get-DirectoryId([object]$Value) {
  $parsed = [guid]::Empty
  if ($null -ne $Value -and [guid]::TryParse([string]$Value, [ref]$parsed) -and $parsed -ne [guid]::Empty) {
    return $parsed.ToString()
  }
  return $null
}

# An address that looks like one, cut to nothing when it does not.
function Get-Address([object]$Value) {
  $text = Get-BoundedText -Value $Value -Max 320
  if ($text -match '^[^\s@]+@[^\s@]+$') { return $text }
  return $null
}

# Each trustee is looked up once. Exchange returns every recipient in the tenant when
# Get-EXORecipient is given an empty identity, so an empty value is never looked up: it is not a
# trustee at all.
$trusteeCache = @{}
function Get-ResolvedTrustee([object]$Raw) {
  $value = Get-BoundedText -Value $Raw -Max 500
  if ([string]::IsNullOrWhiteSpace($value)) { return $null }
  if ($trusteeCache.ContainsKey($value)) { return $trusteeCache[$value] }
  $found = @()
  try {
    $found = @(Get-EXORecipient -Identity $value -Properties PrimarySmtpAddress, DisplayName, ExternalDirectoryObjectId, RecipientTypeDetails -ErrorAction Stop)
  }
  catch {
    # Only "not found" and "matches more than one recipient" make a trustee unresolved. Any other
    # error (a refused cmdlet, a throttled or broken session) would quietly turn real access into
    # unresolved trustees, so the export stops instead, before anything is written.
    $reason = "$($_.CategoryInfo.Category) $($_.FullyQualifiedErrorId) $($_.Exception.Message)"
    $unknownTrustee = $reason -match "ObjectNotFound|ManagementObjectNotFound|Ambiguous|couldn't be found|could not be found|matches multiple"
    if ($reason -match 'CommandNotFound' -or -not $unknownTrustee) {
      Write-Host "Get-EXORecipient failed for the trustee $value : $($_.Exception.Message)" -ForegroundColor Red
      Write-Host "Nothing was written. If the cmdlet is not allowed for your role, check which roles may run it with: Get-ManagementRole -Cmdlet Get-Recipient" -ForegroundColor Red
      exit 1
    }
    $found = @()
  }
  $trustee = [ordered]@{ id = $null; address = $null; name = $value; kind = 'unresolved' }
  if ($found.Count -eq 1) {
    $recipient = $found[0]
    $type = [string](Get-FieldValue -Item $recipient -Name 'RecipientTypeDetails')
    $id = Get-DirectoryId -Value (Get-FieldValue -Item $recipient -Name 'ExternalDirectoryObjectId')
    $kind = 'other'
    if ($type -match 'Group') { $kind = 'group' }
    elseif ($type -match '(Mailbox|MailUser|User)$') { $kind = 'user' }
    # People and groups are matched by id in Prism365; one without an id is shown, not matched.
    if ($null -eq $id -and $kind -ne 'other') { $kind = 'other' }
    $name = Get-BoundedText -Value (Get-FieldValue -Item $recipient -Name 'DisplayName') -Max 500
    if ([string]::IsNullOrWhiteSpace($name)) { $name = $value }
    $trustee = [ordered]@{
      id      = $id
      address = Get-Address -Value (Get-FieldValue -Item $recipient -Name 'PrimarySmtpAddress')
      name    = $name
      kind    = $kind
    }
  }
  $trusteeCache[$value] = $trustee
  return $trustee
}

# The module and the one session this export reads with.
$module = @(Get-Module -Name ExchangeOnlineManagement)
if ($module.Count -eq 0) {
  Write-Host 'ExchangeOnlineManagement is not loaded. Run Connect-ExchangeOnline first, then run this script in the same window.' -ForegroundColor Red
  exit 1
}
$newest = $module[0]
foreach ($loaded in $module) { if ($loaded.Version -gt $newest.Version) { $newest = $loaded } }
if ($newest.Version.Major -lt 3) {
  Write-Host "This script needs ExchangeOnlineManagement 3.0.0 or later; version $($newest.Version) is loaded." -ForegroundColor Red
  exit 1
}
$connected = [System.Collections.Generic.List[object]]::new()
foreach ($connection in @(Get-ConnectionInformation)) {
  if (([string](Get-FieldValue -Item $connection -Name 'State')) -eq 'Connected') { $connected.Add($connection) }
}
if ($connected.Count -ne 1) {
  Write-Host "Found $($connected.Count) Exchange Online connections. Connect exactly once, to the tenant Prism365 reads (Disconnect-ExchangeOnline, then Connect-ExchangeOnline), and run this again." -ForegroundColor Red
  exit 1
}
$tenantId = Get-DirectoryId -Value (Get-FieldValue -Item $connected[0] -Name 'TenantID')
if ($null -eq $tenantId) {
  Write-Host 'The Exchange Online connection does not say which tenant it is for. Update ExchangeOnlineManagement and connect again.' -ForegroundColor Red
  exit 1
}

$outputFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
if ([System.IO.File]::Exists($outputFile) -and -not $Force) {
  Write-Host "$outputFile already exists. Choose another -OutputPath, or add -Force to replace it." -ForegroundColor Red
  exit 1
}

# Taken before anything is read, so the file never claims to be newer than what it holds.
$exportedAt = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", [System.Globalization.CultureInfo]::InvariantCulture)

$types = @('SharedMailbox')
if ($IncludeUserMailboxes) { $types += 'UserMailbox' }
Write-Host "Reading $($types -join ' and ') mailboxes in tenant $tenantId."
try {
  $found = @(Get-EXOMailbox -RecipientTypeDetails $types -ResultSize Unlimited -Properties GrantSendOnBehalfTo -ErrorAction Stop)
}
catch {
  Write-Host "Exchange refused Get-EXOMailbox: $($_.Exception.Message)" -ForegroundColor Red
  Write-Host "Check which roles may read mailboxes and their permissions with: $roleCheck" -ForegroundColor Red
  exit 1
}

# A mailbox's own address always resolves. If it does not, Get-EXORecipient is not working for
# this session, and every trustee would come out unresolved: stop instead of writing that.
foreach ($mailbox in $found) {
  $probe = Get-Address -Value (Get-FieldValue -Item $mailbox -Name 'PrimarySmtpAddress')
  if ($null -eq $probe) { continue }
  if ((Get-ResolvedTrustee -Raw $probe).kind -eq 'unresolved') {
    Write-Host "Get-EXORecipient could not look up the mailbox $probe, so no trustee could be resolved. Nothing was written." -ForegroundColor Red
    Write-Host "Check which roles may run it with: Get-ManagementRole -Cmdlet Get-Recipient (and $roleCheck)" -ForegroundColor Red
    exit 1
  }
  break
}

$mailboxes = [System.Collections.Generic.List[object]]::new()
$grants = [System.Collections.Generic.List[object]]::new()
$seen = @{}
$skipped = 0
$unresolved = 0
$index = 0
foreach ($mailbox in $found) {
  $index += 1
  $id = Get-DirectoryId -Value (Get-FieldValue -Item $mailbox -Name 'ExternalDirectoryObjectId')
  $address = Get-Address -Value (Get-FieldValue -Item $mailbox -Name 'PrimarySmtpAddress')
  $type = [string](Get-FieldValue -Item $mailbox -Name 'RecipientTypeDetails')
  # A mailbox Prism365 could not name by its directory id is left out rather than half written.
  if ($null -eq $id -or $null -eq $address -or $types -notcontains $type) {
    $skipped += 1
    continue
  }
  Write-Progress -Activity 'Reading mailbox permissions' -Status "$index of $($found.Count): $address" -PercentComplete ([int](100 * $index / [Math]::Max(1, $found.Count)))
  $mailboxes.Add([ordered]@{
      id          = $id
      address     = $address
      displayName = Get-BoundedText -Value (Get-FieldValue -Item $mailbox -Name 'DisplayName') -Max 256
      type        = $type
    })

  # The rights on this mailbox as (right, the trustee as Exchange wrote it).
  $pending = [System.Collections.Generic.List[object]]::new()
  try {
    $fullAccess = @(Get-EXOMailboxPermission -Identity $id -ResultSize Unlimited -ErrorAction Stop)
    $sendAs = @(Get-EXORecipientPermission -Identity $id -AccessRights SendAs -ResultSize Unlimited -ErrorAction Stop)
  }
  catch {
    Write-Host "Exchange refused to read the permissions of $address : $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Nothing was written. If the cmdlet is not allowed for your role, check which roles may run it with: $roleCheck" -ForegroundColor Red
    exit 1
  }
  foreach ($entry in $fullAccess) {
    if ((Get-IsTrue (Get-FieldValue -Item $entry -Name 'IsInherited')) -or (Get-IsTrue (Get-FieldValue -Item $entry -Name 'Deny'))) { continue }
    $rights = @(Get-FieldValue -Item $entry -Name 'AccessRights') -join ','
    if ($rights -notmatch '\bFullAccess\b') { continue }
    $raw = Get-BoundedText -Value (Get-FieldValue -Item $entry -Name 'User') -Max 500
    if ($raw -eq 'NT AUTHORITY\SELF') { continue }
    $pending.Add(@('FullAccess', $raw))
  }
  foreach ($entry in $sendAs) {
    if ((Get-IsTrue (Get-FieldValue -Item $entry -Name 'IsInherited')) -or (Get-IsTrue (Get-FieldValue -Item $entry -Name 'Deny'))) { continue }
    $control = [string](Get-FieldValue -Item $entry -Name 'AccessControlType')
    if ($control -ne '' -and $control -ne 'Allow') { continue }
    $rights = @(Get-FieldValue -Item $entry -Name 'AccessRights') -join ','
    if ($rights -ne '' -and $rights -notmatch '\bSendAs\b') { continue }
    $raw = Get-FieldValue -Item $entry -Name 'Trustee'
    if ($null -eq $raw) { $raw = Get-FieldValue -Item $entry -Name 'User' }
    $raw = Get-BoundedText -Value $raw -Max 500
    if ($raw -eq 'NT AUTHORITY\SELF') { continue }
    $pending.Add(@('SendAs', $raw))
  }
  foreach ($raw in @(Get-FieldValue -Item $mailbox -Name 'GrantSendOnBehalfTo')) {
    $pending.Add(@('SendOnBehalf', (Get-BoundedText -Value $raw -Max 500)))
  }

  foreach ($pair in $pending) {
    $trustee = Get-ResolvedTrustee -Raw $pair[1]
    if ($null -eq $trustee) { continue }
    $who = $trustee.id
    if ($null -eq $who) { $who = 'text:' + $trustee.name }
    $key = "$id|$($pair[0])|$who"
    if ($seen.ContainsKey($key)) { continue }
    $seen[$key] = $true
    if ($trustee.kind -eq 'unresolved') { $unresolved += 1 }
    $grants.Add([ordered]@{ mailboxId = $id; right = $pair[0]; trustee = $trustee })
  }
}
Write-Progress -Activity 'Reading mailbox permissions' -Completed

$document = [ordered]@{
  format               = 'prism365.mailbox-access'
  version              = 1
  tenantId             = $tenantId
  exportedAt           = $exportedAt
  includeUserMailboxes = [bool]$IncludeUserMailboxes
  mailboxes            = $mailboxes.ToArray()
  grants               = $grants.ToArray()
  unresolved           = $unresolved
}
$json = ConvertTo-Json -InputObject $document -Depth 6 -Compress
Set-Content -LiteralPath $outputFile -Value $json -Encoding UTF8
Write-Host "Wrote $outputFile with $($mailboxes.Count) mailboxes and $($grants.Count) grants ($unresolved with a trustee that could not be resolved)."
if ($skipped -gt 0) {
  Write-Host "$skipped mailboxes were left out because Exchange returned no directory id or address for them."
}
Write-Host 'Import the file on the Microsoft access page in Prism365.'
