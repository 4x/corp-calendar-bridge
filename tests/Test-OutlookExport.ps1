#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$root = Split-Path -Parent $PSScriptRoot
$directory = Join-Path $PSScriptRoot ('.test-exports-' + [guid]::NewGuid().ToString('N'))
$fixedNow = [datetime]'2026-10-02T12:00:00'

function Get-Date { return $fixedNow }
function Assert($Condition, [string]$Message) {
    if (-not $Condition) { throw "Assertion failed: $Message" }
}
function Appointment([string]$Id, [string]$Start, [string]$End, [int]$Sensitivity = 0, [bool]$AllDay = $false, [int]$MeetingStatus = 1) {
    return [pscustomobject]@{
        Class = 26; GlobalAppointmentID = $Id; Subject = 'Subject only'
        Start = [datetime]$Start; End = [datetime]$End
        StartUTC = [datetime]::SpecifyKind([datetime]$Start, [DateTimeKind]::Unspecified)
        EndUTC = [datetime]::SpecifyKind([datetime]$End, [DateTimeKind]::Unspecified)
        AllDayEvent = $AllDay; Sensitivity = $Sensitivity; MeetingStatus = $MeetingStatus
        Body = 'NEVER EXPORT'; Location = 'NEVER EXPORT'; Attendees = 'NEVER EXPORT'
    }
}
$items = [pscustomobject]@{ IncludeRecurrences = $false; Index = 0; Records = @(); ReadFailure = $false; SortCalled = $false; RestrictCalled = $false; FilterText = '' }
$items | Add-Member ScriptMethod Sort {
    param($Field, $Descending)
    Assert ($Field -eq '[Start]' -and -not $Descending) 'Ascending start sort is required'
    $this.SortCalled = $true
}
$items | Add-Member ScriptMethod Restrict {
    param($Filter)
    Assert ($this.SortCalled -and $this.IncludeRecurrences) 'Sort then expand recurrences before filtering'
    $this.RestrictCalled = $true
    $this.FilterText = $Filter
    return $this
}
$items | Add-Member ScriptMethod GetFirst {
    $this.Index = 0
    if ($this.Records.Length -eq 0) { return $null }
    return $this.Records[0]
}
$items | Add-Member ScriptMethod GetNext {
    if ($this.ReadFailure) { throw 'Simulated Outlook read error' }
    $this.Index++
    if ($this.Index -ge $this.Records.Length) { return $null }
    return $this.Records[$this.Index]
}
$folder = [pscustomobject]@{ Items = $items }
$session = [pscustomobject]@{ Folder = $folder }
$session | Add-Member ScriptMethod GetDefaultFolder { param($Kind); Assert ($Kind -eq 9) 'Calendar folder only'; return $this.Folder }
$fakeOutlook = [pscustomobject]@{ Session = $session }
$fakeOutlook | Add-Member ScriptMethod GetNamespace { param($Kind); Assert ($Kind -eq 'MAPI') 'MAPI namespace'; return $this.Session }
function New-Object {
    param([string]$ComObject)
    Assert ($ComObject -eq 'Outlook.Application') 'Only Outlook application requested'
    return $fakeOutlook
}
function Run-Export {
    param([switch]$BusyOnly)
    & (Join-Path $root 'outlook/Export-OutlookCalendar.ps1') -SourceId client-a -OutputDirectory $directory -BusyOnly:$BusyOnly | Out-Null
}
try {
    $items.Records = @(
        (Appointment 'series' '2026-10-03T09:00:00' '2026-10-03T10:00:00'),
        (Appointment 'private' '2026-10-03T11:00:00' '2026-10-03T12:00:00' 2),
        (Appointment 'cancelled' '2026-10-03T13:00:00' '2026-10-03T14:00:00' 0 $false 5),
        (Appointment 'all-day' '2026-10-04T00:00:00' '2026-10-05T00:00:00' 0 $true),
        (Appointment 'series' '2026-10-10T09:00:00' '2026-10-10T10:00:00')
    )
    Run-Export
    $path = Join-Path $directory 'client-a.snapshot.json'
    $raw = [IO.File]::ReadAllText($path)
    $data = $raw | ConvertFrom-Json
    Assert $items.RestrictCalled 'Restrict called'
    Assert ($items.FilterText -match '\[Start\] <' -and $items.FilterText -match '\[End\] >') 'Bounded overlap filter'
    Assert ($data.events.Length -eq 4) 'Cancelled meeting excluded'
    Assert ($data.events[0].start -eq '2026-10-03T09:00:00Z') 'COM UTC not converted twice'
    Assert ($data.events[1].subject -eq 'Busy') 'Sensitivity-marked subject hidden'
    Assert ($data.events[2].all_day -and $data.events[2].end -eq '2026-10-05') 'All-day exclusive end'
    Assert ($data.events[0].id -ne $data.events[3].id) 'Recurring occurrences have distinct identities'
    Assert ($data.events[0].id -match '^[a-f0-9]{64}$') 'Identifiers hashed'
    Assert ($raw -notmatch 'NEVER EXPORT|GlobalAppointmentID|Attendees|Location|Body') 'No private fields exported'
    Assert ($data.complete -and $data.source_id -eq 'client-a') 'Complete source snapshot'
    Write-Output 'PASS: timing, recurrence identities, cancellation, all-day dates, sensitivity, privacy'

    $oldIds = @($data.events | ForEach-Object { $_.id })
    Run-Export
    $again = [IO.File]::ReadAllText($path) | ConvertFrom-Json
    Assert ((@($again.events | ForEach-Object { $_.id }) -join ',') -eq ($oldIds -join ',')) 'IDs stable across exports'
    Run-Export -BusyOnly
    $busy = [IO.File]::ReadAllText($path) | ConvertFrom-Json
    Assert (@($busy.events | Where-Object { $_.subject -ne 'Busy' }).Length -eq 0) 'BusyOnly hides all titles'
    Write-Output 'PASS: stable identities and BusyOnly mode'

    $previous = [IO.File]::ReadAllText($path)
    $items.ReadFailure = $true
    $failed = $false
    try { Run-Export } catch { $failed = $true }
    Assert $failed 'Read errors propagate'
    Assert ([IO.File]::ReadAllText($path) -eq $previous) 'Read error leaves previous snapshot intact'
    $items.ReadFailure = $false
    $items.Records = @($items.Records[0], $items.Records[0])
    $failed = $false
    try { Run-Export } catch { $failed = $true }
    Assert $failed 'Duplicate identities abort'
    Assert ([IO.File]::ReadAllText($path) -eq $previous) 'Duplicate error leaves previous snapshot intact'
    Write-Output 'PASS: no partial snapshot on read or identity failure'

    $items.Records = @()
    Run-Export
    $empty = [IO.File]::ReadAllText($path) | ConvertFrom-Json
    Assert ($empty.events -is [array] -and $empty.events.Length -eq 0) 'Empty events serialized as an array'
    Write-Output 'PASS: empty calendar snapshot'
    Write-Output 'All mocked Outlook exporter tests passed.'
} finally {
    if (Test-Path -LiteralPath $directory) { Remove-Item -LiteralPath $directory -Recurse -Force }
}
