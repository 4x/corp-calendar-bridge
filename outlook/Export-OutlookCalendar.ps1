#Requires -Version 5.1
[CmdletBinding(DefaultParameterSetName = 'Export')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Diagnose')][switch]$Diagnose,
    [Parameter(Mandatory = $true, ParameterSetName = 'List')][switch]$ListCalendars,
    [Parameter(Mandatory = $true, ParameterSetName = 'Export')]
    [ValidatePattern('^[a-z0-9][a-z0-9_-]{0,39}$')][string]$SourceId,
    [Parameter(ParameterSetName = 'Export')][string]$OutputDirectory = '.\exports',
    [Parameter(ParameterSetName = 'Export')][ValidateRange(0, 90)][int]$DaysBack = 7,
    [Parameter(ParameterSetName = 'Export')][ValidateRange(1, 365)][int]$DaysAhead = 90,
    [Parameter(ParameterSetName = 'Export')][string]$StoreDisplayName,
    [Parameter(ParameterSetName = 'Export')][switch]$BusyOnly
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Release-Com($Object) {
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Object)
    }
}
function Utc-Text([datetime]$Date) {
    return $Date.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", [Globalization.CultureInfo]::InvariantCulture)
}
function Hash-Text([string]$Text) {
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text)))).Replace('-', '').ToLowerInvariant() }
    finally { $sha.Dispose() }
}

if ($Diagnose) {
    $classic = $null -ne (Get-Item 'Registry::HKEY_CLASSES_ROOT\Outlook.Application' -ErrorAction SilentlyContinue)
    $new = @(Get-Process -Name 'olk' -ErrorAction SilentlyContinue).Length -gt 0
    Write-Output "Classic Outlook automation registered: $classic"
    Write-Output "New Outlook currently running (olk.exe): $new"
    Write-Output 'OUTLOOK.EXE / File > Office Account means classic Outlook. olk.exe or no File menu usually means new Outlook.'
    Write-Output 'Registration alone does not prove your account is configured in classic Outlook. Open it, confirm the calendar, then run -ListCalendars.'
    Write-Output 'Browser-only/new Outlook is not supported by this exporter; ask IT about an approved Microsoft Graph integration.'
    return
}

$app = $session = $stores = $store = $folder = $items = $restricted = $item = $null
$temp = $null
try {
    try { $app = New-Object -ComObject Outlook.Application }
    catch { throw 'Classic Outlook automation is unavailable. Run -Diagnose and consult IT; this script cannot automate new Outlook or bypass company restrictions.' }
    $session = $app.GetNamespace('MAPI')
    if ($ListCalendars) {
        $stores = $session.Stores
        for ($i = 1; $i -le $stores.Count; $i++) {
            $store = $stores.Item($i)
            try {
                try { $folder = $store.GetDefaultFolder(9) }
                catch { continue }
                Write-Output $store.DisplayName
            } finally { Release-Com $folder; $folder = $null; Release-Com $store; $store = $null }
        }
        return
    }
    if ($StoreDisplayName) {
        $stores = $session.Stores
        $matches = 0
        for ($i = 1; $i -le $stores.Count; $i++) {
            $store = $stores.Item($i)
            try {
                if ($store.DisplayName -eq $StoreDisplayName) {
                    $matches++
                    if ($matches -gt 1) { throw 'Store name is ambiguous. Use a profile with a uniquely named store.' }
                    $folder = $store.GetDefaultFolder(9)
                }
            } finally { Release-Com $store; $store = $null }
        }
        if ($matches -ne 1) { throw 'No store with that display name was found. Run -ListCalendars.' }
    } else {
        $folder = $session.GetDefaultFolder(9)
        Write-Output 'Using the Outlook profile default calendar. Use -StoreDisplayName to choose an account explicitly.'
    }

    $start = (Get-Date).Date.AddDays(-$DaysBack)
    $end = (Get-Date).Date.AddDays($DaysAhead)
    $culture = [Globalization.CultureInfo]::CurrentCulture
    $filter = "[Start] < '" + $end.ToString('g', $culture) + "' AND [End] > '" + $start.ToString('g', $culture) + "'"
    $items = $folder.Items
    $items.Sort('[Start]', $false)
    $items.IncludeRecurrences = $true
    $restricted = $items.Restrict($filter)
    $events = [Collections.Generic.List[object]]::new()
    $seen = [Collections.Generic.HashSet[string]]::new()
    $visited = 0
    # Never use Count with IncludeRecurrences: a series may have no end date.
    $item = $restricted.GetFirst()
    while ($null -ne $item) {
        try {
            $visited++
            if ($visited -gt 10000) { throw 'Safety limit reached. Reduce the date window. No new snapshot was written.' }
            if ($item.Class -ne 26) { throw 'Unexpected non-appointment item. No partial snapshot will be written.' }
            if ($item.Start -ge $end) { break }
            if ($item.End -gt $start -and $item.Start -lt $end -and $item.MeetingStatus -notin @(5, 7)) {
                $globalId = [string]$item.GlobalAppointmentID
                if ([string]::IsNullOrWhiteSpace($globalId)) { throw 'An appointment has no stable identifier. No partial snapshot will be written.' }
                $allDay = [bool]$item.AllDayEvent
                if ($allDay) {
                    $eventStart = ([datetime]$item.Start).ToString('yyyy-MM-dd')
                    $eventEnd = ([datetime]$item.End).ToString('yyyy-MM-dd')
                } else {
                    # COM exposes UTC values as DateTimeKind.Unspecified; do not convert them from local time again.
                    $eventStart = Utc-Text ([datetime]::SpecifyKind([datetime]$item.StartUTC, [DateTimeKind]::Utc))
                    $eventEnd = Utc-Text ([datetime]::SpecifyKind([datetime]$item.EndUTC, [DateTimeKind]::Utc))
                }
                # Flatten recurrences. A moved occurrence gets a new ID; reconciliation removes the old one.
                $id = Hash-Text ($globalId + '|' + $eventStart)
                if (-not $seen.Add($id)) { throw 'Duplicate appointment identity detected. No partial snapshot will be written.' }
                $subject = [string]$item.Subject
                if ($BusyOnly -or $item.Sensitivity -ne 0) { $subject = 'Busy' }
                if ([string]::IsNullOrWhiteSpace($subject)) { $subject = '(No subject)' }
                $events.Add([ordered]@{ id = $id; subject = $subject; all_day = $allDay; start = $eventStart; end = $eventEnd })
            }
        } finally { Release-Com $item; $item = $null }
        $item = $restricted.GetNext()
    }
    $snapshot = [ordered]@{
        schema_version = 1
        source_id = $SourceId
        generated_at = Utc-Text (Get-Date)
        complete = $true
        window = [ordered]@{ start = Utc-Text $start; end = Utc-Text $end }
        events = @($events.ToArray())
    }
    $directory = [IO.Path]::GetFullPath($OutputDirectory)
    [void][IO.Directory]::CreateDirectory($directory)
    $destination = Join-Path $directory "$SourceId.snapshot.json"
    $temp = Join-Path $directory ([IO.Path]::GetRandomFileName())
    [IO.File]::WriteAllText($temp, ($snapshot | ConvertTo-Json -Depth 6), [Text.UTF8Encoding]::new($false))
    if ([IO.File]::Exists($destination)) { [IO.File]::Replace($temp, $destination, [System.Management.Automation.Language.NullString]::Value) }
    else { [IO.File]::Move($temp, $destination) }
    $temp = $null
    Write-Output "Exported $($events.Count) occurrences to $destination. No body, location, attendee, or account address was exported."
} finally {
    if ($temp -and [IO.File]::Exists($temp)) { [IO.File]::Delete($temp) }
    foreach ($object in @($item, $restricted, $items, $folder, $store, $stores, $session, $app)) { Release-Com $object }
    # Do not quit Outlook: it may belong to the user.
}
