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
function Get-EnvValue([string]$Name) {
    $value = [Environment]::GetEnvironmentVariable($Name)
    if ([string]::IsNullOrEmpty($value)) { return $null }
    return $value
}
function Test-Elevated {
    try {
        $principal = [Security.Principal.WindowsPrincipal]::new([Security.Principal.WindowsIdentity]::GetCurrent())
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch { return $false }
}
function Get-OutlookExecutable {
    $candidates = [Collections.Generic.List[string]]::new()
    foreach ($root in @('ProgramFiles', 'ProgramFiles(x86)')) {
        $base = Get-EnvValue $root
        if (-not $base) { continue }
        foreach ($relative in @('Microsoft Office\root\Office16\OUTLOOK.EXE', 'Microsoft Office\Office16\OUTLOOK.EXE', 'Microsoft Office\root\Office15\OUTLOOK.EXE')) {
            $candidates.Add((Join-Path $base $relative))
        }
    }
    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) { return $candidate }
    }
    return $null
}
function Get-ComErrorDetail($ErrorRecord) {
    $messages = [Collections.Generic.List[string]]::new()
    $exception = $ErrorRecord.Exception
    while ($null -ne $exception) {
        $message = (([string]$exception.Message) -replace '\s+', ' ').Trim()
        if ($message) {
            if ($message.Length -gt 200) { $message = $message.Substring(0, 200) + '...' }
            $messages.Add($message)
        }
        $exception = $exception.InnerException
    }
    $detail = 'HRESULT {0}' -f ('0x{0:X8}' -f $ErrorRecord.Exception.HResult)
    if ($messages.Count) { return $detail + ': ' + ($messages -join ' | ') }
    return $detail
}
function Get-ComRemediation([int]$HResult) {
    $tips = [Collections.Generic.List[string]]::new()
    if ($HResult -eq -2147024891) {
        $tips.Add('Access denied (0x80070005): Outlook and PowerShell are usually running at different privilege levels. Close Outlook completely, then open PowerShell the same way you open Outlook - normally both non-elevated. Running Outlook as administrator while PowerShell is not (or the reverse) causes this.')
    } elseif ($HResult -eq -2146822884) {
        $tips.Add('Generic Outlook error (0x800A03EC): finish any first-run/profile dialogs, make sure the calendar opens normally in Outlook, and check that File > Options > General does not have "Always use New Outlook" enabled.')
    }
    $tips.Add('Bitness must match Outlook: use C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe (64-bit) for 64-bit Outlook, or C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe (32-bit) for 32-bit Outlook. The report above lists both.')
    $tips.Add('Close any lingering Outlook dialog or add-in prompt that blocks automation, then retry.')
    $tips.Add('If company policy blocks Outlook COM automation, ask IT for an approved alternative. Do not weaken security settings to force it.')
    return $tips
}

if ($Diagnose) {
    Write-Output 'Classic Outlook automation report'
    Write-Output '=================================='
    $classic = $null -ne (Get-Item 'Registry::HKEY_CLASSES_ROOT\Outlook.Application' -ErrorAction SilentlyContinue)
    $newOutlook = @(Get-Process -Name 'olk' -ErrorAction SilentlyContinue).Length -gt 0
    $runningOutlook = @(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue).Length -gt 0
    $powerShellBits = if ([IntPtr]::Size -eq 8) { '64-bit' } else { '32-bit' }
    $outlookExe = Get-OutlookExecutable
    $outlookBits = 'not found'
    if ($null -ne $outlookExe) {
        $x86Root = Get-EnvValue 'ProgramFiles(x86)'
        if ($x86Root -and $outlookExe.StartsWith($x86Root, [StringComparison]::OrdinalIgnoreCase)) { $outlookBits = '32-bit' }
        elseif ($null -ne (Get-EnvValue 'ProgramFiles')) { $outlookBits = '64-bit' }
    }
    Write-Output "Classic Outlook automation registered: $classic"
    Write-Output "New Outlook currently running (olk.exe): $newOutlook"
    Write-Output "Classic Outlook currently running: $runningOutlook"
    Write-Output "PowerShell: $powerShellBits"
    Write-Output "OUTLOOK.EXE: $outlookBits"
    if ($null -ne $outlookExe) { Write-Output "Outlook path: $outlookExe" }
    Write-Output "PowerShell elevated: $(Test-Elevated)"
    if ($powerShellBits -ne $outlookBits -and $outlookBits -ne 'not found') { Write-Output 'WARNING: PowerShell and Outlook bitness differ. This alone prevents COM automation.' }
    Write-Output 'OUTLOOK.EXE / File > Office Account means classic Outlook. olk.exe or no File menu usually means new Outlook.'
    $comApp = $null
    $probeSession = $null
    $probeStores = $null
    try {
        $comApp = New-Object -ComObject Outlook.Application
        Write-Output 'COM automation: AVAILABLE'
        try {
            $probeSession = $comApp.GetNamespace('MAPI')
            $probeStores = $probeSession.Stores
            Write-Output "Calendar stores visible: $($probeStores.Count)"
            Write-Output 'Run -ListCalendars to see store names locally, then export with -StoreDisplayName.'
        } catch {
            Write-Output 'Outlook started, but the MAPI namespace could not be read. The profile may be unavailable for this session.'
        }
    } catch {
        $hresult = 0
        try { $hresult = $_.Exception.HResult } catch { $hresult = 0 }
        Write-Output 'COM automation: UNAVAILABLE'
        Write-Output "Detail: $(Get-ComErrorDetail $_)"
        foreach ($tip in (Get-ComRemediation -HResult $hresult)) { Write-Output "  - $tip" }
    } finally {
        foreach ($object in @($probeStores, $probeSession, $comApp)) { Release-Com $object }
    }
    Write-Output 'Note: this report can contain your Windows account name. Redact it before sharing.'
    return
}

$app = $session = $stores = $store = $folder = $items = $restricted = $item = $null
$temp = $null
try {
    try { $app = New-Object -ComObject Outlook.Application }
    catch {
        $code = '0x{0:X8}' -f $_.Exception.HResult
        throw "Classic Outlook automation is unavailable (HRESULT $code). Run -Diagnose for bitness, elevation, and remediation guidance. This script cannot automate new Outlook or bypass company restrictions."
    }
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
