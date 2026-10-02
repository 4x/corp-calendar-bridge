#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$SnapshotPath,
    [Parameter(Mandatory = $true)][ValidatePattern('^[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+$')][string]$Repository,
    [string]$Branch = 'main',
    [Parameter(Mandatory = $true)][switch]$EmployerApproved
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (-not $EmployerApproved) { throw 'Employer approval is required to transfer calendar subjects to GitHub.' }
if ([string]::IsNullOrWhiteSpace($env:CALENDAR_GITHUB_TOKEN)) {
    throw 'Set CALENDAR_GITHUB_TOKEN to a fine-grained token scoped to the private data repository, Contents: read and write.'
}
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$headers = @{
    Authorization = "Bearer $env:CALENDAR_GITHUB_TOKEN"
    Accept = 'application/vnd.github+json'
    'X-GitHub-Api-Version' = '2022-11-28'
    'User-Agent' = 'outlook-calendar-bridge'
}
function GitHub-Request([string]$Method, [string]$Uri, $Body = $null) {
    try {
        if ($null -eq $Body) { return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -MaximumRedirection 0 }
        return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -MaximumRedirection 0 -ContentType 'application/json; charset=utf-8' -Body ([Text.Encoding]::UTF8.GetBytes(($Body | ConvertTo-Json -Depth 8)))
    } catch {
        # Do not echo HTTP bodies or headers: they can contain corporate data or credentials.
        $status = 0
        if ($null -ne $_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode }
        throw "GitHub request failed (HTTP $status). Check connectivity, token permissions, branch, and repository."
    }
}

$metadata = GitHub-Request 'GET' "https://api.github.com/repos/$Repository"
if ($metadata.private -ne $true) { throw 'Refusing to publish to a public repository.' }
$bytes = [IO.File]::ReadAllBytes((Resolve-Path -LiteralPath $SnapshotPath).Path)
if ($bytes.Length -gt 900000) { throw 'Snapshot exceeds the safety limit (900 KB). Reduce the export window.' }
$snapshot = [Text.Encoding]::UTF8.GetString($bytes).TrimStart([char]0xFEFF) | ConvertFrom-Json
$expected = @('complete', 'events', 'generated_at', 'schema_version', 'source_id', 'window')
$actual = @($snapshot.PSObject.Properties.Name | Sort-Object)
if (($actual -join ',') -ne (($expected | Sort-Object) -join ',')) { throw 'Unexpected snapshot fields. Refusing upload.' }
if ($snapshot.schema_version -ne 1 -or $snapshot.complete -ne $true -or $snapshot.source_id -notmatch '^[a-z0-9][a-z0-9_-]{0,39}$') { throw 'Invalid or incomplete snapshot.' }
if ((@($snapshot.window.PSObject.Properties.Name | Sort-Object) -join ',') -ne 'end,start') { throw 'Unexpected window fields.' }
foreach ($event in $snapshot.events) {
    if ((@($event.PSObject.Properties.Name | Sort-Object) -join ',') -ne 'all_day,end,id,start,subject') { throw 'Unexpected event fields. Refusing upload of possible private content.' }
    if ($event.id -notmatch '^[a-f0-9]{64}$' -or $event.subject -isnot [string] -or $event.all_day -isnot [bool]) { throw 'Invalid event.' }
}
$path = "snapshots/$($snapshot.source_id).snapshot.json"
$branchQuery = [Uri]::EscapeDataString($Branch)
# Require an existing branch; no repository/branch creation is performed.
[void](GitHub-Request 'GET' "https://api.github.com/repos/$Repository/branches/$branchQuery")
# Listing the directory avoids treating authorization/network errors as 'file not found'.
$tree = GitHub-Request 'GET' "https://api.github.com/repos/$Repository/git/trees/$branchQuery`?recursive=1"
if ($tree.truncated) { throw 'Repository tree is too large. Use a dedicated small data repository.' }
$existing = @($tree.tree | Where-Object { $_.path -eq $path })
$body = @{
    message = "Refresh calendar snapshot: $($snapshot.source_id)"
    content = [Convert]::ToBase64String($bytes)
    branch = $Branch
}
if ($existing.Length -gt 0) {
    if ($existing[0].type -ne 'blob') { throw 'Snapshot path is not a file.' }
    $body.sha = $existing[0].sha
}
[void](GitHub-Request 'PUT' "https://api.github.com/repos/$Repository/contents/$path" $body)
Write-Output "Published $path to private repository $Repository. GitHub retains previous versions; do not use this as confidential storage."
