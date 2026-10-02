#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$root = Split-Path -Parent $PSScriptRoot
$directory = Join-Path $PSScriptRoot ('.test-publish-' + [guid]::NewGuid().ToString('N'))
$mock = [pscustomobject]@{ Private = $true; Calls = [Collections.Generic.List[object]]::new(); Uploaded = $null }
$previousToken = $env:CALENDAR_GITHUB_TOKEN

function Assert($Condition, [string]$Message) {
    if (-not $Condition) { throw "Assertion failed: $Message" }
}
function Invoke-RestMethod {
    param($Method, $Uri, $Headers, $MaximumRedirection, $ContentType, $Body)
    Assert ($MaximumRedirection -eq 0) 'Redirects disabled to protect token'
    $mock.Calls.Add(@{ Method = $Method; Uri = $Uri })
    if ($Method -eq 'PUT') {
        $mock.Uploaded = [Text.Encoding]::UTF8.GetString($Body) | ConvertFrom-Json
        return @{ content = @{ path = 'snapshots/client-a.snapshot.json' } }
    }
    if ($Uri -match '/branches/') { return @{ name = 'main' } }
    if ($Uri -match '/git/trees/') {
        return @{ truncated = $false; tree = @(@{ path = 'snapshots/client-a.snapshot.json'; type = 'blob'; sha = 'old-sha' }) }
    }
    return @{ private = $mock.Private }
}
function Run-Publish {
    param([switch]$Approved = $true)
    & (Join-Path $root 'outlook/Publish-Snapshot.ps1') -SnapshotPath (Join-Path $directory 'input.json') -Repository 'example/private-data' -EmployerApproved:$Approved | Out-Null
}
function Expect-Failure {
    param([scriptblock]$Action, [string]$Message)
    $failed = $false
    try { & $Action } catch { $failed = $true }
    Assert $failed $Message
}
try {
    [void][IO.Directory]::CreateDirectory($directory)
    $path = Join-Path $directory 'input.json'
    $inputData = [ordered]@{
        schema_version = 1; source_id = 'client-a'; generated_at = '2026-10-02T12:00:00Z'; complete = $true
        window = @{ start = '2026-09-25T00:00:00Z'; end = '2026-12-31T00:00:00Z' }
        events = @(@{ id = ('a' * 64); subject = 'Planning'; all_day = $false; start = '2026-10-03T09:00:00Z'; end = '2026-10-03T10:00:00Z' })
    }
    [IO.File]::WriteAllText($path, ($inputData | ConvertTo-Json -Depth 8), [Text.UTF8Encoding]::new($false))
    $env:CALENDAR_GITHUB_TOKEN = 'fake-test-token'
    Run-Publish
    Assert ($mock.Uploaded.sha -eq 'old-sha') 'Existing blob SHA used for conflict protection'
    Assert ($mock.Uploaded.branch -eq 'main') 'Selected branch used'
    $uploaded = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($mock.Uploaded.content)) | ConvertFrom-Json
    Assert ($uploaded.source_id -eq 'client-a' -and $uploaded.events[0].subject -eq 'Planning') 'Original snapshot uploaded'
    Assert (@($mock.Calls | Where-Object { $_.Uri -match '/contents/snapshots/client-a.snapshot.json$' }).Length -eq 1) 'Only designated snapshot path written'
    Write-Output 'PASS: private snapshot upload, destination and SHA conflict protection'

    $mock.Calls.Clear()
    $mock.Private = $false
    Expect-Failure { Run-Publish } 'Public repo rejected'
    Assert (@($mock.Calls | Where-Object { $_.Method -eq 'PUT' }).Length -eq 0) 'No upload to public repo'
    $mock.Private = $true
    $mock.Calls.Clear()
    Expect-Failure { Run-Publish -Approved:$false } 'Explicit approval required'
    Assert ($mock.Calls.Count -eq 0) 'No API calls without approval'
    $env:CALENDAR_GITHUB_TOKEN = ''
    Expect-Failure { Run-Publish } 'Missing token rejected'
    Assert ($mock.Calls.Count -eq 0) 'No API calls without token'
    $env:CALENDAR_GITHUB_TOKEN = 'fake-test-token'
    Write-Output 'PASS: public repository, missing approval, and missing token blocked'

    $inputData.events[0].body = 'PRIVATE BODY'
    [IO.File]::WriteAllText($path, ($inputData | ConvertTo-Json -Depth 8), [Text.UTF8Encoding]::new($false))
    $mock.Calls.Clear()
    Expect-Failure { Run-Publish } 'Unexpected private event fields rejected'
    Assert (@($mock.Calls | Where-Object { $_.Method -eq 'PUT' }).Length -eq 0) 'No private content uploaded'
    [void]$inputData.events[0].Remove('body')
    $inputData.complete = $false
    [IO.File]::WriteAllText($path, ($inputData | ConvertTo-Json -Depth 8), [Text.UTF8Encoding]::new($false))
    $mock.Calls.Clear()
    Expect-Failure { Run-Publish } 'Incomplete snapshot rejected'
    Assert (@($mock.Calls | Where-Object { $_.Method -eq 'PUT' }).Length -eq 0) 'No partial export uploaded'
    Write-Output 'PASS: privacy allowlist and completeness guard'
    Write-Output 'All mocked GitHub publisher tests passed.'
} finally {
    $env:CALENDAR_GITHUB_TOKEN = $previousToken
    if (Test-Path -LiteralPath $directory) { Remove-Item -LiteralPath $directory -Recurse -Force }
}
