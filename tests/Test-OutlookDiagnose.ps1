#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$root = Split-Path -Parent $PSScriptRoot
$script = Join-Path $root 'outlook/Export-OutlookCalendar.ps1'
$out = [Collections.Generic.List[string]]::new()

function Assert($Condition, [string]$Message) {
    if (-not $Condition) { throw "Assertion failed: $Message" }
}
# Diagnose must never exit with a non-zero result, even when COM creation fails.
function Run-Diagnose {
    $out.Clear()
    & $script -Diagnose | ForEach-Object { $out.Add([string]$_) }
}

# Case 1: COM creation fails the way it does in the field (Outlook running elevated, or bitness mismatch).
function New-Object {
    param([string]$ComObject)
    Assert ($ComObject -eq 'Outlook.Application') 'Only Outlook application requested'
    throw [Runtime.InteropServices.COMException]::new('Exception calling "InvokeMember" on type "Outlook.Application". Exception from HRESULT: 0x800A03EC.', -2146822884)
}
Run-Diagnose
$report = $out -join "`n"
Assert ($report -match 'COM automation: UNAVAILABLE') 'Failure is reported explicitly'
Assert ($report -match '0x800A03EC') 'Underlying HRESULT is surfaced instead of a generic message'
Assert ($report -match 'PowerShell: (64|32)-bit') 'PowerShell bitness reported'
Assert ($report -match 'OUTLOOK.EXE: ') 'Outlook install reported'
Assert ($report -match 'Always use New Outlook') 'New Outlook switch suggested'
Assert ($report -match 'SysWOW64.*powershell\.exe' -and $report -match 'System32.*powershell\.exe') 'Both bitness-specific PowerShell paths given'
Assert ($report -match 'ask IT') 'Policy route mentioned instead of weakening security'
Assert ($report -match 'Redact it before sharing') 'Redaction warning present'
Assert ($report -notmatch 'Exception calling "InvokeMember" on type "Outlook\.Application".{0,40}@') 'No account details assumed'
Write-Output 'PASS: COM failure is diagnosed with HRESULT and remediation'

# Case 2: COM succeeds; the report must be actionable and must not touch accounts.
Remove-Item Function:\New-Object
$store = [pscustomobject]@{ DisplayName = 'Corp Account' }
$folder = [pscustomobject]@{}
$store | Add-Member ScriptMethod GetDefaultFolder { param($Kind); return $folder }
$session = [pscustomobject]@{ Stores = @($store) }
$session | Add-Member ScriptMethod GetDefaultFolder { param($Kind); return $folder }
$fakeOutlook = [pscustomobject]@{ Session = $session }
$fakeOutlook | Add-Member ScriptMethod GetNamespace { param($Kind); Assert ($Kind -eq 'MAPI') 'MAPI namespace'; return $this.Session }
function New-Object {
    param([string]$ComObject)
    Assert ($ComObject -eq 'Outlook.Application') 'Only Outlook application requested'
    return $fakeOutlook
}
Run-Diagnose
$report = $out -join "`n"
Assert ($report -match 'COM automation: AVAILABLE') 'Success reported'
Assert ($report -match 'Calendar stores visible: 1') 'Store count reported without naming accounts'
Assert ($report -match 'ListCalendars') 'Next step suggested'
Assert ($report -notmatch 'Corp Account') 'Diagnose does not print account names'
Write-Output 'PASS: successful automation is diagnosed without exposing accounts'
Write-Output 'All Outlook diagnose tests passed.'