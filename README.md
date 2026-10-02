# Outlook → personal Google Calendar bridge

Combine several corporate calendars in a dedicated **Work overview** calendar in your personal Google account. View it alongside your personal events on your iPhone. The corporate accounts stay on their dedicated PCs.

```text
Corporate PC A: classic Outlook → PowerShell exporter → private GitHub data repo
Corporate PC B: classic Outlook → PowerShell exporter → private GitHub data repo
                                                          ↓
Personal computer: Python importer → personal Google “Work overview” calendar
                                                          ↓
                                                Google Calendar on iPhone
```

**This is a one-way snapshot mirror, not a live or two-way integration.** No accounts, repositories, tokens, or scheduled tasks are created automatically. You must configure and authorize them yourself. It does not circumvent conditional access, Outlook security prompts, or employer controls.

## 1. Get permission first

Obtain permission from **each employer** to copy calendar subjects and times to both GitHub and your personal Google account. Subject lines alone may reveal clients, projects, legal matters, or personal information. Being able to read an event does not authorize exporting it.

- Use a **separate private data repository** with no collaborators, public Pages, or Actions workflows. Keep the code in another repository.
- GitHub is **not a confidential vault**. Every upload creates a commit; old subjects remain in history after events change or are deleted. Making the repo public later exposes that history. Deleting the current file is not sufficient to erase history.
- Google events are marked private, but calendar owners, administrators, or sufficiently privileged collaborators may still see them. Do not share the target calendar.
- Use opaque source names (`client-a`, `client-b`), not corporate account addresses or employer names.
- Private, personal, and confidential Outlook appointments are exported as `Busy`. Add `-BusyOnly` to hide **all** subjects; approval is still needed to transfer availability.
- The code repository ignores exports, local configuration, OAuth credentials, and tokens. Never force-add them. GitHub tokens are read from an environment variable, not the configuration.
- If employer approval is denied, **do not run the upload or Google import**. There is also a local-file input option for approved non-GitHub transfer methods.

## 2. Detect your Outlook version on each corporate PC

Open Outlook. Classic Outlook normally has **File → Office Account**. New Outlook normally has no File menu and may have a “New Outlook” label/toggle. Task Manager → Details shows `OUTLOOK.EXE` for classic and `olk.exe` for new Outlook. A browser tab is Outlook on the web.

Download/copy the `outlook` scripts using an employer-approved method. In **Windows PowerShell 5.1**, from the code folder:

```powershell
.\outlook\Export-OutlookCalendar.ps1 -Diagnose
```

This reports classic COM registration, new Outlook detection, PowerShell and Outlook bitness, whether PowerShell is elevated, and whether COM automation actually starts. It also surfaces the underlying HRESULT and remediation steps when automation fails.

If COM automation reports `UNAVAILABLE`, the usual causes are:

- **Bitness mismatch.** PowerShell and Outlook must match: use `C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe` for 64-bit Outlook, or `C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe` for 32-bit Outlook. The report lists both.
- **Privilege mismatch.** If Outlook runs elevated and PowerShell does not (or the reverse), COM fails with access denied. Close Outlook fully and open PowerShell the same way you open Outlook—normally both without elevation.
- **A blocking dialog.** Finish any first-run, profile, or add-in prompt so the calendar opens normally.
- **New Outlook enabled.** Check `File > Options > General` for `Always use New Outlook`.
- **Company policy.** If policy blocks Outlook automation, ask IT for an approved alternative rather than weakening security settings.

Both Outlook versions can be installed simultaneously, so registration alone is not proof that your account is configured in classic Outlook.

The report can contain your Windows account name; redact it before sharing it.

If classic Outlook is available, open it, wait until the calendar is up to date, then run:

```powershell
.\outlook\Export-OutlookCalendar.ps1 -ListCalendars
```

This lists mailbox/store display names **locally**; they are not uploaded. Confirm the intended account is configured in classic Outlook. A store's default Calendar is supported; shared calendars, sub-calendars, and browser-only/new Outlook are not supported by this initial exporter.

**If you only have new Outlook/web:** stop here and ask IT whether Microsoft Graph access or an approved calendar-export mechanism is allowed. Graph usually requires application authorization and potentially administrator consent. Do not switch clients, change execution policy, install software, or bypass security prompts against company policy.

## 3. Export a calendar on each corporate PC

Choose a distinct, stable source ID per calendar. On PC A:

```powershell
.\outlook\Export-OutlookCalendar.ps1 -SourceId client-a -StoreDisplayName 'DISPLAY NAME FROM LIST'
```

On PC B, use `client-b`. If the Outlook profile contains only the intended account, omit `-StoreDisplayName` to use its default calendar.

Exports are written to `exports/client-a.snapshot.json`. Defaults: 7 days back, 90 days ahead. Customize with `-DaysBack 14 -DaysAhead 180`; add `-BusyOnly` for availability-only output.

The exporter reads only appointment subject, timing, sensitivity, cancellation status, and identity. Output contains:

- Subject (or `Busy` for sensitivity-marked events).
- UTC start/end for timed appointments; date-only start/exclusive end for all-day events.
- SHA-256 hashed occurrence identity; opaque source ID.
- Export timestamp, covered window, schema version, and completion marker.

**Never exported:** event bodies/descriptions, attachments, meeting links from bodies, location, organizer, attendees, or corporate account email addresses. A meeting link typed into the **subject** is still part of the subject; use `-BusyOnly` if this is a concern.

Recurrences are expanded into individual occurrences, including Outlook's edited exceptions, inside the date window. Identical subject/time events from different companies remain distinct. Cancelled meetings are omitted. Deletions are inferred from a complete new snapshot.

Any read failure aborts the export; it does not write a partial snapshot. It leaves the previous file intact. **Upload only if this export succeeded.** Outlook cached/offline calendars can still be outdated even when COM reads succeed—check connectivity and synchronization in Outlook first.

## 4. Configure the private GitHub relay

Create a **private**, dedicated repository, initialize it with a README so `main` exists, and keep it private permanently. Do not put secrets or OAuth files there.

Create separate [fine-grained GitHub personal access tokens](https://github.com/settings/personal-access-tokens):

- On each corporate PC: selected data repository only, **Contents: read and write**, with an expiry. GitHub also grants the necessary Metadata read permission.
- On the personal computer: selected data repository only, **Contents: read**.

The writer can read all snapshots in this repository. If cross-company access is not permitted, a single shared repository is **not appropriate**; use isolated relays or an approved alternative instead. Never place a personal token on a corporate PC without permission.

In the **current PowerShell session**, enter the writer token without placing it in shell history:

```powershell
$secureToken = Read-Host 'GitHub writer token' -AsSecureString
$env:CALENDAR_GITHUB_TOKEN = [System.Net.NetworkCredential]::new('', $secureToken).Password
```

The environment variable contains a plaintext token in process memory; clear it afterward. For unattended use, obtain an approved credential store from IT—do not hardcode a token in a scheduled task or script.

After a successful export and employer approval:

```powershell
.\outlook\Publish-Snapshot.ps1 -SnapshotPath .\exports\client-a.snapshot.json -Repository YOUR-USER/YOUR-PRIVATE-DATA-REPO -EmployerApproved
Remove-Item Env:\CALENDAR_GITHUB_TOKEN
```

The publisher verifies the repository is private, checks the export's field allowlist, and writes only `snapshots/client-a.snapshot.json`. No Git installation is needed. Non-default branches can use `-Branch`. Concurrent upload conflicts fail safely; retry after exporting again. Repository administrators can change privacy or access afterward; the scripts cannot prevent that.

## 5. Configure Google on your personal computer

Requirements: **Python 3.10+**, a browser, and your personal Google account. The importer uses only Python's standard library; no package installation is needed.

1. In [Google Cloud Console](https://console.cloud.google.com/), create a project and enable **Google Calendar API**.
2. Configure Google Auth Platform's app branding and audience. For an External app in Testing, add your personal Google account as a test user.
3. Add only the scope `https://www.googleapis.com/auth/calendar.app.created` to Data Access. This allows secondary app-created calendars rather than your existing personal calendar.
4. Create an OAuth client of type **Desktop app**, download its JSON, and save it as `secrets/client_secret.json` in the code folder. Do not use a service account or a Web application client.
5. Copy [config.example.json](config.example.json) to a local `config.json`. Set your repository, branch, sources, and optional subject prefixes. Set `employer_approved` to `true` **only after approval from all employers**. Remove sample sources you are not using.
6. Optionally set `calendar_timezone` to an IANA name such as `Europe/London` or `America/New_York`. Timed events use UTC so they display in the viewing device's timezone. All-day dates are preserved as Outlook calendar dates.

Paths in the configuration are relative to its folder. Protect the `secrets` folder with your OS's account permissions and disk encryption; files are not encrypted by this tool. Unix token files use mode 0600, but on Windows access depends on inherited ACLs.

From the code folder:

```text
python calendar_bridge.py authorize
python calendar_bridge.py init
```

Authorization opens Google's consent screen. **Select your personal Google account**, not a corporate account. The browser callback listens on a randomly allocated loopback-only port, uses OAuth state checking and PKCE, and times out after three minutes. Tokens are saved locally. `init` creates a dedicated calendar and saves its ID to your configuration; it refuses to create another if an ID already exists. It never targets `primary`.

**OAuth Testing caveat:** Google's refresh tokens for External apps in Testing typically expire after seven days for this Calendar scope. Rerun `authorize` when needed. For long-term operation, review Google's publishing/verification requirements before changing the app's publishing status; authorization can still be revoked later. Keep using the same OAuth project/client for the app-created calendar.

## 6. Preview and apply the combined sync

Set the personal PC's **read-only** GitHub token in `CALENDAR_GITHUB_TOKEN`. In PowerShell use the same secure prompt shown above, but enter the reader token.

```text
python calendar_bridge.py sync
```

This performs a **read-only dry run against Google** and prints only create/update/delete/unchanged counts by source—no subjects. Then:

```text
python calendar_bridge.py sync --apply
```

You may supply a configuration elsewhere with `--config PATH`. `authorize` and `init` are explicitly mutating setup commands; `--apply` is required only for sync.

Behavior and safeguards:

- All configured sources must be present, complete, valid, and fresh (48 hours by default). If any source fails, no Google calendar writes begin.
- All GitHub snapshots are read from one pinned commit; older snapshots than the last successfully applied version are rejected.
- Stable identities avoid duplicates on repeat runs. Title edits update existing copies; moved occurrences are reconciled as old/new occurrences.
- Only events marked as created by this bridge for that source are eligible for changes. Unrelated events and your personal calendar are not changed.
- Deletions apply **only inside the new export window**. Older mirrored events remain as history; events beyond the window are left alone. Shortening the export window does not clear those copies. Removing a source from configuration does not delete its previous copies.
- An empty snapshot that would delete events is blocked. Deletions exceeding 25 events or 50% of the reconciled existing events are blocked. After inspecting the dry run and Outlook, intentionally allow them with `--allow-empty` and/or `--allow-large-delete`. Do not put these overrides in unattended schedules.
- All creations/updates across sources happen before deletions. An upsert failure skips all deletions. The Google API has no transaction: earlier successful writes can remain after later failures. Rerunning reconciles them.
- GET requests retry temporary server errors; writes are not blindly retried. Deterministic event IDs recover interrupted inserts; deleted Google IDs are recreated as new copies.
- Events have no attendees, no reminders, and `sendUpdates=none`. Do not add attendees or confidential content to mirrored events: the mirror is source-controlled and can replace your edits. Editing Google copies does not update Outlook.
- Snapshot completion is a contract, not a cryptographic guarantee. Anyone with data-repository write access can alter snapshots. GitHub's access controls and token protection are part of the trust model.

If approval allows local files instead of GitHub, set each source to e.g. `{"id": "client-a", "prefix": "[A] ", "local_file": "exports/client-a.snapshot.json"}`. Transfer the file to the personal computer through the approved method. No GitHub token is needed when all sources are local. The same employer permission requirement applies to copying the data to Google.

## 7. See it on your iPhone

In Google Calendar on your personal computer, confirm **Work overview** exists and contains the expected events. In the Google Calendar iPhone app, sign into your personal Google account and enable that calendar in the calendar list. Corporate accounts do not need to be added to the phone.

You can also use Apple Calendar with your personal Google account; ensure the new secondary calendar is selected and enabled for sync. The Google Calendar app is the simplest initial verification.

## 8. Optional scheduling, after a successful manual test

- **Corporate PCs:** with IT permission, use Windows Task Scheduler under your signed-in account, **Run only when user is logged on**. Keep classic Outlook open and synchronized. Export then publish only on success. Do not use SYSTEM or “run whether user is logged on or not” for Outlook automation. Microsoft does not recommend unattended/server-side Office automation.
- **Personal computer:** schedule `python calendar_bridge.py sync --apply --config ABSOLUTE_CONFIG_PATH` using your OS scheduler. Ensure that its working directory, Python path, and approved reader-token source are set correctly. A personal PC must be awake and online.
- Start with daily or hourly runs, not near-real-time polling. Schedule corporate exports before personal imports. Expired tokens, sleeping/disconnected PCs, and unsynchronized Outlook will interrupt freshness.
- Error logs contain counts and generic errors rather than subjects/tokens. A stale snapshot causes sync to stop; it **does not remove old Google events**, so treat the overview as potentially outdated until a successful sync.
- Avoid simultaneous importer runs. A lock file prevents overlap. If a process crashes, confirm no importer is running before deleting `secrets/.bridge.lock`.

## Troubleshooting

| Symptom | Action |
| --- | --- |
| Classic automation unavailable | Run `-Diagnose` and read the reported HRESULT. Match PowerShell bitness to Outlook, and run both at the same privilege level. New Outlook/web needs another approved integration. |
| Script execution blocked | Ask IT to approve/sign the script; do not bypass policy. |
| Wrong or empty source calendar | Inspect Outlook, use `-ListCalendars`, specify the intended `-StoreDisplayName`; do not override deletion guards blindly. |
| GitHub HTTP 403/404 | Check token expiry, selected repository, Contents permissions, and branch. Missing/unauthorized sources stop the import. |
| Google authorization fails after a week | Check OAuth Testing status and rerun `authorize`. |
| Google HTTP 403 | Check Calendar API enablement, granted scope, and that the calendar was created with the same OAuth application. |
| Unexpected mass deletion | Review export coverage and recurrence changes in Outlook. Rerun dry run after a fresh export. |
| Times look incorrect | Compare known appointments, including DST and all-day dates, before scheduling. Timed UTC is converted by your calendar app. |

## Validation and limitations

```text
python -m unittest discover -s tests -v
python -m py_compile calendar_bridge.py tests/test_calendar_bridge.py
```

On Windows with PowerShell 5.1, run the mocked exporter and publisher tests (no Outlook, GitHub, or Google account required):

```powershell
.\tests\Test-OutlookExport.ps1
.\tests\Test-SnapshotPublish.ps1
.\tests\Test-OutlookDiagnose.ps1
```

Automated tests exercise privacy allowlists, validation, repeated syncs, changes, source isolation, deletion safety, write recovery, exporter behavior using fake Outlook objects, and publisher safeguards using fake GitHub responses. They do **not** prove that a particular employer's Outlook profile, locale, recurrence cache, security settings, GitHub permissions, or Google account works. Perform an approved manual pilot with a timed event, a recurrence with an edited/deleted occurrence, an all-day event, a sensitivity-marked event, a reschedule, and a deletion before relying on the overview.

No external dependencies or typechecker are configured; Python syntax compilation and tests are the available checks. No live account credentials are included. The default export window limits how far ahead you'll see events; no Microsoft Graph support, automatic credential-store setup, or stale-data notification is implemented yet.

References: [Outlook recurrence filtering](https://learn.microsoft.com/en-us/office/vba/api/outlook.items.includerecurrences), [Google Calendar scopes](https://developers.google.com/workspace/calendar/api/auth), [Google desktop OAuth](https://developers.google.com/identity/protocols/oauth2/native-app).
