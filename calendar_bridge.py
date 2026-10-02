#!/usr/bin/env python3
"""Private Outlook snapshot -> app-created Google calendar. Python 3.10+."""
from __future__ import annotations

import argparse
import base64
import hashlib
import http.server
import json
import os
from pathlib import Path
import re
import secrets
import time
from contextlib import contextmanager
from datetime import date, datetime, timedelta, timezone
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import parse_qs, quote, urlencode
from urllib.request import HTTPRedirectHandler, Request, build_opener
import webbrowser

SCOPE = 'https://www.googleapis.com/auth/calendar.app.created'
OWNER = 'outlook-calendar-bridge-v1'
UTC = timezone.utc
MAX_BYTES = 900_000
SOURCE_RE = re.compile(r'[a-z0-9][a-z0-9_-]{0,39}\Z')
ID_RE = re.compile(r'[a-f0-9]{64}\Z')


class BridgeError(Exception):
    pass


class ApiError(BridgeError):
    def __init__(self, status: int):
        self.status = status
        super().__init__(f'API request failed (HTTP {status}); check authorization, permissions, and service availability.')


class NoRedirect(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        raise BridgeError('Unexpected HTTP redirect refused to protect credentials.')


def request(method: str, url: str, *, token: str | None = None,
            body: Any = None, form: bool = False, github: bool = False) -> Any:
    headers = {'User-Agent': OWNER}
    if token:
        headers['Authorization'] = 'Bearer ' + token
    if github:
        headers.update({'Accept': 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28'})
    data = None
    if body is not None:
        data = (urlencode(body) if form else json.dumps(body)).encode('utf-8')
        headers['Content-Type'] = 'application/x-www-form-urlencoded' if form else 'application/json'
    # Retry read-only calls only. Ambiguous write failures are reconciled on the next run.
    for attempt in range(4):
        try:
            with build_opener(NoRedirect).open(Request(url, data=data, headers=headers, method=method), timeout=30) as response:
                raw = response.read(8_000_001)
            if len(raw) > 8_000_000:
                raise BridgeError('API response exceeds safety limit.')
            return json.loads(raw) if raw else None
        except HTTPError as exc:
            if method == 'GET' and exc.code in (429, 500, 502, 503, 504) and attempt < 3:
                time.sleep(2 ** attempt)
                continue
            raise ApiError(exc.code) from None
        except (URLError, TimeoutError, OSError):
            raise BridgeError('Network request failed; no credentials or response bodies are logged.') from None
    raise BridgeError('Request retries exhausted.')


def read_json(path: Path) -> Any:
    with path.open(encoding='utf-8-sig') as handle:
        return json.load(handle)


def save_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
    temporary = path.with_name(path.name + '.' + secrets.token_hex(8) + '.tmp')
    try:
        fd = os.open(temporary, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
        with os.fdopen(fd, 'w', encoding='utf-8') as handle:
            json.dump(data, handle, indent=2, ensure_ascii=False)
            handle.write('\n')
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


@contextmanager
def lock(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
    try:
        fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
    except FileExistsError:
        raise BridgeError('Another sync may be running. If it crashed, remove secrets/.bridge.lock only after checking no importer is running.') from None
    try:
        os.close(fd)
        yield
    finally:
        path.unlink(missing_ok=True)


def utc(value: Any) -> datetime:
    if not isinstance(value, str):
        raise BridgeError('Expected an ISO timestamp.')
    try:
        result = datetime.fromisoformat(value.replace('Z', '+00:00'))
        if result.tzinfo is None:
            raise ValueError
        return result.astimezone(UTC)
    except ValueError:
        raise BridgeError('Invalid timestamp: timezone is required.') from None


def iso(value: datetime) -> str:
    return value.astimezone(UTC).isoformat().replace('+00:00', 'Z')


def exact_keys(value: Any, keys: set[str]) -> None:
    if not isinstance(value, dict) or set(value) != keys:
        raise BridgeError('Unexpected or missing snapshot fields; refusing potentially private content.')


def validate_snapshot(data: Any, source: str, max_age_hours: float = 48,
                      now: datetime | None = None) -> dict[str, Any]:
    now = now or datetime.now(UTC)
    exact_keys(data, {'schema_version', 'source_id', 'generated_at', 'complete', 'window', 'events'})
    if type(data['schema_version']) is not int or data['schema_version'] != 1 or data['complete'] is not True:
        raise BridgeError('Unsupported or incomplete snapshot.')
    if data['source_id'] != source or not SOURCE_RE.fullmatch(source):
        raise BridgeError('Snapshot source does not match configuration.')
    generated = utc(data['generated_at'])
    if generated > now + timedelta(minutes=5) or now - generated > timedelta(hours=max_age_hours):
        raise BridgeError(f'{source}: snapshot is stale or future-dated; export again. Existing Google events are left untouched.')
    exact_keys(data['window'], {'start', 'end'})
    lower, upper = utc(data['window']['start']), utc(data['window']['end'])
    if not lower <= generated < upper or upper - lower > timedelta(days=457):
        raise BridgeError('Invalid snapshot window.')
    if not isinstance(data['events'], list) or len(data['events']) > 10000:
        raise BridgeError('Invalid event list or excessive event count.')
    seen: set[str] = set()
    for event in data['events']:
        exact_keys(event, {'id', 'subject', 'all_day', 'start', 'end'})
        if not isinstance(event['id'], str) or not ID_RE.fullmatch(event['id']) or event['id'] in seen:
            raise BridgeError('Invalid or duplicate event identity.')
        seen.add(event['id'])
        if not isinstance(event['subject'], str) or not event['subject'].strip() or len(event['subject']) > 4096:
            raise BridgeError('Invalid event subject.')
        if type(event['all_day']) is not bool:
            raise BridgeError('all_day must be boolean.')
        if event['all_day']:
            try:
                if not all(isinstance(event[key], str) and re.fullmatch(r'\d{4}-\d{2}-\d{2}', event[key]) for key in ('start', 'end')):
                    raise ValueError
                start_date, end_date = date.fromisoformat(event['start']), date.fromisoformat(event['end'])
            except ValueError:
                raise BridgeError('Invalid all-day dates.') from None
            if end_date <= start_date or end_date < lower.date() or start_date > upper.date():
                raise BridgeError('All-day event is invalid or outside the export window.')
        else:
            start, end = utc(event['start']), utc(event['end'])
            if end <= start or start >= upper or end <= lower:
                raise BridgeError('Timed event is invalid or outside the export window.')
    return data


def google_id(source: str, event_id: str) -> str:
    # Hex is a subset of Google's base32hex event-ID alphabet.
    return hashlib.sha256((OWNER + '|' + source + '|' + event_id).encode()).hexdigest()


def desired_event(source: str, prefix: str, event: dict) -> dict:
    key = 'date' if event['all_day'] else 'dateTime'
    start = event['start'] if event['all_day'] else iso(utc(event['start']))
    end = event['end'] if event['all_day'] else iso(utc(event['end']))
    return {
        'summary': prefix + event['subject'],
        'start': {key: start}, 'end': {key: end},
        'visibility': 'private', 'transparency': 'opaque',
        'reminders': {'useDefault': False, 'overrides': []},
        'extendedProperties': {'private': {'bridge': OWNER, 'source': source, 'outlook_id': event['id']}},
    }


def owned(event: dict, source: str) -> bool:
    private = event.get('extendedProperties', {}).get('private', {})
    return private.get('bridge') == OWNER and private.get('source') == source


def same_event(existing: dict, desired: dict) -> bool:
    for key in ('summary', 'visibility', 'transparency'):
        if existing.get(key) != desired[key]:
            return False
    reminders = existing.get('reminders', {})
    if reminders.get('useDefault') is not False or reminders.get('overrides', []):
        return False
    for key in ('start', 'end'):
        old, new = existing.get(key, {}), desired[key]
        if 'date' in new:
            if old.get('date') != new['date']:
                return False
        elif 'dateTime' not in old or utc(old['dateTime']) != utc(new['dateTime']):
            return False
    return existing.get('extendedProperties', {}).get('private', {}) == desired['extendedProperties']['private']


def overlaps(event: dict, window: dict) -> bool:
    lower, upper = utc(window['start']), utc(window['end'])
    start, end = event.get('start', {}), event.get('end', {})
    if 'date' in start and 'date' in end:
        # Conservative date-only bounds avoid deletion at a timezone-ambiguous boundary.
        return date.fromisoformat(start['date']) < upper.date() and date.fromisoformat(end['date']) > lower.date()
    if 'dateTime' in start and 'dateTime' in end:
        return utc(start['dateTime']) < upper and utc(end['dateTime']) > lower
    raise BridgeError('A managed Google event has invalid scheduling data; refusing reconciliation.')


def plan(snapshot: dict, existing: list[dict], prefix: str = '') -> dict[str, list]:
    source = snapshot['source_id']
    managed = [event for event in existing if owned(event, source) and event.get('status') != 'cancelled']
    by_id: dict[str, dict] = {}
    duplicates: list[dict] = []
    for event in managed:
        identity = event['extendedProperties']['private'].get('outlook_id')
        if not isinstance(identity, str) or not ID_RE.fullmatch(identity):
            raise BridgeError('Invalid identity on a managed Google event; refusing reconciliation.')
        if identity in by_id:
            duplicates.append(event)
        else:
            by_id[identity] = event
    wanted = {event['id'] for event in snapshot['events']}
    result: dict[str, list] = {'create': [], 'update': [], 'delete': [], 'unchanged': []}
    for event in snapshot['events']:
        body = desired_event(source, prefix, event)
        old = by_id.get(event['id'])
        if old is None:
            result['create'].append({'id': google_id(source, event['id']), **body})
        elif not same_event(old, body):
            result['update'].append((old['id'], body))
        else:
            result['unchanged'].append(old['id'])
    result['delete'] = [event['id'] for event in managed if
                        (event['extendedProperties']['private']['outlook_id'] not in wanted or event in duplicates)
                        and overlaps(event, snapshot['window'])]
    return result


class Google:
    def __init__(self, config: dict, root: Path):
        self.root = root
        self.token_path = root / config.get('token_file', 'secrets/token.json')
        self.client_path = root / config.get('client_secret_file', 'secrets/client_secret.json')
        self.credentials: dict = {}

    def authorize(self) -> None:
        client = read_json(self.client_path).get('installed')
        if not client or not str(client.get('client_id', '')).endswith('.apps.googleusercontent.com'):
            raise BridgeError('Use a Google OAuth Desktop app client JSON file (not a web app or service account).')
        state, verifier = secrets.token_urlsafe(32), secrets.token_urlsafe(64)
        challenge = base64.urlsafe_b64encode(hashlib.sha256(verifier.encode()).digest()).decode().rstrip('=')
        outcome: dict = {}

        class Callback(http.server.BaseHTTPRequestHandler):
            def do_GET(self):
                from urllib.parse import urlsplit
                parsed = urlsplit(self.path)
                query = parse_qs(parsed.query)
                valid = parsed.path == '/callback' and secrets.compare_digest(query.get('state', [''])[0], state)
                if valid:
                    outcome.update(query)
                self.send_response(200 if valid else 400)
                self.send_header('Content-Type', 'text/plain; charset=utf-8')
                self.end_headers()
                self.wfile.write(b'Authorization received. Return to the terminal.' if valid else b'Invalid OAuth callback.')

            def log_message(self, *args):
                pass  # Never log callback authorization codes.

        # Port 0 lets the OS choose a free loopback port, avoiding other threads' servers.
        with http.server.HTTPServer(('127.0.0.1', 0), Callback) as server:
            server.timeout = 1
            redirect = f'http://127.0.0.1:{server.server_port}/callback'
            url = 'https://accounts.google.com/o/oauth2/v2/auth?' + urlencode({
                'client_id': client['client_id'], 'redirect_uri': redirect,
                'response_type': 'code', 'scope': SCOPE, 'access_type': 'offline',
                'prompt': 'consent', 'state': state, 'code_challenge': challenge,
                'code_challenge_method': 'S256',
            })
            print('Authorize your PERSONAL Google account in the browser. If it does not open, use this URL:\n' + url)
            webbrowser.open(url)
            deadline = time.monotonic() + 180
            while not outcome and time.monotonic() < deadline:
                server.handle_request()
        if 'code' not in outcome:
            raise BridgeError('Authorization denied or timed out; try authorize again.')
        credentials = request('POST', 'https://oauth2.googleapis.com/token', form=True, body={
            'client_id': client['client_id'], 'client_secret': client.get('client_secret', ''),
            'code': outcome['code'][0], 'code_verifier': verifier,
            'redirect_uri': redirect, 'grant_type': 'authorization_code',
        })
        if not credentials.get('refresh_token') or SCOPE not in credentials.get('scope', '').split():
            raise BridgeError('Google did not grant the required calendar scope or offline access.')
        credentials['expires_at'] = time.time() + credentials.get('expires_in', 3600)
        save_json(self.token_path, credentials)
        self.credentials = credentials
        print('Google authorization saved locally. Keep the token file private.')

    def access_token(self) -> str:
        if not self.credentials:
            self.credentials = read_json(self.token_path)
        if SCOPE not in self.credentials.get('scope', '').split():
            raise BridgeError('Required scope missing; run authorize again.')
        if self.credentials.get('expires_at', 0) <= time.time() + 60:
            client = read_json(self.client_path)['installed']
            refreshed = request('POST', 'https://oauth2.googleapis.com/token', form=True, body={
                'client_id': client['client_id'], 'client_secret': client.get('client_secret', ''),
                'refresh_token': self.credentials['refresh_token'], 'grant_type': 'refresh_token',
            })
            self.credentials.update(refreshed)
            self.credentials['expires_at'] = time.time() + refreshed.get('expires_in', 3600)
            save_json(self.token_path, self.credentials)
        return self.credentials['access_token']

    def call(self, method: str, path: str, body: Any = None, params: dict | None = None) -> Any:
        url = 'https://www.googleapis.com/calendar/v3/' + path
        if params:
            url += '?' + urlencode(params, doseq=True)
        return request(method, url, token=self.access_token(), body=body)

    def list_events(self, calendar: str, source: str) -> list[dict]:
        result: list[dict] = []
        params: dict[str, Any] = {'privateExtendedProperty': ['bridge=' + OWNER, 'source=' + source],
                                 'showDeleted': 'false', 'maxResults': 2500}
        # No date restriction here: moved appointments must be matched even at their old time.
        for _ in range(100):
            page = self.call('GET', f'calendars/{quote(calendar, safe="")}/events', params=params)
            result.extend(page.get('items', []))
            if not page.get('nextPageToken'):
                return result
            params['pageToken'] = page['nextPageToken']
        raise BridgeError('Too many managed Google events; no partial listing will be used.')


def load_config(path: Path) -> dict:
    config = read_json(path)
    if not isinstance(config, dict) or not isinstance(config.get('sources'), list) or not config['sources']:
        raise BridgeError('Configure at least one source.')
    repo = config.get('github_repository')
    if repo and (not isinstance(repo, str) or not re.fullmatch(r'[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+', repo)):
        raise BridgeError('Invalid GitHub repository name.')
    seen = set()
    for source in config['sources']:
        identity = source.get('id', '')
        if not isinstance(identity, str) or not SOURCE_RE.fullmatch(identity) or identity in seen:
            raise BridgeError('Source IDs must be unique opaque lowercase labels.')
        seen.add(identity)
        if not source.get('local_file') and not repo:
            raise BridgeError('Each source needs a local_file or a GitHub repository.')
        prefix = source.get('prefix', '')
        if not isinstance(prefix, str) or len(prefix) > 100:
            raise BridgeError('Invalid subject prefix.')
    age = config.get('max_snapshot_age_hours', 48)
    if type(age) not in (int, float) or not 1 <= age <= 168:
        raise BridgeError('max_snapshot_age_hours must be between 1 and 168.')
    return config


def fetch_snapshots(config: dict, root: Path) -> list[tuple[dict, dict]]:
    remote = any(not source.get('local_file') for source in config['sources'])
    token = os.environ.get('CALENDAR_GITHUB_TOKEN', '')
    repo = config.get('github_repository')
    if remote:
        if config.get('employer_approved') is not True:
            raise BridgeError('Set employer_approved to true only after every exporting employer permits this transfer.')
        if not token:
            raise BridgeError('Set CALENDAR_GITHUB_TOKEN to a fine-grained Contents: read token for the data repository.')
        metadata = request('GET', f'https://api.github.com/repos/{repo}', token=token, github=True)
        if metadata.get('private') is not True:
            raise BridgeError('Refusing to read corporate snapshots from a public repository.')
        ref = quote(config.get('github_branch', 'main'), safe='')
        branch = request('GET', f'https://api.github.com/repos/{repo}/branches/{ref}', token=token, github=True)
        commit = branch['commit']['sha']  # Pin all source reads to a single revision.
    result = []
    for source in config['sources']:
        if source.get('local_file'):
            path = root / source['local_file']
            if path.stat().st_size > MAX_BYTES:
                raise BridgeError('Snapshot exceeds safety limit.')
            data = read_json(path)
        else:
            path = f'snapshots/{source["id"]}.snapshot.json'
            payload = request('GET', f'https://api.github.com/repos/{repo}/contents/{path}?ref={commit}', token=token, github=True)
            if payload.get('type') != 'file' or payload.get('encoding') != 'base64' or payload.get('size', MAX_BYTES + 1) > MAX_BYTES:
                raise BridgeError('Invalid GitHub snapshot file.')
            raw = base64.b64decode(payload['content'])
            if len(raw) > MAX_BYTES:
                raise BridgeError('Snapshot exceeds safety limit.')
            data = json.loads(raw.decode('utf-8-sig'))
        result.append((source, validate_snapshot(data, source['id'], config.get('max_snapshot_age_hours', 48))))
    return result


def check_deletions(snapshot: dict, actions: dict, allow_empty: bool, allow_large: bool) -> None:
    count = len(actions['delete'])
    if count and not snapshot['events'] and not allow_empty:
        raise BridgeError('Empty snapshot would delete events. Inspect Outlook and rerun with --allow-empty only if intentional.')
    total = count + len(actions['update']) + len(actions['unchanged'])
    if count and (count > 25 or count / max(total, 1) > 0.5) and not allow_large:
        raise BridgeError('Large deletion blocked (>25 events or >50% of reconciled existing events). Inspect the dry run before using --allow-large-delete.')


def apply_plans(google: Google, calendar: str, plans: list[tuple[dict, dict, dict]]) -> None:
    path = f'calendars/{quote(calendar, safe="")}/events'
    # All upserts first. If any fails, no deletions happen on this run.
    for source, _, actions in plans:
        for body in actions['create']:
            try:
                google.call('POST', path, body, {'sendUpdates': 'none'})
            except ApiError as exc:
                if exc.status != 409:
                    raise
                existing = google.call('GET', path + '/' + body['id'])
                if existing.get('status') == 'cancelled':
                    # Deleted Google IDs cannot always be reused. Future runs find this copy by metadata.
                    google.call('POST', path, {**body, 'id': secrets.token_hex(32)}, {'sendUpdates': 'none'})
                else:
                    if not owned(existing, source['id']):
                        raise BridgeError('Google event ID collision; refusing to overwrite an unrelated event.') from None
                    google.call('PUT', path + '/' + body['id'], {key: value for key, value in body.items() if key != 'id'}, {'sendUpdates': 'none'})
        for identity, body in actions['update']:
            # Full replacement safely handles timed <-> all-day changes; never send attendees.
            google.call('PUT', path + '/' + quote(identity, safe=''), body, {'sendUpdates': 'none'})
    for _, _, actions in plans:
        for identity in actions['delete']:
            try:
                google.call('DELETE', path + '/' + quote(identity, safe=''), params={'sendUpdates': 'none'})
            except ApiError as exc:
                if exc.status not in (404, 410):
                    raise


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('command', choices=['authorize', 'init', 'sync'])
    parser.add_argument('--config', type=Path, default=Path('config.json'))
    parser.add_argument('--apply', action='store_true', help='Write to Google; sync defaults to read-only dry run.')
    parser.add_argument('--allow-empty', action='store_true')
    parser.add_argument('--allow-large-delete', action='store_true')
    args = parser.parse_args(argv)
    try:
        config_path = args.config.resolve()
        root = config_path.parent
        config = load_config(config_path)
        google = Google(config, root)
        with lock(root / 'secrets/.bridge.lock'):
            if args.command == 'authorize':
                google.authorize()
                return 0
            if args.command == 'init':
                if config.get('calendar_id'):
                    raise BridgeError('A target calendar is already configured; refusing to create another.')
                calendar = google.call('POST', 'calendars', {
                    'summary': config.get('calendar_name', 'Work overview'),
                    'timeZone': config.get('calendar_timezone', 'Etc/UTC'),
                })
                config['calendar_id'] = calendar['id']
                save_json(config_path, config)
                print('Created a dedicated Google Calendar and saved its ID in your configuration.')
                return 0
            calendar = config.get('calendar_id')
            if not calendar or calendar == 'primary':
                raise BridgeError('Run init to create the dedicated target calendar. The primary calendar is never supported.')
            # Validate ALL inputs before Google requests or any calendar mutation.
            snapshots = fetch_snapshots(config, root)
            state_path = root / 'secrets/sync_state.json'
            state = read_json(state_path) if state_path.exists() else {}
            last = state.get(calendar, {})
            for source, snapshot in snapshots:
                if source['id'] in last and utc(snapshot['generated_at']) < utc(last[source['id']]):
                    raise BridgeError('An older snapshot would roll back the calendar; export again.')
            plans = []
            for source, snapshot in snapshots:
                actions = plan(snapshot, google.list_events(calendar, source['id']), source.get('prefix', ''))
                print(source['id'] + ': ' + ', '.join(f'{key}={len(value)}' for key, value in actions.items()))
                plans.append((source, snapshot, actions))
            for _, snapshot, actions in plans:
                check_deletions(snapshot, actions, args.allow_empty, args.allow_large_delete)
            if not args.apply:
                print('DRY RUN: no calendar changes. Rerun with --apply after reviewing counts.')
                return 0
            apply_plans(google, calendar, plans)
            state[calendar] = {**last, **{source['id']: snapshot['generated_at'] for source, snapshot in snapshots}}
            save_json(state_path, state)
            print('Sync complete. Only bridge-managed events were changed; no invitations or reminders were sent.')
            return 0
    except (BridgeError, OSError, ValueError, KeyError, TypeError):
        # Specific bridge messages are safe; other failures may contain file contents or credentials.
        import sys
        exc = sys.exc_info()[1]
        message = str(exc) if isinstance(exc, BridgeError) else 'Invalid/missing configuration, credentials, data, or local file access. Check setup instructions.'
        print('Error: ' + message, file=sys.stderr)
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
