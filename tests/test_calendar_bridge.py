from datetime import datetime, timedelta, timezone
import base64
import contextlib
import io
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import calendar_bridge as bridge

NOW = datetime(2026, 10, 2, 12, tzinfo=timezone.utc)


def event(identity='a', all_day=False):
    return {
        'id': identity * 64, 'subject': 'Planning', 'all_day': all_day,
        'start': '2026-10-03' if all_day else '2026-10-03T09:00:00Z',
        'end': '2026-10-04' if all_day else '2026-10-03T10:00:00Z',
    }


def snapshot(events=None):
    return {
        'schema_version': 1, 'source_id': 'client-a', 'generated_at': bridge.iso(NOW), 'complete': True,
        'window': {'start': '2026-09-25T00:00:00Z', 'end': '2026-12-31T00:00:00Z'},
        'events': [event()] if events is None else events,
    }


def existing(value=None, identity='google-1', source='client-a'):
    return {'id': identity, 'status': 'confirmed', **bridge.desired_event(source, '[A] ', value or event())}


class ValidationTests(unittest.TestCase):
    def validate(self, value):
        return bridge.validate_snapshot(value, 'client-a', now=NOW)

    def test_accepts_timed_and_all_day(self):
        self.validate(snapshot([event(), event('b', True)]))

    def test_accepts_utc_window_with_local_all_day_date_at_edge(self):
        value = snapshot([event(all_day=True)])
        value['window']['start'] = '2026-10-01T23:00:00Z'
        self.validate(value)

    def test_rejects_private_fields_at_every_level(self):
        for location in ('root', 'event', 'window'):
            value = snapshot()
            target = value if location == 'root' else value['events'][0] if location == 'event' else value['window']
            target['body'] = 'Should never leave Outlook'
            with self.subTest(location=location), self.assertRaises(bridge.BridgeError):
                self.validate(value)

    def test_rejects_stale_future_incomplete_or_wrong_source(self):
        for key, value in [('generated_at', bridge.iso(NOW - timedelta(days=3))),
                           ('generated_at', bridge.iso(NOW + timedelta(hours=1))),
                           ('complete', False), ('source_id', 'client-b'), ('schema_version', True)]:
            data = snapshot()
            data[key] = value
            with self.subTest(key=key, value=value), self.assertRaises(bridge.BridgeError):
                self.validate(data)

    def test_rejects_duplicate_invalid_ids_and_non_boolean_all_day(self):
        for events in ([event(), event()], [{**event(), 'id': 'abc'}], [{**event(), 'all_day': 1}]):
            with self.assertRaises(bridge.BridgeError):
                self.validate(snapshot(events))

    def test_rejects_bad_dates_and_out_of_range_timed_events(self):
        for changes in ({'start': '2026-10-03T09:00:00'}, {'end': '2026-10-03T08:00:00Z'},
                        {'start': '2027-01-01T09:00:00Z', 'end': '2027-01-01T10:00:00Z'},
                        {'all_day': True, 'start': '2026-02-30', 'end': '2026-03-01'}):
            with self.subTest(changes=changes), self.assertRaises(bridge.BridgeError):
                self.validate(snapshot([{**event(), **changes}]))


class PlanningTests(unittest.TestCase):
    def test_repeat_sync_is_idempotent(self):
        actions = bridge.plan(snapshot(), [existing()], '[A] ')
        self.assertEqual(actions['unchanged'], ['google-1'])
        self.assertFalse(actions['create'] or actions['update'] or actions['delete'])

    def test_google_omitted_empty_overrides_does_not_cause_rewrite(self):
        value = existing()
        value['reminders'] = {'useDefault': False}
        self.assertEqual(len(bridge.plan(snapshot(), [value], '[A] ')['unchanged']), 1)

    def test_offset_timestamp_is_equivalent(self):
        value = existing()
        value['start']['dateTime'] = '2026-10-03T11:00:00+02:00'
        self.assertEqual(len(bridge.plan(snapshot(), [value], '[A] ')['unchanged']), 1)

    def test_subject_change_updates(self):
        value = existing()
        value['summary'] = 'Old subject'
        self.assertEqual(len(bridge.plan(snapshot(), [value], '[A] ')['update']), 1)

    def test_time_change_removes_old_occurrence_and_creates_new(self):
        changed = {**event('b'), 'start': '2026-10-03T11:00:00Z', 'end': '2026-10-03T12:00:00Z'}
        actions = bridge.plan(snapshot([changed]), [existing()], '[A] ')
        self.assertEqual(len(actions['create']), 1)
        self.assertEqual(actions['delete'], ['google-1'])

    def test_update_matches_managed_identity_outside_window(self):
        value = existing()
        value['start'] = {'dateTime': '2027-01-01T09:00:00Z'}
        value['end'] = {'dateTime': '2027-01-01T10:00:00Z'}
        actions = bridge.plan(snapshot(), [value], '[A] ')
        self.assertEqual(len(actions['update']), 1)
        self.assertEqual(actions['create'], [])

    def test_never_deletes_personal_other_source_or_outside_window(self):
        personal = {'id': 'personal'}
        other = existing(source='client-b')
        old = existing(identity='past')
        old['start'] = {'dateTime': '2026-01-01T09:00:00Z'}
        old['end'] = {'dateTime': '2026-01-01T10:00:00Z'}
        self.assertEqual(bridge.plan(snapshot([]), [personal, other, old])['delete'], [])

    def test_duplicate_managed_events_are_removed(self):
        actions = bridge.plan(snapshot(), [existing(), existing(identity='duplicate')], '[A] ')
        self.assertEqual(actions['delete'], ['duplicate'])

    def test_all_day_end_date_is_exclusive_and_contains_no_private_content(self):
        body = bridge.desired_event('client-a', '[A] ', event(all_day=True))
        self.assertEqual(body['start'], {'date': '2026-10-03'})
        self.assertEqual(body['end'], {'date': '2026-10-04'})
        for key in ('description', 'location', 'attendees', 'organizer', 'recurrence'):
            self.assertNotIn(key, body)
        self.assertFalse(body['reminders']['useDefault'])
        self.assertEqual(body['visibility'], 'private')

    def test_source_namespaces_do_not_collide(self):
        self.assertNotEqual(bridge.google_id('client-a', 'a' * 64), bridge.google_id('client-b', 'a' * 64))


class SafetyTests(unittest.TestCase):
    def test_empty_and_mass_deletion_require_explicit_flags(self):
        actions = {'delete': ['x'], 'update': [], 'unchanged': []}
        with self.assertRaises(bridge.BridgeError):
            bridge.check_deletions(snapshot([]), actions, False, True)
        with self.assertRaises(bridge.BridgeError):
            bridge.check_deletions(snapshot([]), actions, True, False)
        bridge.check_deletions(snapshot([]), actions, True, True)

    def test_more_than_25_deletions_blocked_even_if_low_percentage(self):
        actions = {'delete': list(range(26)), 'update': [], 'unchanged': list(range(200))}
        with self.assertRaises(bridge.BridgeError):
            bridge.check_deletions(snapshot(), actions, False, False)

    def test_failed_upsert_prevents_all_deletions(self):
        class FailingGoogle:
            calls = []

            def call(self, method, *args, **kwargs):
                self.calls.append(method)
                raise bridge.ApiError(503)

        google = FailingGoogle()
        actions = {'create': [{'id': 'new'}], 'update': [], 'delete': ['old']}
        with self.assertRaises(bridge.ApiError):
            bridge.apply_plans(google, 'dedicated', [({'id': 'client-a'}, snapshot(), actions)])
        self.assertEqual(google.calls, ['POST'])

    def test_conflicting_unrelated_google_event_is_not_overwritten(self):
        class ConflictGoogle:
            calls = []

            def call(self, method, *args, **kwargs):
                self.calls.append(method)
                if method == 'POST':
                    raise bridge.ApiError(409)
                return {'id': 'unrelated', 'status': 'confirmed'}

        google = ConflictGoogle()
        actions = {'create': [{'id': 'new'}], 'update': [], 'delete': []}
        with self.assertRaises(bridge.BridgeError):
            bridge.apply_plans(google, 'dedicated', [({'id': 'client-a'}, snapshot(), actions)])
        self.assertEqual(google.calls, ['POST', 'GET'])

    def test_cancelled_id_can_be_recreated_without_touching_tombstone(self):
        class TombstoneGoogle:
            calls = []

            def call(self, method, path, body=None, params=None):
                self.calls.append((method, body))
                if method == 'POST' and len(self.calls) == 1:
                    raise bridge.ApiError(409)
                if method == 'GET':
                    return {'id': 'old', 'status': 'cancelled'}
                return body

        google = TombstoneGoogle()
        body = {'id': 'old', **bridge.desired_event('client-a', '', event())}
        bridge.apply_plans(google, 'dedicated', [({'id': 'client-a'}, snapshot(), {'create': [body], 'update': [], 'delete': []})])
        self.assertEqual([method for method, _ in google.calls], ['POST', 'GET', 'POST'])
        self.assertNotEqual(google.calls[-1][1]['id'], 'old')

    def test_timed_to_all_day_uses_full_replacement(self):
        class RecordingGoogle:
            calls = []

            def call(self, method, *args, **kwargs):
                self.calls.append(method)

        google = RecordingGoogle()
        actions = {'create': [], 'update': [('old', bridge.desired_event('client-a', '', event(all_day=True)))], 'delete': []}
        bridge.apply_plans(google, 'dedicated', [({'id': 'client-a'}, snapshot(), actions)])
        self.assertEqual(google.calls, ['PUT'])

    def test_atomic_json_and_exclusive_lock(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'secrets/token.json'
            bridge.save_json(path, {'test': 1})
            bridge.save_json(path, {'test': 2})
            self.assertEqual(bridge.read_json(path), {'test': 2})
            lockpath = path.parent / '.bridge.lock'
            with bridge.lock(lockpath):
                with self.assertRaises(bridge.BridgeError):
                    with bridge.lock(lockpath):
                        pass
            self.assertFalse(lockpath.exists())

    def test_config_rejects_duplicate_sources(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'config.json'
            path.write_text(json.dumps({'sources': [{'id': 'client-a', 'local_file': 'a'}, {'id': 'client-a', 'local_file': 'b'}]}))
            with self.assertRaises(bridge.BridgeError):
                bridge.load_config(path)

    def test_fetch_validates_all_sources_before_google_write(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / 'a.json').write_text(json.dumps(snapshot()))
            invalid = snapshot()
            invalid['source_id'] = 'client-b'
            invalid['complete'] = False
            (root / 'b.json').write_text(json.dumps(invalid))
            config = {'sources': [{'id': 'client-a', 'local_file': 'a.json'}, {'id': 'client-b', 'local_file': 'b.json'}]}
            validate = bridge.validate_snapshot
            with patch.object(bridge, 'validate_snapshot', side_effect=lambda data, source, age: validate(data, source, age, NOW)):
                with self.assertRaises(bridge.BridgeError):
                    bridge.fetch_snapshots(config, root)


class ImporterTests(unittest.TestCase):
    def test_github_public_repo_rejected_before_download(self):
        config = {'github_repository': 'example/data', 'employer_approved': True, 'sources': [{'id': 'client-a'}]}
        with patch.dict(bridge.os.environ, {'CALENDAR_GITHUB_TOKEN': 'fake'}), patch.object(bridge, 'request', return_value={'private': False}) as api:
            with self.assertRaises(bridge.BridgeError):
                bridge.fetch_snapshots(config, Path('.'))
            self.assertEqual(api.call_count, 1)

    def test_github_sources_pinned_to_one_commit(self):
        config = {'github_repository': 'example/data', 'employer_approved': True,
                  'sources': [{'id': 'client-a'}, {'id': 'client-b'}]}
        other = snapshot()
        other['source_id'] = 'client-b'
        def payload(data):
            raw = json.dumps(data).encode()
            return {'type': 'file', 'encoding': 'base64', 'size': len(raw), 'content': base64.b64encode(raw).decode()}
        responses = [{'private': True}, {'commit': {'sha': 'fixed-commit'}}, payload(snapshot()), payload(other)]
        validate = bridge.validate_snapshot
        with patch.dict(bridge.os.environ, {'CALENDAR_GITHUB_TOKEN': 'fake'}), patch.object(bridge, 'request', side_effect=responses) as api, patch.object(bridge, 'validate_snapshot', side_effect=lambda data, source, age: validate(data, source, age, NOW)):
            self.assertEqual(len(bridge.fetch_snapshots(config, Path('.'))), 2)
            for call in api.call_args_list[2:]:
                self.assertIn('?ref=fixed-commit', call.args[1])
                self.assertEqual(call.args[0], 'GET')

    def test_sync_defaults_to_dry_run(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            config = {'calendar_id': 'dedicated', 'sources': [{'id': 'client-a', 'local_file': 'a.json'}]}
            bridge.save_json(root / 'config.json', config)
            with patch.object(bridge, 'fetch_snapshots', return_value=[(config['sources'][0], snapshot())]), patch.object(bridge.Google, 'list_events', return_value=[]), patch.object(bridge, 'apply_plans') as apply, contextlib.redirect_stdout(io.StringIO()) as output:
                self.assertEqual(bridge.main(['sync', '--config', str(root / 'config.json')]), 0)
                apply.assert_not_called()
                self.assertIn('DRY RUN', output.getvalue())
                self.assertNotIn('Planning', output.getvalue())
                self.assertFalse((root / 'secrets/sync_state.json').exists())

    def test_older_snapshot_cannot_roll_back_applied_calendar(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            config = {'calendar_id': 'dedicated', 'sources': [{'id': 'client-a', 'local_file': 'a.json'}]}
            bridge.save_json(root / 'config.json', config)
            bridge.save_json(root / 'secrets/sync_state.json', {'dedicated': {'client-a': bridge.iso(NOW + timedelta(minutes=1))}})
            with patch.object(bridge, 'fetch_snapshots', return_value=[(config['sources'][0], snapshot())]), patch.object(bridge.Google, 'list_events') as listing, patch.object(bridge, 'apply_plans') as apply, contextlib.redirect_stderr(io.StringIO()):
                self.assertEqual(bridge.main(['sync', '--apply', '--config', str(root / 'config.json')]), 1)
                listing.assert_not_called()
                apply.assert_not_called()


if __name__ == '__main__':
    unittest.main()
