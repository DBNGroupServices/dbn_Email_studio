"""Version 2 features: voice profiles, quick intents, variants, metadata,
security disclosure guard, certificate authority renewal and new endpoints."""
from datetime import timedelta
import json
from pathlib import Path
import sys
import tempfile
import unittest
import urllib.request
from unittest.mock import patch

from cryptography.hazmat.primitives.asymmetric import padding

import app
import office_tls
import profiles

ROOT = Path(__file__).resolve().parents[1]
CAN_STORE = sys.platform == 'win32' or __import__('os').environ.get('DBN_EMAIL_TEST_PLAINTEXT_STORE') == '1'


def context_of(request):
    return json.loads(request['input'][0]['content'][0]['text'])


class PromptTests(unittest.TestCase):
    def test_profile_intent_sliders_and_meta_reach_the_prompt(self):
        profile = profiles.validate({'name': 'Tenant', 'guidance': 'Reassuring', 'greeting': 'Hi {first name},', 'signoff': 'Cheers,',
                                     'formality': 2, 'warmth': 5, 'directness': 3, 'examples': ['Hi Jo, all sorted. Cheers, D'],
                                     'signature': 'Fictional Name\nFictional Org'})
        request, _ = app.build_request({'mode': 'reply', 'source': 'Fictional email', 'intent': 'acknowledge', '_profile': profile,
                                        'sliders': {'formality': 1}, 'tone': 'Use profile',
                                        'meta': {'from': {'name': 'Jo Citizen', 'email': 'jo@example.com'}, 'to': [{'name': 'A', 'email': 'a@example.com'}], 'reply_all': True}},
                                       app.DEFAULTS, app.MODELS)
        ctx = context_of(request)
        self.assertEqual(ctx['voice_profile']['example_emails'], ['Hi Jo, all sorted. Cheers, D'])
        self.assertEqual(ctx['tone_controls'], {'formality': 1, 'warmth': 5, 'directness': 3})
        self.assertIn('Acknowledge', ctx['quick_intent'])
        self.assertEqual(ctx['email_metadata']['from']['name'], 'Jo Citizen')
        self.assertTrue(ctx['email_metadata']['reply_all'])
        self.assertEqual(ctx['sender']['signature'], 'Fictional Name\nFictional Org')
        self.assertIn('Security disclosure guard', request['instructions'])
        self.assertEqual(request['text']['format']['schema'], app.SCHEMA)

    def test_signature_can_be_left_to_outlook(self):
        profile = profiles.validate({'name': 'X', 'signature': 'Sig', 'include_signature': False})
        request, _ = app.build_request({'mode': 'compose', 'points': 'Fictional', '_profile': profile}, app.DEFAULTS | {'signature': 'Saved sig'}, app.MODELS)
        self.assertEqual(context_of(request)['sender']['signature'], '')
        self.assertIn('signature_note', context_of(request)['sender'])

    def test_variants_use_variant_schema_and_are_bounded(self):
        request, _ = app.build_request({'mode': 'reply', 'source': 'Fictional', 'variants': 3}, app.DEFAULTS, app.MODELS)
        self.assertEqual(request['text']['format']['schema'], app.VARIANT_SCHEMA)
        self.assertEqual(context_of(request)['number_of_variants'], 3)
        for bad in (0, 4, 'many'):
            with self.assertRaises(app.UserError):
                app.build_request({'mode': 'reply', 'source': 'Fictional', 'variants': bad}, app.DEFAULTS, app.MODELS)

    def test_invalid_intent_tone_and_meta_rejected(self):
        with self.assertRaises(app.UserError):
            app.build_request({'mode': 'reply', 'source': 'x', 'intent': 'send-now'}, app.DEFAULTS, app.MODELS)
        with self.assertRaises(app.UserError):
            app.build_request({'mode': 'reply', 'source': 'x', 'tone': 'Sarcastic'}, app.DEFAULTS, app.MODELS)
        with self.assertRaises(app.UserError):
            app.build_request({'mode': 'reply', 'source': 'x', 'meta': {'to': 'not-a-list'}}, app.DEFAULTS, app.MODELS)

    def test_security_guard_can_be_turned_off(self):
        request, _ = app.build_request({'mode': 'reply', 'source': 'x'}, app.DEFAULTS | {'security_guard': False}, app.MODELS)
        self.assertNotIn('Security disclosure guard', request['instructions'])
        assist, _ = app.build_assist_request({'action': 'summary', 'source': 'x'}, app.DEFAULTS, app.MODELS)
        self.assertIn('Security disclosure guard', assist['instructions'])

    def test_variant_response_is_normalised(self):
        payload = {'variants': [{'label': 'Direct', 'subject': 'S1', 'body': 'B1'}, {'label': 'Warm', 'subject': 'S2', 'body': 'B2'}], 'checks': ['c']}
        response = {'status': 'completed', 'output': [{'type': 'message', 'content': [{'type': 'output_text', 'text': json.dumps(payload)}]}]}
        result = app.read_response(response)
        self.assertEqual((result['subject'], result['body']), ('S1', 'B1'))
        self.assertEqual([a['label'] for a in result['alternatives']], ['Direct', 'Warm'])
        bad = {'status': 'completed', 'output': [{'type': 'message', 'content': [{'type': 'output_text', 'text': json.dumps({'variants': [], 'checks': []})}]}]}
        with self.assertRaises(app.UserError):
            app.read_response(bad)


class ProfileTests(unittest.TestCase):
    def test_validation(self):
        for bad in ({}, {'name': ''}, {'name': 'x', 'formality': 9}, {'name': 'x', 'formality': True},
                    {'name': 'x', 'examples': ['a'] * 4}, {'name': 'x', 'length': 'Epic'}, {'name': 'x', 'id': '../etc'},
                    {'name': 'x' * 81}):
            with self.assertRaises(profiles.ProfileError, msg=bad):
                profiles.validate(bad)
        p = profiles.validate({'name': ' Mine ', 'examples': [' ', 'Hello']})
        self.assertEqual(p['name'], 'Mine')
        self.assertEqual(p['examples'], ['Hello'])
        self.assertTrue(p['id'])

    @unittest.skipUnless(CAN_STORE, 'needs Windows DPAPI or the test store')
    def test_store_round_trip_and_defaults(self):
        with tempfile.TemporaryDirectory() as directory:
            store = app.Store(Path(directory))
            defaults = profiles.load(store)
            self.assertGreaterEqual(len(defaults), 4)
            saved = profiles.save(store, {'name': 'New one', 'warmth': 5})
            self.assertEqual(len(saved), len(defaults) + 1)
            new_id = saved[-1]['id']
            profiles.save(store, {'id': new_id, 'name': 'Renamed'})
            self.assertEqual(profiles.find(store, new_id)['name'], 'Renamed')
            self.assertNotIn(b'Renamed', (Path(directory) / 'profiles.dat').read_bytes())
            left = profiles.delete(store, new_id)
            self.assertIsNone(profiles.find(store, new_id))
            self.assertEqual(len(left), len(defaults))


@unittest.skipUnless(CAN_STORE, 'needs Windows DPAPI or the test store')
class CertificateTests(unittest.TestCase):
    def test_ca_is_name_constrained_and_leaf_renews_without_new_ca(self):
        with tempfile.TemporaryDirectory() as directory:
            store = app.Store(Path(directory))
            saved, public = office_tls.prepare(store)
            self.assertTrue(public.exists())
            ca = office_tls._cert(saved['ca_cert'])
            constraints = ca.extensions.get_extension_for_class(office_tls.x509.NameConstraints)
            self.assertTrue(constraints.critical)
            leaf = office_tls._cert(saved['leaf_cert'])
            self.assertLessEqual((office_tls._expiry(leaf) - office_tls._now()).days, office_tls.LEAF_DAYS)
            # Same day: nothing changes.
            again, _ = office_tls.prepare(store)
            self.assertEqual(again['leaf_cert'], saved['leaf_cert'])
            # Near expiry: the leaf renews, the trusted CA stays the same.
            later = office_tls._now() + timedelta(days=office_tls.LEAF_DAYS - 10)
            renewed, _ = office_tls.prepare(store, now=later)
            self.assertNotEqual(renewed['leaf_cert'], saved['leaf_cert'])
            self.assertEqual(renewed['ca_cert'], saved['ca_cert'])
            # The renewed leaf chains to the CA.
            new_leaf = office_tls._cert(renewed['leaf_cert'])
            self.assertEqual(new_leaf.issuer, ca.subject)
            ca.public_key().verify(new_leaf.signature, new_leaf.tbs_certificate_bytes, padding.PKCS1v15(), new_leaf.signature_hash_algorithm)


@unittest.skipUnless(CAN_STORE, 'needs Windows DPAPI or the test store')
class EndpointTests(unittest.TestCase):
    def setUp(self):
        (ROOT / '.test-data').mkdir(exist_ok=True)
        self.temp = tempfile.TemporaryDirectory(dir=ROOT / '.test-data')
        self.server = app.AppServer(('127.0.0.1', 0), app.Store(Path(self.temp.name)))
        __import__('threading').Thread(target=self.server.serve_forever, daemon=True).start()
        self.base = f'http://127.0.0.1:{self.server.server_port}'

    def tearDown(self):
        self.server.shutdown(); self.server.server_close(); self.temp.cleanup()

    def call(self, path, data=None):
        request = urllib.request.Request(self.base + '/api/' + path, data=None if data is None else json.dumps(data).encode(),
                                         headers={'X-App-Token': self.server.token, 'Content-Type': 'application/json'},
                                         method='GET' if data is None else 'POST')
        try:
            with urllib.request.urlopen(request, timeout=5) as response:
                return response.status, json.load(response)
        except urllib.error.HTTPError as error:
            return error.code, json.load(error)

    def test_profiles_templates_autostart_and_meta(self):
        status, items = self.call('profiles')
        self.assertEqual(status, 200)
        status, items = self.call('profiles', {'name': 'Fictional voice', 'formality': 5})
        self.assertEqual(status, 200)
        created = items[-1]['id']
        status, error = self.call('profiles', {'name': ''})
        self.assertEqual(status, 400)
        self.assertIn('name', error['error'])
        status, items = self.call('profiles/delete', {'id': created})
        self.assertNotIn(created, [p['id'] for p in items])
        status, templates = self.call('templates', {'name': 'T', 'body': 'Fictional body'})
        status, templates = self.call('templates/delete', {'id': templates[0]['id']})
        self.assertEqual(templates, [])
        status, auto = self.call('autostart')
        self.assertEqual(status, 200)
        self.assertIn('enabled', auto)
        if not auto['supported']:
            status, _ = self.call('autostart', {'enabled': True})
            self.assertEqual(status, 400)
        status, meta = self.call('meta')
        self.assertIn('acknowledge', meta['intents'])
        status, saved = self.call('settings', {'security_guard': False})
        self.assertEqual(status, 200)
        self.assertFalse(saved['settings']['security_guard'])
        status, _ = self.call('settings', {'security_guard': 'no'})
        self.assertEqual(status, 400)

    def test_generate_uses_saved_profile(self):
        status, items = self.call('profiles', {'name': 'Fictional voice', 'guidance': 'UNIQUE-GUIDANCE'})
        ident = items[-1]['id']
        captured = {}
        def fake(key, path, payload=None, plan=False):
            captured['payload'] = payload
            text = json.dumps({'subject': 'S', 'body': 'B', 'checks': []})
            return {'status': 'completed', 'output': [{'type': 'message', 'content': [{'type': 'output_text', 'text': text}]}]}
        with patch.object(app, 'openai_request', fake), patch.object(self.server.store, 'key', lambda: 'sk-test'), \
             patch.object(app.Store, 'settings', lambda s: app.DEFAULTS | {'provider': 'api'}):
            status, result = self.call('generate', {'mode': 'reply', 'source': 'Fictional', 'profile_id': ident})
        self.assertEqual(status, 200, result)
        self.assertIn('UNIQUE-GUIDANCE', captured['payload']['input'][0]['content'][0]['text'])

    def test_static_draft_tools_served(self):
        with urllib.request.urlopen(self.base + '/office/draft-tools.js', timeout=5) as response:
            self.assertIn(b'toHtml', response.read())


if __name__ == '__main__':
    unittest.main()
