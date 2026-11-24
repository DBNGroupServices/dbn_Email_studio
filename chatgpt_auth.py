"""Official local-client Sign in with ChatGPT; credentials stay in DPAPI storage."""
import base64
from contextlib import contextmanager
import hashlib
import json
import re
import secrets
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid

import jwt

ISSUER = 'https://auth.openai.com'
RESOURCE = 'https://api.openai.com/v1'
SCOPES = 'openid profile email offline_access resource.invoke chatgpt.tokens.use.direct'
USAGE_URL = 'https://chatgpt.com/#settings/Usage'


class AuthError(Exception):
    pass


@contextmanager
def session_lock(store):
    """Serialize rotating-token refresh/sign-out across app processes on Windows."""
    if not hasattr(store, 'directory'):
        yield
        return
    import msvcrt
    store.directory.mkdir(parents=True, exist_ok=True)
    with (store.directory / 'chatgpt-session.lock').open('a+b') as handle:
        if handle.seek(0, 2) == 0:
            handle.write(b'0');handle.flush()
        for attempt in range(300):
            try:
                handle.seek(0);msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
                break
            except OSError:
                time.sleep(.1)
        else:raise AuthError('Another Email Studio connection update is still running. Try again shortly.')
        try:yield
        finally:
            handle.seek(0);msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)


def request_json(url, form=None, bearer=None):
    headers = {'User-Agent': 'DBN-Email-Studio/ChatGPT'}
    if bearer:
        headers['Authorization'] = 'Bearer ' + bearer
    if form is not None:
        headers['Content-Type'] = 'application/x-www-form-urlencoded'
    request = urllib.request.Request(url, data=urllib.parse.urlencode(form).encode() if form is not None else None, headers=headers)
    try:
        with urllib.request.urlopen(request, timeout=25) as response:
            body = response.read(2 * 1024 * 1024)
            return json.loads(body) if body else {}
    except urllib.error.HTTPError as error:
        try:
            code = json.loads(error.read(10000)).get('error', '')
        except (ValueError, AttributeError):
            code = ''
        if code in ('invalid_grant', 'invalid_token'):
            raise AuthError('Your ChatGPT connection needs a fresh sign-in. Choose Continue with ChatGPT.') from None
        raise AuthError(f'OpenAI could not complete this connection step (HTTP {error.code}). Try again. Your existing connection has been kept.') from None
    except (OSError, ValueError):
        raise AuthError('Could not reach OpenAI securely. Check your connection and try again. Your saved connection has been kept.') from None


def public_state(store):
    saved = store.read('chatgpt', {})
    profile = saved.get('profiles', {}).get(saved.get('active', ''), {})
    enabled = bool(profile.get('access_token') and 'chatgpt.tokens.use.direct' in profile.get('scopes', []))
    return {'signed_in': bool(profile.get('access_token')), 'enabled': enabled,
            'email': profile.get('email', ''), 'name': profile.get('name', ''),
            'models': profile.get('models', {}) if enabled else {},
            'accounts': [{'id': key, 'email': value.get('email', ''), 'name': value.get('name', '')}
                         for key, value in saved.get('profiles', {}).items()],
            'active': saved.get('active', ''), 'welcomed': saved.get('welcomed', False)}


class ChatGPTAuth:
    def __init__(self, store, redirect_uri):
        self.store = store
        self.redirect_uri = redirect_uri
        self.lock = threading.RLock()
        self.pending = None
        self.message = ''
        self.discovery_cache = None
        self.jwks_cache = None
        self.jwks_time = 0

    def discovery(self):
        if self.discovery_cache is None:
            data = request_json(ISSUER + '/.well-known/openid-configuration')
            if data.get('issuer') != ISSUER:
                raise AuthError('OpenAI sign-in configuration could not be verified.')
            for field in ('authorization_endpoint', 'token_endpoint', 'jwks_uri', 'revocation_endpoint'):
                parsed = urllib.parse.urlparse(data.get(field, ''))
                if parsed.scheme != 'https' or parsed.netloc != 'auth.openai.com':
                    raise AuthError('OpenAI sign-in endpoint could not be verified.')
            self.discovery_cache = data
        return self.discovery_cache

    def status(self):
        with self.lock:
            if self.pending and self.pending['expires'] < time.time():
                self.pending = None
                self.message = 'Sign-in expired. Choose Continue with ChatGPT to try again.'
            return {**public_state(self.store), 'pending': bool(self.pending), 'message': self.message}

    def start(self, account='', new=False):
        with self.lock:
            saved = self.store.read('chatgpt', {})
            if not saved.get('host_id'):
                saved['host_id'] = 'urn:uuid:' + str(uuid.uuid4())
                self.store.write('chatgpt', saved)
            profiles = saved.get('profiles', {})
            selected = '' if new else account or saved.get('active', '')
            if selected and selected not in profiles:
                raise AuthError('Choose a saved ChatGPT account or add another account.')
            profile = profiles.get(selected, {})
            verifier = secrets.token_urlsafe(64)
            pending = {'state': secrets.token_urlsafe(32), 'nonce': secrets.token_urlsafe(32),
                       'verifier': verifier, 'expires': time.time() + 600, 'client_id': selected,
                       'subject': profile.get('subject'), 'redirect_uri': self.redirect_uri}
            params = {'client_id': selected or 'dynamic_agent_client', 'ext_agent_host_id': saved['host_id'],
                      'response_type': 'code', 'redirect_uri': self.redirect_uri, 'scope': SCOPES,
                      'resource': RESOURCE, 'state': pending['state'], 'nonce': pending['nonce'],
                      'code_challenge_method': 'S256',
                      'code_challenge': base64.urlsafe_b64encode(hashlib.sha256(verifier.encode()).digest()).rstrip(b'=').decode()}
            if not selected:
                params['agent_name_hint'] = 'DBN Email Studio'
            # No retained ID token is put in a URL; the normal account selector is used.
            url = self.discovery()['authorization_endpoint'] + '?' + urllib.parse.urlencode(params)
            self.pending = pending
            self.message = 'Complete sign-in and enable use of your ChatGPT plan in your browser.'
            return url

    def validate_identity(self, token, client_id, nonce=None, subject=None):
        try:
            header = jwt.get_unverified_header(token)
            if header.get('alg') != 'RS256':
                raise ValueError()
            if not self.jwks_cache or time.time() - self.jwks_time > 3600 or not any(k.get('kid') == header.get('kid') for k in self.jwks_cache.get('keys', [])):
                self.jwks_cache = request_json(self.discovery()['jwks_uri'])
                self.jwks_time = time.time()
            key = next(k for k in self.jwks_cache.get('keys', []) if k.get('kid') == header.get('kid') and k.get('kty') == 'RSA')
            claims = jwt.decode(token, jwt.PyJWK.from_dict(key).key, algorithms=['RS256'],
                                issuer=ISSUER, audience=client_id, leeway=5,
                                options={'require': ['sub', 'exp', 'iat']})
            if not isinstance(claims['sub'], str) or not claims['sub'] or (nonce is not None and not secrets.compare_digest(str(claims.get('nonce', '')), nonce)) or (subject and claims['sub'] != subject):
                raise ValueError()
            return claims
        except (jwt.PyJWTError, ValueError, TypeError, KeyError, StopIteration):
            raise AuthError('The ChatGPT sign-in identity could not be verified. Your existing account has not been changed.') from None

    def callback(self, query):
        values = urllib.parse.parse_qs(query, keep_blank_values=True)
        with self.lock:
            pending = self.pending
            if not pending or pending['expires'] < time.time() or len(values.get('state', [])) != 1 or not secrets.compare_digest(values['state'][0], pending['state']):
                raise AuthError('This sign-in callback is invalid or expired. Start sign-in again from Email Studio.')
            self.pending = None  # One-use even if exchange fails; replay cannot replace credentials.
            try:
                if values.get('error'):
                    raise AuthError('Sign-in was cancelled or permission was declined. Your existing connection has been kept.')
                if len(values.get('code', [])) != 1 or len(values.get('client_id', [])) > 1:
                    raise AuthError('OpenAI sign-in was incomplete. Start sign-in again.')
                issued = values.get('client_id', [pending['client_id']])[0]
                if not re.fullmatch(r'oaiapp_[A-Za-z0-9_-]{1,200}', issued) or (pending['client_id'] and issued != pending['client_id']):
                    raise AuthError('The sign-in registration did not match. Start sign-in again.')
                tokens = request_json(self.discovery()['token_endpoint'], {'grant_type': 'authorization_code',
                    'client_id': issued, 'code': values['code'][0], 'code_verifier': pending['verifier'],
                    'redirect_uri': pending['redirect_uri'], 'resource': RESOURCE})
                claims = self.validate_identity(tokens.get('id_token', ''), issued, pending['nonce'], pending['subject'])
                scopes = tokens.get('scope', '').split()
                saved = self.store.read('chatgpt', {})
                profile = {'client_id': issued, 'subject': claims['sub'], 'email': str(claims.get('email', ''))[:320],
                           'name': str(claims.get('name', ''))[:200], 'scopes': scopes, 'models': {},
                           'access_token': tokens.get('access_token', ''), 'refresh_token': tokens.get('refresh_token', ''),
                           'id_token': tokens.get('id_token', ''), 'expires_at': time.time() + int(tokens.get('expires_in', 3600))}
                if tokens.get('token_type', '').lower() != 'bearer' or not profile['access_token']:
                    raise AuthError('OpenAI did not return a usable connection. Start sign-in again.')
                saved.setdefault('profiles', {})[issued] = profile
                saved['active'] = issued
                self.store.write('chatgpt', saved)
                settings = self.store.settings()
                settings['provider'] = 'chatgpt'
                self.store.write('settings', settings)
                self.message = 'Signed in. ChatGPT plan permission is enabled.' if 'chatgpt.tokens.use.direct' in scopes else 'Signed in, but ChatGPT plan permission was not granted. Continue with ChatGPT to enable it.'
                if 'chatgpt.tokens.use.direct' in scopes:
                    try:
                        self.models()
                    except AuthError:
                        self.message = 'Signed in. Could not load models yet; use Refresh models in Settings.'
                return self.message
            except AuthError as error:
                self.message = str(error)
                raise

    def access_token(self):
        with self.lock, session_lock(self.store):
            saved = self.store.read('chatgpt', {})
            profile = saved.get('profiles', {}).get(saved.get('active', ''), {})
            if not profile.get('access_token') or 'chatgpt.tokens.use.direct' not in profile.get('scopes', []):
                raise AuthError('Choose Continue with ChatGPT in Settings and enable use of your ChatGPT plan.')
            if profile.get('expires_at', 0) > time.time() + 90:
                return profile['access_token']
            if not profile.get('refresh_token'):
                raise AuthError('Your ChatGPT session has expired. Choose Continue with ChatGPT to reconnect.')
            tokens = request_json(self.discovery()['token_endpoint'], {'grant_type': 'refresh_token',
                'client_id': profile['client_id'], 'refresh_token': profile['refresh_token'], 'resource': RESOURCE})
            if not tokens.get('access_token') or tokens.get('token_type', '').lower() != 'bearer':
                raise AuthError('ChatGPT could not renew the connection. Sign in again.')
            if tokens.get('id_token'):
                self.validate_identity(tokens['id_token'], profile['client_id'], subject=profile['subject'])
            if 'scope' in tokens:
                profile['scopes'] = tokens['scope'].split()
                if 'chatgpt.tokens.use.direct' not in profile['scopes']:
                    raise AuthError('ChatGPT plan permission is no longer enabled. Sign in again to grant it.')
            for field in ('access_token', 'refresh_token', 'id_token'):
                if tokens.get(field):profile[field] = tokens[field]
            profile['expires_at'] = time.time() + int(tokens.get('expires_in', 3600))
            self.store.write('chatgpt', saved)
            return profile['access_token']

    def models(self):
        with self.lock:
            catalog = request_json(RESOURCE + '/models', bearer=self.access_token())
            models = {m['slug']: str(m.get('display_name', m['slug']))[:150] for m in catalog.get('models', [])
                      if m.get('visibility') == 'list' and isinstance(m.get('slug'), str) and re.fullmatch(r'[A-Za-z0-9._-]{1,100}', m['slug'])}
            if not models:
                raise AuthError('No models are available for this ChatGPT account. Check your plan or workspace permissions.')
            saved = self.store.read('chatgpt', {})
            saved['profiles'][saved['active']]['models'] = models
            self.store.write('chatgpt', saved)
            settings = self.store.settings()
            if settings['model'] not in models:
                settings['model'] = 'gpt-6.1-sol' if 'gpt-6.1-sol' in models else next(iter(models))
                self.store.write('settings', settings)
            return models

    def sign_out(self):
        with self.lock, session_lock(self.store):
            self.pending = None
            saved = self.store.read('chatgpt', {})
            profile = saved.get('profiles', {}).get(saved.get('active', ''), {})
            revoked = True
            if profile.get('refresh_token'):
                revoked = False
                for attempt in range(2):
                    try:
                        request_json(self.discovery()['revocation_endpoint'], {'token': profile['refresh_token'],
                            'token_type_hint': 'refresh_token', 'client_id': profile['client_id']})
                        revoked = True
                        break
                    except AuthError:
                        if attempt == 0:time.sleep(.5)
            for field in ('access_token', 'refresh_token', 'id_token', 'models'):
                profile.pop(field, None)
            self.store.write('chatgpt', saved)
            self.message = 'Signed out.' if revoked else 'Signed out locally. Remote sign-out could not be confirmed; disconnect Email Studio in ChatGPT Settings → Security and login.'
            return self.message


def read_stream(response, raise_error):
    """Require a terminal completion; never return partial text after a failed stream."""
    lines = []
    total = 0
    items = {}
    texts = {}
    refusals = {}
    for raw in response:
        total += len(raw)
        if total > 8 * 1024 * 1024:
            raise AuthError('The AI response was too large. Try a shorter request.')
        line = raw.decode('utf-8').rstrip('\r\n')
        if line:
            if line.startswith('data:'):lines.append(line[5:].lstrip())
            continue
        if not lines:continue
        data = '\n'.join(lines)
        lines = []
        if data == '[DONE]':continue
        try:event = json.loads(data)
        except ValueError:raise AuthError('The AI stream was unreadable. Your inputs have been kept.') from None
        kind = event.get('type')
        index = event.get('output_index', 0)
        part_index = event.get('content_index', 0)
        key = (index, part_index)
        if kind == 'response.output_item.done':
            items[index] = event['item']
        elif kind == 'response.output_text.delta':
            texts[key] = texts.get(key, '') + event.get('delta', '')
        elif kind == 'response.output_text.done':
            texts[key] = event.get('text', texts.get(key, ''))
        elif kind == 'response.refusal.delta':
            refusals[key] = refusals.get(key, '') + event.get('delta', '')
        elif kind == 'response.refusal.done':
            refusals[key] = event.get('refusal', refusals.get(key, ''))
        if kind == 'response.completed':
            result = event['response']
            # The plan route may omit output from the terminal envelope. Its
            # completed item/text events carry the content. Never use it until
            # this terminal event confirms success.
            if not result.get('output'):
                output = []
                for output_index in sorted(set(items) | {k[0] for k in texts} | {k[0] for k in refusals}):
                    item = items.get(output_index)
                    if item and (item.get('type') != 'message' or item.get('content')):
                        output.append(item)
                        continue
                    content = []
                    for content_index in sorted({k[1] for k in texts if k[0] == output_index} | {k[1] for k in refusals if k[0] == output_index}):
                        slot = (output_index, content_index)
                        if slot in refusals:content.append({'type':'refusal','refusal':refusals[slot]})
                        elif slot in texts:content.append({'type':'output_text','text':texts[slot]})
                    output.append({'type':'message','role':'assistant','content':content})
                result['output'] = output
            return result
        if kind in ('response.failed', 'error'):
            error = event.get('response', {}).get('error') or event.get('error') or event
            raise_error(error.get('code', ''), response.headers.get('x-request-id', ''), getattr(response, 'status', 200))
        if kind == 'response.incomplete':
            raise AuthError('ChatGPT could not finish this draft. Your inputs have been kept; try a shorter request.')
    raise AuthError('The ChatGPT connection ended before the draft completed. Your inputs have been kept; try again.')
