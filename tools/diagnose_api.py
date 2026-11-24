"""Local support check. Never prints API keys or user email content."""
import json
from pathlib import Path
import sys
import urllib.request
import urllib.error

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import app

store = app.Store(Path(sys.argv[1]) if len(sys.argv) > 1 else app.DATA_DIR)
key = store.key()
model = store.settings()['model']
print(json.dumps({'saved_key': bool(key), 'model': model, 'source_version': app.VERSION}), flush=True)
try:
    with urllib.request.urlopen('http://127.0.0.1:8767/health', timeout=3) as response:
        print('Running app:', json.dumps(json.load(response)), flush=True)
except Exception as error:
    print('Running app unavailable:', type(error).__name__, flush=True)

def check(path, payload=None):
    request = urllib.request.Request('https://api.openai.com/v1/' + path,
        data=json.dumps(payload).encode() if payload is not None else None,
        headers={'Authorization': 'Bearer ' + key, 'Content-Type': 'application/json', 'User-Agent': 'DBN-Email-Studio/Support'})
    try:
        with urllib.request.urlopen(request, timeout=45) as response:
            value=json.load(response)
            print(json.dumps({'path': path, 'status': response.status,
                'request_id': response.headers.get('x-request-id'),
                'organisation': response.headers.get('openai-organization'),
                'project': response.headers.get('openai-project'),
                'response_status': value.get('status'), 'model': value.get('model', value.get('id'))}), flush=True)
    except urllib.error.HTTPError as error:
        details=json.loads(error.read(20000)).get('error', {})
        print(json.dumps({'path': path, 'http': error.code, 'code': details.get('code'),
            'type': details.get('type'), 'message': str(details.get('message', '')).replace(key, '[REDACTED]'),
            'request_id': error.headers.get('x-request-id'),
            'organisation': error.headers.get('openai-organization'),
            'project': error.headers.get('openai-project')}), flush=True)

if key:
    check('models/' + model)
    request, _ = app.build_request({'mode': 'compose', 'points': 'Fictional test: thank Sam for the update. No signature needed.'}, app.DEFAULTS | {'model': model})
    check('responses', request)
