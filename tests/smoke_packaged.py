"""Smoke test the built EXE, with isolated Windows-encrypted test data."""
import base64
import json
import os
from pathlib import Path
import subprocess
import sys
import time
import urllib.request

ROOT=Path(__file__).resolve().parents[1]
data_dir=(ROOT/'.test-data'/'packaged-smoke').resolve()
assert data_dir.is_relative_to(ROOT)
data_dir.mkdir(parents=True,exist_ok=True)
environment=os.environ.copy();environment['DBN_EMAIL_DATA_DIR']=str(data_dir)
environment.pop('OPENAI_API_KEY',None)
executable=Path(sys.argv[1]).resolve() if len(sys.argv)>1 else ROOT/'dist'/'DBN Email Studio.exe'
process=subprocess.Popen([str(executable),'--port','8772','--no-browser'],env=environment,creationflags=subprocess.CREATE_NO_WINDOW)
url='http://127.0.0.1:8772'
try:
    for attempt in range(60):
        if process.poll() is not None:raise RuntimeError('Packaged app exited before starting.')
        try:
            with urllib.request.urlopen(url+'/health',timeout=1) as response:
                health=json.load(response)
                assert health['app']=='dbn-email-studio-v1'
                assert health['version']=='1.5.2'
            break
        except (OSError,ValueError):time.sleep(.5)
    else:raise RuntimeError('Packaged app did not start within 30 seconds.')
    with urllib.request.urlopen(url,timeout=5) as response:html=response.read().decode()
    token=html.split('name="app-token" content="')[1].split('"')[0]
    assert 'upload-progress' in html
    assert 'Great emails.' in html and 'i-sparkles' in html
    assert 'outlook-import' in html and 'outlook-open' in html
    assert 'outlook-picker-dialog' in html and 'outlook-picker-import' in html
    def call(path,data=None):
        req=urllib.request.Request(url+'/api/'+path,data=json.dumps(data).encode() if data is not None else None,headers={'X-App-Token':token,'Content-Type':'application/json'})
        with urllib.request.urlopen(req,timeout=30) as response:return json.load(response)
    assert not call('settings')['has_key']
    diagnostic=call('network',{})
    print('Packaged HTTPS diagnostic:',diagnostic,flush=True)
    assert diagnostic['ok'], 'The packaged app cannot reach OpenAI over verified HTTPS.'
    misnamed=call('inspect',{'files':[{'name':'image.png','data':base64.b64encode(b'\xff\xd8\xff\xe0JPEG fixture').decode()}]})
    assert misnamed['files'][0]['mime']=='image/jpeg'
    for name in ('Example incoming email.eml','Example Outlook email.msg'):
        path=ROOT/'examples'/name
        parsed=call('inspect',{'files':[{'name':name,'data':base64.b64encode(path.read_bytes()).decode()}]})
        assert 'Friday, 30 October' in parsed['files'][0]['preview']
        print('Packaged reader OK:',name,flush=True)
    if '--outlook-browse' in sys.argv:
        listing=call('outlook/browse',{})
        assert listing['picker'] and listing['folder']
        assert 'entry_id' not in json.dumps(listing) and 'store_id' not in json.dumps(listing)
        print('Packaged live Outlook picker: headers=',len(listing['emails']),'mailboxes=',len(listing['mailboxes']),flush=True)
    if len(sys.argv)>2 and not sys.argv[2].startswith('--'):
        supplied=Path(sys.argv[2])
        parsed=call('inspect',{'files':[{'name':supplied.name,'data':base64.b64encode(supplied.read_bytes()).decode()}]})
        assert not parsed['warnings']
        assert len(parsed['files'])==9
        print('Supplied email regression: body plus 8 images loaded, no warnings.',flush=True)
    call('settings',{'api_key':'sk-fictional-packaged-test-key','name':'QA fixture'})
    assert call('settings')['has_key']
    assert b'sk-fictional-packaged-test-key' not in (data_dir/'credentials.dat').read_bytes()
    call('settings',{'remove_key':True})
    assert not call('settings')['has_key']
    print('Packaged app startup, assets, Outlook readers and encrypted storage: PASS',flush=True)
finally:
    try:
        request=urllib.request.Request(url+'/api/shutdown',data=b'{}',headers={'X-App-Token':token,'Content-Type':'application/json'})
        urllib.request.urlopen(request,timeout=2).close()
        process.wait(timeout=8)
    except Exception:
        process.terminate()
