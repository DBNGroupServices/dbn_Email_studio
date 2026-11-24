"""Disposable browser QA server with simulated AI. Never included in the app."""
import json
from pathlib import Path
import sys
import time

sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import app
import base64

data=Path(__file__).resolve().parents[1]/'.test-data'/'browser-qa'
store=app.Store(data)
store.save_settings({'provider':'chatgpt','api_key':'sk-fictional-browser-test-key','name':'Sam','save_history':True})
store.write('chatgpt', {})

def simulated_ai(key,path,payload=None,**kwargs):
    if path.startswith('models/'):
        return {'id':path[7:]}
    context=json.loads(payload['input'][0]['content'][0]['text'])
    if context['requested_revision']:
        result={'subject':context['current_draft']['subject'],'body':context['current_draft']['body']+'\n\n[Simulated revision applied]','checks':[]}
    else:
        result={'subject':'Re: Security coverage for tenant event','body':'Hi Alex,\n\nThank you for your enquiry. We will check availability for two security officers on Friday, 30 October, from 5:00 pm to 9:00 pm and prepare a quote.\n\nPlease confirm the event location and any special access requirements.\n\nKind regards,\nSam','checks':[]}
    return {'status':'completed','output':[{'type':'message','content':[{'type':'output_text','text':json.dumps(result)}]}]}

app.openai_request=simulated_ai
fixture=Path(__file__).resolve().parents[1]/'examples'/'Example Outlook email.msg'
app.outlook_bridge.import_selected_email=lambda source=None: {'file':{'name':fixture.name,'data':base64.b64encode(fixture.read_bytes()).decode()},'entry_id':(source or {}).get('entry_id','fictional-entry'),'store_id':'fictional-store'}
def fake_folder(name):
    return {'entry_id':name,'store_id':'fictional-store','name':name,'path':'Fictional mailbox'+('\\'+name if name!='Fictional mailbox' else '')}
def fake_listing(folder=None,query='',offset=0):
    current=(folder or {}).get('entry_id','Inbox')
    messages=[{'entry_id':f'fictional-{index}','store_id':'fictional-store',
        'subject':'Security coverage for tenant event' if index==0 else ('<img src=x onerror=alert(1)>' if index==1 else f'Fictional email {index+1:02}'),
        'sender':'Alex' if index<3 else 'Sam','received':'2026-10-28T08:00:00','size':11000000 if index==2 else 87000,
        'has_attachments':index%2==0,'unread':index%3==0} for index in range(75)] if current!='Fictional mailbox' else []
    if query:messages=[row for row in messages if query.lower() in (row['subject']+' '+row['sender']).lower()]
    return {'folder':fake_folder(current),'parent':fake_folder('Fictional mailbox') if current!='Fictional mailbox' else None,
        'inbox':fake_folder('Inbox'),'mailboxes':[fake_folder('Fictional mailbox')],
        'folders':[fake_folder('Inbox'),fake_folder('Sent Items')] if current=='Fictional mailbox' else [],
        'emails':messages[offset:offset+50],'has_more':len(messages)>offset+50}
app.outlook_bridge.list_emails=fake_listing
app.outlook_bridge.open_draft=lambda subject,body,source=None: {'ok':True,'action':'reply' if source else 'new'}
server=app.AppServer(('127.0.0.1',8771),store)
if '--signed-in' in sys.argv:
    store.write('chatgpt', {'active':'oaiapp_fixture','welcomed':False,'profiles':{'oaiapp_fixture':{
        'email':'sam@example.com','name':'Sam','access_token':'fictional-access',
        'scopes':['chatgpt.tokens.use.direct'],'models':{'gpt-6.1-sol':'GPT-6.1 Sol'}}}})
    server.chatgpt.models=lambda: {'gpt-6.1-sol':'GPT-6.1 Sol'}
    server.chatgpt.access_token=lambda: 'fictional-access'
print('TEST ONLY: simulated AI and Outlook, localhost:8771',flush=True)
try:server.serve_forever()
finally:server.server_close()
