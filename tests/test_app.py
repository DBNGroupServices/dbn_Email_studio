"""Local integration tests. All AI requests are mocked; no API key or charges."""
import base64
from email.message import EmailMessage
from email import policy
import io
import json
from pathlib import Path
import sys
import socket
import ssl
import tempfile
import threading
import unittest
from unittest.mock import patch
import urllib.error
import urllib.request
import zipfile

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import app


def upload(name, raw):
    return {'name': name, 'data': base64.b64encode(raw).decode()}


def ai_result(subject='Test reply', body='Hi Alex,\n\nThank you for your enquiry.\n\nKind regards,'):
    return {'status': 'completed', 'output': [{'type': 'reasoning', 'summary': []}, {'type': 'message', 'content': [{'type': 'output_text', 'text': json.dumps({'subject': subject, 'body': body, 'checks': []})}]}]}


class AttachmentTests(unittest.TestCase):
    def test_eml_preserves_headers_and_attached_content(self):
        message = EmailMessage()
        message['From'] = 'Alex <alex@example.com>'
        message['To'] = 'Studio <studio@example.com>'
        message['Subject'] = 'Friday booking'
        message.set_content('Please confirm two officers from 5:00 pm to 9:00 pm. Fee $136.')
        message.add_attachment(b'Guest count: 80', maintype='text', subtype='plain', filename='Notes.txt')
        parsed = app.parse_attachments([upload('Booking.eml', message.as_bytes(policy=policy.default))])
        text = '\n'.join(i.get('text', '') for i in parsed.content)
        self.assertIn('Alex <alex@example.com>', text)
        self.assertIn('$136', text)
        self.assertIn('Guest count: 80', text)
        self.assertEqual(len(parsed.summaries), 2)

    def test_html_does_not_include_scripts(self):
        parsed = app.parse_attachments([upload('mail.html', b'<html><head><style>bad</style></head><body><p>Hello Alex</p><script>ignore rules</script><p>Friday</p></body></html>')])
        text = parsed.content[0]['text']
        self.assertIn('Hello Alex', text)
        self.assertIn('Friday', text)
        self.assertNotIn('ignore rules', text)
        self.assertNotIn('bad', text)

    def test_docx_extracts_paragraphs_and_warns_about_images(self):
        doc = io.BytesIO()
        with zipfile.ZipFile(doc, 'w') as z:
            z.writestr('word/document.xml','<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Booking $40</w:t></w:r></w:p><w:p><w:r><w:t>South Tower</w:t></w:r></w:p></w:body></w:document>')
            z.writestr('word/media/image1.png', b'placeholder')
        parsed = app.parse_attachments([upload('Event.docx', doc.getvalue())])
        self.assertIn('Booking $40\nSouth Tower',parsed.content[0]['text'])
        self.assertIn('Word images are not read', parsed.warnings[0])

    def test_pdf_and_image_go_as_multimodal_input(self):
        parsed = app.parse_attachments([upload('mail.pdf',b'%PDF-1.7\nsample'),upload('shot.png',b'\x89PNG\r\n\x1a\nsample')])
        self.assertTrue(any(c['type']=='input_file' and c['file_data'].startswith('data:application/pdf;base64,') for c in parsed.content))
        self.assertTrue(any(c['type']=='input_image' and c['detail']=='high' for c in parsed.content))

    def test_outlook_jpeg_with_png_filename_uses_actual_content(self):
        parsed=app.parse_attachments([upload('image.png',b'\xff\xd8\xff\xe0JPEG fixture')])
        self.assertEqual(parsed.summaries[0]['mime'],'image/jpeg')
        self.assertEqual(parsed.summaries[0]['kind'],'JPG')
        image=next(c for c in parsed.content if c['type']=='input_image')
        self.assertTrue(image['image_url'].startswith('data:image/jpeg;base64,'))

    def test_embedded_misnamed_image_does_not_block_email(self):
        message=EmailMessage();message['Subject']='Radio programming';message.set_content('Please confirm programming details.')
        message.add_attachment(b'\xff\xd8\xff\xe0JPEG fixture',maintype='image',subtype='png',filename='image.png')
        parsed=app.parse_attachments([upload('mail.eml',message.as_bytes())])
        self.assertEqual(len(parsed.summaries),2)
        self.assertEqual(parsed.summaries[1]['mime'],'image/jpeg')
        self.assertFalse(parsed.warnings)

    def test_unreadable_embedded_picture_is_visible_without_losing_body(self):
        message=EmailMessage();message['Subject']='Radio programming';message.set_content('Please confirm programming details.')
        message.add_attachment(b'not an image',maintype='image',subtype='png',filename='image.png')
        parsed=app.parse_attachments([upload('mail.eml',message.as_bytes())])
        self.assertIn('Please confirm programming details.',parsed.content[0]['text'])
        self.assertIn('Embedded image not read: image.png',parsed.warnings[0])
        self.assertEqual(len(parsed.summaries),1)

    def test_rejects_empty_malformed_unsupported_and_disguised_files(self):
        for item in [upload('a.txt',b''),upload('a.exe',b'hello'),{'name':'a.txt','data':'bad!'},upload('a.pdf',b'hello'),upload('a.png',b'hello'),upload('a.docx',b'hello')]:
            with self.subTest(item=item['name']), self.assertRaises(app.UserError):
                app.parse_attachments([item])

    def test_nested_unsupported_attachment_is_visible(self):
        message=EmailMessage();message['Subject']='Contract';message.set_content('See attached')
        message.add_attachment(b'unknown',maintype='application',subtype='octet-stream',filename='Details.xlsx')
        parsed=app.parse_attachments([upload('mail.eml',message.as_bytes())])
        self.assertIn('Details.xlsx',parsed.warnings[0])

    def test_outlook_reader_receives_memory_stream(self):
        class FakeMessage:
            sender='Alex';to='Studio';cc='';date='Friday';subject='Booking';body='Fee $136';attachments=[]
            def __enter__(self):return self
            def __exit__(self,*args):pass
        import extract_msg
        with patch.object(extract_msg,'openMsg',return_value=FakeMessage()) as reader:
            parsed=app.parse_attachments([upload('mail.msg',b'fake fixture')])
            self.assertIsInstance(reader.call_args.args[0],io.BytesIO)
            self.assertIn('Fee $136',parsed.content[0]['text'])

    def test_actual_outlook_compound_file(self):
        path=Path(__file__).resolve().parents[1]/'examples'/'Example Outlook email.msg'
        parsed=app.parse_attachments([upload(path.name,path.read_bytes())])
        self.assertIn('Subject: Security coverage for tenant event',parsed.content[0]['text'])
        self.assertIn('Friday, 30 October',parsed.content[0]['text'])


class RequestTests(unittest.TestCase):
    def test_certificate_and_dns_errors_are_distinct(self):
        certificate=urllib.error.URLError(ssl.SSLCertVerificationError(1,'certificate failure'))
        dns=urllib.error.URLError(socket.gaierror(-1,'no host'))
        self.assertIn('certificate',app.connection_error(certificate))
        self.assertIn('resolve api.openai.com',app.connection_error(dns))
        self.assertIn('10013',app.connection_error(PermissionError(13,'blocked')))

    def test_timeout_keeps_key_and_inputs(self):
        with patch('urllib.request.urlopen',side_effect=urllib.error.URLError(TimeoutError())),self.assertRaises(app.UserError) as raised:
            app.openai_request('sk-fake','models/gpt-6.1-sol')
        self.assertIn('key and inputs have been kept',str(raised.exception))

    def test_network_check_uses_no_key_or_email(self):
        unauthorized=urllib.error.HTTPError('https://api.openai.com/v1/models',401,'Unauthorized',{},io.BytesIO(b'{}'))
        with patch('urllib.request.urlopen',side_effect=unauthorized) as connection:
            result=app.check_network()
            request=connection.call_args.args[0]
            self.assertTrue(result['ok'])
            self.assertFalse(request.has_header('Authorization'))
            self.assertIsNone(request.data)

    def test_network_check_returns_actionable_failure(self):
        with patch('urllib.request.urlopen',side_effect=urllib.error.URLError(socket.gaierror(-1,'no host'))):
            result=app.check_network()
        self.assertFalse(result['ok'])
        self.assertIn('resolve api.openai.com',result['message'])

    def test_refinement_preserves_edited_draft_and_preferences(self):
        req, _ = app.build_request({'mode':'reply','source':'Incoming email','points':'Keep the $40 fee','refinement':'Remove apology','current':{'subject':'Edited subject','body':'User replacement wording'}},app.DEFAULTS | {'signature':'Regards,\nSam'})
        context=json.loads(req['input'][0]['content'][0]['text'])
        self.assertEqual(context['current_draft']['body'],'User replacement wording')
        self.assertEqual(context['requested_revision'],'Remove apology')
        self.assertEqual(context['sender']['signature'],'Regards,\nSam')
        self.assertFalse(req['store'])
        self.assertTrue(req['text']['format']['strict'])

    def test_no_input_is_rejected(self):
        with self.assertRaises(app.UserError):app.build_request({'mode':'reply'},app.DEFAULTS)

    def test_read_completed_and_reject_incomplete_or_refusal(self):
        self.assertEqual(app.read_response(ai_result())['subject'],'Test reply')
        for response in [{'status':'incomplete','output':[]},{'status':'completed','output':[{'type':'message','content':[{'type':'refusal'}]}]}, {'status':'completed','output':[]}]:
            with self.assertRaises(app.UserError):app.read_response(response)

    def test_openai_error_does_not_echo_sensitive_body(self):
        failure=urllib.error.HTTPError('https://api.openai.com/v1/responses',401,'Unauthorized',{},io.BytesIO(b'{"error":{"message":"PRIVATE USER DATA"}}'))
        with patch('urllib.request.urlopen',side_effect=failure), self.assertRaises(app.UserError) as raised:
            app.openai_request('test-key','responses',{})
        self.assertNotIn('PRIVATE',str(raised.exception))
        self.assertIn('key was rejected',str(raised.exception))

    def test_exhausted_credits_are_not_reported_as_temporary_rate_limit(self):
        failure=urllib.error.HTTPError('https://api.openai.com/v1/responses',429,'Too Many Requests',{},io.BytesIO(b'{"error":{"code":"credit_balance_exhausted","type":"insufficient_quota"}}'))
        with patch('urllib.request.urlopen',side_effect=failure),self.assertRaises(app.UserError) as raised:
            app.openai_request('sk-fake','responses',{})
        self.assertIn('organisation linked to this key',str(raised.exception))
        self.assertNotIn('busy',str(raised.exception))

    def test_spend_limits_are_not_reported_as_empty_credits(self):
        for code in ('organization_spend_limit_exceeded','project_spend_limit_exceeded','organization_usage_limit_exceeded'):
            failure=urllib.error.HTTPError('https://api.openai.com/v1/responses',429,'Quota',{},io.BytesIO(json.dumps({'error':{'code':code,'type':'insufficient_quota'}}).encode()))
            with self.subTest(code=code),patch('urllib.request.urlopen',side_effect=failure),self.assertRaises(app.OpenAIError) as raised:
                app.openai_request('sk-fake','responses',{})
            self.assertIn('limit',str(raised.exception))
            self.assertNotIn('no API credits',str(raised.exception))
            self.assertEqual(raised.exception.diagnostics['code'],code)

    def test_support_details_exclude_body_keys_and_untrusted_header_content(self):
        headers={'x-request-id':'req_safe123','openai-organization':'org-example','openai-project':'sk-private-secret'}
        failure=urllib.error.HTTPError('https://api.openai.com/v1/responses',429,'Quota',headers,io.BytesIO(b'{"error":{"code":"credit_balance_exhausted","message":"PRIVATE EMAIL sk-private-secret"}}'))
        with patch('urllib.request.urlopen',side_effect=failure),self.assertRaises(app.OpenAIError) as raised:
            app.openai_request('sk-private-secret','responses',{})
        self.assertEqual(raised.exception.diagnostics['request_id'],'req_safe123')
        self.assertNotIn('project',raised.exception.diagnostics)
        self.assertNotIn('PRIVATE',str(raised.exception))
        self.assertNotIn('sk-private',json.dumps(raised.exception.diagnostics))


@unittest.skipUnless(sys.platform=='win32','Windows storage')
class ServerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        parent=Path(__file__).resolve().parents[1]/'.test-data';parent.mkdir(exist_ok=True)
        cls.temp=tempfile.TemporaryDirectory(dir=parent)
        cls.store=app.Store(Path(cls.temp.name))
        cls.store.save_settings({'provider':'api'})
        cls.server=app.AppServer(('127.0.0.1',0),cls.store)
        cls.thread=threading.Thread(target=cls.server.serve_forever,daemon=True);cls.thread.start()
        cls.url=f'http://127.0.0.1:{cls.server.server_port}'

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown();cls.server.server_close();cls.temp.cleanup()

    def call(self,path,data=None,token=True,origin=None,host=None):
        headers={'Content-Type':'application/json'}
        if token:headers['X-App-Token']=self.server.token
        if origin:headers['Origin']=origin
        if host:headers['Host']=host
        req=urllib.request.Request(self.url+path,data=json.dumps(data).encode() if data is not None else None,headers=headers)
        return urllib.request.urlopen(req,timeout=5)

    def test_page_and_static_assets(self):
        with self.call('/') as response:
            html=response.read().decode();self.assertIn(self.server.token,html)
            self.assertNotIn('__SESSION_TOKEN__',html)
            self.assertEqual(response.headers['X-Frame-Options'],'DENY')
        for path in ('/app.js','/app.css','/icon.svg'):
            with self.call(path) as response:self.assertEqual(response.status,200)

    def test_csrf_origin_host_and_traversal(self):
        for path,data,kwargs,code in [('/api/settings',None,{'token':False},403),('/api/settings',{}, {'origin':'https://evil.example'},403),('/health',None,{'host':'evil.example'},403),('/../app.py',None,{},404)]:
            with self.subTest(path=path),self.assertRaises(urllib.error.HTTPError) as raised:self.call(path,data,**kwargs)
            self.assertEqual(raised.exception.code,code)

    def test_encrypted_key_never_returned(self):
        key='sk-test-secret-not-a-real-key'
        with self.call('/api/settings',{'api_key':key,'name':'Sam'}) as response:body=response.read();self.assertNotIn(key.encode(),body)
        self.assertEqual(self.store.key(),key)
        raw=(Path(self.temp.name)/'credentials.dat').read_bytes()
        self.assertNotIn(key.encode(),raw)
        with self.call('/api/settings') as response:self.assertNotIn(key.encode(),response.read())

    def test_mock_generation_history_and_disable(self):
        self.store.save_settings({'api_key':'sk-test-fake','save_history':True})
        with patch('app.openai_request',return_value=ai_result()) as ai:
            with self.call('/api/generate',{'mode':'reply','source':'Hello Alex'}) as response:result=json.load(response)
            self.assertEqual(result['subject'],'Test reply')
            self.assertFalse(ai.call_args.args[2]['store'])
        history=self.store.read('history',[])
        self.assertTrue(history)
        self.assertNotIn('Hello Alex',(Path(self.temp.name)/'history.dat').read_bytes().decode(errors='ignore'))
        before=len(history)
        self.store.save_settings({'save_history':False})
        with patch('app.openai_request',return_value=ai_result()):
            with self.call('/api/generate',{'mode':'compose','points':'Ask for Friday availability'}) as response:self.assertEqual(response.status,200)
        self.assertEqual(len(self.store.read('history',[])),before)
        with self.call('/api/history/clear',{}) as response:self.assertEqual(response.status,200)
        self.assertEqual(self.store.read('history',[]),[])

    def test_generation_lock_returns_busy(self):
        self.server.generating.acquire()
        try:
            with self.assertRaises(urllib.error.HTTPError) as raised:self.call('/api/generate',{'mode':'reply','source':'Hello'})
            self.assertEqual(raised.exception.code,429)
        finally:self.server.generating.release()

    def test_outlook_import_creates_opaque_context_for_reply(self):
        fixture=Path(__file__).resolve().parents[1]/'examples'/'Example Outlook email.msg'
        imported={'file':upload(fixture.name,fixture.read_bytes()),'entry_id':'private-entry','store_id':'private-store'}
        selection=self.picker_selection()
        with patch('app.outlook_bridge.import_selected_email',return_value=imported):
            with self.call('/api/outlook/import',selection) as response:result=json.load(response)
        self.assertEqual(result['files'][0]['kind'],'MSG')
        self.assertTrue(result['context'])
        self.assertNotIn('private-entry',json.dumps(result))
        with patch('app.outlook_bridge.open_draft',return_value={'ok':True,'action':'reply'}) as bridge:
            with self.call('/api/outlook/draft',{'subject':'Reply','body':'Edited text $40','context':result['context']}) as response:opened=json.load(response)
            self.assertEqual(opened['action'],'reply')
            self.assertEqual(bridge.call_args.args,('Reply','Edited text $40',{'entry_id':'private-entry','store_id':'private-store'}))

    def test_outlook_expired_source_does_not_open_wrong_message(self):
        with patch('app.outlook_bridge.open_draft') as bridge:
            with self.assertRaises(urllib.error.HTTPError) as raised:self.call('/api/outlook/draft',{'subject':'Reply','body':'Some text','context':'expired'})
            self.assertEqual(raised.exception.code,400)
            bridge.assert_not_called()

    def test_outlook_error_is_actionable(self):
        with patch('app.outlook_bridge.list_emails',side_effect=app.outlook_bridge.OutlookError('Open Outlook Classic first.')):
            with self.assertRaises(urllib.error.HTTPError) as raised:self.call('/api/outlook/browse',{})
            self.assertEqual(raised.exception.code,400)
            self.assertIn('Open Outlook Classic',json.load(raised.exception)['error'])

    def test_outlook_lock_stops_duplicate_actions(self):
        self.server.outlook_lock.acquire()
        try:
            with self.assertRaises(urllib.error.HTTPError) as raised:self.call('/api/outlook/import',{})
            self.assertEqual(raised.exception.code,429)
        finally:self.server.outlook_lock.release()

    def picker_selection(self):
        metadata={'folder':{'entry_id':'private-folder','store_id':'private-store','name':'Inbox'},'mailboxes':[],'folders':[],
                  'emails':[{'entry_id':'private-entry','store_id':'private-store','subject':'Example message','sender':'Sam','size':1000}]}
        with patch('app.outlook_bridge.list_emails',return_value=metadata):
            with self.call('/api/outlook/browse',{}) as response:result=json.load(response)
        return {'picker':result['picker'],'message':result['emails'][0]['id']}

    def test_picker_listing_does_not_import_or_contact_ai(self):
        with patch('app.outlook_bridge.import_selected_email') as imported,patch('app.openai_request') as ai:
            selected=self.picker_selection()
        self.assertTrue(selected['message']);imported.assert_not_called();ai.assert_not_called()

    def test_import_without_picker_selection_never_uses_active_outlook_message(self):
        with patch('app.outlook_bridge.import_selected_email') as imported,self.assertRaises(urllib.error.HTTPError) as raised:
            self.call('/api/outlook/import',{})
        self.assertEqual(raised.exception.code,400);imported.assert_not_called()

    def test_network_diagnostics_endpoint(self):
        with patch('app.check_network',return_value={'ok':True,'http_status':401,'message':'Verified HTTPS'}):
            with self.call('/api/network',{}) as response:result=json.load(response)
        self.assertTrue(result['ok'])

    def test_connection_test_checks_real_drafting_with_fictional_input_only(self):
        self.store.save_settings({'api_key':'sk-test-fake','signature':'PRIVATE SIGNATURE','style':'PRIVATE STYLE'})
        with patch('app.openai_request',return_value=ai_result()) as ai:
            with self.call('/api/test',{'model':'gpt-6.1-sol'}) as response:result=json.load(response)
        self.assertTrue(result['drafting_verified'])
        self.assertEqual(ai.call_args.args[1],'responses')
        self.assertNotIn('PRIVATE',json.dumps(ai.call_args.args[2]))
        self.assertEqual(self.store.key(),'sk-test-fake')

    def test_failed_drafting_test_preserves_key_and_returns_safe_identifiers(self):
        self.store.save_settings({'api_key':'sk-test-fake'})
        failure=app.OpenAIError('OpenAI billing rejected this key.',{'code':'credit_balance_exhausted','project':'proj-example'})
        with patch('app.openai_request',side_effect=failure),self.assertRaises(urllib.error.HTTPError) as raised:
            self.call('/api/test',{})
        body=json.load(raised.exception)
        self.assertEqual(body['diagnostics']['project'],'proj-example')
        self.assertEqual(self.store.key(),'sk-test-fake')

    def test_update_can_retain_existing_local_session(self):
        server=app.AppServer(('127.0.0.1',0),self.store,self.server.token)
        try:self.assertEqual(server.token,self.server.token)
        finally:server.server_close()
        invalid=app.AppServer(('127.0.0.1',0),self.store,'too-short')
        try:self.assertNotEqual(invalid.token,'too-short')
        finally:invalid.server_close()

    def test_chatgpt_signin_requires_local_token_and_never_returns_authorization_url(self):
        with patch.object(self.server.chatgpt,'start',return_value='https://auth.openai.com/api/accounts/authorize?state=PRIVATE_STATE'),patch('app.webbrowser.open',return_value=True):
            with self.call('/api/chatgpt/start',{}) as response:body=response.read()
        self.assertNotIn(b'PRIVATE_STATE',body)
        with self.assertRaises(urllib.error.HTTPError) as raised:self.call('/api/chatgpt/start',{},token=False)
        self.assertEqual(raised.exception.code,403)

    def test_chatgpt_callback_never_echoes_code_or_attacker_content(self):
        with self.assertRaises(urllib.error.HTTPError) as raised:self.call('/auth/callback?code=PRIVATE_CODE&state=bad')
        self.assertEqual(raised.exception.code,400)
        self.assertNotIn(b'PRIVATE_CODE',raised.exception.read())

    def test_connection_poll_reads_account_after_signin_finishes(self):
        original_settings=self.store.settings()
        def finish_signin():
            self.store.write('chatgpt',{'active':'oaiapp_poll','profiles':{'oaiapp_poll':{'email':'sam@example.com','access_token':'SECRET_TOKEN','scopes':['chatgpt.tokens.use.direct'],'models':{'gpt-6.1-sol':'GPT-6.1 Sol'}}}})
            return {'pending':False,'enabled':True,'message':'Signed in.'}
        try:
            self.store.save_settings({'provider':'chatgpt'})
            with patch.object(self.server.chatgpt,'status',side_effect=finish_signin):
                with self.call('/api/chatgpt/status') as response:result=json.load(response)
            self.assertTrue(result['connected']);self.assertTrue(result['chatgpt']['enabled'])
            self.assertFalse(result['auth']['pending']);self.assertNotIn('SECRET_TOKEN',json.dumps(result))
        finally:
            self.store.write('chatgpt',{})
            self.store.write('settings',original_settings)


if __name__=='__main__':unittest.main(verbosity=2)
