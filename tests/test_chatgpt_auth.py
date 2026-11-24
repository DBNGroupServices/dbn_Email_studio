"""OAuth verification and subscription routing tests; no real accounts or requests."""
import base64
import hashlib
import io
import json
from pathlib import Path
import sys
import tempfile
import threading
import time
import unittest
import urllib.error
import urllib.parse
from unittest.mock import patch

from cryptography.hazmat.primitives.asymmetric import rsa
import jwt

sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import app
import chatgpt_auth as auth

DISCOVERY={'issuer':auth.ISSUER,'authorization_endpoint':auth.ISSUER+'/api/accounts/authorize',
           'token_endpoint':auth.ISSUER+'/api/accounts/oauth/token','revocation_endpoint':auth.ISSUER+'/api/accounts/oauth/revoke',
           'jwks_uri':auth.ISSUER+'/.well-known/jwks.json'}


class MemoryStore:
    def __init__(self):self.data={}
    def read(self,name,default):return json.loads(json.dumps(self.data.get(name,default)))
    def write(self,name,value):self.data[name]=json.loads(json.dumps(value))
    def settings(self):return app.DEFAULTS|self.read('settings',{})


class SignInTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.private=rsa.generate_private_key(public_exponent=65537,key_size=2048)
        cls.jwk=json.loads(jwt.algorithms.RSAAlgorithm.to_jwk(cls.private.public_key()))|{'kid':'fixture-key','alg':'RS256'}

    def setUp(self):
        self.store=MemoryStore();self.client=auth.ChatGPTAuth(self.store,'http://127.0.0.1:8767/auth/callback')
        self.client.discovery_cache=DISCOVERY
        self.client.jwks_cache={'keys':[self.jwk]};self.client.jwks_time=time.time()

    def token(self,**changes):
        return jwt.encode({'iss':auth.ISSUER,'aud':'oaiapp_fixture','sub':'fixture-user','iat':int(time.time()),'exp':int(time.time())+3600,
            'nonce':self.client.pending['nonce'] if self.client.pending else 'fixture-nonce',
            'email':'sam@example.com',**changes},self.private,algorithm='RS256',headers={'kid':'fixture-key'})

    def start(self):return urllib.parse.parse_qs(urllib.parse.urlparse(self.client.start()).query)

    def callback(self):
        return urllib.parse.urlencode({'state':self.client.pending['state'],'code':'fictional-code','client_id':'oaiapp_fixture'})

    def fake_tokens(self,**changes):
        return {'id_token':self.token(),'access_token':'fictional-access','refresh_token':'fictional-refresh',
                'token_type':'Bearer','expires_in':3600,'scope':auth.SCOPES,**changes}

    def test_registration_pkce_host_stability_and_no_secrets_in_public_status(self):
        params=self.start();pending=self.client.pending.copy();host=params['ext_agent_host_id']
        self.assertEqual(params['client_id'],['dynamic_agent_client'])
        self.assertEqual(params['resource'],[auth.RESOURCE])
        self.assertIn('chatgpt.tokens.use.direct',params['scope'][0])
        challenge=base64.urlsafe_b64encode(hashlib.sha256(pending['verifier'].encode()).digest()).rstrip(b'=').decode()
        self.assertEqual(params['code_challenge'],[challenge])
        self.assertEqual(self.start()['ext_agent_host_id'],host)
        self.assertNotEqual(self.client.pending['state'],pending['state'])
        self.assertNotIn(pending['verifier'],json.dumps(self.client.status()))

    def test_callback_validates_identity_then_saves_without_exposing_tokens(self):
        self.start();query=self.callback();tokens=self.fake_tokens()
        catalog={'models':[{'visibility':'list','slug':'gpt-6.1-sol','display_name':'GPT-6.1 Sol'}]}
        with patch.object(auth,'request_json',side_effect=[tokens,catalog]) as request:
            self.client.callback(query)
        exchange=request.call_args_list[0].args[1]
        self.assertEqual(exchange['client_id'],'oaiapp_fixture')
        self.assertEqual(exchange['redirect_uri'],'http://127.0.0.1:8767/auth/callback')
        self.assertNotIn('client_secret',exchange)
        public=self.client.status();self.assertTrue(public['enabled'])
        self.assertNotIn('fictional-access',json.dumps(public))
        self.assertNotIn('fictional-refresh',json.dumps(public))
        self.assertEqual(public['models'],{'gpt-6.1-sol':'GPT-6.1 Sol'})
        with self.assertRaises(auth.AuthError):self.client.callback(query)

    def test_invalid_state_does_not_exchange_or_consume_valid_attempt(self):
        self.start()
        with patch.object(auth,'request_json') as request,self.assertRaises(auth.AuthError):
            self.client.callback('state=wrong&code=x&client_id=oaiapp_fixture')
        request.assert_not_called();self.assertIsNotNone(self.client.pending)

    def test_expired_denied_and_wrong_client_callbacks_keep_existing_credentials(self):
        for mode in ('expired','denied','wrong-client'):
            self.start();query=self.callback();before=self.store.read('chatgpt',{})
            if mode=='expired':self.client.pending['expires']=0
            if mode=='denied':query+='&error=access_denied'
            if mode=='wrong-client':
                self.client.pending['client_id']='oaiapp_original'
            with self.subTest(mode=mode),patch.object(auth,'request_json') as request,self.assertRaises(auth.AuthError):self.client.callback(query)
            request.assert_not_called();self.assertEqual(self.store.read('chatgpt',{}),before)

    def test_signed_token_requires_issuer_audience_nonce_and_expiry(self):
        self.start()
        for changes in ({'iss':'https://attacker.example'},{'aud':'other-app'},{'nonce':'wrong'},{'exp':int(time.time())-60},{'sub':''}):
            with self.subTest(changes=changes),self.assertRaises(auth.AuthError):
                self.client.validate_identity(self.token(**changes),'oaiapp_fixture',self.client.pending['nonce'])
        with self.assertRaises(auth.AuthError):
            self.client.validate_identity(jwt.encode({'sub':'fake'},'key',algorithm='HS256'),'oaiapp_fixture','nonce')

    def test_missing_plan_permission_never_enables_inference(self):
        self.start();query=self.callback();tokens=self.fake_tokens(scope='openid email profile')
        with patch.object(auth,'request_json',return_value=tokens):self.client.callback(query)
        self.assertTrue(self.client.status()['signed_in']);self.assertFalse(self.client.status()['enabled'])
        with self.assertRaises(auth.AuthError):self.client.access_token()

    def saved_profile(self):
        profile={'client_id':'oaiapp_fixture','subject':'fixture-user','email':'sam@example.com','access_token':'old-access',
                 'refresh_token':'old-refresh','id_token':'old-id','expires_at':0,'scopes':auth.SCOPES.split(),'models':{}}
        self.store.write('chatgpt',{'host_id':'urn:uuid:fixture','active':'oaiapp_fixture','profiles':{'oaiapp_fixture':profile}})

    def test_refresh_rotation_is_serialized_and_keeps_registration(self):
        self.saved_profile()
        with patch.object(auth,'request_json',return_value={'access_token':'new-access','refresh_token':'new-refresh','token_type':'Bearer','expires_in':3600}) as request:
            results=[]
            threads=[threading.Thread(target=lambda:results.append(self.client.access_token())) for _ in range(3)]
            for thread in threads:thread.start()
            for thread in threads:thread.join()
        self.assertEqual(results,['new-access']*3);request.assert_called_once()
        self.assertEqual(request.call_args.args[1]['client_id'],'oaiapp_fixture')
        self.assertNotIn('scope',request.call_args.args[1])
        saved=self.store.read('chatgpt',{})['profiles']['oaiapp_fixture'];self.assertEqual(saved['refresh_token'],'new-refresh')

    def test_signout_revokes_and_clears_tokens_but_retains_host_and_client(self):
        self.saved_profile()
        with patch.object(auth,'request_json',return_value={}) as request:self.client.sign_out()
        self.assertEqual(request.call_args.args[1]['token_type_hint'],'refresh_token')
        self.assertFalse(self.client.status()['signed_in'])
        self.assertEqual(self.store.read('chatgpt',{})['host_id'],'urn:uuid:fixture')
        self.assertIn('oaiapp_fixture',self.store.read('chatgpt',{})['profiles'])


class SubscriptionRequestTests(unittest.TestCase):
    def response(self,*events):
        stream=io.BytesIO(b''.join(('data: '+json.dumps(event)+'\n\n').encode() for event in events))
        stream.headers={'x-request-id':'req_fixture'}
        return stream

    def test_subscription_request_uses_public_endpoint_stream_and_supported_fields(self):
        completed={'type':'response.completed','response':{'status':'completed','output':[]}}
        with patch('urllib.request.urlopen',return_value=self.response(completed)) as request:
            value=app.openai_request('fictional-oauth','responses',{'model':'test-model','input':[], 'max_output_tokens':500,'temperature':1,'previous_response_id':'private'},plan=True)
        sent=request.call_args.args[0];payload=json.loads(sent.data)
        self.assertEqual(sent.full_url,'https://api.openai.com/v1/responses')
        self.assertTrue(payload['stream']);self.assertFalse(payload['store'])
        self.assertNotIn('max_output_tokens',payload);self.assertNotIn('temperature',payload);self.assertNotIn('previous_response_id',payload)
        self.assertEqual(value['status'],'completed')

    def test_partial_stream_failed_limit_and_disconnect_never_return_a_draft(self):
        delta={'type':'response.output_text.delta','delta':'partial draft'}
        failed={'type':'response.failed','response':{'error':{'code':'subscription_sharing_usage_limit_exceeded'}}}
        with self.assertRaises(app.OpenAIError) as raised:auth.read_stream(self.response(delta,failed),app.plan_failure)
        self.assertEqual(raised.exception.diagnostics['code'],'subscription_sharing_usage_limit_exceeded')
        with self.assertRaises(auth.AuthError):auth.read_stream(self.response(delta),app.plan_failure)

    def test_completed_envelope_without_output_uses_finished_stream_items(self):
        draft={'subject':'Thanks','body':'Hi Sam, thank you for the update.','checks':[]}
        item={'type':'message','content':[{'type':'output_text','text':json.dumps(draft)}]}
        result=auth.read_stream(self.response({'type':'response.output_item.done','output_index':1,'item':item},
            {'type':'response.completed','response':{'status':'completed','output':[]}}),app.plan_failure)
        self.assertEqual(app.read_response(result),draft)

    def test_completed_envelope_without_items_reassembles_text_deltas(self):
        text=json.dumps({'subject':'Thanks','body':'Thank you, Sam.','checks':[]})
        events=[{'type':'response.output_text.delta','output_index':0,'content_index':0,'delta':chunk} for chunk in (text[:20],text[20:])]
        events.append({'type':'response.completed','response':{'status':'completed','output':[]}})
        result=auth.read_stream(self.response(*events),app.plan_failure)
        self.assertEqual(app.read_response(result)['body'],'Thank you, Sam.')

    def test_streamed_refusal_is_not_mistaken_for_a_finished_draft(self):
        result=auth.read_stream(self.response({'type':'response.refusal.done','refusal':'No'},
            {'type':'response.completed','response':{'status':'completed','output':[]}}),app.plan_failure)
        with self.assertRaises(app.UserError):app.read_response(result)

    def test_subscription_429_never_suggests_api_credits_or_falls_back(self):
        failure=urllib.error.HTTPError(auth.RESOURCE+'/responses',429,'Limit',{},io.BytesIO(b'{"error":{"code":"subscription_sharing_usage_limit_exceeded"}}'))
        with patch('urllib.request.urlopen',side_effect=failure) as request,self.assertRaises(app.OpenAIError) as raised:
            app.openai_request('fictional-oauth','responses',{},plan=True)
        request.assert_called_once();self.assertIn('Manage usage',str(raised.exception));self.assertNotIn('API credit',str(raised.exception))

    def test_new_install_defaults_to_subscription_and_secrets_stay_encrypted(self):
        root=Path(__file__).resolve().parents[1]/'.test-data'
        with tempfile.TemporaryDirectory(dir=root) as temp:
            store=app.Store(Path(temp));self.assertEqual(store.settings()['provider'],'chatgpt')
            store.write('chatgpt',{'active':'oaiapp_x','profiles':{'oaiapp_x':{'access_token':'SECRET_ACCESS','refresh_token':'SECRET_REFRESH','scopes':auth.SCOPES.split(),'models':{}}}})
            self.assertNotIn(b'SECRET',(Path(temp)/'chatgpt.dat').read_bytes())
            public=store.snapshot();self.assertTrue(public['connected']);self.assertNotIn('SECRET',json.dumps(public))


if __name__=='__main__':unittest.main()
