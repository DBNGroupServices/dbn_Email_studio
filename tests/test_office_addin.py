import base64
import json
from pathlib import Path
import ssl
import socket
import tempfile
import threading
import unittest
import urllib.request
from unittest.mock import patch

import app
import office_tls


class OfficeAddinTests(unittest.TestCase):
    def test_assistant_tasks_are_not_misread_as_email_replies(self):
        for action in ('summary','actions','proofread','translate','followup','meeting-notes'):
            request,_=app.build_assist_request({'action':action,'source':'Fictional text','language':'French'},app.DEFAULTS,app.MODELS)
            self.assertIn('untrusted',request['instructions'])
            self.assertEqual(request['text']['format']['schema'],app.SCHEMA)
            if action=='translate':self.assertIn('French',request['instructions'])
        with self.assertRaises(app.UserError):app.build_assist_request({'action':'send','source':'text'},app.DEFAULTS,app.MODELS)

    def test_office_tls_keeps_keys_encrypted_and_validates_loopback_https(self):
        root=Path(__file__).resolve().parents[1]
        with tempfile.TemporaryDirectory(dir=root/'.test-data') as directory:
            store=app.Store(Path(directory));main=app.AppServer(('127.0.0.1',0),store)
            try:
                with patch.object(office_tls,'PORT',0):status=office_tls.start(main,app.AppServer)
                public=office_tls.public_status(main)
                self.assertTrue(status['running']);self.assertFalse(public['trust_installed_by_app'])
                self.assertNotIn('key',public)
                raw=(Path(directory)/(office_tls.STORE_KEY+'.dat')).read_bytes();self.assertNotIn(b'PRIVATE KEY',raw)
                saved=store.read(office_tls.STORE_KEY,{})
                # Trust only the CA, exactly as Windows does after installation.
                context=ssl.create_default_context(cadata=base64.b64decode(saved['ca_cert']).decode())
                self.assertEqual(context.verify_mode,ssl.CERT_REQUIRED)
                secure=main.office_server
                self.assertIs(secure.chatgpt,main.chatgpt)
                url=f'https://localhost:{secure.server_port}'
                # Browsers preconnect before sending a TLS handshake. One idle
                # client must not prevent Outlook from opening another request.
                with socket.create_connection(('127.0.0.1',secure.server_port),timeout=2):
                    with urllib.request.urlopen(url+'/health',context=context,timeout=2) as response:
                        self.assertEqual(response.status,200)
                with urllib.request.urlopen(url+'/office/taskpane.html',context=context,timeout=5) as response:
                    body=response.read().decode();self.assertIn(main.token,body);self.assertNotIn('__SESSION_TOKEN__',body)
                    self.assertIsNone(response.headers.get('X-Frame-Options'))
                    self.assertIn('appsforoffice.microsoft.com',response.headers['Content-Security-Policy'])
                with urllib.request.urlopen(url+'/icons/email-blue-v2-32.png',context=context,timeout=5) as response:
                    self.assertIn('max-age',response.headers['Cache-Control'])
                    self.assertTrue(response.read().startswith(b'\x89PNG'))
                for path in ('/office/taskpane.html?_host_Info=Outlook%24Win32%2416.02%24en-US',
                             '/office/commands.html?_host_Info=Outlook%24Web',
                             '/office/taskpane.js?v=1'):
                    with urllib.request.urlopen(url+path,context=context,timeout=5) as response:
                        self.assertEqual(response.status,200)
                        self.assertIn('https://outlook.cloud.microsoft',response.headers['Content-Security-Policy'])
                        if 'taskpane.html' in path:
                            self.assertIn(main.token,response.read().decode())
                with self.assertRaises(urllib.error.HTTPError) as denied:
                    urllib.request.urlopen(url+'/api/settings?_host_Info=ignored',context=context,timeout=5)
                self.assertEqual(denied.exception.code,403)
                with urllib.request.urlopen(url+'/',context=context,timeout=5) as response:
                    self.assertEqual(response.headers['X-Frame-Options'],'DENY')
            finally:
                if main.office_server:main.office_server.shutdown();main.office_server.server_close()
                main.server_close()

    def test_manifest_requests_current_item_not_whole_mailbox(self):
        import xml.etree.ElementTree as ET
        source=(Path(__file__).resolve().parents[1]/'office-addin'/'manifest.xml').read_text()
        root=ET.fromstring(source)
        self.assertEqual(root.find('{http://schemas.microsoft.com/office/appforoffice/1.1}Permissions').text,'ReadWriteItem')
        self.assertNotIn('ReadWriteMailbox',source)
        self.assertIn('MessageReadCommandSurface',source);self.assertIn('MessageComposeCommandSurface',source)
        self.assertIn('https://localhost:8788',source)


if __name__=='__main__':unittest.main()
