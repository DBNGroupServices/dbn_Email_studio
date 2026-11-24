import base64
import json
import os
from pathlib import Path
import subprocess
import sys
import unittest
from unittest.mock import patch

sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import outlook_bridge as bridge


class OutlookBridgeTests(unittest.TestCase):
    def test_reply_is_data_not_executable_shell_text(self):
        result=subprocess.CompletedProcess([],0,stdout=b'{"ok":true,"action":"new"}',stderr=b'')
        body='Hello $(Get-Secret) <script>test</script> $40'
        with patch.object(bridge,'powershell_path',return_value=Path('powershell.exe')),patch.object(bridge.subprocess,'run',return_value=result) as run:
            bridge.open_draft('Subject',body)
            arguments=run.call_args.args[0]
            self.assertNotIn(body,' '.join(arguments))
            self.assertEqual(json.loads(run.call_args.kwargs['input'])['body'],body)
            self.assertEqual(base64.b64decode(arguments[-1]).decode('utf-16-le'),bridge.DRAFT_SCRIPT)
            self.assertNotIn('-ExecutionPolicy',arguments)

    def test_controlled_not_running_error(self):
        result=subprocess.CompletedProcess([],0,stdout=b'{"ok":false,"code":"not_running"}',stderr=b'')
        with patch.object(bridge,'powershell_path',return_value=Path('powershell.exe')),patch.object(bridge.subprocess,'run',return_value=result),self.assertRaises(bridge.OutlookError) as raised:
            bridge.open_draft('Subject','Body')
        self.assertIn('Open Outlook Classic',str(raised.exception))

    def test_timeout_keeps_draft_safe(self):
        with patch.object(bridge,'powershell_path',return_value=Path('powershell.exe')),patch.object(bridge.subprocess,'run',side_effect=subprocess.TimeoutExpired('powershell.exe',45)),self.assertRaises(bridge.OutlookError) as raised:
            bridge.open_draft('Subject','Body')
        self.assertIn('draft text is still here',str(raised.exception))

    def test_import_copy_is_removed_after_loading(self):
        paths=[]
        def export_fixture(script,payload=None,environment=None):
            path=Path(environment['DBN_EMAIL_OUTLOOK_FILE']);paths.append(path)
            path.write_bytes(b'fictional msg export')
            return {'ok':True,'entry_id':'entry','store_id':'store','subject':'Test: email'}
        with patch.object(bridge,'run_script',side_effect=export_fixture):
            result=bridge.import_selected_email({'entry_id':'chosen-entry','store_id':'chosen-store'})
        self.assertEqual(base64.b64decode(result['file']['data']),b'fictional msg export')
        self.assertEqual(result['file']['name'],'Test_ email.msg')
        self.assertFalse(paths[0].exists())
        self.assertFalse(paths[0].parent.exists())

    @unittest.skipUnless(sys.platform=='win32','Windows PowerShell syntax validation')
    def test_native_powershell_parses_both_scripts(self):
        # Parse only; do not connect to Outlook or touch a mailbox in the tests.
        parser=r'$text=[Console]::In.ReadToEnd(); $tokens=$null; $errors=$null; [void][System.Management.Automation.Language.Parser]::ParseInput($text,[ref]$tokens,[ref]$errors); if($errors.Count){exit 1}'
        encoded=base64.b64encode(parser.encode('utf-16-le')).decode()
        for script in (bridge.IMPORT_SCRIPT,bridge.DRAFT_SCRIPT,bridge.LIST_SCRIPT):
            result=subprocess.run([str(bridge.powershell_path()),'-NoProfile','-NonInteractive','-EncodedCommand',encoded],input=script.encode('ascii'),capture_output=True,timeout=15,creationflags=subprocess.CREATE_NO_WINDOW)
            self.assertEqual(result.returncode,0,result.stderr.decode(errors='replace'))

    def test_import_requires_explicit_message_and_ignores_outlook_selection(self):
        with patch.object(bridge,'run_script') as run,self.assertRaises(bridge.OutlookError):bridge.import_selected_email()
        run.assert_not_called()
        self.assertIn('GetItemFromID',bridge.IMPORT_SCRIPT)
        self.assertNotIn('ActiveWindow',bridge.IMPORT_SCRIPT)
        self.assertNotIn('.Selection',bridge.IMPORT_SCRIPT)

    def test_listing_passes_search_as_data_and_never_opens_message_bodies(self):
        query="O'Brien $(Get-Secret) <img>"
        with patch.object(bridge,'run_script',return_value={'ok':True}) as run:bridge.list_emails({'entry_id':'folder','store_id':'store'},query,50)
        self.assertEqual(run.call_args.args[1]['query'],query)
        self.assertNotIn(query,run.call_args.args[0])
        self.assertNotIn('.Body',bridge.LIST_SCRIPT);self.assertNotIn('.HTMLBody',bridge.LIST_SCRIPT)
        self.assertNotIn('.Display(',bridge.LIST_SCRIPT);self.assertNotIn('.SaveAs(',bridge.LIST_SCRIPT)
        self.assertNotIn('.UnRead =',bridge.LIST_SCRIPT)
        for query,offset in [('x'*201,0),('',-1),('',True),('',1.5)]:
            with self.subTest(offset=offset),self.assertRaises(bridge.OutlookError):bridge.list_emails(query=query,offset=offset)


class EmailPickerTests(unittest.TestCase):
    def fixture(self):
        return {'folder':{'entry_id':'private-folder','store_id':'private-store','name':'Inbox','path':'Mailbox\\Inbox'},
                'parent':None,'inbox':None,'mailboxes':[],'folders':[],
                'emails':[{'entry_id':'private-email','store_id':'private-store','subject':"O'Brien <img>",'sender':'Sam','received':'2026-10-28T08:00:00','size':1000}],
                'has_more':True}

    def test_headers_use_opaque_ids_and_imports_the_exact_chosen_email(self):
        picker=bridge.EmailPicker()
        with patch.object(bridge,'list_emails',return_value=self.fixture()):result=picker.browse()
        self.assertNotIn('private-email',json.dumps(result));self.assertNotIn('private-folder',json.dumps(result));self.assertNotIn('private-store',json.dumps(result))
        self.assertEqual(result['emails'][0]['subject'],"O'Brien <img>")
        with patch.object(bridge,'import_selected_email',return_value={'file':{}}) as run:
            picker.import_email(result['picker'],result['emails'][0]['id'])
        self.assertEqual(run.call_args.args[0],{'entry_id':'private-email','store_id':'private-store'})

    def test_missing_forged_expired_or_previous_page_selection_cannot_import(self):
        picker=bridge.EmailPicker()
        with patch.object(bridge,'list_emails',return_value=self.fixture()):first=picker.browse()
        with patch.object(bridge,'import_selected_email') as run:
            for token,message in [('',first['emails'][0]['id']),(first['picker'],'private-email'),('forged',first['emails'][0]['id'])]:
                with self.subTest(message=message),self.assertRaises(bridge.OutlookError):picker.import_email(token,message)
            with patch.object(bridge,'list_emails',return_value=self.fixture()):picker.browse(first['picker'],first['folder']['id'],offset=50)
            with self.assertRaises(bridge.OutlookError):picker.import_email(first['picker'],first['emails'][0]['id'])
            picker.expires=0
            with self.assertRaises(bridge.OutlookError):picker.browse(first['picker'])
        run.assert_not_called()

    def test_folder_navigation_passes_only_registered_outlook_ids(self):
        picker=bridge.EmailPicker()
        with patch.object(bridge,'list_emails',return_value=self.fixture()) as run:
            first=picker.browse()
            picker.browse(first['picker'],first['folder']['id'],"Sam's update",50)
        self.assertEqual(run.call_args.args,({'entry_id':'private-folder','store_id':'private-store'},"Sam's update",50))
        with patch.object(bridge,'list_emails') as run,self.assertRaises(bridge.OutlookError):picker.browse(first['picker'],'private-folder')
        run.assert_not_called()


if __name__=='__main__':unittest.main(verbosity=2)
