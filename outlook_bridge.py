"""On-demand Outlook Classic integration. No mailbox polling and no sending."""
from __future__ import annotations

import base64
import json
import os
from pathlib import Path
import re
import secrets
import subprocess
import sys
import tempfile
import time


class OutlookError(Exception):
    pass


PREAMBLE = r'''
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$utf8 = New-Object System.Text.UTF8Encoding($false)
[Console]::InputEncoding = $utf8
[Console]::OutputEncoding = $utf8
function Result($value) { [Console]::Out.WriteLine(($value | ConvertTo-Json -Compress -Depth 6)) }
try { $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') }
catch { Result @{ok=$false; code='not_running'}; exit 0 }
'''

IMPORT_SCRIPT = PREAMBLE + r'''
$item = $null; $session = $null; $parent = $null
try {
    $payload = [Console]::In.ReadToEnd() | ConvertFrom-Json
    if (!$payload.entry_id -or !$payload.store_id) { Result @{ok=$false; code='choose_email'}; exit 0 }
    $session = $outlook.Session
    try { $item = $session.GetItemFromID($payload.entry_id, $payload.store_id) }
    catch { Result @{ok=$false; code='source_unavailable'}; exit 0 }
    if ($null -eq $item -or $item.Class -ne 43) { Result @{ok=$false; code='not_email'}; exit 0 }
    if ($item.Size -gt 10485760) { Result @{ok=$false; code='too_large'}; exit 0 }
    $item.SaveAs($env:DBN_EMAIL_OUTLOOK_FILE, 9)
    $storeId = ''
    try { $parent = $item.Parent; $storeId = [string]$parent.StoreID } catch {}
    Result @{ok=$true; entry_id=[string]$item.EntryID; store_id=$storeId; subject=[string]$item.Subject}
}
catch { Result @{ok=$false; code='import_failed'} }
finally {
    foreach ($obj in @($parent,$item,$session,$outlook)) {
        if ($null -ne $obj -and [Runtime.InteropServices.Marshal]::IsComObject($obj)) {
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } catch {}
        }
    }
}
'''

LIST_SCRIPT = PREAMBLE + r"""
function Release($obj) {
    if ($null -ne $obj -and [Runtime.InteropServices.Marshal]::IsComObject($obj)) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } catch {}
    }
}
function FolderInfo($value) {
    if ($null -eq $value) { return $null }
    return @{entry_id=[string]$value.EntryID; store_id=[string]$value.StoreID; name=[string]$value.Name; path=[string]$value.FolderPath}
}
function Cell($row, $name, $fallback) {
    try { $value=$row.Item($name); if ($null -eq $value -or $value -is [DBNull]) { return $fallback }; return $value }
    catch { return $fallback }
}
$session=$null; $folder=$null; $explorer=$null; $parent=$null; $stores=$null
$currentStore=$null; $inbox=$null; $children=$null; $table=$null; $columns=$null
try {
    $payload = [Console]::In.ReadToEnd() | ConvertFrom-Json
    $session = $outlook.Session
    if ($payload.entry_id) {
        try { $folder = $session.GetFolderFromID($payload.entry_id, $payload.store_id) }
        catch { Result @{ok=$false; code='folder_unavailable'}; exit 0 }
    } else {
        try { $explorer=$outlook.ActiveExplorer(); if ($null -ne $explorer) { $folder=$explorer.CurrentFolder } } catch {}
        if ($null -eq $folder -or $folder.DefaultItemType -ne 0) { Release $folder; $folder=$session.GetDefaultFolder(6) }
    }
    $roots = New-Object 'System.Collections.Generic.List[object]'
    $stores = $session.Stores
    foreach ($mailStore in $stores) {
        $root=$null
        try { $root=$mailStore.GetRootFolder(); $roots.Add((FolderInfo $root)) }
        catch {} finally { Release $root; Release $mailStore }
        if ($roots.Count -ge 100) { break }
    }
    $subfolders = New-Object 'System.Collections.Generic.List[object]'
    $children=$folder.Folders
    foreach ($child in $children) {
        try { if ($child.DefaultItemType -eq 0) { $subfolders.Add((FolderInfo $child)) } }
        finally { Release $child }
        if ($subfolders.Count -ge 250) { break }
    }
    $parentInfo=$null
    try { $parent=$folder.Parent; if ($parent.Class -eq 2) { $parentInfo=FolderInfo $parent } } catch {}
    try { $currentStore=$folder.Store; $inbox=$currentStore.GetDefaultFolder(6) } catch {}
    $filter = '@SQL="http://schemas.microsoft.com/mapi/proptag/0x001A001F" LIKE ''IPM.Note%'''
    if ($payload.query) {
        $term=([string]$payload.query).Replace("'", "''")
        $filter += ' AND ("http://schemas.microsoft.com/mapi/proptag/0x0037001F" LIKE ''%' + $term + '%'' OR "http://schemas.microsoft.com/mapi/proptag/0x0C1A001F" LIKE ''%' + $term + '%'')'
    }
    $table=$folder.GetTable($filter, 0)
    $columns=$table.Columns
    [void]$columns.RemoveAll()
    foreach ($name in @('EntryID','Subject','SenderName','ReceivedTime','UnRead','Size','http://schemas.microsoft.com/mapi/proptag/0x0E1B000B')) { [void]$columns.Add($name) }
    [void]$table.Sort('[ReceivedTime]', $true)
    $remaining=[int]$payload.offset
    while ($remaining -gt 0 -and !$table.EndOfTable) {
        $chunk=[Math]::Min(500,$remaining); [void]$table.GetArray($chunk); $remaining-=$chunk
    }
    $messages=New-Object 'System.Collections.Generic.List[object]'
    while (!$table.EndOfTable -and $messages.Count -lt 50) {
        $row=$null
        try {
            $row=$table.GetNextRow()
            $received=''
            try { $value=Cell $row 'ReceivedTime' $null; if ($null -ne $value) { $received=([datetime]$value).ToString('o') } } catch {}
            $messages.Add(@{entry_id=[string](Cell $row 'EntryID' ''); store_id=[string]$folder.StoreID;
                subject=[string](Cell $row 'Subject' ''); sender=[string](Cell $row 'SenderName' ''); received=$received;
                unread=[bool](Cell $row 'UnRead' $false); size=[long](Cell $row 'Size' 0);
                has_attachments=[bool](Cell $row 'http://schemas.microsoft.com/mapi/proptag/0x0E1B000B' $false)})
        } finally { Release $row }
    }
    Result @{ok=$true; folder=(FolderInfo $folder); parent=$parentInfo; inbox=(FolderInfo $inbox);
        mailboxes=@($roots.ToArray()); folders=@($subfolders.ToArray()); emails=@($messages.ToArray()); has_more=(!$table.EndOfTable)}
}
catch { Result @{ok=$false; code='browse_failed'} }
finally { foreach ($obj in @($columns,$table,$children,$inbox,$currentStore,$parent,$folder,$explorer,$stores,$session,$outlook)) { Release $obj } }
"""

DRAFT_SCRIPT = PREAMBLE + r'''
$original = $null; $draft = $null; $session = $null
try {
    $payload = [Console]::In.ReadToEnd() | ConvertFrom-Json
    $action = 'new'
    if ($payload.entry_id) {
        $session = $outlook.Session
        try { $original = $session.GetItemFromID($payload.entry_id, $payload.store_id) }
        catch { Result @{ok=$false; code='source_unavailable'}; exit 0 }
        if ($null -eq $original -or $original.Class -ne 43) { Result @{ok=$false; code='source_unavailable'}; exit 0 }
        $draft = $original.Reply()
        $action = 'reply'
    } else { $draft = $outlook.CreateItem(0) }
    $draft.Subject = [string]$payload.subject
    [void]$draft.Display($false)
    $text = [System.Net.WebUtility]::HtmlEncode([string]$payload.body)
    $text = $text -replace "\r?\n", '<br>'
    $fragment = '<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt">' + $text + '</div><br>'
    $existing = [string]$draft.HTMLBody
    $bodyTag = [regex]::Match($existing, '(?is)<body\b[^>]*>')
    if ($bodyTag.Success) { $draft.HTMLBody = $existing.Insert($bodyTag.Index + $bodyTag.Length, $fragment) }
    else { $draft.HTMLBody = $fragment + $existing }
    Result @{ok=$true; action=$action}
}
catch { Result @{ok=$false; code='draft_failed'} }
finally {
    foreach ($obj in @($draft,$original,$session,$outlook)) {
        if ($null -ne $obj -and [Runtime.InteropServices.Marshal]::IsComObject($obj)) {
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } catch {}
        }
    }
}
'''

ERRORS = {
    'not_running': 'Open Outlook Classic, then try again. This connection does not support the new Outlook app.',
    'choose_email': 'Choose an email in the import picker first.',
    'folder_unavailable': 'That Outlook folder is no longer available. Reopen the email picker and choose another folder.',
    'browse_failed': 'Outlook could not list this folder. Check for an Outlook prompt, then try another folder or save the email as MSG.',
    'select_one': 'Select exactly one email in the Outlook Classic message list, or open the email in its own window, then try again.',
    'not_email': 'The selected Outlook item is not an email. Select one email and try again.',
    'too_large': 'The selected email is over 10 MB. Save a smaller copy or attach only the relevant files.',
    'import_failed': 'Outlook could not provide the selected email. Check for an Outlook prompt, or save the email as MSG and attach it. Protected messages may restrict export.',
    'source_unavailable': 'The imported Outlook email is no longer available in its original location. Import it again, or copy your draft into Outlook.',
    'draft_failed': 'Outlook could not finish opening the draft. Check any Outlook prompt and open compose window. Your email text is still here; you can copy it instead.'
}


def powershell_path():
    path = Path(os.environ.get('SystemRoot', r'C:\Windows')) / 'System32' / 'WindowsPowerShell' / 'v1.0' / 'powershell.exe'
    if sys.platform != 'win32' or not path.is_file():
        raise OutlookError('The Outlook Classic connection requires Windows and Windows PowerShell.')
    return path


def run_script(script, payload=None, environment=None):
    command = [str(powershell_path()), '-NoLogo', '-NoProfile', '-NonInteractive', '-STA', '-WindowStyle', 'Hidden',
               '-EncodedCommand', base64.b64encode(script.encode('utf-16-le')).decode('ascii')]
    try:
        result = subprocess.run(command, input=json.dumps(payload or {}, ensure_ascii=True).encode('ascii'),
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=45, env=environment,
            creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
        if result.returncode:
            raise OutlookError('The Outlook Classic connection could not start. Your organisation may restrict local Office automation. Save the email as MSG and attach it instead.')
        decoded = json.loads(result.stdout.decode('utf-8-sig').strip())
        if not isinstance(decoded, dict):
            raise ValueError()
        if not decoded.get('ok'):
            raise OutlookError(ERRORS.get(decoded.get('code'), 'Outlook could not complete this action. Your email text is still here.'))
        return decoded
    except subprocess.TimeoutExpired:
        raise OutlookError('Outlook did not respond in time. Check Outlook for a prompt, then try again. Your draft text is still here.') from None
    except (OSError, UnicodeError, ValueError):
        raise OutlookError('Could not read the Outlook Classic connection response. Save the email as MSG and attach it instead.') from None


def import_selected_email(source=None):
    if not source or not source.get('entry_id') or not source.get('store_id'):
        raise OutlookError('Choose an email in the import picker first.')
    # One temporary export, deleted immediately after reading; never a mailbox scan.
    with tempfile.TemporaryDirectory(prefix='DBN-Email-Outlook-') as directory:
        target = Path(directory) / 'selected-email.msg'
        environment = os.environ.copy()
        environment['DBN_EMAIL_OUTLOOK_FILE'] = str(target)
        metadata = run_script(IMPORT_SCRIPT, source, environment=environment)
        if not target.is_file() or not 0 < target.stat().st_size <= 10 * 1024 * 1024:
            raise OutlookError('Outlook did not provide a readable email under 10 MB. Save it as MSG and attach it instead.')
        data = target.read_bytes()
    subject = str(metadata.get('subject', '')).strip()
    filename = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', subject).strip(' .')[:140] or 'Outlook email'
    return {'file': {'name': filename + '.msg', 'data': base64.b64encode(data).decode('ascii')},
            'entry_id': str(metadata.get('entry_id', '')), 'store_id': str(metadata.get('store_id', ''))}


def list_emails(folder=None, query='', offset=0):
    if not isinstance(query, str) or len(query) > 200 or not isinstance(offset, int) or isinstance(offset, bool) or not 0 <= offset <= 1000000:
        raise OutlookError('Enter a shorter search or reopen the email picker.')
    return run_script(LIST_SCRIPT, {**(folder or {}), 'query':query, 'offset':offset})


class EmailPicker:
    """Short-lived, local metadata references; Outlook IDs never enter the UI."""
    def __init__(self):
        self.token = ''
        self.expires = 0
        self.folders = {}
        self.messages = {}

    def validate(self, token):
        if not token or not secrets.compare_digest(self.token, token) or self.expires < time.monotonic():
            raise OutlookError('This email picker has expired. Close it and choose Import from Outlook Classic again.')

    def browse(self, token='', folder='', query='', offset=0):
        if not token:
            self.token=secrets.token_urlsafe(24);self.folders={};self.messages={}
            self.expires=time.monotonic()+1200
        else:self.validate(token)
        reference=self.folders.get(folder) if folder else None
        if folder and reference is None:
            raise OutlookError('Choose a folder from the email picker.')
        result=list_emails(reference,query,offset)
        def folder_info(info):
            if not info or not info.get('entry_id') or not info.get('store_id'):return None
            identity={k:info[k] for k in ('entry_id','store_id')}
            key=next((k for k,v in self.folders.items() if v==identity),None)
            if not key:
                if len(self.folders)>=2048:raise OutlookError('Reopen the email picker to continue browsing folders.')
                key=secrets.token_urlsafe(24);self.folders[key]=identity
            return {'id':key,'name':str(info.get('name',''))[:300],'path':str(info.get('path',''))[:1500]}
        messages={};rows=[]
        for item in result.get('emails',[])[:50]:
            if not item.get('entry_id') or not item.get('store_id'):continue
            key=secrets.token_urlsafe(24)
            messages[key]={k:item[k] for k in ('entry_id','store_id')}
            rows.append({'id':key,'subject':str(item.get('subject',''))[:1000],'sender':str(item.get('sender',''))[:500],
                         'received':str(item.get('received',''))[:80],'size':int(item.get('size',0)),
                         'unread':bool(item.get('unread')),'has_attachments':bool(item.get('has_attachments'))})
        self.messages=messages
        return {'picker':self.token,'folder':folder_info(result.get('folder')),'parent':folder_info(result.get('parent')),
                'inbox':folder_info(result.get('inbox')),'mailboxes':[x for x in (folder_info(i) for i in result.get('mailboxes',[])[:100]) if x],
                'folders':[x for x in (folder_info(i) for i in result.get('folders',[])[:250]) if x],
                'emails':rows,'offset':offset,'has_more':bool(result.get('has_more')),'query':query}

    def import_email(self, token, message):
        self.validate(token)
        reference=self.messages.get(message)
        if not reference:raise OutlookError('Select an email from the current page before importing.')
        return import_selected_email(reference)


def open_draft(subject, body, source=None):
    return run_script(DRAFT_SCRIPT, {'subject': subject, 'body': body, 'entry_id': (source or {}).get('entry_id', ''),
                                    'store_id': (source or {}).get('store_id', '')})
