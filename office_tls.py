"""Loopback-only HTTPS for the personal Outlook add-in. Never installs trust."""
import base64
from datetime import datetime, timedelta, timezone
import hashlib
import ipaddress
from pathlib import Path
import ssl
import tempfile
import threading

from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.x509.oid import NameOID, ExtendedKeyUsageOID

PORT = 8788


class OfficeSetupError(Exception):
    pass


def prepare(store):
    saved = store.read('office_tls', {})
    if not saved:
        key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
        name = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, 'DBN Email Studio localhost')])
        now = datetime.now(timezone.utc)
        cert = (x509.CertificateBuilder().subject_name(name).issuer_name(name).public_key(key.public_key())
                .serial_number(x509.random_serial_number()).not_valid_before(now-timedelta(minutes=5))
                .not_valid_after(now+timedelta(days=365))
                .add_extension(x509.SubjectAlternativeName([x509.DNSName('localhost'),x509.IPAddress(ipaddress.ip_address('127.0.0.1'))]),False)
                .add_extension(x509.BasicConstraints(ca=False,path_length=None),True)
                .add_extension(x509.ExtendedKeyUsage([ExtendedKeyUsageOID.SERVER_AUTH]),False)
                .sign(key,hashes.SHA256()))
        saved = {'key':base64.b64encode(key.private_bytes(serialization.Encoding.PEM,serialization.PrivateFormat.PKCS8,serialization.NoEncryption())).decode(),
                 'cert':base64.b64encode(cert.public_bytes(serialization.Encoding.PEM)).decode(),
                 'der':base64.b64encode(cert.public_bytes(serialization.Encoding.DER)).decode()}
        store.write('office_tls',saved)
    store.directory.mkdir(parents=True,exist_ok=True)
    public = store.directory/'DBN Email Studio localhost.cer'
    public.write_bytes(base64.b64decode(saved['der']))
    return saved, public


def public_status(server):
    saved = server.store.read('office_tls', {})
    return {'prepared':bool(saved),'running':bool(getattr(server,'office_server',None) or getattr(server,'secure',False)),
            'url':f'https://localhost:{PORT}/office/taskpane.html',
            'certificate_path':str(server.store.directory/'DBN Email Studio localhost.cer'),
            'sha256':hashlib.sha256(base64.b64decode(saved['der'])).hexdigest().upper() if saved else '',
            'trust_installed_by_app':False}


def start(server, factory):
    if getattr(server,'office_server',None):return public_status(server)
    saved, _ = prepare(server.store)
    context=ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    context.minimum_version=ssl.TLSVersion.TLSv1_2
    # OpenSSL needs files; keep the plaintext key only in a short-lived user temp directory.
    with tempfile.TemporaryDirectory(prefix='DBN-Office-TLS-') as directory:
        key=Path(directory)/'key.pem';cert=Path(directory)/'cert.pem'
        key.write_bytes(base64.b64decode(saved['key']));cert.write_bytes(base64.b64decode(saved['cert']))
        context.load_cert_chain(cert,key)
    try: secure=factory(('127.0.0.1',PORT),server.store,server.token)
    except OSError:raise OfficeSetupError('The local Outlook add-in port is in use. Close the other local add-in server and try again.') from None
    secure.secure=True
    secure.chatgpt=server.chatgpt
    secure.generating=server.generating
    secure.outlook_lock=server.outlook_lock
    secure.outlook_sources=server.outlook_sources
    secure.outlook_picker=server.outlook_picker
    secure.ssl_context=context
    server.office_server=secure
    threading.Thread(target=secure.serve_forever,daemon=True).start()
    return public_status(server)
