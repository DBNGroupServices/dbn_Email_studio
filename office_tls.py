"""Loopback-only HTTPS for the Outlook add-in. Never installs trust itself.

Version 2 uses a small per-user certificate authority (CA) that is valid for ten
years and signs a short-lived localhost server certificate. Windows trusts only
the CA (CurrentUser root store, installed by the setup script after the user
approves its fingerprint). The server certificate renews itself automatically
well before expiry, so the add-in no longer stops loading when a certificate
lapses — previously a common cause of Outlook's "can't load the add-in" error.

Private keys are stored only inside the DPAPI-encrypted app store. The CA key is
constrained to signing localhost/127.0.0.1 server certificates (name
constraints), so trusting it cannot be used to impersonate other websites.
"""
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
STORE_KEY = 'office_tls_v2'
LEGACY_KEY = 'office_tls'
CA_DAYS = 3650
LEAF_DAYS = 397          # Browser/WebView maximum for server certificates.
RENEW_BEFORE_DAYS = 45   # Renew silently when this close to expiry.
CA_NAME = 'DBN Email Studio Local CA'
CA_FILE = 'DBN Email Studio Local CA.cer'


class OfficeSetupError(Exception):
    pass


def _now():
    return datetime.now(timezone.utc)


def _pem_key(key):
    return base64.b64encode(key.private_bytes(serialization.Encoding.PEM, serialization.PrivateFormat.PKCS8,
                                              serialization.NoEncryption())).decode()


def _load_key(encoded):
    return serialization.load_pem_private_key(base64.b64decode(encoded), password=None)


def _cert(encoded):
    return x509.load_pem_x509_certificate(base64.b64decode(encoded))


def _make_ca():
    key = rsa.generate_private_key(public_exponent=65537, key_size=3072)
    name = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, CA_NAME),
                      x509.NameAttribute(NameOID.ORGANIZATION_NAME, 'DBN Email Studio (this computer only)')])
    now = _now()
    permitted = [x509.DNSName('localhost'), x509.IPAddress(ipaddress.ip_network('127.0.0.1/32'))]
    cert = (x509.CertificateBuilder().subject_name(name).issuer_name(name).public_key(key.public_key())
            .serial_number(x509.random_serial_number())
            .not_valid_before(now - timedelta(minutes=5)).not_valid_after(now + timedelta(days=CA_DAYS))
            .add_extension(x509.BasicConstraints(ca=True, path_length=0), critical=True)
            .add_extension(x509.KeyUsage(digital_signature=False, content_commitment=False, key_encipherment=False,
                                         data_encipherment=False, key_agreement=False, key_cert_sign=True,
                                         crl_sign=True, encipher_only=False, decipher_only=False), critical=True)
            .add_extension(x509.NameConstraints(permitted_subtrees=permitted, excluded_subtrees=None), critical=True)
            .add_extension(x509.SubjectKeyIdentifier.from_public_key(key.public_key()), critical=False)
            .sign(key, hashes.SHA256()))
    return key, cert


def _make_leaf(ca_key, ca_cert):
    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    now = _now()
    cert = (x509.CertificateBuilder()
            .subject_name(x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, 'localhost')]))
            .issuer_name(ca_cert.subject).public_key(key.public_key())
            .serial_number(x509.random_serial_number())
            .not_valid_before(now - timedelta(minutes=5)).not_valid_after(now + timedelta(days=LEAF_DAYS))
            .add_extension(x509.SubjectAlternativeName([x509.DNSName('localhost'),
                                                        x509.IPAddress(ipaddress.ip_address('127.0.0.1'))]), False)
            .add_extension(x509.BasicConstraints(ca=False, path_length=None), True)
            .add_extension(x509.KeyUsage(digital_signature=True, content_commitment=False, key_encipherment=True,
                                         data_encipherment=False, key_agreement=False, key_cert_sign=False,
                                         crl_sign=False, encipher_only=False, decipher_only=False), True)
            .add_extension(x509.ExtendedKeyUsage([ExtendedKeyUsageOID.SERVER_AUTH]), False)
            .add_extension(x509.AuthorityKeyIdentifier.from_issuer_public_key(ca_key.public_key()), False)
            .sign(ca_key, hashes.SHA256()))
    return key, cert


def _expiry(cert):
    value = getattr(cert, 'not_valid_after_utc', None)
    return value if value is not None else cert.not_valid_after.replace(tzinfo=timezone.utc)


def prepare(store, now=None):
    """Create or renew certificates. Returns (saved, ca_certificate_path)."""
    now = now or _now()
    saved = store.read(STORE_KEY, {})
    changed = False
    if not saved.get('ca_key') or _expiry(_cert(saved['ca_cert'])) - now < timedelta(days=RENEW_BEFORE_DAYS):
        ca_key, ca_cert = _make_ca()
        saved = {'ca_key': _pem_key(ca_key), 'ca_cert': base64.b64encode(ca_cert.public_bytes(serialization.Encoding.PEM)).decode()}
        changed = True
    ca_key = _load_key(saved['ca_key'])
    ca_cert = _cert(saved['ca_cert'])
    if changed or not saved.get('leaf_key') or _expiry(_cert(saved['leaf_cert'])) - now < timedelta(days=RENEW_BEFORE_DAYS):
        leaf_key, leaf_cert = _make_leaf(ca_key, ca_cert)
        saved['leaf_key'] = _pem_key(leaf_key)
        saved['leaf_cert'] = base64.b64encode(leaf_cert.public_bytes(serialization.Encoding.PEM)).decode()
        changed = True
    if changed:
        store.write(STORE_KEY, saved)
    store.directory.mkdir(parents=True, exist_ok=True)
    public = store.directory / CA_FILE
    der = ca_cert.public_bytes(serialization.Encoding.DER)
    if not public.exists() or public.read_bytes() != der:
        public.write_bytes(der)
    return saved, public


def _context(saved):
    context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    context.minimum_version = ssl.TLSVersion.TLSv1_2
    # OpenSSL needs files; keep the plaintext key only in a short-lived user temp directory.
    with tempfile.TemporaryDirectory(prefix='DBN-Office-TLS-') as directory:
        key = Path(directory) / 'key.pem'
        chain = Path(directory) / 'chain.pem'
        key.write_bytes(base64.b64decode(saved['leaf_key']))
        chain.write_bytes(base64.b64decode(saved['leaf_cert']) + base64.b64decode(saved['ca_cert']))
        context.load_cert_chain(chain, key)
    return context


def public_status(server):
    saved = server.store.read(STORE_KEY, {})
    status = {'prepared': bool(saved), 'running': bool(getattr(server, 'office_server', None) or getattr(server, 'secure', False)),
              'url': f'https://localhost:{PORT}/office/taskpane.html', 'port': PORT,
              'certificate_path': str(server.store.directory / CA_FILE),
              'sha256': '', 'thumbprint': '', 'server_certificate_expires': '', 'ca_expires': '',
              'legacy_certificate_present': bool(server.store.read(LEGACY_KEY, {})),
              'trust_installed_by_app': False}
    if saved:
        ca = _cert(saved['ca_cert'])
        der = ca.public_bytes(serialization.Encoding.DER)
        status.update(sha256=hashlib.sha256(der).hexdigest().upper(), thumbprint=hashlib.sha1(der).hexdigest().upper(),
                      ca_expires=_expiry(ca).isoformat(), server_certificate_expires=_expiry(_cert(saved['leaf_cert'])).isoformat())
    return status


def _renewal_loop(server, stop):
    """Check once a day; swap the TLS context in place when the leaf renews."""
    while not stop.wait(24 * 3600):
        try:
            before = server.store.read(STORE_KEY, {}).get('leaf_cert')
            saved, _ = prepare(server.store)
            if saved.get('leaf_cert') != before and server.office_server is not None:
                server.office_server.ssl_context = _context(saved)
        except Exception:  # Never stop serving because of a failed background check.
            pass


def start(server, factory):
    if getattr(server, 'office_server', None):
        return public_status(server)
    saved, _ = prepare(server.store)
    context = _context(saved)
    try:
        secure = factory(('127.0.0.1', PORT), server.store, server.token)
    except OSError:
        raise OfficeSetupError('The Outlook add-in port (8788) is in use. Close any other copy of DBN Email Studio and try again.') from None
    secure.secure = True
    for name in ('chatgpt', 'generating', 'outlook_lock', 'outlook_sources', 'outlook_picker'):
        setattr(secure, name, getattr(server, name))
    secure.ssl_context = context
    server.office_server = secure
    stop = threading.Event()
    server.office_renewal_stop = stop
    threading.Thread(target=secure.serve_forever, daemon=True).start()
    threading.Thread(target=_renewal_loop, args=(server, stop), daemon=True).start()
    return public_status(server)
