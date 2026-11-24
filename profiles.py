"""Voice profiles: reusable "who am I writing as, and how" presets.

A profile bundles tone controls (formality, warmth, directness), greeting and
sign-off habits, an optional signature, rules to follow or avoid, and up to three
example emails the AI imitates for rhythm and vocabulary. Profiles are stored
encrypted with the rest of the app data.
"""
import uuid

MAX_PROFILES = 30
LEVELS = ('formality', 'warmth', 'directness')
LENGTHS = ('Brief', 'Standard', 'Detailed')
FORMATS = ('Paragraphs', 'Bullet points')
TEXT_LIMITS = {'name': 80, 'purpose': 300, 'guidance': 4000, 'greeting': 200, 'signoff': 200,
               'signature': 2000, 'avoid': 2000}

DEFAULT_PROFILES = [
    {'id': 'builtin-professional', 'name': 'Professional (default)', 'purpose': 'Everyday business correspondence.',
     'guidance': 'Clear, courteous and efficient. Lead with the answer or next step. Short paragraphs.',
     'greeting': 'Hi {first name},', 'signoff': 'Kind regards,', 'formality': 3, 'warmth': 3, 'directness': 4,
     'length': 'Standard', 'format': 'Paragraphs', 'avoid': 'Corporate jargon, over-apologising, exclamation marks.'},
    {'id': 'builtin-tenant', 'name': 'Tenant & occupier relations', 'purpose': 'Replies to tenants, occupiers and visitors.',
     'guidance': 'Service-focused and reassuring. Acknowledge the concern, explain what will happen and when, and make it easy to reply. '
                 'Avoid internal jargon and avoid committing to outcomes the notes do not confirm.',
     'greeting': 'Hi {first name},', 'signoff': 'Kind regards,', 'formality': 3, 'warmth': 4, 'directness': 3,
     'length': 'Standard', 'format': 'Paragraphs', 'avoid': 'Blame, speculation about causes, internal procedure detail.'},
    {'id': 'builtin-contractor', 'name': 'Contractors & suppliers — firm', 'purpose': 'Instructions, compliance and performance matters.',
     'guidance': 'Firm, specific and unambiguous. State requirements, deadlines and consequences only where supplied. '
                 'Use numbered steps for instructions. Polite but not soft.',
     'greeting': 'Hi {first name},', 'signoff': 'Regards,', 'formality': 4, 'warmth': 2, 'directness': 5,
     'length': 'Standard', 'format': 'Paragraphs', 'avoid': 'Hedging language, vague timeframes.'},
    {'id': 'builtin-executive', 'name': 'Executive / client briefing', 'purpose': 'Updates to senior stakeholders, owners and clients.',
     'guidance': 'Bottom line first, then key facts, risk and recommended decision. Concise, confident and neutral. '
                 'Use short bullet points for multiple facts.',
     'greeting': 'Hi {first name},', 'signoff': 'Kind regards,', 'formality': 4, 'warmth': 2, 'directness': 5,
     'length': 'Brief', 'format': 'Paragraphs', 'avoid': 'Operational minutiae, emotive language.'},
    {'id': 'builtin-team', 'name': 'Internal team — direct', 'purpose': 'Quick notes to staff and colleagues.',
     'guidance': 'Friendly, plain and direct. Get to the point; clear actions and owners.',
     'greeting': 'Hi {first name},', 'signoff': 'Thanks,', 'formality': 2, 'warmth': 4, 'directness': 4,
     'length': 'Brief', 'format': 'Paragraphs', 'avoid': ''},
]


class ProfileError(ValueError):
    pass


def _text(data, key):
    value = data.get(key, '')
    if value is None:
        value = ''
    if not isinstance(value, str) or len(value) > TEXT_LIMITS[key]:
        raise ProfileError(f'Profile {key} must be text under {TEXT_LIMITS[key]:,} characters.')
    return value.strip()


def _level(value, default=3):
    if value in (None, ''):
        return default
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not 1 <= value <= 5:
        raise ProfileError('Tone controls must be between 1 and 5.')
    return int(round(value))


def validate(data):
    """Return a clean profile dict or raise ProfileError (subclass of ValueError)."""
    if not isinstance(data, dict):
        raise ProfileError('Invalid profile.')
    profile = {key: _text(data, key) for key in TEXT_LIMITS}
    if not profile['name']:
        raise ProfileError('Give the profile a name.')
    ident = data.get('id') or ''
    if not isinstance(ident, str) or len(ident) > 64 or (ident and not all(c.isalnum() or c in '-_' for c in ident)):
        raise ProfileError('Invalid profile.')
    profile['id'] = ident or uuid.uuid4().hex
    for key in LEVELS:
        profile[key] = _level(data.get(key))
    length = data.get('length') or 'Standard'
    fmt = data.get('format') or 'Paragraphs'
    if length not in LENGTHS or fmt not in FORMATS:
        raise ProfileError('Choose a valid length and layout.')
    profile['length'], profile['format'] = length, fmt
    examples = data.get('examples') or []
    if not isinstance(examples, list) or len(examples) > 3 or not all(isinstance(e, str) and len(e) <= 5000 for e in examples):
        raise ProfileError('Add up to three example emails of up to 5,000 characters each.')
    profile['examples'] = [e.strip() for e in examples if e.strip()]
    profile['include_signature'] = data.get('include_signature', True) is not False
    return profile


def load(store):
    saved = store.read('profiles', None)
    if saved is None:
        return [dict(p, examples=[], include_signature=True, signature='') for p in DEFAULT_PROFILES]
    return saved


def find(store, ident):
    if not ident:
        return None
    return next((p for p in load(store) if p.get('id') == ident), None)


def save(store, data):
    profile = validate(data)
    records = load(store)
    for index, existing in enumerate(records):
        if existing['id'] == profile['id']:
            records[index] = profile
            break
    else:
        if len(records) >= MAX_PROFILES:
            raise ProfileError(f'You can save up to {MAX_PROFILES} profiles.')
        records.append(profile)
    store.write('profiles', records)
    return records


def delete(store, ident):
    records = [p for p in load(store) if p.get('id') != ident]
    store.write('profiles', records)
    return records


def sliders(value, profile):
    base = {key: profile.get(key, 3) for key in LEVELS}
    if value is None:
        return base
    if not isinstance(value, dict):
        raise ProfileError('Invalid tone controls.')
    return {key: _level(value.get(key), base[key]) for key in LEVELS}


def prompt_view(profile):
    view = {k: profile[k] for k in ('name', 'purpose', 'guidance', 'greeting', 'signoff', 'avoid', 'length', 'format') if profile.get(k)}
    if profile.get('examples'):
        view['example_emails'] = profile['examples']
    return view
