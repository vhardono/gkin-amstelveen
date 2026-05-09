"""
Microsoft Graph email reader for GKIN collecte information.
Uses OAuth2 device-code flow — works with free Outlook.com accounts.

Known email sources:
  Tikkie collecte : from fokkedj@gmail.com, subject "Tikkie Collecte <date>"
  OLE QR code     : from pmcvb.gkin@gmail.com, subject "QR Code OLE <date>"

Required environment variable (add to .env):
    IMAP_EMAIL  – your Outlook.com address, e.g. vega@outlook.com
"""

import json
import os
import re
import tempfile
from datetime import datetime, timedelta
from typing import Any, Dict, Optional

import msal
import requests

# Microsoft's well-known public client ID for Office apps — works for personal accounts
_PUBLIC_CLIENT_ID = '9e5f94bc-e8a4-4e73-b8be-63364c29d753'
_AUTHORITY        = 'https://login.microsoftonline.com/common'
_SCOPES           = ['Mail.Read']
_GRAPH_BASE       = 'https://graph.microsoft.com/v1.0'

TOKEN_CACHE_PATH  = os.path.join(os.path.dirname(__file__), '..', '.msal_token_cache.json')
UPLOAD_DIR        = os.path.join(os.path.dirname(__file__), '..', 'output', '_uploads')

# In-memory store for device flow (avoids filesystem writes on Railway)
_device_flow_store: Dict[str, Any] = {}

# Keywords to match collecte-related email subjects
COLLECTE_SUBJECT_PAT = re.compile(
    r'(tikkie|collecte|betaalverzoek|donatie|ole|dankoffer|qr)',
    re.IGNORECASE
)
TIKKIE_PAT = re.compile(r'https?://tikkie\.me/\S+', re.IGNORECASE)
ING_PAT    = re.compile(r'https?://\S*(?:ing\.nl|payreq)\S+', re.IGNORECASE)
URL_PAT    = re.compile(r'https?://\S{15,}', re.IGNORECASE)
AMOUNT_PAT = re.compile(r'€\s*(\d{1,6}[.,]\d{2})')


def _load_cache() -> msal.SerializableTokenCache:
    """Load token cache from env var (Railway) or local file (dev)."""
    cache = msal.SerializableTokenCache()
    # Prefer env var — set by Railway as a persistent environment variable
    env_cache = os.getenv('MSAL_TOKEN_CACHE', '').strip()
    if env_cache:
        try:
            cache.deserialize(env_cache)
        except Exception as e:
            print(f"Warning: MSAL_TOKEN_CACHE is malformed, ignoring: {e}")
    elif os.path.exists(TOKEN_CACHE_PATH):
        with open(TOKEN_CACHE_PATH, 'r') as f:
            cache.deserialize(f.read())
    return cache


def _save_cache(cache: msal.SerializableTokenCache):
    """Persist token cache to local file (dev). On Railway, export via /email-token-export."""
    if cache.has_state_changed:
        try:
            os.makedirs(os.path.dirname(os.path.abspath(TOKEN_CACHE_PATH)), exist_ok=True)
            with open(TOKEN_CACHE_PATH, 'w') as f:
                f.write(cache.serialize())
        except OSError:
            pass  # read-only filesystem on Railway — token exported via endpoint instead


def get_token_cache_json() -> str:
    """Return current serialized token cache (for copying into Railway env var)."""
    cache = _load_cache()
    return cache.serialize()


def _make_app(cache: msal.SerializableTokenCache) -> msal.PublicClientApplication:
    return msal.PublicClientApplication(
        _PUBLIC_CLIENT_ID,
        authority=_AUTHORITY,
        token_cache=cache,
    )


class OutlookCollecteReader:
    """Fetch collecte data from Outlook.com via Microsoft Graph OAuth2."""

    def __init__(self):
        self.email_addr = os.getenv('IMAP_EMAIL', '')

    def is_authenticated(self) -> bool:
        """True if a cached token exists and can be refreshed silently."""
        cache = _load_cache()
        app   = _make_app(cache)
        accounts = app.get_accounts()
        if not accounts:
            return False
        result = app.acquire_token_silent(_SCOPES, account=accounts[0])
        _save_cache(cache)
        return result is not None and 'access_token' in result

    def start_device_flow(self) -> Dict[str, str]:
        """Initiate device-code login. Returns {user_code, verification_uri, message}."""
        cache = _load_cache()
        app   = _make_app(cache)
        flow  = app.initiate_device_flow(scopes=_SCOPES)
        if 'user_code' not in flow:
            raise RuntimeError(f"Device flow failed: {flow.get('error_description', flow)}")
        # Store flow in a temp file so acquire_token_by_device_flow can be called later
        _device_flow_store['flow'] = flow
        # Also try file for local dev (ignore errors on Railway)
        try:
            flow_path = os.path.join(os.path.dirname(TOKEN_CACHE_PATH), '.msal_device_flow.json')
            with open(flow_path, 'w') as f:
                json.dump(flow, f)
        except OSError:
            pass
        return {
            'user_code':        flow['user_code'],
            'verification_uri': flow['verification_uri'],
            'message':          flow['message'],
        }

    def complete_device_flow(self) -> bool:
        """Poll for completion of device-code login. Returns True if successful."""
        # Try in-memory first (Railway), then fall back to file (local dev)
        flow = _device_flow_store.get('flow')
        if not flow:
            flow_path = os.path.join(os.path.dirname(TOKEN_CACHE_PATH), '.msal_device_flow.json')
            if not os.path.exists(flow_path):
                return False
            with open(flow_path) as f:
                flow = json.load(f)
        cache = _load_cache()
        app   = _make_app(cache)
        result = app.acquire_token_by_device_flow(flow)
        _save_cache(cache)
        _device_flow_store.pop('flow', None)
        flow_path = os.path.join(os.path.dirname(TOKEN_CACHE_PATH), '.msal_device_flow.json')
        if os.path.exists(flow_path):
            try:
                os.remove(flow_path)
            except OSError:
                pass
        return 'access_token' in result

    def _get_token(self) -> str:
        cache    = _load_cache()
        app      = _make_app(cache)
        accounts = app.get_accounts()
        if not accounts:
            raise ValueError('Niet ingelogd. Gebruik de "Inloggen" knop eerst.')
        result = app.acquire_token_silent(_SCOPES, account=accounts[0])
        _save_cache(cache)
        if not result or 'access_token' not in result:
            raise ValueError('Token verlopen. Log opnieuw in via de "Inloggen" knop.')
        return result['access_token']

    def _graph_get(self, path: str, params: Dict = None) -> Dict:
        token = self._get_token()
        resp  = requests.get(
            f'{_GRAPH_BASE}{path}',
            headers={'Authorization': f'Bearer {token}'},
            params=params,
            timeout=20,
        )
        resp.raise_for_status()
        return resp.json()

    def fetch_collecte_data(self, target_date: datetime = None, since_days: int = 60) -> Dict[str, Any]:
        """Search inbox for Tikkie and OLE emails matching target_date.

        Tikkie: from fokkedj@gmail.com, subject contains 'Tikkie Collecte'
        OLE:    from pmcvb.gkin@gmail.com, subject contains 'QR Code OLE'
        Both also contain an image attachment (QR) and a payment URL.
        """
        import base64

        since = (datetime.utcnow() - timedelta(days=since_days)).strftime('%Y-%m-%dT00:00:00Z')

        result: Dict[str, Any] = {
            'dankoffer_url': '', 'ole_url': '',
            'dankoffer_qr':  '', 'ole_qr':  '',
            'collecte_contant': '', 'collecte_bonnen': '',
            'collecte_bank':    '', 'collecte_tikkie': '',
            'collecte_ole':     '',
            'emails_found': 0,
            'source_subjects': [],
            'not_found': [],
        }

        # Build date variants to match against subject, e.g. "10-5-2026", "10 mei 2026", "10 mei"
        date_variants = []
        if target_date:
            NL_MONTHS = ['','januari','februari','maart','april','mei','juni',
                         'juli','augustus','september','oktober','november','december']
            d, m, y = target_date.day, target_date.month, target_date.year
            date_variants = [
                f"{d}-{m}-{y}",          # 10-5-2026
                f"{d}-{m:02d}-{y}",      # 10-05-2026
                f"{d}/{m}/{y}",           # 10/5/2026
                f"{d} {NL_MONTHS[m]}",    # 10 mei
                f"{d} {NL_MONTHS[m]} {y}", # 10 mei 2026
            ]

        def _date_in_subject(subject: str) -> bool:
            if not date_variants:
                return True  # no date filter — accept all
            sl = subject.lower()
            return any(v.lower() in sl for v in date_variants)

        def _search_messages(subject_keyword: str, sender: str) -> list:
            """Fetch messages matching subject keyword, filter by sender client-side."""
            try:
                data = self._graph_get('/me/messages', params={
                    '$filter':  (
                        f"receivedDateTime ge {since} and "
                        f"contains(subject,'{subject_keyword}')"
                    ),
                    '$top':     10,
                    '$select':  'id,subject,receivedDateTime,from,body,hasAttachments',
                    '$orderby': 'receivedDateTime desc',
                })
                msgs = data.get('value', [])
                # Filter by sender (Graph filter on 'from' is unreliable for personal)
                return [
                    m for m in msgs
                    if sender.lower() in m.get('from', {}).get('emailAddress', {}).get('address', '').lower()
                ]
            except Exception:
                return []

        def _extract_url(body: str) -> str:
            for pat in [TIKKIE_PAT, ING_PAT, URL_PAT]:
                m = pat.search(body)
                if m:
                    return m.group(0).rstrip('.,)>')
            return ''

        def _save_first_image(msg_id: str, prefix: str) -> str:
            try:
                # Fetch attachment list with contentType included
                atts = self._graph_get(
                    f"/me/messages/{msg_id}/attachments",
                    params={'$select': 'id,name,contentType,size,isInline'}
                ).get('value', [])
                for att in atts:
                    ct = att.get('contentType', '')
                    # Accept any image attachment (inline or regular)
                    if not ct.startswith('image/'):
                        continue
                    # Fetch full attachment with contentBytes
                    full = self._graph_get(f"/me/messages/{msg_id}/attachments/{att['id']}")
                    cb = full.get('contentBytes', '')
                    if not cb:
                        continue
                    ext = os.path.splitext(att.get('name', 'qr.png'))[1] or '.png'
                    os.makedirs(UPLOAD_DIR, exist_ok=True)
                    tmp = tempfile.NamedTemporaryFile(
                        delete=False, suffix=ext, dir=UPLOAD_DIR, prefix=prefix + '_'
                    )
                    tmp.write(base64.b64decode(cb))
                    tmp.close()
                    return os.path.basename(tmp.name)  # return filename only
            except Exception as e:
                print(f'[email_reader] _save_first_image error: {e}')
            return ''

        # --- Fetch Tikkie collecte email ---
        tikkie_msgs = _search_messages('Tikkie Collecte', 'fokkedj@gmail.com')
        tikkie_match = next((m for m in tikkie_msgs if _date_in_subject(m.get('subject',''))), None)
        if tikkie_match:
            result['emails_found'] += 1
            result['source_subjects'].append(tikkie_match.get('subject',''))
            body = re.sub(r'<[^>]+>', ' ', tikkie_match.get('body', {}).get('content', ''))
            result['dankoffer_url'] = _extract_url(body)
            # Always try attachments — hasAttachments may miss inline images
            result['dankoffer_qr'] = _save_first_image(tikkie_match['id'], 'dankoffer')
        else:
            result['not_found'].append('Tikkie Collecte e-mail niet gevonden voor deze datum')

        # --- Fetch OLE QR email ---
        ole_msgs = _search_messages('QR Code OLE', 'pmcvb.gkin@gmail.com')
        ole_match = next((m for m in ole_msgs if _date_in_subject(m.get('subject',''))), None)
        if ole_match:
            result['emails_found'] += 1
            result['source_subjects'].append(ole_match.get('subject',''))
            body = re.sub(r'<[^>]+>', ' ', ole_match.get('body', {}).get('content', ''))
            result['ole_url'] = _extract_url(body)
            # Always try attachments — OLE QR is inline (isInline=True, hasAttachments=False)
            result['ole_qr'] = _save_first_image(ole_match['id'], 'ole')
        else:
            result['not_found'].append('QR Code OLE e-mail niet gevonden voor deze datum')

        return result

    def fetch_opbrengst_data(self, target_date: datetime = None, since_days: int = 60) -> Dict[str, Any]:
        """Fetch collecte opbrengsten from emails received in the 7 days before target_date.

        Reguliere: fokkedj@gmail.com,    subject like "20260503 Collecte opbrengst"
        OLE:       pmcvb.gkin@gmail.com, subject like "Opbrengst collecte OLE 3 mei 2026"

        Returns a list of opbrengst entries (one per service found) in result['entries'],
        plus flattened fields for the most-recent reguliere entry for backward compatibility.
        """
        NL_MONTHS = ['','januari','februari','maart','april','mei','juni',
                     'juli','augustus','september','oktober','november','december']

        # Window: 7 days before target_date (or last 14 days if no date)
        if target_date:
            window_end   = target_date
            window_start = target_date - timedelta(days=7)
        else:
            window_end   = datetime.utcnow()
            window_start = window_end - timedelta(days=14)

        since = window_start.strftime('%Y-%m-%dT00:00:00Z')
        until = window_end.strftime('%Y-%m-%dT23:59:59Z')

        result: Dict[str, Any] = {
            'collecte_contant': '', 'collecte_bonnen': '',
            'collecte_bank':    '', 'collecte_tikkie': '',
            'collecte_ole':     '',
            'bezoekers_volwassenen': '', 'bezoekers_kinderen': '',
            'emails_found': 0, 'source_subjects': [], 'not_found': [],
            'entries': [],   # list of per-service dicts for multi-table rendering
        }

        def _clean(body_html: str) -> str:
            t = re.sub(r'<[^>]+>', ' ', body_html)
            t = t.replace('&nbsp;', ' ').replace('&amp;', '&')
            return re.sub(r'\s+', ' ', t).strip()

        def _extract(label_pat: str, body: str) -> str:
            # Email format: "Eur  181,35 collecte fysiek ..."  → amount BEFORE label
            m = re.search(r'Eur\s+([\d]{1,6}[,.]\d{2})\s+' + label_pat, body, re.IGNORECASE)
            return m.group(1).replace('.', ',') if m else ''

        def _extract_service_date(subject: str) -> str:
            """Try to pull a human-readable service date from the subject."""
            # "20260503 Collecte opbrengst" → "3 mei 2026"
            m = re.match(r'(\d{4})(\d{2})(\d{2})', subject.strip())
            if m:
                y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
                return f"{d} {NL_MONTHS[mo]} {y}"
            # "Opbrengst collecte OLE 3 mei 2026" — take everything after last keyword
            m2 = re.search(r'OLE\s+(.+)', subject, re.IGNORECASE)
            if m2:
                return m2.group(1).strip()
            return subject

        # --- Reguliere collecte: all matching emails in window ---
        msgs = self._graph_get('/me/messages', params={
            '$filter': (f"receivedDateTime ge {since} and receivedDateTime le {until}"
                        f" and contains(subject,'Collecte opbrengst')"),
            '$top': 10, '$select': 'id,subject,from,body,receivedDateTime',
            '$orderby': 'receivedDateTime asc',
        }).get('value', [])
        fokke_msgs = [m for m in msgs if 'fokkedj@gmail.com' in
                      m.get('from', {}).get('emailAddress', {}).get('address', '').lower()]

        for msg in fokke_msgs:
            result['emails_found'] += 1
            result['source_subjects'].append(msg.get('subject', ''))
            raw_html = msg.get('body', {}).get('content', '')
            body = _clean(raw_html)  # single-line, used for _extract and bezoekers
            # Preserve line breaks for extra-item parsing
            body_lines = re.sub(r'<[^>]+>', ' ', raw_html)
            body_lines = body_lines.replace('&nbsp;', ' ').replace('&amp;', '&')
            body_lines = re.sub(r'[ \t]+', ' ', body_lines)  # collapse spaces but keep newlines
            service_date = _extract_service_date(msg.get('subject', ''))

            entry = {
                'type': 'regulier',
                'service_date': service_date,
                'subject': msg.get('subject', ''),
                'collecte_contant': _extract(r'collecte\s+fysiek', body),
                'collecte_bonnen':  _extract(r'collectebonnen\s+fysiek', body),
                'collecte_bank':    _extract(r'collecte\s+via\s+Bank', body),
                'collecte_tikkie':  _extract(r'collecte\s+via\s+tikkie', body),
                'collecte_ole':     '',
                'bezoekers_volwassenen': '',
                'bezoekers_kinderen': '',
                'extra_items': [],  # additional collecte lines after the ===== separator
            }
            bm = re.search(r'(\d+)\s+Volwassenen[,\s]+(\d+)\s+Kind', body, re.IGNORECASE)
            if bm:
                entry['bezoekers_volwassenen'] = bm.group(1)
                entry['bezoekers_kinderen']    = bm.group(2)

            # Parse extra items after the second ===== block
            # Format: "Eur  200,00 Description text"
            # They appear after the bezoekers line and before any OLE section
            extra_section = re.split(r'={3,}', body_lines)
            # Take the block(s) after the first ===== (which ends the reguliere section)
            for block in extra_section[1:]:
                for line_m in re.finditer(
                    r'Eur\s+([\d.,]+)\s+([^\n\r]+)', block, re.IGNORECASE
                ):
                    amount = line_m.group(1).strip().replace('.', ',')
                    desc   = re.sub(r'\s+', ' ', line_m.group(2)).strip()
                    # Skip OLE lines — those are handled separately
                    if desc and not re.search(r'\bOLE\b', desc, re.IGNORECASE):
                        entry['extra_items'].append({'desc': desc, 'amount': amount})

            result['entries'].append(entry)

        if not fokke_msgs:
            result['not_found'].append('Collecte opbrengst e-mail niet gevonden in de afgelopen 7 dagen')

        # --- OLE collecte: all matching emails in window ---
        msgs2 = self._graph_get('/me/messages', params={
            '$filter': (f"receivedDateTime ge {since} and receivedDateTime le {until}"
                        f" and contains(subject,'Opbrengst collecte OLE')"),
            '$top': 10, '$select': 'id,subject,from,body,receivedDateTime',
            '$orderby': 'receivedDateTime asc',
        }).get('value', [])
        ole_msgs = [m for m in msgs2 if 'pmcvb.gkin@gmail.com' in
                    m.get('from', {}).get('emailAddress', {}).get('address', '').lower()]

        for msg in ole_msgs:
            result['emails_found'] += 1
            result['source_subjects'].append(msg.get('subject', ''))
            body = _clean(msg.get('body', {}).get('content', ''))
            service_date = _extract_service_date(msg.get('subject', ''))

            ole_amount = ''
            om = re.search(r'bedraagt\s*[€]?\s*([\d.,]+)', body, re.IGNORECASE)
            if om:
                raw = om.group(1).rstrip(',-').replace('.', ',')
                if ',' not in raw:
                    raw += ',00'
                ole_amount = raw

            # Try to merge OLE amount into an existing reguliere entry for same date
            merged = False
            for entry in result['entries']:
                if entry['type'] == 'regulier' and entry['service_date'] == service_date:
                    entry['collecte_ole'] = ole_amount
                    merged = True
                    break
            if not merged:
                result['entries'].append({
                    'type': 'ole',
                    'service_date': service_date,
                    'subject': msg.get('subject', ''),
                    'collecte_contant': '', 'collecte_bonnen': '',
                    'collecte_bank': '', 'collecte_tikkie': '',
                    'collecte_ole': ole_amount,
                    'bezoekers_volwassenen': '', 'bezoekers_kinderen': '',
                })

        if not ole_msgs:
            result['not_found'].append('OLE collecte opbrengst e-mail niet gevonden in de afgelopen 7 dagen')

        # Flatten most-recent entry into top-level fields for backward compat
        if result['entries']:
            last = result['entries'][-1]
            for k in ('collecte_contant','collecte_bonnen','collecte_bank',
                      'collecte_tikkie','collecte_ole',
                      'bezoekers_volwassenen','bezoekers_kinderen'):
                result[k] = last.get(k, '')

        return result

    def fetch_liederen(self, target_date: datetime = None, since_days: int = 60) -> Dict[str, Any]:
        """Fetch liederen from aveenliederen@gmail.com, subject contains 'Liederen'.
        Parses the email body for 7 song lines matching labels like '1e lied', '2e lied', etc.
        Returns dict with songs (list of 7 strings), source_subject, not_found.
        """
        NL_MONTHS = ['','januari','februari','maart','april','mei','juni',
                     'juli','augustus','september','oktober','november','december']

        since = (datetime.utcnow() - timedelta(days=since_days)).strftime('%Y-%m-%dT00:00:00Z')

        result: Dict[str, Any] = {
            'songs': ['', '', '', '', '', '', ''],
            'source_subject': '',
            'not_found': [],
        }

        date_variants = []
        if target_date:
            d, m, y = target_date.day, target_date.month, target_date.year
            date_variants = [
                f"{d}-{m}-{y}",
                f"{d:02d}-{m:02d}-{y}",
                f"{d} {NL_MONTHS[m]}",
                f"{d} {NL_MONTHS[m]} {y}",
                f"{d}/{m}/{y}",
            ]

        def _date_matches(text: str) -> bool:
            if not date_variants:
                return True
            tl = text.lower()
            return any(v.lower() in tl for v in date_variants)

        try:
            msgs = self._graph_get('/me/messages', params={
                '$filter': f"receivedDateTime ge {since} and contains(subject,'Liederen')",
                '$top': 20,
                '$select': 'id,subject,from,body,receivedDateTime',
                '$orderby': 'receivedDateTime desc',
            }).get('value', [])
        except Exception as e:
            result['not_found'].append(f'Fout bij ophalen e-mails: {e}')
            return result

        sender_msgs = [
            m for m in msgs
            if 'aveenliederen@gmail.com' in
               m.get('from', {}).get('emailAddress', {}).get('address', '').lower()
        ]

        # Find email matching target date (subject or body)
        match_msg = None
        for msg in sender_msgs:
            subject = msg.get('subject', '')
            body_html = msg.get('body', {}).get('content', '')
            body_text = re.sub(r'<[^>]+>', ' ', body_html)
            body_text = body_text.replace('&nbsp;', ' ').replace('&amp;', '&')
            body_text = re.sub(r'\s+', ' ', body_text)
            if _date_matches(subject) or _date_matches(body_text):
                match_msg = msg
                break

        if not match_msg:
            result['not_found'].append('Liederen e-mail niet gevonden voor deze datum')
            return result

        result['source_subject'] = match_msg.get('subject', '')

        # Parse body — strip HTML, preserve line structure
        body_html = match_msg.get('body', {}).get('content', '')
        # Replace block-level tags with newlines before stripping
        body_text = re.sub(r'<br\s*/?>|</p>|</div>|</li>|</tr>', '\n', body_html, flags=re.IGNORECASE)
        body_text = re.sub(r'<[^>]+>', '', body_text)
        body_text = body_text.replace('&nbsp;', ' ').replace('&amp;', '&').replace('&gt;', '>').replace('&lt;', '<')
        lines = [l.strip() for l in body_text.splitlines() if l.strip()]

        # Labels to search for (flexible matching)
        SONG_PATTERNS = [
            r'1[ée]\s*lied',        # 1e lied / 1ée lied
            r'2[ée]\s*lied',
            r'3[ée]\s*lied',
            r'4[ée]\s*lied',
            r'5[ée]\s*lied',
            r'6[ée]\s*lied',
            r'7[ée]\s*lied',
        ]

        songs = [''] * 7
        for i, pat in enumerate(SONG_PATTERNS):
            for line in lines:
                if re.search(pat, line, re.IGNORECASE):
                    # Extract value after the label+colon/dash
                    val = re.split(r'[:–\-]', line, maxsplit=1)
                    if len(val) > 1:
                        songs[i] = val[1].strip()
                    else:
                        # Label takes up whole line — value is on next line
                        idx = lines.index(line)
                        if idx + 1 < len(lines):
                            songs[i] = lines[idx + 1].strip()
                    break

        result['songs'] = songs
        return result

    def fetch_overdenking(self, target_date: datetime = None, since_days: int = 14) -> Dict[str, Any]:
        """Fetch overdenking from scribagkin@gmail.com, subject contains 'Overdenking'.
        Attachment is a .docx with structure:
          [0] "Overdenking, 10 mei 2026"
          [1] "Thema title\\nSchriftlezing ref"
          [3..n-2] body paragraphs
          [last non-empty] "ds. Firstname Lastname"
        Returns dict with predikant, thema, schriftlezing, content, not_found.
        """
        import base64
        from io import BytesIO
        try:
            from docx import Document as _Document
        except ImportError:
            return {'error': 'python-docx not installed', 'not_found': ['python-docx niet beschikbaar']}

        NL_MONTHS = ['','januari','februari','maart','april','mei','juni',
                     'juli','augustus','september','oktober','november','december']

        since = (datetime.utcnow() - timedelta(days=since_days)).strftime('%Y-%m-%dT00:00:00Z')

        result: Dict[str, Any] = {
            'predikant': '', 'thema': '', 'schriftlezing': '', 'content': '',
            'not_found': [],
        }

        # Date variants for subject/attachment matching
        date_variants = []
        if target_date:
            d, m, y = target_date.day, target_date.month, target_date.year
            date_variants = [
                f"{d}-{m}-{y}",           # 10-5-2026
                f"{d:02d}-{m:02d}-{y}",   # 10-05-2026
                f"{d} {NL_MONTHS[m]}",     # 10 mei
                f"{d} {NL_MONTHS[m]} {y}", # 10 mei 2026
                f"{d}/{m}/{y}",            # 10/5/2026
            ]

        def _date_matches(text: str) -> bool:
            if not date_variants:
                return True
            tl = text.lower()
            return any(v.lower() in tl for v in date_variants)

        # Search emails
        msgs = self._graph_get('/me/messages', params={
            '$filter': f"receivedDateTime ge {since} and contains(subject,'Overdenking')",
            '$top': 10,
            '$select': 'id,subject,from,hasAttachments,receivedDateTime',
            '$orderby': 'receivedDateTime desc',
        }).get('value', [])

        scriba_msgs = [m for m in msgs if 'scribagkin@gmail.com' in
                       m.get('from', {}).get('emailAddress', {}).get('address', '').lower()]

        # Find best match: subject contains date, or attachment name contains date
        match_msg = None
        match_att = None

        for msg in scriba_msgs:
            subject = msg.get('subject', '')
            # Check subject first
            if _date_matches(subject):
                # Get attachments
                atts = self._graph_get(
                    f"/me/messages/{msg['id']}/attachments",
                    params={'$select': 'id,name,contentType,size'}
                ).get('value', [])
                docx_atts = [a for a in atts if 'wordprocessingml' in a.get('contentType', '')
                             or a.get('name','').lower().endswith('.docx')]
                if docx_atts:
                    match_msg = msg
                    match_att = docx_atts[0]
                    break
            # Also check attachment names if subject has no date
            if not match_msg and msg.get('hasAttachments'):
                atts = self._graph_get(
                    f"/me/messages/{msg['id']}/attachments",
                    params={'$select': 'id,name,contentType,size'}
                ).get('value', [])
                for a in atts:
                    if ('wordprocessingml' in a.get('contentType', '')
                            or a.get('name','').lower().endswith('.docx')):
                        if _date_matches(a.get('name', '')):
                            match_msg = msg
                            match_att = a
                            break
                if match_msg:
                    break

        if not match_msg or not match_att:
            result['not_found'].append('Overdenking e-mail niet gevonden voor deze datum')
            return result

        # Fetch attachment bytes
        full = self._graph_get(f"/me/messages/{match_msg['id']}/attachments/{match_att['id']}")
        cb = full.get('contentBytes', '')
        if not cb:
            result['not_found'].append('Overdenking bijlage kon niet worden geladen')
            return result

        # Parse docx
        doc = _Document(BytesIO(base64.b64decode(cb)))
        paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

        # Extract predikant — last non-empty paragraph, usually "ds. Firstname Lastname"
        predikant = ''
        for p in reversed(paras):
            if re.search(r'\bds\.?\b|\bprop\.?\b|\bdrs\.?\b', p, re.IGNORECASE) or \
               re.search(r'^(ds|prop|drs)[\s.]', p, re.IGNORECASE):
                predikant = p
                break
        if not predikant and paras:
            predikant = paras[-1]

        # Extract thema + schriftlezing from paragraph [1]
        # Format: '"Thema title"\nSchriftlezing ref'  or two separate paragraphs
        thema = ''
        schriftlezing = ''
        if len(paras) > 1:
            second = paras[1]
            if '\n' in second:
                parts = [x.strip() for x in second.split('\n', 1)]
                thema = parts[0].strip('""\u201c\u201d')
                schriftlezing = parts[1]
            else:
                thema = second.strip('""\u201c\u201d')
                # Look for schriftlezing in next paragraph (bible ref pattern)
                if len(paras) > 2 and re.search(r'\d+:\d+', paras[2]):
                    schriftlezing = paras[2]

        # Body: all paragraphs between [2] and the predikant line
        body_paras = []
        for p in paras[2:]:
            if p == predikant:
                break
            if not re.search(r'\d+:\d+', p) or len(p) > 30:  # skip lone bible refs already captured
                body_paras.append(p)
        content = '\n\n'.join(body_paras)

        result['predikant']    = predikant
        result['thema']        = thema
        result['schriftlezing'] = schriftlezing
        result['content']      = content
        result['source_subject'] = match_msg.get('subject', '')
        return result
