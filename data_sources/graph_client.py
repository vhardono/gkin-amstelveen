"""
SharePoint/Graph storage client for GKIN Amstelveen.

Provides a minimal Dropbox-compatible API over Microsoft Graph Drive items so
callers can switch storage backends via STORAGE_BACKEND without code changes.

Required env vars:
    MICROSOFT_CLIENT_ID / TENANT_ID / CLIENT_SECRET  (already used for SSO)
    SHAREPOINT_SITE_HOST   e.g. gkinportal.sharepoint.com
    SHAREPOINT_SITE_PATH   e.g. /sites/kerkbode  (default: root site)
    SHAREPOINT_BASE_PATH   folder prefix inside the drive, e.g. /Kerkbode GKIN Amstelveen
    SHAREPOINT_DRIVE_NAME  optional; default document library if unset
"""

import os
import requests
from urllib.parse import quote

GRAPH_ROOT = 'https://graph.microsoft.com/v1.0'
GRAPH_SCOPE = ['https://graph.microsoft.com/.default']

SHAREPOINT_SITE_HOST = os.environ.get('SHAREPOINT_SITE_HOST', '')
SHAREPOINT_SITE_PATH = os.environ.get('SHAREPOINT_SITE_PATH', '') or '/'
SHAREPOINT_BASE_PATH = os.environ.get('SHAREPOINT_BASE_PATH', '').rstrip('/')
SHAREPOINT_DRIVE_NAME = os.environ.get('SHAREPOINT_DRIVE_NAME', '')

_msal_app = None
_drive_cache = {}

SHAREPOINT_MAP_JSON = os.environ.get('SHAREPOINT_MAP', '')
_PATH_RULES = None


def _rules():
    """Path remap rules from SHAREPOINT_MAP JSON:
    {"<prefix>": "<site>|<new prefix>", "*<suffix>": "<site>|<new prefix>"}
    Prefix rules replace the matched prefix; *-rules match on path suffix."""
    global _PATH_RULES
    if _PATH_RULES is None:
        try:
            import json as _j
            _PATH_RULES = _j.loads(SHAREPOINT_MAP_JSON) if SHAREPOINT_MAP_JSON else {}
        except Exception:
            _PATH_RULES = {}
    return _PATH_RULES


def _norm_site(s):
    s = (s or '').strip() or SHAREPOINT_SITE_PATH
    if s and not s.startswith('/'):
        s = '/sites/' + s
    return s or '/'


def _token():
    global _msal_app
    import msal
    if _msal_app is None:
        _msal_app = msal.ConfidentialClientApplication(
            os.environ.get('MICROSOFT_CLIENT_ID', ''),
            authority=f"https://login.microsoftonline.com/{os.environ.get('MICROSOFT_TENANT_ID', '')}",
            client_credential=os.environ.get('MICROSOFT_CLIENT_SECRET', ''))
    result = _msal_app.acquire_token_for_client(scopes=GRAPH_SCOPE)
    if 'access_token' not in result:
        raise RuntimeError(f"Graph token failed: {result.get('error')}: {result.get('error_description')}")
    return result['access_token']


def _headers():
    return {'Authorization': f'Bearer {_token()}'}


def _drive(site_path=None):
    """Resolve the SharePoint drive (document library) id, cached per site."""
    key = _norm_site(site_path)
    if key in _drive_cache:
        return _drive_cache[key]
    if not SHAREPOINT_SITE_HOST:
        raise RuntimeError('SHAREPOINT_SITE_HOST is not set')
    site_path = key
    r = requests.get(f'{GRAPH_ROOT}/sites/{SHAREPOINT_SITE_HOST}:{site_path}',
                     headers=_headers(), timeout=30)
    r.raise_for_status()
    site_id = r.json()['id']
    if SHAREPOINT_DRIVE_NAME:
        r = requests.get(f'{GRAPH_ROOT}/sites/{site_id}/drives',
                         headers=_headers(), timeout=30)
        r.raise_for_status()
        drives = r.json().get('value', [])
        match = next((d for d in drives if d.get('name') == SHAREPOINT_DRIVE_NAME), None)
        if not match:
            raise RuntimeError(f"Drive '{SHAREPOINT_DRIVE_NAME}' not found on site")
        _drive_cache[key] = match['id']
    else:
        r = requests.get(f'{GRAPH_ROOT}/sites/{site_id}/drive',
                         headers=_headers(), timeout=30)
        r.raise_for_status()
        _drive_cache[key] = r.json()['id']
    return _drive_cache[key]


def _resolve_path(dbx_path: str):
    """Map a Dropbox-style path to (site_path, sharepoint_path).

    '/# Kerkbode GKIN Amstelveen/x.y' -> strips '# ' then applies SHAREPOINT_MAP
    rules (longest prefix first, then *-suffix), else SHAREPOINT_BASE_PATH + path.
    """
    p = (dbx_path or '').lstrip('/')
    if p.startswith('# '):
        # Drop the Dropbox root folder segment entirely ('# Kerkbode GKIN Amstelveen/')
        p = p.split('/', 1)[1] if '/' in p else p[2:]
    rules = _rules()
    for prefix in sorted((k for k in rules if not k.startswith('*')), key=len, reverse=True):
        if p.startswith(prefix):
            site, _, new_prefix = rules[prefix].partition('|')
            return _norm_site(site), new_prefix.rstrip('/') + '/' + p[len(prefix):]
    for star in rules:
        if not star.startswith('*'):
            continue
        suffix = star[1:]
        if p.endswith(suffix):
            site, _, new_prefix = rules[star].partition('|')
            return _norm_site(site), new_prefix.rstrip('/') + '/' + os.path.basename(p)
    p = f'{SHAREPOINT_BASE_PATH}/{p}' if SHAREPOINT_BASE_PATH else '/' + p
    return _norm_site(SHAREPOINT_SITE_PATH), p


def sp_full_path(dbx_path: str) -> str:
    return _resolve_path(dbx_path)[1]


def _item_url(path: str, suffix: str = '') -> str:
    site, sp_path = _resolve_path(path)
    return f"{GRAPH_ROOT}/drives/{_drive(site)}/root:{quote(sp_path, safe='/')}{suffix}"


class _Resp:
    """Minimal stand-in for dropbox files_download response."""
    def __init__(self, content):
        self.content = content


class GraphEntry:
    """Minimal stand-in for dropbox.files.FileMetadata."""
    def __init__(self, item, path_display):
        self.name = item.get('name', '')
        self.path_display = path_display
        self.path_lower = path_display.lower()
        self.is_folder = 'folder' in item
        self.size = item.get('size', 0)
        self.server_modified = item.get('lastModifiedDateTime', '')


class GraphDrive:
    """Dropbox-compatible facade over a SharePoint drive."""

    def files_download(self, path):
        r = requests.get(_item_url(path, ':/content'), headers=_headers(), timeout=60)
        r.raise_for_status()
        return None, _Resp(r.content)

    def files_upload(self, data, path, mode=None, **kwargs):
        url = _item_url(path, ':/content')
        if isinstance(data, (str, bytes)) and len(data) <= 3_500_000:
            r = requests.put(url, data=data, headers=_headers(), timeout=120)
            r.raise_for_status()
            return None
        # Resumable upload for larger files
        r = requests.post(_item_url(path, ':/createUploadSession'),
                          json={}, headers=_headers(), timeout=30)
        r.raise_for_status()
        upload_url = r.json()['uploadUrl']
        size = len(data)
        chunk = 8 * 1024 * 1024
        sent = 0
        while sent < size:
            end = min(sent + chunk, size)
            part = requests.put(upload_url, data=data[sent:end],
                                headers={'Content-Length': str(end - sent),
                                         'Content-Range': f'bytes {sent}-{end-1}/{size}'},
                                timeout=180)
            part.raise_for_status()
            sent = end
        return None

    def files_list_folder(self, path, include_non_downloadable_files=False):
        items, url = [], _item_url(path, ':/children')
        while url:
            r = requests.get(url, headers=_headers(), timeout=30)
            r.raise_for_status()
            payload = r.json()
            items.extend(payload.get('value', []))
            url = payload.get('@odata.nextLink')
        base = sp_full_path(path).rstrip('/')
        entries = [GraphEntry(i, f'{base}/{i.get("name", "")}') for i in items]

        class _Result:
            pass
        res = _Result()
        res.entries = entries
        res.has_more = False
        return res

    def files_delete_v2(self, path):
        r = requests.delete(_item_url(path), headers=_headers(), timeout=30)
        r.raise_for_status()
        return None

    def files_get_metadata(self, path):
        r = requests.get(_item_url(path), headers=_headers(), timeout=30)
        r.raise_for_status()
        return GraphEntry(r.json(), sp_full_path(path))
