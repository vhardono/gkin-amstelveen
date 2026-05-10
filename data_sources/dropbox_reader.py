"""
Dropbox Excel Reader
Reads Takenrooster Excel file from Dropbox for church bulletin data.
Extracts dates, predikant, and OvD (with full name lookup from People tab).
"""

import os
import pandas as pd
import dropbox
from dropbox.exceptions import ApiError
from io import BytesIO
from datetime import datetime
from typing import Dict, List, Any, Optional
from config import Config

TAKENROOSTER_PATH = '/# Kerkbode GKIN Amstelveen/Rooster/Takenrooster_GKIN_Amstelveen_2026.xlsx'
MEDEDELINGEN_PATH_TEMPLATE = '/# Kerkbode GKIN Amstelveen/{year}/Mededelingen Overzicht.xlsx'


class DropboxExcelReader:
    def __init__(self):
        """Initialize Dropbox client using refresh token for long-lived access"""
        if not Config.DROPBOX_REFRESH_TOKEN:
            raise ValueError("DROPBOX_REFRESH_TOKEN not configured")

        self.dbx = dropbox.Dropbox(
            oauth2_refresh_token=Config.DROPBOX_REFRESH_TOKEN,
            app_key=Config.DROPBOX_APP_KEY,
            app_secret=Config.DROPBOX_APP_SECRET,
        )
        self._people_map: Optional[Dict[str, Dict]] = None

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def get_takenrooster(self) -> Dict[str, Any]:
        """Read the Takenrooster and return structured data.

        Returns dict with:
            'dates'   – list of available service dates (datetime objects)
            'entries' – list of dicts with keys: date, dag, predikant, ovd, opmerking
                        where ovd is the full name with salutation from People tab.
        """
        try:
            _, response = self.dbx.files_download(TAKENROOSTER_PATH)
            content = BytesIO(response.content)

            people_df = pd.read_excel(content, sheet_name='People', header=0)
            self._build_people_map(people_df)

            content.seek(0)
            current_df = pd.read_excel(content, sheet_name='CURRENT', header=None)

            entries = self._parse_current_sheet(current_df)
            dates = [e['date'] for e in entries]

            print(f"Takenrooster: {len(entries)} services loaded")
            return {
                'source': TAKENROOSTER_PATH,
                'dates': dates,
                'entries': entries,
            }

        except Exception as e:
            print(f"Error reading takenrooster: {e}")
            return {'error': str(e), 'dates': [], 'entries': []}

    def get_mededelingen(self, mededelingen_date: datetime = None) -> Dict[str, Any]:
        """Read Mededelingen Overzicht from the Output tab.

        The year in the file path is derived from the selected mededelingen_date.

        Returns dict with:
            'regionale_nl'  – Regionale Mededelingen (Nederlands) from B2
            'landelijke_nl' – Landelijke Mededelingen (Nederlands) from B3
            'regionale_id'  – Regionale Mededelingen (Bahasa Indonesia) from C2
            'landelijke_id' – Landelijke Mededelingen (Bahasa Indonesia) from C3
        """
        if mededelingen_date is None:
            year = datetime.now().year
        else:
            year = mededelingen_date.year

        path = MEDEDELINGEN_PATH_TEMPLATE.format(year=year)
        try:
            print(f"Reading Mededelingen Overzicht: {path}")
            _, response = self.dbx.files_download(path)
            df = pd.read_excel(BytesIO(response.content), sheet_name='Output', header=None)

            regionale_nl = str(df.iloc[1, 1]).strip() if pd.notna(df.iloc[1, 1]) else ''
            landelijke_nl = str(df.iloc[2, 1]).strip() if pd.notna(df.iloc[2, 1]) else ''
            regionale_id = str(df.iloc[1, 2]).strip() if pd.notna(df.iloc[1, 2]) else ''
            landelijke_id = str(df.iloc[2, 2]).strip() if pd.notna(df.iloc[2, 2]) else ''

            print(f"Mededelingen loaded: regionale={len(regionale_nl)} chars, landelijke={len(landelijke_nl)} chars")
            return {
                'source': path,
                'regionale_nl': regionale_nl,
                'landelijke_nl': landelijke_nl,
                'regionale_id': regionale_id,
                'landelijke_id': landelijke_id,
            }

        except Exception as e:
            print(f"Error reading mededelingen: {e}")
            return {'error': str(e), 'regionale_nl': '', 'landelijke_nl': '',
                    'regionale_id': '', 'landelijke_id': ''}

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _build_people_map(self, df: pd.DataFrame) -> None:
        """Build a lookup dict: short_name -> {first, last, title}.
        Also builds _predikant_email_map: full_name -> email from rows
        where Short Name is empty but col F (index 5) has a full name.
        """
        self._people_map = {}
        self._predikant_email_map = {}
        for _, row in df.iterrows():
            def _s(col):
                v = row.get(col, '') if isinstance(col, str) else (row.iloc[col] if col < len(row) else '')
                s = str(v).strip() if pd.notna(v) else ''
                return '' if s.lower() == 'nan' else s
            short = _s('Short Name')
            email = _s('Email') or _s(3)
            full  = _s(5)   # col F = full predikant name (may exist even with short name)
            # Index col F full name -> email for predikant lookup regardless of short name
            if full and email:
                self._predikant_email_map[full.lower()] = email
            if not short:
                continue
            self._people_map[short.lower()] = {
                'first_name': _s('First Name'),
                'last_name':  _s('Last Name'),
                'title':      _s('Title'),
                'email':      email,
            }

    def _resolve_name(self, short_name: str) -> str:
        """Convert a short name to 'title First Last' using People tab.
        Tracks unresolved names in _unresolved_names list.
        Lookup is case-insensitive."""
        short_name = short_name.strip()
        if not short_name or not self._people_map:
            if short_name and hasattr(self, '_unresolved_names'):
                self._unresolved_names.append(short_name)
            return short_name
        person = self._people_map.get(short_name.lower())
        if person:
            parts = [person['title'], person['first_name'], person['last_name']]
            return ' '.join(p for p in parts if p and p.lower() != 'nan')
        # Name not found in People tab
        if hasattr(self, '_unresolved_names'):
            self._unresolved_names.append(short_name)
        return short_name

    def _resolve_email(self, short_name: str) -> str:
        """Return email for a short name from People tab.
        Lookup is case-insensitive."""
        short_name = short_name.strip()
        if not short_name or not self._people_map:
            return ''
        person = self._people_map.get(short_name.lower())
        return person.get('email', '') if person else ''

    def _parse_current_sheet(self, df: pd.DataFrame) -> List[Dict[str, Any]]:
        """Parse the CURRENT sheet into a list of service entries."""
        # Find the header row (contains 'DAG' and 'DATUM')
        header_idx = None
        for i in range(min(10, len(df))):
            row_vals = [str(v).strip().upper() for v in df.iloc[i] if pd.notna(v)]
            if 'DAG' in row_vals and 'DATUM' in row_vals:
                header_idx = i
                break

        if header_idx is None:
            print("Could not find header row in CURRENT sheet")
            return []

        # Map column names to indices
        headers = [str(v).strip().upper() if pd.notna(v) else '' for v in df.iloc[header_idx]]
        col = {}
        for idx, h in enumerate(headers):
            h_norm = h.strip().upper().replace('\n', ' ')
            # Canonical mappings
            if h_norm in ('DAG', 'DATUM', 'PREDIKANT', 'OPMERKING', 'TIJD',
                          'BEAMER', 'MUZIEK', 'MULTIMEDIA'):
                col[h_norm] = idx
            elif h_norm in ('OVD', 'OVD.', 'OV D'):
                col['OVD'] = idx
            elif h_norm.startswith('1E') or h_norm in ('1E ONTV', '1E OUDERLING',
                                                        'EERSTE OUDERLING', '1EO'):
                col['1EO'] = idx
        print(f'Takenrooster columns found: {col}')

        entries = []
        for i in range(header_idx + 1, len(df)):
            row = df.iloc[i]
            datum_val = row.iloc[col.get('DATUM', 1)]
            if pd.isna(datum_val):
                continue

            # Parse date
            if isinstance(datum_val, datetime):
                date_obj = datum_val
            else:
                try:
                    date_obj = pd.to_datetime(datum_val)
                except Exception:
                    continue

            def _cell(key, default_col):
                idx2 = col.get(key, default_col)
                v = row.iloc[idx2] if idx2 < len(row) else None
                return str(v).strip() if pd.notna(v) else ''

            dag       = _cell('DAG', 0)
            predikant = _cell('PREDIKANT', 4)
            ovd_short = _cell('OVD', 5)
            opmerking = _cell('OPMERKING', 3)
            eo1_short    = _cell('1EO', 6)    # G = index 6 (1e ONTV)
            beamer_short = _cell('BEAMER', 9)  # J = index 9 (BEAMER)
            # Parse TIJD: may be a time object or string like '10:30:00'
            tijd_raw = _cell('TIJD', 2)
            import datetime as _dt
            if hasattr(tijd_raw, 'strftime'):
                tijd = tijd_raw.strftime('%H:%M')
            elif tijd_raw:
                try:
                    t = _dt.time.fromisoformat(str(tijd_raw).split('.')[0])
                    tijd = t.strftime('%H:%M')
                except Exception:
                    tijd = str(tijd_raw)[:5]
            else:
                tijd = '10:30'

            def _resolve_list(key, default_col):
                raw = _cell(key, default_col)
                if not raw or raw == '-':
                    return ''
                names = [n.strip() for n in raw.split(',') if n.strip() and n.strip() != '-']
                resolved = []
                for n in names:
                    r = self._resolve_name(n)
                    resolved.append(r or n)
                return ', '.join(resolved)

            def _resolve_name_list(raw_value):
                """Resolve a comma-separated list of names like _resolve_list but returns individual unresolved tracking."""
                if not raw_value or raw_value == '-':
                    return ''
                names = [n.strip() for n in raw_value.split(',') if n.strip() and n.strip() != '-']
                resolved = []
                for n in names:
                    r = self._resolve_name(n)
                    resolved.append(r or n)
                return ', '.join(resolved)

            # Track unresolved names for this entry
            self._unresolved_names = []

            muziek      = _resolve_list('MUZIEK', 10)
            voorzangers = _resolve_list('VOORZANGERS', 11)
            multimedia  = _resolve_list('MULTIMEDIA', 12)
            knd_raw     = _cell('KND', 7)
            tieners_raw = _cell('TIENERS', 8)
            knd         = _resolve_name_list(knd_raw) if knd_raw and knd_raw != '-' else ''
            tieners     = _resolve_name_list(tieners_raw) if tieners_raw and tieners_raw != '-' else ''

            # Resolve OVD/1EO/Beamer names (tracks unresolved)
            ovd_full    = self._resolve_name(ovd_short) or ovd_short
            eo1_full    = self._resolve_name(eo1_short) or eo1_short
            beamer_full = self._resolve_name(beamer_short) or beamer_short

            # Get unique unresolved names
            unresolved = list(dict.fromkeys(self._unresolved_names))  # preserve order, remove duplicates

            # Resolve predikant email from People tab
            pred_email  = self._predikant_email_map.get(predikant.lower(), '')
            ovd_email   = self._resolve_email(ovd_short)
            eo1_email   = self._resolve_email(eo1_short)
            beamer_email = self._resolve_email(beamer_short)

            entries.append({
                'date':            date_obj,
                'dag':             dag,
                'predikant':       predikant,
                'predikant_email': pred_email,
                'ovd':             ovd_full,
                'ovd_email':       ovd_email,
                '1eo':             eo1_full,
                '1eo_email':       eo1_email,
                'beamer':          beamer_full,
                'beamer_email':    beamer_email,
                'opmerking':       opmerking,
                'tijd':            tijd,
                'muziek':          muziek,
                'voorzangers':     voorzangers,
                'multimedia':      multimedia,
                'knd':             knd,
                'tieners':         tieners,
                'unresolved_names': unresolved,
            })

        return entries
