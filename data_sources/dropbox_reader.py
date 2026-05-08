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
        """Build a lookup dict: short_name -> {first, last, title}."""
        self._people_map = {}
        for _, row in df.iterrows():
            short = str(row.get('Short Name', '')).strip()
            if not short or short.lower() == 'nan':
                continue
            def _s(col):
                v = row.get(col, '')
                s = str(v).strip() if pd.notna(v) else ''
                return '' if s.lower() == 'nan' else s
            self._people_map[short] = {
                'first_name': _s('First Name'),
                'last_name':  _s('Last Name'),
                'title':      _s('Title'),
            }

    def _resolve_name(self, short_name: str) -> str:
        """Convert a short name to 'title First Last' using People tab."""
        short_name = short_name.strip()
        if not short_name or not self._people_map:
            return short_name
        person = self._people_map.get(short_name)
        if person:
            parts = [person['title'], person['first_name'], person['last_name']]
            return ' '.join(p for p in parts if p and p.lower() != 'nan')
        return short_name

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

            entries.append({
                'date':      date_obj,
                'dag':       dag,
                'predikant': predikant,
                'ovd':       self._resolve_name(ovd_short) or ovd_short,
                '1eo':       self._resolve_name(eo1_short) or eo1_short,
                'beamer':    self._resolve_name(beamer_short) or beamer_short,
                'opmerking': opmerking,
            })

        return entries
