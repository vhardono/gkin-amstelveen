"""
Preekbevestiging Document Generator
Copies the template docx and fills in dynamic values.
"""

import copy
import os
import re
from datetime import datetime, timedelta

from docx import Document
from docx.oxml.ns import qn

DUTCH_MONTHS = ['januari','februari','maart','april','mei','juni',
                'juli','augustus','september','oktober','november','december']
DUTCH_DAYS   = ['maandag','dinsdag','woensdag','donderdag',
                'vrijdag','zaterdag','zondag']

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__),
                             '20260329_Preekbevestiging zr B T Sari.docx')


def _fmt_date_long(iso_str: str) -> str:
    d = datetime.strptime(iso_str, '%Y-%m-%d')
    return f"Zondag {d.day} {DUTCH_MONTHS[d.month-1]} {d.year}"


def _fmt_date_today() -> str:
    d = datetime.now()
    return f"Amstelveen, {d.day} {DUTCH_MONTHS[d.month-1]} {d.year}"


def _deadline_tue(iso_str: str) -> str:
    d = datetime.strptime(iso_str, '%Y-%m-%d')
    tue = d - timedelta(days=5)
    return f"dinsdag {tue.day} {DUTCH_MONTHS[tue.month-1]} {tue.year}"


def _replace_in_para(para, old: str, new: str):
    """Replace text in a paragraph preserving run formatting of the first run."""
    full = para.text
    if old not in full:
        return False
    # Rebuild runs: clear all, put replacement in first run
    new_text = full.replace(old, new)
    for i, run in enumerate(para.runs):
        if i == 0:
            run.text = new_text
        else:
            run.text = ''
    return True


def _replace_in_cell(cell, old: str, new: str):
    for para in cell.paragraphs:
        _replace_in_para(para, old, new)


def generate(entry: dict, iso_date: str, output_path: str, songs: list = None):
    doc = Document(TEMPLATE_PATH)

    predikant  = entry.get('predikant', '')
    ovd        = entry.get('ovd', '')
    eo1        = entry.get('1eo', '')
    tijd       = entry.get('tijd', '10:30') or '10:30'

    # Build greeting name: "zr. B. T. Sari" style — use full predikant name
    # Extract title + rest: e.g. "ds. C. de Jonge" -> greeting "ds. De Jonge"
    parts = predikant.strip().split()
    title = parts[0] if parts else ''
    last  = parts[-1] if len(parts) > 1 else predikant
    last_cap = last[0].upper() + last[1:] if last else last
    greeting_name = f"{title} {last_cap}".strip()

    date_long   = _fmt_date_long(iso_date)     # "Zondag 29 maart 2026"
    date_today  = _fmt_date_today()            # "Amstelveen, 15 maart 2026"
    deadline    = _deadline_tue(iso_date)      # "dinsdag 24 maart 2026"

    # --- Substitution map ---
    subs = {
        # Header date
        'Amstelveen, 15 maart 2026': date_today,
        # Addressee
        'Zr. B. T. Sari':           predikant,
        # Salutation
        'Geachte zr. B. T. Sari,':  f'Geachte {greeting_name},',
        # Service date in body
        'Zondag 29 maart 2026 om 10.30 uur': f'{date_long} om {tijd} uur',
        # Deadline in body
        'uiterlijk dinsdag 24 maart 2026 het invulformulier': f'uiterlijk {deadline} het invulformulier',
        # Footer deadline
        'uiterlijk dinsdag 24 maart 2026)': f'uiterlijk {deadline})',
    }

    # Apply to all paragraphs
    for para in doc.paragraphs:
        for old, new in subs.items():
            _replace_in_para(para, old, new)

    # Apply to all table cells
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for old, new in subs.items():
                    _replace_in_cell(cell, old, new)
                # Table 1 specific dynamic fields
                txt = cell.text.strip()
                if txt == 'Zondag 29 maart 2026':
                    _replace_in_cell(cell, txt, date_long)
                elif txt == '10:30 uur':
                    _replace_in_cell(cell, txt, f'{tijd} uur')
                elif txt == 'zr. B. T. Sari':
                    _replace_in_cell(cell, txt, predikant)
                elif txt == 'br. Hamra Simatupang':
                    _replace_in_cell(cell, txt, ovd)
                elif txt == 'zr. Joyce Uning':
                    _replace_in_cell(cell, txt, eo1)

    # Fill song table (Table 2) with provided songs
    if songs and len(doc.tables) > 2:
        song_table = doc.tables[2]  # 0-indexed: Table 2 is the songs table
        song_labels = [
            '1e lied (Intochtslied)', '2e lied', '3e lied', '4e lied',
            '5e lied', '6e lied (Dankoffer)', '7e lied (Slotlied)'
        ]
        for ri, row in enumerate(song_table.rows):
            if ri < len(songs) and songs[ri]:
                # Value goes in col index 2
                cells = row.cells
                if len(cells) > 2:
                    for para in cells[2].paragraphs:
                        for run in para.runs:
                            run.text = ''
                        if para.runs:
                            para.runs[0].text = songs[ri]
                        else:
                            para.add_run(songs[ri])

    doc.save(output_path)
