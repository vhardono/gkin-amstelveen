#!/usr/bin/env python3
"""
Emergency manual liturgie generator.

This script uses the exact same logic as the webapp (app._run_liturgi) to
generate LiturgieA, LiturgieB and LiturgieP from a local Excel file and an
optional Preek .docx.  It can be run manually when the webapp is unavailable.

How to run:
    cd /Users/vega/Library/CloudStorage/Dropbox/working\ folder
    python3 /path/to/emergency_liturgie.py

The script will look for files in ./file mingguan/ and copy the generated
outputs back into that folder.
"""

import os
import re
import shutil
import sys
import tempfile
from pathlib import Path

# ---------------------------------------------------------------------------
# Locate the project that contains the shared liturgie logic
# ---------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parent
POSSIBLE_PROJECT_DIRS = [
    str(SCRIPT_DIR),
    '/Users/vega/CascadeProjects/gkin-project',
]
PROJECT_DIR = None
for d in POSSIBLE_PROJECT_DIRS:
    if os.path.isfile(os.path.join(d, 'app.py')):
        PROJECT_DIR = d
        break

if not PROJECT_DIR:
    print('FOUT: kan app.py niet vinden. Plaats dit script naast gkin-project of pas PROJECT_DIR aan.')
    sys.exit(1)

sys.path.insert(0, PROJECT_DIR)
import app

# ---------------------------------------------------------------------------
# Locate the working folder (where file mingguan/ lives)
# ---------------------------------------------------------------------------
WORKING_DIR = str(SCRIPT_DIR)
FILE_MINGGUAN = os.path.join(WORKING_DIR, 'file mingguan')

if not os.path.isdir(FILE_MINGGUAN):
    DROPBOX_FALLBACK = '/Users/vega/Library/CloudStorage/Dropbox/working folder'
    if os.path.isdir(os.path.join(DROPBOX_FALLBACK, 'file mingguan')):
        WORKING_DIR = DROPBOX_FALLBACK
        FILE_MINGGUAN = os.path.join(WORKING_DIR, 'file mingguan')
    else:
        print(f'FOUT: geen "file mingguan" map gevonden bij {WORKING_DIR} of {DROPBOX_FALLBACK}')
        sys.exit(1)

print(f'Werkmap: {WORKING_DIR}')
print(f'file mingguan: {FILE_MINGGUAN}')
print(f'Project: {PROJECT_DIR}')
print()

# ---------------------------------------------------------------------------
# Helpers to list and pick the latest files
# ---------------------------------------------------------------------------
def _date_key(name):
    m = re.search(r'(\d{4})(\d{2})(\d{2})', name)
    return m.group(0) if m else '00000000'


def list_main_files():
    files = []
    for f in os.listdir(FILE_MINGGUAN):
        low = f.lower()
        if low.startswith('main liturgy file') and low.endswith('.xlsx'):
            files.append(f)
    files.sort(key=_date_key, reverse=True)
    return files


def list_preek_files():
    files = []
    for f in os.listdir(FILE_MINGGUAN):
        low = f.lower()
        if low.startswith('preek') and low.endswith('.docx'):
            files.append(f)
    files.sort(key=_date_key, reverse=True)
    return files


def choose_file(files, label, optional=False):
    if not files:
        if optional:
            print(f'Geen {label} bestanden gevonden; wordt overgeslagen.')
            return None
        print(f'FOUT: geen {label} bestand gevonden in {FILE_MINGGUAN}')
        sys.exit(1)

    print(f'Beschikbare {label} bestanden:')
    for i, f in enumerate(files, 1):
        marker = ' (laatste)' if i == 1 else ''
        print(f'  {i}. {f}{marker}')

    if optional:
        print('  0. Geen preek gebruiken')

    prompt = f'Kies {label} nummer (Enter = 1): '
    raw = input(prompt).strip()
    if raw == '':
        return files[0]
    try:
        idx = int(raw)
    except ValueError:
        idx = -1

    if optional and idx == 0:
        return None
    if 1 <= idx <= len(files):
        return files[idx - 1]
    print('Ongeldige keuze, gebruik de laatste.')
    return files[0]


# ---------------------------------------------------------------------------
# User choices
# ---------------------------------------------------------------------------
print('Welke bestanden wil je genereren?')
print('  1 = LiturgieA (Word)')
print('  2 = LiturgieB (Word)')
print('  3 = LiturgieP (PowerPoint)')
print("  'all' of leeg = alles")
choice_raw = input('Keuze: ').strip().lower()

if choice_raw in ('', 'all', 'alles', '1,2,3', '123'):
    want_a = want_b = want_p = True
else:
    parts = {c.strip() for c in choice_raw.split(',')}
    want_a = '1' in parts
    want_b = '2' in parts
    want_p = '3' in parts
    if not (want_a or want_b or want_p):
        print('Geen geldige keuze; alles wordt gegenereerd.')
        want_a = want_b = want_p = True

print(f'-> LiturgieA: {want_a} | LiturgieB: {want_b} | LiturgieP: {want_p}')
print()

main_files = list_main_files()
main_file = choose_file(main_files, 'Main Liturgy')
main_path = os.path.join(FILE_MINGGUAN, main_file)

preek_files = list_preek_files()
preek_file = choose_file(preek_files, 'Preek', optional=True)
preek_path = os.path.join(FILE_MINGGUAN, preek_file) if preek_file else None

med_raw = input('Mededelingen toevoegen? (y/n, default n): ').strip().lower()
include_mededelingen = med_raw.startswith('y')
mededelingen_language = None
if include_mededelingen:
    med_lang = input('Taal (nl/id, default id): ').strip().lower()
    mededelingen_language = med_lang if med_lang in ('nl', 'id') else 'id'

# ---------------------------------------------------------------------------
# Run the same generator as the webapp
# ---------------------------------------------------------------------------
print()
print('Bezig met genereren, dit kan even duren...')

with open(main_path, 'rb') as f:
    excel_bytes = f.read()

preek_bytes = None
if preek_path:
    with open(preek_path, 'rb') as f:
        preek_bytes = f.read()

work_dir = tempfile.mkdtemp(prefix='liturgi_manual_')
try:
    result = app._run_liturgi(
        excel_bytes,
        preek_bytes,
        work_dir,
        include_mededelingen=include_mededelingen,
        mededelingen_language=mededelingen_language,
    )
finally:
    # The generated files are already copied; cleanup is safe after that
    pass

# ---------------------------------------------------------------------------
# Copy the requested outputs to the file mingguan folder
# ---------------------------------------------------------------------------
print()
print('Gegenereerde bestanden:')

copied = []

def _copy_output(key, wanted, label):
    if not wanted:
        return
    src = result.get(key)
    if not src or not os.path.isfile(src):
        print(f'  ! {label} niet gevonden in output')
        return
    dst = os.path.join(FILE_MINGGUAN, os.path.basename(src))
    shutil.copy2(src, dst)
    print(f'  -> {dst}')
    copied.append(dst)

_copy_output('liturgieA', want_a, 'LiturgieA')
_copy_output('liturgieB', want_b, 'LiturgieB')
_copy_output('liturgieP', want_p, 'LiturgieP')

if not copied:
    print('  Geen bestanden gekopieerd.')

warnings = result.get('warnings') or []
if warnings:
    print()
    print('Waarschuwingen:')
    for w in warnings:
        print(f'  - {w}')

print()
print('Klaar.')

# Clean up the temp work directory
shutil.rmtree(work_dir, ignore_errors=True)
