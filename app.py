"""
GKIN Amstelveen Mededelingen Generator – Web App
"""

import base64
import builtins
import os
import shutil
import tempfile
import threading
import traceback
import uuid
from datetime import datetime
from typing import Dict, Optional
from flask import Flask, render_template, request, send_file, jsonify, session, redirect, url_for
from werkzeug.utils import secure_filename

import re
import requests as _requests
from bs4 import BeautifulSoup as _BS
import pandas as pd
import dropbox

from data_sources.dropbox_reader import DropboxExcelReader
from data_sources.scipio_scraper import ScipioScraper
from data_sources.preekroster_scraper import PreekrosterScraper
from data_sources.email_reader import OutlookCollecteReader, get_token_cache_json
from bulletin_generator import BulletinGenerator
from voorlees_generator import VoorleesGenerator

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10MB max upload
app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'gkin-amstelveen-secret-2026')
SITE_PASSWORD = os.environ.get('SITE_PASSWORD', '')

def _password_required(f):
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        if SITE_PASSWORD and not session.get('authenticated'):
            return redirect(url_for('login_page', next=request.path))
        return f(*args, **kwargs)
    return decorated

@app.route('/login', methods=['GET', 'POST'])
def login_page():
    error = None
    if request.method == 'POST':
        if request.form.get('password') == SITE_PASSWORD:
            session['authenticated'] = True
            next_url = request.args.get('next') or '/'
            return redirect(next_url)
        error = 'Onjuist wachtwoord. Probeer opnieuw.'
    return render_template('login.html', error=error,
                           next=request.args.get('next', '/'))

@app.route('/logout')
def logout():
    session.clear()
    return redirect('/')

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), 'output', '_uploads')
os.makedirs(UPLOAD_DIR, exist_ok=True)

# Cache takenrooster on startup
_takenrooster_cache = None

# Cache bijbelbasics KND data (slug -> {title, verse, date_day, date_month})
_knd_cache = None
_KND_DUTCH_MONTHS = {
    'januari':1,'februari':2,'maart':3,'april':4,'mei':5,'juni':6,
    'juli':7,'augustus':8,'september':9,'oktober':10,'november':11,'december':12
}


def _load_knd_cache():
    global _knd_cache
    if _knd_cache is not None:
        return _knd_cache
    try:
        r = _requests.get('https://www.bijbelbasics.nl/programma/', timeout=10)
        soup = _BS(r.text, 'html.parser')
        lines = [l.strip() for l in soup.get_text(separator='\n').split('\n') if l.strip()]
        date_pat = re.compile(
            r'^(?:zondag|zaterdag|vrijdag)\s+(\d+)\s+(\w+)\s+(\d{4})$', re.I)
        verse_start = re.compile(r'^([A-Z][a-zA-Zë]+)\s+(\d+:\d+)$')
        verse_end   = re.compile(r'^-\s*(\d+:\d+)$')
        entries = []
        for i, line in enumerate(lines):
            m = date_pat.match(line)
            if not m:
                continue
            day       = int(m.group(1))
            month_num = _KND_DUTCH_MONTHS.get(m.group(2).lower(), 0)
            year      = int(m.group(3))
            if not month_num:
                continue
            # title is the line before the date
            title = lines[i - 1] if i > 0 else ''
            # verse: next two lines are "Book X:Y" and "- X:Z"
            verse = ''
            if i + 2 < len(lines):
                ms = verse_start.match(lines[i + 1])
                me = verse_end.match(lines[i + 2])
                if ms and me:
                    verse = f"{ms.group(1)} {ms.group(2)} - {me.group(1)}"
                elif ms:
                    verse = f"{ms.group(1)} {ms.group(2)}"
            entries.append({'title': title, 'verse': verse,
                            'day': day, 'month': month_num, 'year': year})
        _knd_cache = entries
    except Exception as e:
        # Don't cache on error - allow retry on next request
        print(f'[KND] Error loading from bijbelbasics.nl: {e}')
        return []
    return _knd_cache


def _get_takenrooster():
    global _takenrooster_cache
    if _takenrooster_cache is None:
        reader = DropboxExcelReader()
        _takenrooster_cache = reader.get_takenrooster()
    return _takenrooster_cache


def _get_takenrooster_and_render_page():
    taken = _get_takenrooster()
    dutch_months = [
        'januari', 'februari', 'maart', 'april', 'mei', 'juni',
        'juli', 'augustus', 'september', 'oktober', 'november', 'december'
    ]
    dutch_days = ['maandag', 'dinsdag', 'woensdag', 'donderdag',
                  'vrijdag', 'zaterdag', 'zondag']

    today = datetime.now().date()
    dates = []
    for entry in taken['entries']:
        d = entry['date']
        if hasattr(d, 'date'):
            d_date = d.date()
        else:
            d_date = d
        if d_date < today:
            continue
        label = (f"{dutch_days[d.weekday()].capitalize()} {d.day} {dutch_months[d.month - 1]} {d.year}"
                 f"  —  {entry['predikant']}")
        if entry.get('opmerking'):
            opm = entry['opmerking'].split('OLE')[0].strip().rstrip(',').strip()
            if opm:
                label += f" ({opm})"
            elif 'OLE' in entry.get('opmerking', ''):
                label += " (OLE)"
        dates.append({
            'value': d.strftime('%Y-%m-%d'),
            'label': label,
            'predikant': entry['predikant'],
            'ovd': entry.get('ovd', ''),
            'predikant_email': entry.get('predikant_email', ''),
            'beamer_email':    entry.get('beamer_email', ''),
            'beamer':          entry.get('beamer', ''),
            '1eo':             entry.get('1eo', ''),
            '1eo_email':       entry.get('1eo_email', ''),
        })

    return render_template('mededelingen.html', dates=dates)


@app.route('/')
def home():
    return render_template('home.html')


@app.route('/mededelingen')
@_password_required
def mededelingen_index():
    return _get_takenrooster_and_render_page()


@app.route('/preekbevestiging')
@_password_required
def preekbevestiging_index():
    taken = _get_takenrooster()
    dutch_months = [
        'januari', 'februari', 'maart', 'april', 'mei', 'juni',
        'juli', 'augustus', 'september', 'oktober', 'november', 'december'
    ]
    dutch_days = ['maandag', 'dinsdag', 'woensdag', 'donderdag',
                  'vrijdag', 'zaterdag', 'zondag']
    today = datetime.now().date()
    dates = []
    for entry in taken['entries']:
        d = entry['date']
        d_date = d.date() if hasattr(d, 'date') else d
        if d_date < today:
            continue
        label = (f"{dutch_days[d.weekday()].capitalize()} {d.day} {dutch_months[d.month - 1]} {d.year}"
                 f"  —  {entry['predikant']}")
        opm = entry.get('opmerking', '')
        if opm:
            opm_clean = opm.split('OLE')[0].strip().rstrip(',').strip()
            if opm_clean:
                label += f" ({opm_clean})"
            elif 'OLE' in opm:
                label += " (OLE)"
        dates.append({
            'value':        d.strftime('%Y-%m-%d'),
            'label':        label,
            'predikant':    entry['predikant'],
            'ovd':          entry.get('ovd', ''),
            'ovd_email':    entry.get('ovd_email', ''),
            '1eo':          entry.get('1eo', ''),
            '1eo_email':    entry.get('1eo_email', ''),
            'beamer':       entry.get('beamer', ''),
            'beamer_email': entry.get('beamer_email', ''),
            'dag':          entry.get('dag', ''),
            'tijd':            entry.get('tijd', '10:30') or '10:30',
            'opmerking':       opm,
            'predikant_email': entry.get('predikant_email', ''),
            'muziek':          entry.get('muziek', ''),
            'voorzangers':  entry.get('voorzangers', ''),
            'multimedia':   entry.get('multimedia', ''),
            'knd':          entry.get('knd', ''),
            'tieners':      entry.get('tieners', ''),
            'unresolved_names': entry.get('unresolved_names', []),
        })
    return render_template('preekbevestiging.html', dates=dates)


@app.route('/generate-preekbevestiging', methods=['POST'])
def generate_preekbevestiging():
    from preekbevestiging_generator import generate as gen_doc
    data     = request.get_json(force=True)
    iso_date = data.get('date', '')
    if not iso_date:
        return jsonify({'error': 'no date'}), 400
    taken = _get_takenrooster()
    entry = next((e for e in taken['entries']
                  if e['date'].strftime('%Y-%m-%d') == iso_date), None)
    if not entry:
        return jsonify({'error': 'date not found'}), 404
    d = datetime.strptime(iso_date, '%Y-%m-%d')
    filename = f"Preekbevestiging_{d.strftime('%Y%m%d')}_{entry.get('predikant','').replace(' ','_')}.docx"
    out_path = os.path.join(UPLOAD_DIR, filename)
    songs = data.get('songs', [])
    gen_doc(entry, iso_date, out_path, songs=songs)
    return send_file(out_path, as_attachment=True,
                     download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')


@app.route('/fetch-liederen', methods=['POST'])
def fetch_liederen():
    from data_sources.email_reader import OutlookCollecteReader
    data     = request.get_json(force=True)
    iso_date = data.get('date', '')
    if not iso_date:
        return jsonify({'error': 'no date'}), 400
    try:
        d = datetime.strptime(iso_date, '%Y-%m-%d')
    except ValueError:
        return jsonify({'error': 'invalid date'}), 400
    try:
        reader = OutlookCollecteReader()
        if not reader.is_authenticated():
            return jsonify({'error': 'not_authenticated', 'songs': ['','','','','','',''], 'not_found': ['Niet ingelogd bij Outlook']}), 200
        result = reader.fetch_liederen(target_date=d)
        return jsonify(result)
    except Exception as e:
        return jsonify({'error': str(e), 'songs': ['','','','','','',''], 'not_found': [str(e)]}), 200


@app.route('/fetch-language', methods=['POST'])
def fetch_language():
    from data_sources.preekroster_scraper import PreekrosterScraper
    data     = request.get_json(force=True)
    iso_date = data.get('date', '')
    if not iso_date:
        return jsonify({'error': 'no date'}), 400
    try:
        d = datetime.strptime(iso_date, '%Y-%m-%d')
    except ValueError:
        return jsonify({'error': 'invalid date'}), 400

    dutch_months_short = ['jan','feb','mrt','apr','mei','jun',
                          'jul','aug','sep','okt','nov','dec']
    date_key = f"{d.day} {dutch_months_short[d.month - 1]}"  # e.g. "17 mei"

    try:
        from datetime import timedelta
        scraper = PreekrosterScraper()
        # Use d - 6 days as mededelingen_date so d always falls inside the window
        roster  = scraper.get_preekroster(mededelingen_date=d - timedelta(days=6))
        for entry in roster.get('am_table', []):
            raw = entry.get('date', '')
            # raw may be "28 jun" or "28 jun (Pinksteren)"
            if raw.startswith(date_key):
                return jsonify({'language': entry.get('language', '')})
    except Exception:
        pass
    return jsonify({'language': ''})


@app.route('/fetch-knd-thema', methods=['POST'])
def fetch_knd_thema():
    data = request.get_json(force=True)
    date_str = data.get('date', '')
    if not date_str:
        return jsonify({'error': 'no date'}), 400
    try:
        d = datetime.strptime(date_str, '%Y-%m-%d')
    except ValueError:
        return jsonify({'error': 'invalid date'}), 400
    entries = _load_knd_cache()
    for entry in entries:
        if entry['day'] == d.day and entry['month'] == d.month:
            return jsonify({'title': entry['title'], 'verse': entry['verse']})
    return jsonify({'title': '', 'verse': ''})


@app.route('/uploads/<path:filename>')
def serve_upload(filename):
    """Serve files from the uploads directory (for QR image previews)."""
    return send_file(os.path.join(UPLOAD_DIR, filename))


@app.route('/generate', methods=['POST'])
def generate():
    date_str = request.form.get('date')
    if not date_str:
        return jsonify({'error': 'No date selected'}), 400

    selected_date = datetime.strptime(date_str, '%Y-%m-%d')

    try:
        # 1. Get takenrooster entry
        taken = _get_takenrooster()
        entry = None
        for e in taken['entries']:
            if e['date'] == selected_date:
                entry = e
                break
        if entry is None:
            return jsonify({'error': f'No takenrooster entry for {date_str}'}), 404

        # 2. Mededelingen from Dropbox
        reader = DropboxExcelReader()
        meded = reader.get_mededelingen(mededelingen_date=selected_date)

        # 3. Birthdays from Scipio
        scipio = ScipioScraper()
        bdays = scipio.get_birthday_list(mededelingen_date=selected_date)

        # 4. Preekroster from GKIN website
        preek = PreekrosterScraper()
        roster = preek.get_preekroster(mededelingen_date=selected_date)

        # 5. Collect user-supplied manual fields from form
        def fv(key): return request.form.get(key, '').strip()

        # Parse activiteiten rows (datum_0, activiteit_0, locatie_0, ...)
        activiteiten = []
        i = 0
        while request.form.get(f'act_datum_{i}') is not None:
            row = {
                'datum': fv(f'act_datum_{i}'),
                'tijd': fv(f'act_tijd_{i}'),
                'activiteit': fv(f'act_activiteit_{i}'),
                'olv': fv(f'act_olv_{i}'),
                'locatie': fv(f'act_locatie_{i}'),
            }
            if any(row.values()):
                activiteiten.append(row)
            i += 1

        # Handle QR image uploads — prefer file upload, fall back to email-fetched path
        def save_upload(field_name, fallback_field):
            f = request.files.get(field_name)
            if f and f.filename:
                ext = os.path.splitext(secure_filename(f.filename))[1].lower()
                path = os.path.join(UPLOAD_DIR, f'{uuid.uuid4().hex}{ext}')
                f.save(path)
                return path
            # Use email-fetched filename if present
            fname = fv(fallback_field)
            if fname:
                full = os.path.join(UPLOAD_DIR, os.path.basename(fname))
                if os.path.exists(full):
                    return full
            return ''

        user_data = {
            'overdenking_predikant':      fv('overdenking_predikant'),
            'overdenking_thema':          fv('overdenking_thema'),
            'overdenking_schriftlezing':  fv('overdenking_schriftlezing'),
            'overdenking_content':        fv('overdenking_content'),
            'collecte_contant':           fv('collecte_contant'),
            'collecte_bonnen':            fv('collecte_bonnen'),
            'collecte_bank':              fv('collecte_bank'),
            'collecte_tikkie':            fv('collecte_tikkie'),
            'collecte_ole':               fv('collecte_ole'),
            'bezoekers_volwassenen':      fv('bezoekers_volwassenen'),
            'bezoekers_kinderen':         fv('bezoekers_kinderen'),
            'dankoffer_url':              fv('dankoffer_url'),
            'dankoffer_qr':               save_upload('dankoffer_qr_file', 'dankoffer_qr_path'),
            'ole_url':                    fv('ole_url'),
            'ole_qr':                     save_upload('ole_qr_file', 'ole_qr_path'),
            'activiteiten':               activiteiten,
            'opbrengst_entries':          [],  # filled below if email entries were fetched
        }

        # Re-assemble multi-entry opbrengst from hidden JSON field (set by JS after email fetch)
        import json as _json
        opbrengst_json = request.form.get('opbrengst_entries_json', '')
        if opbrengst_json:
            try:
                user_data['opbrengst_entries'] = _json.loads(opbrengst_json)
            except Exception:
                pass

        # 6. Generate bulletin
        gen = BulletinGenerator()
        filepath = gen.generate(selected_date, entry, taken['entries'],
                                meded, bdays, roster, user_data)

        filename = os.path.basename(filepath)
        return send_file(filepath, as_attachment=True, download_name=filename,
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/generate-voorlees', methods=['POST'])
def generate_voorlees():
    date_str = request.form.get('date')
    if not date_str:
        return jsonify({'error': 'No date selected'}), 400
    selected_date = datetime.strptime(date_str, '%Y-%m-%d')
    try:
        taken = _get_takenrooster()
        entry = None
        for e in taken['entries']:
            if e['date'] == selected_date:
                entry = e
                break
        if entry is None:
            return jsonify({'error': f'Geen rooster gevonden voor {date_str}'}), 404

        reader = DropboxExcelReader()
        meded  = reader.get_mededelingen(mededelingen_date=selected_date)

        def fv(key): return request.form.get(key, '').strip()

        def save_upload_v(field_name, fallback_field):
            f = request.files.get(field_name)
            if f and f.filename:
                ext  = os.path.splitext(secure_filename(f.filename))[1].lower()
                path = os.path.join(UPLOAD_DIR, f'{uuid.uuid4().hex}{ext}')
                f.save(path)
                return path
            fname = fv(fallback_field)
            if fname:
                full = os.path.join(UPLOAD_DIR, os.path.basename(fname))
                if os.path.exists(full):
                    return full
            return ''

        import json as _json
        opbrengst_entries = []
        opbrengst_json = request.form.get('opbrengst_entries_json', '')
        if opbrengst_json:
            try:
                opbrengst_entries = _json.loads(opbrengst_json)
            except Exception:
                pass

        user_data = {
            'collecte_contant':      fv('collecte_contant'),
            'collecte_bonnen':       fv('collecte_bonnen'),
            'collecte_bank':         fv('collecte_bank'),
            'collecte_tikkie':       fv('collecte_tikkie'),
            'collecte_ole':          fv('collecte_ole'),
            'bezoekers_volwassenen': fv('bezoekers_volwassenen'),
            'bezoekers_kinderen':    fv('bezoekers_kinderen'),
            'dankoffer_url':         fv('dankoffer_url'),
            'dankoffer_qr':          save_upload_v('dankoffer_qr_file', 'dankoffer_qr_path'),
            'ole_url':               fv('ole_url'),
            'ole_qr':                save_upload_v('ole_qr_file', 'ole_qr_path'),
            'opbrengst_entries':     opbrengst_entries,
        }

        # Extract welkomstwoord paragraphs from the mededelingen template
        welkom_paras = _extract_welkom_paragraphs(selected_date, entry, meded)

        gen      = VoorleesGenerator()
        filepath = gen.generate(selected_date, entry, meded, user_data,
                                welkom_paragraphs=welkom_paras)
        filename = os.path.basename(filepath)
        return send_file(filepath, as_attachment=True, download_name=filename,
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


def _extract_welkom_paragraphs(selected_date: datetime, entry: dict, meded: dict) -> list:
    """Read the Welkomstwoord paragraphs from the mededelingen Word template."""
    from bulletin_generator import BulletinGenerator, TEMPLATE_PATH
    from docx import Document as _Document
    try:
        doc = _Document(TEMPLATE_PATH)
        paras = []
        in_welkom = False
        for p in doc.paragraphs:
            if p.style.name.startswith('Heading') and 'welkomstwoord' in p.text.lower():
                in_welkom = True
                continue
            if in_welkom:
                if p.style.name.startswith('Heading'):
                    break
                txt = p.text.strip()
                if txt:
                    paras.append(txt)
        # Fill in dynamic date/predikant/ovd from the template text
        predikant = entry.get('predikant', '')
        ovd = entry.get('ovd', '')
        dutch_months = ['januari','februari','maart','april','mei','juni',
                        'juli','augustus','september','oktober','november','december']
        dutch_days = ['maandag','dinsdag','woensdag','donderdag','vrijdag','zaterdag','zondag']
        date_str = f"{selected_date.day} {dutch_months[selected_date.month-1]} {selected_date.year}"
        day_name = dutch_days[selected_date.weekday()]
        result = []
        for p in paras:
            if 'Vandaag,' in p:
                p = f"Vandaag, {day_name} {date_str}, gaat voor {predikant}. De ouderling van dienst is {ovd}. Als u vragen heeft, kunt u de ouderling van dienst aanspreken."
            result.append(p)
        return result
    except Exception:
        return []


@app.route('/get-mededelingen', methods=['POST'])
def get_mededelingen_data():
    """Fetch mededelingen for a date and parse activities from both sections."""
    date_str = request.form.get('date')
    if not date_str:
        return jsonify({'error': 'No date'}), 400
    selected_date = datetime.strptime(date_str, '%Y-%m-%d')
    try:
        reader = DropboxExcelReader()
        meded = reader.get_mededelingen(mededelingen_date=selected_date)
        activities = _parse_activities_from_mededelingen(meded, selected_date)
        return jsonify({
            'regionale_nl': meded.get('regionale_nl', ''),
            'landelijke_nl': meded.get('landelijke_nl', ''),
            'activities': activities,
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


def _parse_activities_from_mededelingen(meded: dict, selected_date: datetime = None) -> list:
    """Parse activity rows from regionale + landelijke mededelingen blocks.

    Scans each paragraph block for Dutch date mentions (e.g. '21 juni', 'zondag 24 mei 2026').
    Each block that contains a date becomes one activity row.
    """
    import re

    MONTHS = r'(?:januari|februari|maart|april|mei|juni|juli|augustus|september|oktober|november|december|jan|feb|mrt|apr|jun|jul|aug|sep|okt|nov|dec)'
    DAYS   = r'(?:maandag|dinsdag|woensdag|donderdag|vrijdag|zaterdag|zondag|ma|di|wo|do|vr|za|zo)'

    # Matches: optional day-name, day-number, month-name, optional year
    date_pat = re.compile(
        rf'(?:{DAYS}\s+)?(\d{{1,2}}\s+{MONTHS}\w*(?:\s+\d{{4}})?)',
        re.IGNORECASE
    )
    # Matches time like "10.30 uur" or "10:30 uur"
    time_pat = re.compile(r'\b(\d{1,2}[:.]\d{2})\s*uur\b', re.IGNORECASE)
    # Matches "o.l.v. Name", "onder leiding van Name", "geleid door Name", "waarin Name voorgaat"
    # Stop at: " op [weekday/date]", " voorgaat", " spreekt", " zal", comma, newline
    olv_pat  = re.compile(
        r'(?:o\.?l\.?v\.?|onder leiding van|geleid door|waarbij\s+|waarin\s+)'
        r'(.*?)'
        r'(?=\s+(?:voorgaat|spreekt|zal\b|op\s+(?:zondag|maandag|dinsdag|woensdag|donderdag|vrijdag|zaterdag|\d))|[,\n]|$)',
        re.IGNORECASE
    )
    # Location: "in VenueName Number" or "te City" — require uppercase start to avoid "in rouw", "bijwonen"
    loc_pat  = re.compile(
        r'(?:\bin\s+([A-Z][A-Za-zÀ-ÿ\s]+\s+\d+)'  # "in Bouwerij 52" — uppercase start
        r'|\bte\s+([A-Z][A-Za-zÀ-ÿ\s]{2,20}?)(?=[,.\n ]|$))'  # "te Den Haag"
    )

    # Keywords that indicate a block is NOT a scheduled activity
    SKIP_KEYWORDS = re.compile(
        r'\b(overlijden|uitvaart|geboren|condoleance|in rouw|bijzondere gelegenheid|leden in)\b',
        re.IGNORECASE
    )

    activities = []
    sources = [meded.get('regionale_nl', ''), meded.get('landelijke_nl', '')]

    for source in sources:
        blocks = [b.strip() for b in source.split('\n\n') if b.strip()]
        for block in blocks:
            # Skip obituaries and non-activity blocks
            if SKIP_KEYWORDS.search(block):
                continue

            flat = ' '.join(block.split('\n'))

            dm = date_pat.search(flat)
            if not dm:
                continue

            datum = re.sub(r'\s+\d{4}$', '', dm.group(1).strip())
            # Shorten month names to 3-letter abbreviations: "juni" → "jun", "maart" → "mrt"
            MONTH_ABBR = {'januari':'jan','februari':'feb','maart':'mrt','april':'apr',
                          'juni':'jun','juli':'jul','augustus':'aug','september':'sep',
                          'oktober':'okt','november':'nov','december':'dec'}
            for full, abbr in MONTH_ABBR.items():
                datum = re.sub(rf'\b{full}\b', abbr, datum, flags=re.IGNORECASE)

            # Filter out dates in the past relative to selected_date
            if selected_date:
                MONTH_NUM = {'jan':1,'feb':2,'mrt':3,'mar':3,'apr':4,'mei':5,'jun':6,
                             'jul':7,'aug':8,'sep':9,'okt':10,'nov':11,'dec':12}
                dm2 = re.match(r'(\d{1,2})\s+(\w+)', datum)
                if dm2:
                    try:
                        day = int(dm2.group(1))
                        mon = MONTH_NUM.get(dm2.group(2).lower()[:3], 0)
                        yr  = selected_date.year
                        act_date = datetime(yr, mon, day) if mon else None
                        if act_date and act_date.date() < selected_date.date():
                            continue
                    except Exception:
                        pass

            # Time
            tm = time_pat.search(flat)
            tijd = tm.group(1).replace('.', ':') if tm else ''

            # o.l.v.
            ovm = olv_pat.search(flat)
            if ovm:
                raw_olv = ovm.group(1).strip()
                # Replace abbreviation dots temporarily (titles + single/double initials)
                raw_olv = re.sub(r'\b(ds|br|zr|dr|drs|mr|ir|prof)\.\s*', r'\1___', raw_olv, flags=re.IGNORECASE)
                raw_olv = re.sub(r'\b([A-Z])\.\s*', r'\1___', raw_olv)  # initials like S.M.
                raw_olv = re.split(r',|\.\s+', raw_olv)[0]
                raw_olv = raw_olv.replace('___', '. ').strip().rstrip('–- ').strip()
                olv = raw_olv
            else:
                olv = ''

            # Activity name: first line heading
            first_line = block.split('\n')[0].strip().rstrip(':–-').strip()
            activiteit = first_line

            # Location: prefer "in StreetName Number" then "te City"
            locatie = ''
            for lm in loc_pat.finditer(flat):
                cand = (lm.group(1) or lm.group(2) or '').strip().rstrip('.,')
                # Only keep short, clean strings
                if cand and len(cand) <= 30:
                    locatie = cand
                    break

            activities.append({
                'datum': datum,
                'tijd': tijd,
                'activiteit': activiteit,
                'olv': olv,
                'locatie': locatie,
            })

    return activities


@app.route('/fetch-email-collecte', methods=['POST'])
def fetch_email_collecte():
    """Fetch collecte URLs, QR images and amounts from recent Outlook emails."""
    try:
        reader = OutlookCollecteReader()
        if not reader.is_authenticated():
            return jsonify({
                'error': 'login_required',
                'message': 'Nog niet ingelogd bij Outlook. Klik op "Inloggen bij Outlook" eerst.'
            }), 401
        # Parse selected date from request
        date_str = request.json.get('date', '') if request.is_json else ''
        target_date = None
        if date_str:
            try:
                target_date = datetime.strptime(date_str, '%Y-%m-%d')
            except ValueError:
                pass
        data = reader.fetch_collecte_data(target_date=target_date, since_days=60)
        return jsonify(data)
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/fetch-email-overdenking', methods=['POST'])
def fetch_email_overdenking():
    """Fetch overdenking content from scriba email attachment."""
    try:
        reader = OutlookCollecteReader()
        if not reader.is_authenticated():
            return jsonify({'error': 'login_required'}), 401
        date_str = request.json.get('date', '') if request.is_json else ''
        target_date = None
        if date_str:
            try:
                target_date = datetime.strptime(date_str, '%Y-%m-%d')
            except ValueError:
                pass
        data = reader.fetch_overdenking(target_date=target_date, since_days=14)
        return jsonify(data)
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/fetch-email-opbrengst', methods=['POST'])
def fetch_email_opbrengst():
    """Fetch collecte opbrengsten amounts and bezoekers from Outlook emails."""
    try:
        reader = OutlookCollecteReader()
        if not reader.is_authenticated():
            return jsonify({'error': 'login_required'}), 401
        date_str = request.json.get('date', '') if request.is_json else ''
        target_date = None
        if date_str:
            try:
                target_date = datetime.strptime(date_str, '%Y-%m-%d')
            except ValueError:
                pass
        data = reader.fetch_opbrengst_data(target_date=target_date, since_days=60)
        return jsonify(data)
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/email-login-start', methods=['POST'])
def email_login_start():
    """Start device-code OAuth2 flow; returns user_code and verification_uri."""
    try:
        reader = OutlookCollecteReader()
        flow   = reader.start_device_flow()
        return jsonify(flow)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/email-login-complete', methods=['POST'])
def email_login_complete():
    """Poll for completed device-code login."""
    try:
        reader  = OutlookCollecteReader()
        success = reader.complete_device_flow()
        if success:
            return jsonify({'status': 'ok'})
        return jsonify({'status': 'pending'}), 202
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/email-auth-status', methods=['GET'])
def email_auth_status():
    """Check whether a valid Outlook token is already cached."""
    try:
        reader = OutlookCollecteReader()
        return jsonify({'authenticated': reader.is_authenticated()})
    except Exception as e:
        return jsonify({'authenticated': False, 'error': str(e)})


@app.route('/email-token-export', methods=['GET'])
def email_token_export():
    """Export current MSAL token cache as JSON string."""
    try:
        token_json = get_token_cache_json()
        if not token_json or token_json == '{}':
            return jsonify({'error': 'Geen token gevonden — log eerst in bij Outlook.'}), 404
        return jsonify({'token_cache': token_json,
                        'instructions': 'Gebruik /email-token-download voor een schoon tekstbestand.'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/email-token-download', methods=['GET'])
def email_token_download():
    """Download the raw MSAL token cache as a plain .txt file.
    Open the file and paste its entire contents into Railway > Variables > MSAL_TOKEN_CACHE.
    """
    try:
        from flask import Response
        token_json = get_token_cache_json()
        if not token_json or token_json == '{}':
            return jsonify({'error': 'Geen token gevonden — log eerst in bij Outlook.'}), 404
        return Response(
            token_json,
            mimetype='text/plain',
            headers={'Content-Disposition': 'attachment; filename=msal_token_cache.txt'}
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/refresh-dates', methods=['POST'])
def refresh_dates():
    """Force refresh of takenrooster cache."""
    global _takenrooster_cache
    _takenrooster_cache = None
    return jsonify({'status': 'ok'})


# ---------------------------------------------------------------------------
# Liturgie routes
# ---------------------------------------------------------------------------

_LITURGIE_BASE    = os.path.dirname(__file__)
_LITURGIE_CACHE   = os.path.join(_LITURGIE_BASE, 'bible_cache')
_LITURGIE_LOGO    = os.path.join(_LITURGIE_BASE, 'logo.png')
_LITURGIE_PHONE   = os.path.join(_LITURGIE_BASE, 'telephone.gif')
_LITURGIE_LOCK    = threading.Lock()
_LITURGIE_FILE_LOCKS: dict = {}

DROPBOX_APP_KEY_L      = os.getenv('DROPBOX_APP_KEY', '')
DROPBOX_APP_SECRET_L   = os.getenv('DROPBOX_APP_SECRET', '')
DROPBOX_REFRESH_L      = os.getenv('DROPBOX_REFRESH_TOKEN', '')
BIBLE_DROPBOX_PATH     = '/working folder/bible'
LOGO_DROPBOX_PATH_L    = '/working folder/logo.png'
PHONE_DROPBOX_PATH_L   = '/working folder/telephone.gif'


def _get_dbx_liturgie():
    import dropbox
    return dropbox.Dropbox(
        oauth2_refresh_token=DROPBOX_REFRESH_L,
        app_key=DROPBOX_APP_KEY_L,
        app_secret=DROPBOX_APP_SECRET_L,
    )


def _ensure_liturgie_file(remote_path: str, local_path: str):
    """Download a single file from Dropbox if not cached locally."""
    if os.path.exists(local_path):
        return
    try:
        dbx = _get_dbx_liturgie()
        _, resp = dbx.files_download(remote_path)
        os.makedirs(os.path.dirname(local_path), exist_ok=True)
        with open(local_path, 'wb') as f:
            f.write(resp.content)
    except Exception as e:
        print(f'Warning: could not download {remote_path}: {e}')


def _ensure_bible_file(filename: str):
    """Download a single bible chapter file from Dropbox on demand."""
    os.makedirs(_LITURGIE_CACHE, exist_ok=True)
    local_path = os.path.join(_LITURGIE_CACHE, filename)
    if os.path.exists(local_path):
        return
    with _LITURGIE_LOCK:
        if filename not in _LITURGIE_FILE_LOCKS:
            _LITURGIE_FILE_LOCKS[filename] = threading.Lock()
    with _LITURGIE_FILE_LOCKS[filename]:
        if os.path.exists(local_path):
            return
        dbx = _get_dbx_liturgie()
        dropbox_path = f'{BIBLE_DROPBOX_PATH}/{filename}'
        try:
            _, resp = dbx.files_download(dropbox_path)
            with open(local_path, 'wb') as f:
                f.write(resp.content)
        except Exception as e:
            raise FileNotFoundError(f'Bijbelbestand niet gevonden in Dropbox: {filename} ({e})')


def _run_liturgi(excel_bytes: bytes, preek_bytes, work_dir: str) -> dict:
    _ensure_liturgie_file(LOGO_DROPBOX_PATH_L, _LITURGIE_LOGO)
    _ensure_liturgie_file(PHONE_DROPBOX_PATH_L, _LITURGIE_PHONE)
    os.makedirs(_LITURGIE_CACHE, exist_ok=True)

    file_mingguan = os.path.join(work_dir, 'file mingguan')
    os.makedirs(file_mingguan, exist_ok=True)

    with open(os.path.join(file_mingguan, 'Main Liturgy file.xlsx'), 'wb') as f:
        f.write(excel_bytes)
    if preek_bytes:
        with open(os.path.join(file_mingguan, 'Preek.docx'), 'wb') as f:
            f.write(preek_bytes)

    for name, src in [('bible', _LITURGIE_CACHE), ('logo.png', _LITURGIE_LOGO), ('telephone.gif', _LITURGIE_PHONE)]:
        link = os.path.join(work_dir, name)
        if not os.path.exists(link) and os.path.exists(src):
            os.symlink(src, link)

    src_path = os.path.join(_LITURGIE_BASE, 'liturgi_core.py')
    with open(src_path, 'r', encoding='utf-8') as fh:
        patched = fh.read().replace(
            'dir_path = os.path.dirname(os.path.realpath(__file__))',
            f'dir_path = {repr(work_dir)}'
        )

    ns = {
        '__file__': src_path,
        '__name__': '__liturgi_run__',
        '__builtins__': builtins,
        '_ensure_bible_file': _ensure_bible_file,
    }
    try:
        exec(compile(patched, src_path, 'exec'), ns)
    except SystemExit:
        pass

    result = {}
    for fname in os.listdir(file_mingguan):
        if fname.startswith('LiturgieA') and fname.endswith('.docx'):
            result['liturgieA'] = os.path.join(file_mingguan, fname)
            result['liturgieA_name'] = fname
        elif fname.startswith('LiturgieB') and fname.endswith('.docx'):
            result['liturgieB'] = os.path.join(file_mingguan, fname)
            result['liturgieB_name'] = fname
        elif fname.startswith('LiturgieP') and fname.endswith('.pptx'):
            result['liturgieP'] = os.path.join(file_mingguan, fname)
            result['liturgieP_name'] = fname
    return result


@app.route('/liturgie')
def liturgie_index():
    return render_template('liturgie.html')


@app.route('/liturgie/generate', methods=['POST'])
def liturgie_generate():
    excel_file = request.files.get('excel_file')
    if not excel_file or not excel_file.filename:
        return jsonify({'error': 'Excel bestand is verplicht.'}), 400

    excel_bytes = excel_file.read()
    preek_file  = request.files.get('preek_file')
    preek_bytes = preek_file.read() if preek_file and preek_file.filename else None
    want_a = request.form.get('want_a', '1') == '1'
    want_b = request.form.get('want_b', '1') == '1'
    want_p = request.form.get('want_p', '1') == '1'

    work_dir = tempfile.mkdtemp(prefix='liturgi_')
    try:
        result = _run_liturgi(excel_bytes, preek_bytes, work_dir)
        if not result:
            return jsonify({'error': 'Geen output bestanden gegenereerd.'}), 500

        files_out = []
        for key, want in [('liturgieA', want_a), ('liturgieB', want_b), ('liturgieP', want_p)]:
            if want and key in result:
                with open(result[key], 'rb') as fh:
                    files_out.append({'name': result[f'{key}_name'], 'data': base64.b64encode(fh.read()).decode('ascii')})

        if not files_out:
            return jsonify({'error': 'Geen geselecteerde bestanden gevonden in output.'}), 500
        return jsonify(files_out)

    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


# ---------------------------------------------------------------------------
# Liturgie Auto-fill Endpoint
# ---------------------------------------------------------------------------

DANKOFFER_DROPBOX_PATH = '/working folder/Dankoffer.xlsx'

def _get_dankoffer_verse(dbx, service_date: datetime, mark_as_used: bool = True) -> Optional[Dict]:
    """
    Read Dankoffer.xlsx from Dropbox and return the next available verse.

    Expected Dankoffer.xlsx format:
    - Column A: Bible verse reference (e.g., "Psalmen 50:14-15")
    - Column B: Bible text (the actual verse content)
    - Column C: Date Used (blank = not used yet, or shows date when used)

    Logic:
    1. Find the first row where Column C (Date Used) is blank/empty
    2. Return that verse
    3. If mark_as_used=True, update Column C with the service date
    4. If all verses are used, reset all dates and use the first verse

    Returns dict with: book, chapter, verse_start, verse_end, full_text, row_index, total_count
    """
    try:
        _, resp = dbx.files_download(DANKOFFER_DROPBOX_PATH)
        df = pd.read_excel(BytesIO(resp.content), header=None)

        # Read verses and their used dates from Column C (index 2)
        verses_data = []
        for idx, row in df.iterrows():
            verse = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
            # Column C (index 2) contains Date Used
            date_used = str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) else ''

            if verse and verse.lower() not in ('nan', 'none', ''):
                verses_data.append({
                    'verse': verse,
                    'date_used': date_used if date_used.lower() not in ('nan', 'none', '') else '',
                    'row_idx': idx
                })

        if not verses_data:
            return None

        # Format the service date for comparison (YYYY-MM-DD)
        service_date_str = service_date.strftime('%Y-%m-%d')

        # STEP 1: Check if this service date is already assigned to a verse
        existing_assignment = None
        for v in verses_data:
            if v['date_used'] == service_date_str:
                existing_assignment = v
                break

        if existing_assignment:
            # This date already has a verse assigned - use it again
            selected = existing_assignment
            reset_needed = False
            already_assigned = True
        else:
            # STEP 2: Find the first unused verse (blank date)
            already_assigned = False
            unused_verses = [v for v in verses_data if not v['date_used']]

            if unused_verses:
                # Use the first unused verse
                selected = unused_verses[0]
                reset_needed = False
            else:
                # STEP 3: All verses are used - find the one with the OLDEST date
                # Parse dates and find the minimum (oldest)
                dated_verses = []
                for v in verses_data:
                    try:
                        if v['date_used']:
                            parsed_date = datetime.strptime(v['date_used'], '%Y-%m-%d')
                            dated_verses.append({**v, 'parsed_date': parsed_date})
                    except ValueError:
                        # If date format is invalid, treat as very old
                        dated_verses.append({**v, 'parsed_date': datetime(1900, 1, 1)})

                if dated_verses:
                    # Sort by date and pick the oldest
                    dated_verses.sort(key=lambda x: x['parsed_date'])
                    selected = dated_verses[0]
                    reset_needed = True  # Indicates we're reusing an old verse
                else:
                    # Fallback - should not happen
                    selected = verses_data[0]
                    reset_needed = True

        verse_text = selected['verse']

        # Parse verse text like "Psalmen 50:14-15" or "Psalmen 50:14"
        import re
        match = re.match(r'^(.+?)\s+(\d+):(\d+)(?:-(\d+))?$', verse_text.strip())
        if not match:
            return None

        book = match.group(1).strip()
        chapter = match.group(2)
        verse_start = match.group(3)
        verse_end = match.group(4)  # None if single verse

        # Calculate unused count (for display purposes)
        if already_assigned:
            unused_count = len([v for v in verses_data if not v['date_used']])
        else:
            unused_count = len(unused_verses) if not reset_needed else 0

        result = {
            'book': book,
            'chapter': chapter,
            'verse_start': verse_start,
            'verse_end': verse_end,
            'full_text': verse_text,
            'row_index': selected['row_idx'] + 1,  # 1-indexed for user display
            'total_count': len(verses_data),
            'unused_count': unused_count,
            'reset_needed': reset_needed,
            'already_assigned': already_assigned,
            'date_assigned': selected['date_used'] if already_assigned else None
        }

        # Update the Dankoffer.xlsx to mark this verse as used (or update date if reusing)
        if mark_as_used:
            try:
                _mark_dankoffer_verse_as_used(dbx, selected['row_idx'], service_date, reset_needed)
                result['marked_as_used'] = True
            except Exception as e:
                print(f'[Dankoffer] Warning: Could not mark verse as used: {e}')
                result['marked_as_used'] = False

        return result

    except Exception as e:
        print(f'[Dankoffer] Error reading dankoffer file: {e}')
        return None


def _mark_dankoffer_verse_as_used(dbx, row_idx: int, service_date: datetime, reset_all: bool = False):
    """
    Update Dankoffer.xlsx in Dropbox to mark a verse as used (Column C).
    Updates just the selected row with the new date.
    If reset_all=True, it means we're reusing an old verse (just updating its date).
    """
    try:
        from io import BytesIO
        import openpyxl

        # Download current file
        _, resp = dbx.files_download(DANKOFFER_DROPBOX_PATH)
        wb = openpyxl.load_workbook(BytesIO(resp.content))
        ws = wb.active

        # Mark the selected verse as used (Column C = column 3)
        # row_idx is 0-indexed, Excel rows are 1-indexed
        excel_row = row_idx + 1  # Adjust if your file has a header row
        date_str = service_date.strftime('%Y-%m-%d')
        ws.cell(row=excel_row, column=3).value = date_str

        # Save to BytesIO
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        # Upload back to Dropbox
        dbx.files_upload(output.read(), DANKOFFER_DROPBOX_PATH, mode=dropbox.files.WriteMode('overwrite'))
        print(f'[Dankoffer] Marked verse at row {excel_row}, column C as used on {date_str}')

    except Exception as e:
        print(f'[Dankoffer] Error updating dankoffer file: {e}')
        raise


def _preview_dankoffer_verse(dbx, service_date: datetime) -> Optional[Dict]:
    """Preview which verse will be used without marking it as used."""
    return _get_dankoffer_verse(dbx, service_date, mark_as_used=False)


def _parse_verse_reference(verse_text: str) -> Optional[Dict]:
    """Parse a verse reference like 'Psalmen 50:14-15' or 'Psalmen 50:14'."""
    import re
    match = re.match(r'^(.+?)\s+(\d+):(\d+)(?:-(\d+))?$', verse_text.strip())
    if not match:
        return None

    return {
        'book': match.group(1).strip(),
        'chapter': match.group(2),
        'verse_start': match.group(3),
        'verse_end': match.group(4),  # None if single verse
    }


@app.route('/liturgie/fill-data', methods=['POST'])
def liturgie_fill_data():
    """
    Auto-populate Main Liturgy file.xlsx with:
    - People on duty from Takenrooster (B4-B12)
    - Dankoffer verse from Dankoffer.xlsx (B21-E21)
    - Tikkie link from email (if available)

    Uses openpyxl to preserve formatting, images, and other content.
    Only modifies specific cells without changing the rest of the file.

    Returns a new Excel file with populated data and alert info.
    """
    excel_file = request.files.get('excel_file')
    if not excel_file or not excel_file.filename:
        return jsonify({'error': 'Excel bestand is verplicht.'}), 400

    excel_bytes = excel_file.read()

    try:
        # Use openpyxl to preserve formatting and images
        from io import BytesIO
        import openpyxl

        wb = openpyxl.load_workbook(BytesIO(excel_bytes))
        ws = wb['Data'] if 'Data' in wb.sheetnames else wb.active

        # Get date from B3 (row 3, column 2 in openpyxl 1-indexed)
        date_cell = ws.cell(row=3, column=2)
        date_val = date_cell.value

        if not date_val:
            return jsonify({'error': 'Geen datum gevonden in cel B3.'}), 400

        # Parse the date
        if isinstance(date_val, datetime):
            service_date = date_val
        else:
            try:
                service_date = datetime.strptime(str(date_val), '%d-%m-%Y') if '-' in str(date_val) else datetime.strptime(str(date_val), '%Y-%m-%d')
            except ValueError:
                try:
                    service_date = pd.to_datetime(str(date_val), dayfirst=True).to_pydatetime()
                except Exception:
                    return jsonify({'error': f'Ongeldige datum in cel B3: {date_val}'}), 400

        # Get takenrooster entry for this date
        taken = _get_takenrooster()
        entry = None
        for e in taken.get('entries', []):
            entry_date = e['date']
            if hasattr(entry_date, 'date'):
                entry_date = entry_date.date()
            else:
                entry_date = entry_date
            if entry_date == service_date.date():
                entry = e
                break

        if not entry:
            return jsonify({
                'error': f'Geen dienst gevonden in takenrooster voor {service_date.strftime("%d-%m-%Y")}'
            }), 404

        # Track what was already filled vs what we populated
        alerts = {
            'already_filled': [],
            'auto_populated': []
        }

        # Helper function to get cell value as string
        def get_cell_value(row, col):
            val = ws.cell(row=row, column=col).value
            return str(val).strip() if val else ''

        # Helper function to set cell value while preserving formatting
        def set_cell_value(row, col, value, field_name):
            cell = ws.cell(row=row, column=col)
            current_val = str(cell.value).strip() if cell.value else ''

            if current_val and current_val.lower() not in ('nan', 'none', ''):
                alerts['already_filled'].append(f'{field_name}: {current_val}')
                return False
            elif value:
                cell.value = value
                alerts['auto_populated'].append(f'{field_name}: {value}')
                return True
            return False

        # Define the fields to populate (excel row, field name, takenrooster key, excel column)
        # B4-B12 correspond to rows 4-12, column 2
        field_mapping = [
            (4, 'Voorganger', 'predikant', 2),      # B4
            (5, 'OvD', 'ovd', 2),                   # B5
            (6, '1e Ontvangst', '1eo', 2),          # B6
            (7, 'Muzikanten', 'muziek', 2),         # B7
            (8, 'Voorzangers', 'voorzangers', 2),   # B8
            (9, 'Beamer', 'beamer', 2),             # B9
            (10, 'Geluid', 'multimedia', 2),        # B10
            (11, 'KND', 'knd', 2),                  # B11
            (12, 'Tieners', 'tieners', 2),          # B12
        ]

        # Populate B4-B12
        for row, field_name, taken_key, col in field_mapping:
            new_val = str(entry.get(taken_key, '')).strip()
            if new_val:
                set_cell_value(row, col, new_val, field_name)

        # Get dankoffer verse from Dropbox
        dbx = _get_dbx_liturgie()
        dankoffer = _get_dankoffer_verse(dbx, service_date)

        # Row 21 (dankoffer) - B21-E21
        dankoffer_row = 21
        if dankoffer:
            # B21: Book name
            set_cell_value(dankoffer_row, 2, dankoffer['book'], f'Dankoffer boek (B21) (rij {dankoffer["row_index"]} uit Dankoffer.xlsx)')

            # C21: Chapter (H.S. / pasal)
            set_cell_value(dankoffer_row, 3, dankoffer['chapter'], 'Dankoffer hoofdstuk (C21)')

            # D21: Start verse (ayat)
            set_cell_value(dankoffer_row, 4, dankoffer['verse_start'], 'Dankoffer begin vers (D21)')

            # E21: End verse (ayat) - only if there's an end verse
            if dankoffer['verse_end']:
                set_cell_value(dankoffer_row, 5, dankoffer['verse_end'], 'Dankoffer eind vers (E21)')

        # Try to get Tikkie link from email if available
        try:
            reader = OutlookCollecteReader()
            if reader.is_authenticated():
                email_data = reader.fetch_collecte_data(target_date=service_date, since_days=60)
                tikkie_url = email_data.get('tikkie_url', '')
                if tikkie_url:
                    # Find Tikkie link row by looking for "Tikkie link" label in column A
                    tikkie_row = None
                    for row in range(1, ws.max_row + 1):
                        label = str(ws.cell(row=row, column=1).value).strip().lower() if ws.cell(row=row, column=1).value else ''
                        if 'tikkie' in label or 'qr_link' in label:
                            tikkie_row = row
                            break

                    if tikkie_row is not None:
                        set_cell_value(tikkie_row, 2, tikkie_url, 'Tikkie link')
        except Exception as e:
            print(f'[Liturgie Fill] Could not fetch Tikkie link: {e}')
            pass  # Non-critical, continue without Tikkie

        # Save the modified Excel preserving all formatting
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        # Encode for response
        excel_b64 = base64.b64encode(output.read()).decode('ascii')

        # Prepare dankoffer info for response
        dankoffer_info = None
        if dankoffer:
            dankoffer_info = {
                'verse': dankoffer['full_text'],
                'row_index': dankoffer['row_index'],
                'total_count': dankoffer['total_count'],
                'unused_count': dankoffer['unused_count'],
                'reset_needed': dankoffer.get('reset_needed', False),
                'marked_as_used': dankoffer.get('marked_as_used', False),
                'already_assigned': dankoffer.get('already_assigned', False),
                'date_assigned': dankoffer.get('date_assigned')
            }

        return jsonify({
            'excel_data': excel_b64,
            'filename': f'Main_Liturgy_file_{service_date.strftime("%Y%m%d")}_filled.xlsx',
            'alerts': alerts,
            'service_date': service_date.strftime('%d-%m-%Y'),
            'dankoffer_verse': dankoffer['full_text'] if dankoffer else None,
            'dankoffer_info': dankoffer_info
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


# ---------------------------------------------------------------------------
# Dankoffer Preview Endpoint
# ---------------------------------------------------------------------------

@app.route('/liturgie/preview-dankoffer', methods=['POST'])
def preview_dankoffer():
    """
    Preview which dankoffer verse will be used for a given date.
    Does NOT mark the verse as used - just shows what would be selected.

    Request body: { "date": "2026-05-11" } (ISO format)
    """
    data = request.get_json(force=True)
    date_str = data.get('date', '')

    if not date_str:
        return jsonify({'error': 'Geen datum opgegeven.'}), 400

    try:
        service_date = datetime.strptime(date_str, '%Y-%m-%d')
    except ValueError:
        return jsonify({'error': 'Ongeldige datum formaat. Gebruik YYYY-MM-DD.'}), 400

    try:
        dbx = _get_dbx_liturgie()
        dankoffer = _preview_dankoffer_verse(dbx, service_date)

        if not dankoffer:
            return jsonify({
                'error': 'Geen dankoffer verzen gevonden in Dankoffer.xlsx.'
            }), 404

        return jsonify({
            'verse': dankoffer['full_text'],
            'row_index': dankoffer['row_index'],
            'total_count': dankoffer['total_count'],
            'unused_count': dankoffer['unused_count'],
            'reset_needed': dankoffer.get('reset_needed', False),
            'service_date': service_date.strftime('%d-%m-%Y'),
            'message': f'Vers {dankoffer["row_index"]} van {dankoffer["total_count"]} zal gebruikt worden.' +
                       (' (ALLE verzen zijn al gebruikt - cyclus wordt gereset!)' if dankoffer.get('reset_needed') else '')
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True, port=5000)
