"""
GKIN Amstelveen Mededelingen Generator – Web App
"""

import os
import uuid
from datetime import datetime
from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename

from data_sources.dropbox_reader import DropboxExcelReader
from data_sources.scipio_scraper import ScipioScraper
from data_sources.preekroster_scraper import PreekrosterScraper
from data_sources.email_reader import OutlookCollecteReader, get_token_cache_json
from bulletin_generator import BulletinGenerator

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10MB max upload

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), 'output', '_uploads')
os.makedirs(UPLOAD_DIR, exist_ok=True)

# Cache takenrooster on startup
_takenrooster_cache = None


def _get_takenrooster():
    global _takenrooster_cache
    if _takenrooster_cache is None:
        reader = DropboxExcelReader()
        _takenrooster_cache = reader.get_takenrooster()
    return _takenrooster_cache


@app.route('/')
def index():
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
        })

    return render_template('index.html', dates=dates)


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
        activities = _parse_activities_from_mededelingen(meded)
        return jsonify({
            'regionale_nl': meded.get('regionale_nl', ''),
            'landelijke_nl': meded.get('landelijke_nl', ''),
            'activities': activities,
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


def _parse_activities_from_mededelingen(meded: dict) -> list:
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
    """Export current MSAL token cache as JSON string.
    Copy the value into Railway's MSAL_TOKEN_CACHE environment variable
    so the token survives redeployments.
    """
    try:
        token_json = get_token_cache_json()
        if not token_json or token_json == '{}':
            return jsonify({'error': 'Geen token gevonden — log eerst in bij Outlook.'}), 404
        return jsonify({'token_cache': token_json,
                        'instructions': 'Kopieer token_cache en plak het in Railway > Variables > MSAL_TOKEN_CACHE'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/refresh-dates', methods=['POST'])
def refresh_dates():
    """Force refresh of takenrooster cache."""
    global _takenrooster_cache
    _takenrooster_cache = None
    return jsonify({'status': 'ok'})


if __name__ == '__main__':
    app.run(debug=True, port=5000)
