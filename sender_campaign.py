"""
Sender.net Campaign Automation for GKIN Amstelveen
"""

import os
import re
import base64
from datetime import datetime
from typing import Dict, List, Optional, Any
import requests


class SenderCampaignGenerator:
    """Handles Sender.net API interactions."""

    def __init__(self, api_token: Optional[str] = None, sender_email: Optional[str] = None, sender_name: Optional[str] = None):
        self.api_token = api_token or os.environ.get('SENDER_API_KEY') or os.environ.get('SENDER_API_TOKEN')
        self.sender_email = sender_email or os.environ.get('SENDER_SENDER_EMAIL', 'newsletter@gkin-amstelveen.top')
        self.sender_name = sender_name or os.environ.get('SENDER_SENDER_NAME', 'GKIN Amstelveen')
        self.base_url = "https://api.sender.net/v2"
        self.headers = {
            "Authorization": f"Bearer {self.api_token}",
            "Content-Type": "application/json",
            "Accept": "application/json"
        }

    def _make_request(self, method: str, endpoint: str, data: Optional[Dict] = None) -> Dict:
        """Make authenticated request to Sender API."""
        url = f"{self.base_url}{endpoint}"
        try:
            if method == "GET":
                resp = requests.get(url, headers=self.headers, timeout=30)
            elif method == "POST":
                resp = requests.post(url, headers=self.headers, json=data, timeout=30)
            else:
                raise ValueError(f"Unsupported HTTP method: {method}")

            resp.raise_for_status()
            return resp.json() if resp.text else {}
        except requests.exceptions.RequestException as e:
            error_response = {'error': str(e), 'status': getattr(e.response, 'status_code', None)}
            if e.response is not None:
                try:
                    error_response['details'] = e.response.json()
                except:
                    error_response['response_text'] = e.response.text[:500]
            return error_response

    def get_lists(self) -> List[Dict]:
        """Fetch available subscriber groups."""
        result = self._make_request("GET", "/groups")
        return result.get('data', [])

    def upload_file(self, file_path: str) -> Dict:
        """Upload a file to Sender."""
        try:
            with open(file_path, 'rb') as f:
                files = {'file': (os.path.basename(file_path), f)}
                resp = requests.post(
                    f"{self.base_url}/files",
                    headers={"Authorization": f"Bearer {self.api_token}"},
                    files=files,
                    timeout=60
                )
                resp.raise_for_status()
                result = resp.json()
                return {
                    'success': True,
                    'url': result.get('data', {}).get('url'),
                    'file_id': result.get('data', {}).get('id')
                }
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def schedule_campaign(self, campaign_id: str, schedule_time: str) -> Dict:
        """Schedule a campaign. schedule_time must be 'Y-m-d H:i:s' UTC."""
        return self._make_request("POST", f"/campaigns/{campaign_id}/schedule",
                                  {"schedule_time": schedule_time})

    def create_campaign(self, name: str, subject: str, html_content: str,
                       list_ids: Optional[List[str]] = None,
                       scheduled_at: Optional[str] = None,
                       preheader: Optional[str] = None) -> Dict:
        """Create campaign, then schedule it if scheduled_at provided.
        scheduled_at should be ISO 8601 with timezone e.g. '2026-05-16T09:00:00+02:00'.
        """
        if not list_ids:
            groups = self.get_lists()
            if groups:
                list_ids = [groups[0].get('id')]

        data = {
            "title": name,
            "subject": subject,
            "content_type": "html",
            "content": html_content,
            "from": self.sender_name,
            "reply_to": self.sender_email,
            "groups": list_ids
        }

        if preheader:
            data["preheader"] = preheader

        result = self._make_request("POST", "/campaigns", data)

        # If created successfully and schedule requested, call schedule endpoint
        if scheduled_at and result.get('data', {}).get('id'):
            campaign_id = result['data']['id']
            try:
                from datetime import timezone
                # Handle JS toISOString() format '2026-05-18T07:00:00.000Z' and '+02:00' variants
                sa = scheduled_at.replace('Z', '+00:00')
                dt = datetime.fromisoformat(sa)
                dt_utc = dt.astimezone(timezone.utc)
                schedule_time = dt_utc.strftime('%Y-%m-%d %H:%M:%S')
                sched_result = self.schedule_campaign(campaign_id, schedule_time)
                print(f"[Sender] Schedule result for {campaign_id}: {sched_result}")
                result['scheduled'] = sched_result.get('success', False)
                result['schedule_time'] = schedule_time
            except Exception as e:
                print(f"[Sender] Schedule error: {e}")
                result['schedule_error'] = str(e)

        return result

    def generate_html(self, service_date: datetime, predikant: str,
                     theme: str = "", bible_verse: str = "",
                     youtube_link: str = "", liturgie_url: str = "",
                     collecte_url: str = "", qr_image_url: str = "",
                     ole_location: str = "", ole_time: str = "10:00",
                     collecte_ovv: str = "") -> str:
        """Generate HTML email for OLE service."""
        months = ['januari', 'februari', 'maart', 'april', 'mei', 'juni',
                  'juli', 'augustus', 'september', 'oktober', 'november', 'december']
        date_str = f"{service_date.day} {months[service_date.month - 1]} {service_date.year}"

        LOCATION_MAP = {
            'AM': 'vanuit de Marcuskerk in Amstelveen',
            'DH': 'vanuit de Marcuskerk in Den Haag',
            'TB': 'vanuit de Pauluskerk te Tilburg'
        }
        # Reverse map: full name -> code (handle both old and new formats)
        REVERSE_MAP = {
            'Kerkgebouw in Amstelveen': 'AM',
            'vanuit de Marcuskerk in Amstelveen': 'AM',
            'Amstelveen': 'AM',
            'Kerkgebouw in Den Haag': 'DH',
            'vanuit de Marcuskerk in Den Haag': 'DH',
            'Den Haag': 'DH',
            'Pauluskerk te Tilburg': 'TB',
            'vanuit de Pauluskerk te Tilburg': 'TB',
            'Tilburg': 'TB'
        }
        # Resolve location code (handle both code 'DH' and full name)
        location_code = REVERSE_MAP.get(ole_location, ole_location).upper() if ole_location else ''
        if location_code not in LOCATION_MAP:
            location_code = ole_location  # fallback
        location_display = f" ({location_code}-OLE)" if location_code else ""
        location_body = LOCATION_MAP.get(location_code, location_code) if location_code else ""
        time_clean = ole_time.replace('u', '').replace('U', '') if ole_time else "10:00"
        location_ole_tag = f"{location_code}-OLE" if location_code else "OLE"
        date_numeric = f"{service_date.day:02d}-{service_date.month:02d}-{service_date.year}"

        # Handle QR image — embed as base64 so Sender.net can render it
        qr_img = ""
        if qr_image_url:
            if qr_image_url.startswith('data:'):
                # Already a data URI (from website scraper)
                qr_img = f'<img src="{qr_image_url}" style="max-width:200px;border-radius:8px;">'
            elif qr_image_url.startswith('http'):
                # Remote URL — fetch and embed as base64
                try:
                    import requests as _req
                    r = _req.get(qr_image_url, timeout=10)
                    r.raise_for_status()
                    ext = qr_image_url.rsplit('.', 1)[-1].lower().split('?')[0]
                    mime = {'jpg': 'image/jpeg', 'jpeg': 'image/jpeg', 'gif': 'image/gif'}.get(ext, 'image/png')
                    b64 = base64.b64encode(r.content).decode('utf-8')
                    qr_img = f'<img src="data:{mime};base64,{b64}" style="max-width:200px;border-radius:8px;">'
                except Exception:
                    qr_img = f'<img src="{qr_image_url}" style="max-width:200px;border-radius:8px;">'
            elif qr_image_url.startswith('/uploads/'):
                try:
                    upload_dir = os.environ.get('UPLOAD_DIR', '/app/uploads')
                    full_path = os.path.join(upload_dir, qr_image_url.replace('/uploads/', ''))
                    if os.path.exists(full_path):
                        with open(full_path, 'rb') as f:
                            b64 = base64.b64encode(f.read()).decode('utf-8')
                            ext = os.path.splitext(full_path)[1].lower()
                            mime = 'image/png' if ext == '.png' else 'image/jpeg'
                            qr_img = f'<img src="data:{mime};base64,{b64}" style="max-width:200px;border-radius:8px;">'
                except Exception:
                    pass

        bible_suffix = f" genomen uit {bible_verse}" if bible_verse else ""
        theme_html = f'<p style="margin:0 0 10px 0;">Het thema is: <strong>"{theme}"</strong>{bible_suffix}.</p>' if theme else ""
        collecte_ovv_text = collecte_ovv if collecte_ovv else f"Collecte OLE {date_numeric}"
        youtube_href = ('https://www.' + re.sub(r'^https?://(?:www\.)?', '', youtube_link)) if youtube_link else '#'

        return f"""<!DOCTYPE html>
<html lang="nl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>GKIN Amstelveen {date_str} (OLE): Online Landelijke Eredienst</title>
<style>
@media only screen and (max-width:640px) {{
  .outer-table {{ width:100% !important; }}
  .content-cell {{ padding:0 20px !important; }}
  .footer-cell {{ padding:0 20px 20px 20px !important; }}
  .btn-td {{ display:block !important; width:100% !important; padding-bottom:10px; }}
  .btn-spacer {{ display:none !important; }}
  .footer-col {{ display:block !important; width:100% !important; padding-bottom:16px; }}
  .footer-spacer {{ display:none !important; }}
}}
</style>
</head>
<body style="margin:0;padding:0;background:#ffffff;">
<table width="100%" cellpadding="0" cellspacing="0" bgcolor="#ffffff"><tr><td align="center">
<table class="outer-table" width="640" cellpadding="0" cellspacing="0" style="background:#ffffff;max-width:640px;">
<!-- Header: Logo left, Title right -->
<tr>
    <td style="background:#ffffff;padding:20px 50px;">
        <table width="100%" cellpadding="0" cellspacing="0">
            <tr>
                <td align="left" width="60">
                    <img src="https://gkin.org/main/images/banners/logo.png" alt="GKIN Logo" style="height:60px;width:auto;">
                </td>
                <td align="right" style="font-family:'Inter',Arial,sans-serif;">
                    <h1 style="color:#000000;margin:0;font-size:18px;font-weight:bold;">GKIN Amstelveen</h1>
                </td>
            </tr>
        </table>
    </td>
</tr>
<!-- Separator line -->
<tr><td class="content-cell" style="padding:0 50px;">
    <table width="100%" cellpadding="0" cellspacing="0" style="border-top:1px solid #EAECED;">
        <tr><td height="20" style="line-height:20px;"></td></tr>
    </table>
</td></tr>
<!-- Main Content -->
<tr><td class="content-cell" style="padding:0 50px;font-family:'Inter',Arial,sans-serif;color:#515856;font-size:16px;line-height:137%;">
<p style="margin:0 0 10px 0;">Beste broeders en zusters,<br></p>
<p style="margin:0 0 10px 0;">Op {date_str} zal {predikant} voorgaan in de Online Landelijke Eredienst (OLE) van GKIN {location_body}, aanvang {time_clean} uur. </p>
{theme_html}
<p style="margin:0 0 25px 0;">De dienst wordt live uitgezonden. De liturgie van de dienst kunt u hieronder vinden. Door op de link te klikken kunt u het bestand bekijken en downloaden. Via de eveneens hieronder vermelde link kunt u de dienst online volgen.</p>
<!-- Buttons -->
<table width="100%" cellpadding="0" cellspacing="0" style="margin:0 0 25px 0;">
    <tr>
        <td class="btn-td" width="250" valign="top">
            <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:separate;">
                <tr>
                    <th align="center" style="background-color:#000000;border-radius:6px;padding:10px 25px;">
                        <a href="{liturgie_url}" style="display:block;font-family:'Inter',Arial,sans-serif;font-size:14px;color:#ffffff;text-decoration:none;line-height:16px;font-weight:normal;">
                            Liturgie<br>({location_ole_tag})
                        </a>
                    </th>
                </tr>
            </table>
        </td>
        <td class="btn-spacer" width="40" style="line-height:20px;"></td>
        <td class="btn-td" width="250" valign="top">
            <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:separate;">
                <tr>
                    <th align="center" style="background-color:#000000;border-radius:6px;padding:10px 25px;">
                        <a href="{youtube_href}" style="display:block;font-family:'Inter',Arial,sans-serif;font-size:14px;color:#ffffff;text-decoration:none;line-height:16px;font-weight:normal;">
                            Webvideo<br>({location_ole_tag})
                        </a>
                    </th>
                </tr>
            </table>
        </td>
    </tr>
</table>
<!-- Collecte -->
<p style="margin:0 0 10px 0;">In deze dienst wordt er 1 keer gecollecteerd.<br><br>De collecte is bestemd voor Landelijke kas (OLE). U kunt dit overmaken via: <a href="{collecte_url}" style="color:#2CB191;text-decoration:underline;">{collecte_url}</a></p>
<p style="margin:0 0 25px 0;">of door overmaking aan GEREJA KRISTEN INDONESIA NEDERLAND, IBAN: NL19 INGB 0002 6182 90 o.v.v. {collecte_ovv_text}.</p>
<!-- QR Code centered -->
<table width="100%" cellpadding="0" cellspacing="0" style="margin:10px 0;">
    <tr>
        <td width="160"></td>
        <td width="320" align="center">
            {qr_img}
        </td>
        <td width="160"></td>
    </tr>
</table>
<!-- Signature -->
<p style="margin:0 0 10px 0;">Wij wensen u allen een gezegende dienst toe.<br></p>
<p style="margin:0 0 10px 0;">Met broederlijke groet in Christus,</p>
<p style="margin:0 0 10px 0;"><br></p>
<p style="margin:0 0 10px 0;">Namens de landelijke kerkenraad GKIN,<br></p>
<p style="margin:0 0 20px 0;">Vega Hardono, Regiosecretaris (AM)<br></p>
</td></tr>
<!-- Footer separator -->
<tr><td class="content-cell" style="padding:0 50px;">
    <table width="100%" cellpadding="0" cellspacing="0" style="border-top:1px solid #EAECED;">
        <tr><td height="20" style="line-height:20px;"></td></tr>
    </table>
</td></tr>
<!-- Footer -->
<tr><td class="footer-cell" style="padding:0 50px 20px 50px;font-family:'Inter',Arial,sans-serif;color:#515856;font-size:14px;line-height:150%;">
    <table width="100%" cellpadding="0" cellspacing="0">
        <tr>
            <td class="footer-col" align="left" width="250" valign="top">
                <p style="margin:0 0 6px 0;"><strong>GKIN Amstelveen</strong></p>
                <p style="margin:0;">Bouwerij 52<br>1185XX Amstelveen</p>
            </td>
            <td class="footer-spacer" width="40"></td>
            <td class="footer-col" align="left" width="250" valign="top">
                <p style="margin:0 0 6px 0;">Wilt u deze e-mails niet meer ontvangen?</p>
                <a href="{{{{unsubscribe_link}}}}">uitschrijven</a>
            </td>
        </tr>
    </table>
</td></tr>
</table>
</td></tr></table>
</body>
</html>"""

    def generate_pm_html(self, service_date: datetime, am_predikant: str,
                         mededelingen_url: str = "", preek_am_url: str = "",
                         ole_location: str = "", ole_predikant: str = "",
                         youtube_link: str = "", preek_ole_url: str = "") -> str:
        """Generate HTML email for post-service Preek & Mededelingen mailing.

        Layout rules:
        - If ole_location == 'AM' (or empty): 3-button row (160px each):
            Mededelingen (AM) | Preek (AM) + am_predikant | Webvideo (AM-OLE)
        - If ole_location is DH/TB (non-AM OLE): 2×2 grid (250px each):
            Row 1: Mededelingen (AM) | Preek (AM) + am_predikant
            Row 2: Preek ({loc}-OLE) "kunt u hier later terugvinden" | Webvideo ({loc}-OLE)
        """
        months = ['januari', 'februari', 'maart', 'april', 'mei', 'juni',
                  'juli', 'augustus', 'september', 'oktober', 'november', 'december']
        date_str = f"{service_date.day} {months[service_date.month - 1]} {service_date.year}"
        nl_days = ['maandag', 'dinsdag', 'woensdag', 'donderdag', 'vrijdag', 'zaterdag', 'zondag']
        day_name = nl_days[service_date.weekday()]

        LOCATION_MAP = {
            'AM': 'vanuit de Marcuskerk in Amstelveen',
            'DH': 'vanuit de Marcuskerk in Den Haag',
            'TB': 'vanuit de Pauluskerk te Tilburg',
        }
        REVERSE_MAP = {
            'Kerkgebouw in Amstelveen': 'AM',
            'vanuit de Marcuskerk in Amstelveen': 'AM',
            'Amstelveen': 'AM',
            'Kerkgebouw in Den Haag': 'DH',
            'vanuit de Marcuskerk in Den Haag': 'DH',
            'Den Haag': 'DH',
            'Pauluskerk te Tilburg': 'TB',
            'vanuit de Pauluskerk te Tilburg': 'TB',
            'Tilburg': 'TB'
        }
        loc_code = REVERSE_MAP.get(ole_location, ole_location).upper() if ole_location else 'AM'
        if loc_code not in LOCATION_MAP:
            loc_code = 'AM'

        youtube_href = ('https://www.' + re.sub(r'^https?://(?:www\.)?', '', youtube_link)) if youtube_link else '#'

        # ------------------------------------------------------------------ #
        # Button helpers
        # ------------------------------------------------------------------ #
        def btn(width: int, line1: str, line2: str, href: str) -> str:
            """Two-line button. Pass line2='' for a blank second line that still occupies height."""
            line2_html = (f'<span style="display:block;line-height:16px;">{line2}</span>'
                          if line2
                          else '<span style="display:block;line-height:16px;visibility:hidden;">&#160;</span>')
            return f"""<td class="btn-td" width="{width}" valign="top">
            <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:separate;">
                <tr>
                    <th align="center" style="background-color:#000000;border-radius:6px;padding:10px 25px;">
                        <a href="{href}" target="_blank" style="display:block;font-family:'Inter',Arial,sans-serif;font-size:14px;color:#ffffff;text-decoration:none;font-weight:normal;">
                            <span style="display:block;line-height:16px;">{line1}</span>
                            {line2_html}
                        </a>
                    </th>
                </tr>
            </table>
        </td>"""

        spacer = '<td class="btn-spacer" width="30" style="line-height:20px;"></td>'
        row_gap = '<tr><td colspan="3" height="10" style="line-height:10px;font-size:10px;">&nbsp;</td></tr>'

        # ------------------------------------------------------------------ #
        # Decide layout
        # ------------------------------------------------------------------ #
        am_only = (loc_code == 'AM')

        if am_only:
            # 3 buttons, 160px each
            btn_width = 160
            buttons_html = f"""<table width="100%" cellpadding="0" cellspacing="0" style="margin:0 0 25px 0;">
    <tr>
        {btn(btn_width, 'Mededelingen', '(AM)', mededelingen_url or '#')}
        {spacer}
        {btn(btn_width, 'Preek (AM)', am_predikant, preek_am_url or '#')}
        {spacer}
        {btn(btn_width, 'Webvideo', '(AM-OLE)', youtube_href)}
    </tr>
</table>"""
        else:
            # 2×2 grid
            btn_width = 250
            loc_tag = f"{loc_code}-OLE"
            preek_ole_href = preek_ole_url or '#'
            buttons_html = f"""<table width="100%" cellpadding="0" cellspacing="0" style="margin:0 0 25px 0;">
    <tr>
        {btn(btn_width, 'Mededelingen', '(AM)', mededelingen_url or '#')}
        {spacer}
        {btn(btn_width, 'Preek (AM)', am_predikant, preek_am_url or '#')}
    </tr>
    {row_gap}
    <tr>
        {btn(btn_width, f'Preek ({loc_tag})', ole_predikant or 'kunt u hier later terugvinden', preek_ole_href)}
        {spacer}
        {btn(btn_width, 'Webvideo', f'({loc_tag})', youtube_href)}
    </tr>
</table>"""

        return f"""<!DOCTYPE html>
<html lang="nl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>GKIN Amstelveen {date_str}: Preek &amp; mededelingen</title>
<style>
@media only screen and (max-width:640px) {{
  .outer-table {{ width:100% !important; }}
  .content-cell {{ padding:0 20px !important; }}
  .footer-cell {{ padding:0 20px 20px 20px !important; }}
  .btn-td {{ display:block !important; width:100% !important; padding-bottom:10px; }}
  .btn-spacer {{ display:none !important; }}
  .footer-col {{ display:block !important; width:100% !important; padding-bottom:16px; }}
  .footer-spacer {{ display:none !important; }}
}}
</style>
</head>
<body style="margin:0;padding:0;background:#ffffff;">
<table width="100%" cellpadding="0" cellspacing="0" bgcolor="#ffffff"><tr><td align="center">
<table class="outer-table" width="640" cellpadding="0" cellspacing="0" style="background:#ffffff;max-width:640px;">
<!-- Header -->
<tr>
    <td style="background:#ffffff;padding:20px 50px;">
        <table width="100%" cellpadding="0" cellspacing="0">
            <tr>
                <td align="left" width="60">
                    <img src="https://gkin.org/main/images/banners/logo.png" alt="GKIN Logo" style="height:60px;width:auto;">
                </td>
                <td align="right" style="font-family:'Inter',Arial,sans-serif;">
                    <h1 style="color:#000000;margin:0;font-size:18px;font-weight:bold;">GKIN Amstelveen</h1>
                </td>
            </tr>
        </table>
    </td>
</tr>
<!-- Separator -->
<tr><td class="content-cell" style="padding:0 50px;">
    <table width="100%" cellpadding="0" cellspacing="0" style="border-top:1px solid #EAECED;">
        <tr><td height="20" style="line-height:20px;"></td></tr>
    </table>
</td></tr>
<!-- Main Content -->
<tr><td class="content-cell" style="padding:0 50px;font-family:'Inter',Arial,sans-serif;color:#515856;font-size:16px;line-height:137%;">
<p style="margin:0 0 10px 0;">Beste gemeenteleden en belangstellenden,</p>
<p style="margin:0 0 10px 0;"><br></p>
<h3 style="font-family:'Inter',Arial,sans-serif;color:#000000;font-size:18px;line-height:125%;font-weight:bold;margin-bottom:8px;">Mededelingen en preek van deze week</h3>
<p style="margin:0 0 25px 0;">Bijgaand de volgende stukken van {day_name} {date_str}<br>(door op de link te klikken kunt u het bestand bekijken en&nbsp;downloaden):</p>
<!-- Buttons -->
{buttons_html}
</td></tr>
<!-- Separator -->
<tr><td class="content-cell" style="padding:0 50px;">
    <table width="100%" cellpadding="0" cellspacing="0" style="border-top:1px solid #EAECED;">
        <tr><td height="20" style="line-height:20px;"></td></tr>
    </table>
</td></tr>
<!-- Website GKIN section -->
<tr><td class="content-cell" style="padding:0 50px;font-family:'Inter',Arial,sans-serif;color:#515856;font-size:16px;line-height:137%;">
<h3 style="font-family:'Inter',Arial,sans-serif;color:#000000;font-size:18px;line-height:125%;font-weight:bold;margin-bottom:8px;">Website GKIN</h3>
<p style="margin:0 0 10px 0;">Voor actuele informatie over de bijeenkomsten en online diensten verwijzen wij u naar de website van GKIN:&nbsp;<a href="https://www.gkin.org" style="color:#2CB191;text-decoration:underline;">www.gkin.org</a></p>
<p style="margin:0 0 10px 0;"><br></p>
<p style="margin:0 0 10px 0;">Wij wensen u allen een fijne week.<br><br></p>
<p style="margin:0 0 10px 0;">Met broederlijke groet in Christus,<br>Namens de landelijke kerkenraad GKIN,</p>
<p style="margin:0 0 10px 0;"><br></p>
<p style="margin:0 0 20px 0;">Vega Hardono, Regiosecretaris (AM)<br></p>
</td></tr>
<!-- Footer separator -->
<tr><td class="content-cell" style="padding:0 50px;">
    <table width="100%" cellpadding="0" cellspacing="0" style="border-top:1px solid #EAECED;">
        <tr><td height="20" style="line-height:20px;"></td></tr>
    </table>
</td></tr>
<!-- Footer -->
<tr><td class="footer-cell" style="padding:0 50px 20px 50px;font-family:'Inter',Arial,sans-serif;color:#515856;font-size:14px;line-height:150%;">
    <table width="100%" cellpadding="0" cellspacing="0">
        <tr>
            <td class="footer-col" align="left" width="250" valign="top">
                <p style="margin:0 0 6px 0;"><strong>GKIN Amstelveen</strong></p>
                <p style="margin:0;">Bouwerij 52<br>1185XX Amstelveen</p>
            </td>
            <td class="footer-spacer" width="40"></td>
            <td class="footer-col" align="left" width="250" valign="top">
                <p style="margin:0 0 6px 0;">Wilt u deze e-mails niet meer ontvangen? U kunt zich hier:</p>
                <a href="{{{{unsubscribe_link}}}}" style="color:#515856;text-decoration:underline;">uitschrijven</a>
            </td>
        </tr>
    </table>
</td></tr>
</table>
</td></tr></table>
</body>
</html>"""
