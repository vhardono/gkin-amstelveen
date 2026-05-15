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

    def create_campaign(self, name: str, subject: str, html_content: str,
                       list_ids: Optional[List[str]] = None,
                       scheduled_at: Optional[str] = None) -> Dict:
        """Create a new email campaign."""
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
        if scheduled_at:
            data['scheduled_at'] = scheduled_at

        return self._make_request("POST", "/campaigns", data)

    def generate_html(self, service_date: datetime, predikant: str,
                     theme: str = "", bible_verse: str = "",
                     youtube_link: str = "", liturgie_url: str = "",
                     collecte_url: str = "", qr_image_url: str = "",
                     ole_location: str = "", ole_time: str = "10:00") -> str:
        """Generate HTML email for OLE service."""
        months = ['januari', 'februari', 'maart', 'april', 'mei', 'juni',
                  'juli', 'augustus', 'september', 'oktober', 'november', 'december']
        date_str = f"{service_date.day} {months[service_date.month - 1]} {service_date.year}"

        LOCATION_MAP = {
            'AM': 'Kerkgebouw in Amstelveen',
            'DH': 'Kerkgebouw in Den Haag',
            'TB': 'Pauluskerk te Tilburg'
        }
        # Reverse map: full name -> code
        REVERSE_MAP = {
            'Kerkgebouw in Amstelveen': 'AM',
            'Kerkgebouw in Den Haag': 'DH',
            'Pauluskerk te Tilburg': 'TB'
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

        theme_html = f'<p>Het thema is: <strong>"{theme}"</strong>{f" uit {bible_verse}" if bible_verse else ""}.</p>' if theme else ""

        return f"""<!DOCTYPE html>
<html lang="nl">
<head><meta charset="UTF-8"><title>GKIN OLE{location_display}: {date_str}, {time_clean}u</title></head>
<body style="margin:0;padding:0;background:#ffffff;">
<table width="100%" cellpadding="0" cellspacing="0" bgcolor="#ffffff"><tr><td align="center">
<table width="640" cellpadding="0" cellspacing="0" style="background:#ffffff;">
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
<tr><td style="padding:0 50px;">
    <table width="100%" cellpadding="0" cellspacing="0" style="border-top:1px solid #EAECED;">
        <tr><td height="20" style="line-height:20px;"></td></tr>
    </table>
</td></tr>
<!-- Main Content -->
<tr><td style="padding:0 50px;font-family:'Inter',Arial,sans-serif;color:#515856;font-size:16px;line-height:137%;">
<p style="margin:0 0 10px 0;">Beste broeders en zusters,<br></p>
<p style="margin:0 0 10px 0;">Op {date_str} zal {predikant} voorgaan in de Online Landelijke Eredienst (OLE) van GKIN {location_body}, aanvang {time_clean} uur. </p>
{theme_html}
<p style="margin:0 0 25px 0;">De dienst wordt live uitgezonden. De liturgie van de dienst kunt u hieronder vinden. Door op de link te klikken kunt u het bestand bekijken en downloaden. Via de eveneens hieronder vermelde link kunt u de dienst online volgen.</p>
<!-- Buttons -->
<table width="100%" cellpadding="0" cellspacing="0" style="margin:0 0 25px 0;">
    <tr>
        <td width="250" valign="top">
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
        <td width="40" style="line-height:20px;"></td>
        <td width="250" valign="top">
            <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:separate;">
                <tr>
                    <th align="center" style="background-color:#000000;border-radius:6px;padding:10px 25px;">
                        <a href="{('https://www.' + re.sub(r'^https?://(?:www\.)?', '', youtube_link)) if youtube_link else '#'}" style="display:block;font-family:'Inter',Arial,sans-serif;font-size:14px;color:#ffffff;text-decoration:none;line-height:16px;font-weight:normal;">
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
<p style="margin:0 0 25px 0;">of door overmaking aan GEREJA KRISTEN INDONESIA NEDERLAND, IBAN: NL19 INGB 0002 6182 90 o.v.v. Collecte OLE {date_numeric}.</p>
<!-- QR Code centered -->
<table width="100%" cellpadding="0" cellspacing="0" style="margin:25px 0;">
    <tr>
        <td width="160"></td>
        <td width="320" align="center">
            {qr_img}
        </td>
        <td width="160"></td>
    </tr>
</table>
<!-- Signature -->
<p style="margin:0 0 10px 0;"><br></p>
<p style="margin:0 0 10px 0;">Wij wensen u allen een gezegende dienst toe.<br></p>
<p style="margin:0 0 10px 0;">Met broederlijke groet in Christus,</p>
<p style="margin:0 0 10px 0;"><br></p>
<p style="margin:0 0 10px 0;">Namens de landelijke kerkenraad GKIN,<br></p>
<p style="margin:0 0 20px 0;">Vega Hardono, Regiosecretaris (AM)<br></p>
</td></tr>
<!-- Footer separator -->
<tr><td style="padding:0 50px;">
    <table width="100%" cellpadding="0" cellspacing="0" style="border-top:1px solid #EAECED;">
        <tr><td height="20" style="line-height:20px;"></td></tr>
    </table>
</td></tr>
<!-- Footer -->
<tr><td style="padding:0 50px 20px 50px;font-family:'Inter',Arial,sans-serif;color:#515856;font-size:14px;line-height:150%;">
    <table width="100%" cellpadding="0" cellspacing="0">
        <tr>
            <td align="left" width="250" valign="top">
                <p style="margin:0 0 6px 0;"><strong>GKIN Amstelveen</strong></p>
                <p style="margin:0;">Bouwerij 52<br>1185XX Amstelveen</p>
            </td>
            <td width="40"></td>
            <td align="left" width="250" valign="top">
                <p style="margin:0 0 6px 0;">Wilt u deze e-mails niet meer ontvangen?</p>
                <a href="{{{{unsubscribe_link}}}}">{{{{unsubscribe_text}}}}</a>
            </td>
        </tr>
    </table>
</td></tr>
</table>
</td></tr></table>
</body>
</html>"""
