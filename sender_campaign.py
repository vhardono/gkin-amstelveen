"""
Sender.net Campaign Automation for GKIN Amstelveen
"""

import os
import base64
from datetime import datetime
from typing import Dict, List, Optional, Any
import requests


class SenderCampaignGenerator:
    """Handles Sender.net API interactions."""

    def __init__(self, api_token: Optional[str] = None, sender_email: Optional[str] = None, sender_name: Optional[str] = None):
        self.api_token = api_token or os.environ.get('SENDER_API_TOKEN')
        self.sender_email = sender_email or os.environ.get('SENDER_SENDER_EMAIL', 'noreply-am@gkin.org')
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
        """Fetch available subscriber lists."""
        result = self._make_request("GET", "/subscriber-lists")
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
            lists = self.get_lists()
            if lists:
                list_ids = [lists[0].get('id')]

        data = {
            "name": name,
            "subject": subject,
            "type": "regular",
            "content": html_content,
            "from": {"email": self.sender_email, "name": self.sender_name},
            "lists": list_ids
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
            'AM': 'in Amstelveen',
            'DH': 'vanuit de Marcuskerk in Den Haag',
            'TB': 'vanuit de Pauluskerk te Tilburg'
        }
        location_display = f" ({ole_location})" if ole_location else ""
        location_body = LOCATION_MAP.get(ole_location, ole_location) if ole_location else ""
        time_clean = ole_time.replace('u', '').replace('U', '') if ole_time else "10:00"

        # Handle QR image
        qr_img = ""
        if qr_image_url:
            if qr_image_url.startswith('/uploads/'):
                try:
                    upload_dir = os.environ.get('UPLOAD_DIR', '/app/uploads')
                    full_path = os.path.join(upload_dir, qr_image_url.replace('/uploads/', ''))
                    if os.path.exists(full_path):
                        with open(full_path, 'rb') as f:
                            b64 = base64.b64encode(f.read()).decode('utf-8')
                            ext = os.path.splitext(full_path)[1].lower()
                            mime = 'image/png' if ext == '.png' else 'image/jpeg'
                            qr_img = f'<img src="data:{mime};base64,{b64}" style="max-width:200px;border-radius:8px;">'
                except:
                    pass
            if not qr_img:
                qr_img = f'<img src="{qr_image_url}" style="max-width:200px;border-radius:8px;">'

        theme_html = f'<p>Het thema is: <strong>"{theme}"</strong>{f" uit {bible_verse}" if bible_verse else ""}.</p>' if theme else ""

        return f"""<!DOCTYPE html>
<html lang="nl">
<head><meta charset="UTF-8"><title>GKIN OLE{location_display}: {date_str}, {time_clean}u</title></head>
<body style="margin:0;padding:0;background:#ffffff;">
<table width="100%" cellpadding="0" cellspacing="0" bgcolor="#ffffff"><tr><td align="center">
<table width="640" cellpadding="0" cellspacing="0" style="background:#ffffff;">
<!-- Header with Logo -->
<tr>
    <td style="background:#000000;padding:30px;text-align:center;">
        <img src="https://gkin.org/main/images/banners/logo.png" alt="GKIN Logo" style="height:60px;width:auto;margin-bottom:10px;">
        <h1 style="color:#ffffff;font-family:Arial,sans-serif;margin:0;font-size:24px;font-weight:bold;">GKIN Amstelveen</h1>
        <p style="color:#cccccc;font-family:Arial,sans-serif;margin:5px 0 0 0;font-size:12px;">Online Landelijke Eredienst (OLE)</p>
    </td>
</tr>
<tr><td style="padding:40px 50px;font-family:Arial,sans-serif;color:#515856;font-size:16px;line-height:1.5;">
<p>Beste broeders en zusters,</p>
<p>Op {date_str} zal {predikant} voorgaan in de Online Landelijke Eredienst (OLE) van GKIN {location_body}, aanvang {time_clean} uur.</p>
{theme_html}
<p>De dienst wordt live uitgezonden. De liturgie vindt u hieronder:</p>
<table width="100%" style="margin:25px 0;"><tr>
<td style="padding:10px;"><a href="{liturgie_url}" style="display:inline-block;background:#000000;color:#ffffff;padding:15px 30px;text-decoration:none;border-radius:6px;">Liturgie ({ole_location or 'OLE'})</a></td>
<td style="padding:10px;"><a href="{youtube_link}" style="display:inline-block;background:#000000;color:#ffffff;padding:15px 30px;text-decoration:none;border-radius:6px;">Webvideo ({ole_location or 'OLE'})</a></td>
</tr></table>
<table width="100%" style="background:#f8f9fa;padding:25px;border-radius:8px;margin-bottom:25px;">
<tr><td>
<h3 style="color:#000000;margin-bottom:15px;">Collecte</h3>
<p style="font-size:14px;">De collecte kan via de QR-code of deze link:</p>
<p><a href="{collecte_url}" style="color:#000000;text-decoration:underline;">{collecte_url}</a></p>
{qr_img}
</td></tr>
</table>
<p style="font-size:14px;">De mededelingenbladen zijn in de bijlage te vinden.</p>
<p>Namens de kerkenraad,<br><br>GKIN Amstelveen</p>
</td></tr>
<tr><td style="background:#f8f9fa;padding:25px;text-align:center;font-size:12px;color:#6c757d;">
GKIN Amstelveen | kerkenraad@gkin.nl | www.gkin.nl
</td></tr>
</table>
</td></tr></table>
</body>
</html>"""
