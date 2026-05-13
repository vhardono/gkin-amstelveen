"""
Brevo (Sendinblue) Campaign Automation for GKIN Amstelveen
Creates email campaigns based on mededelingen and liturgie data.
"""

import os
import base64
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Any
import requests


class BrevoCampaignGenerator:
    """Handles Brevo API interactions for campaign creation."""

    def __init__(self, api_key: Optional[str] = None, sender_email: Optional[str] = None, sender_name: Optional[str] = None):
        self.api_key = api_key or os.environ.get('BREVO_API_KEY')
        self.sender_email = sender_email or os.environ.get('BREVO_SENDER_EMAIL', 'kerkenraad@gkin.nl')
        self.sender_name = sender_name or os.environ.get('BREVO_SENDER_NAME', 'GKIN Amstelveen')
        self.base_url = "https://api.brevo.com/v3"
        self.headers = {
            "api-key": self.api_key,
            "Content-Type": "application/json",
            "Accept": "application/json"
        }

    def _make_request(self, method: str, endpoint: str, data: Optional[Dict] = None) -> Dict:
        """Make authenticated request to Brevo API."""
        url = f"{self.base_url}{endpoint}"
        try:
            if method == "GET":
                resp = requests.get(url, headers=self.headers, timeout=30)
            elif method == "POST":
                resp = requests.post(url, headers=self.headers, json=data, timeout=30)
            elif method == "PUT":
                resp = requests.put(url, headers=self.headers, json=data, timeout=30)
            else:
                raise ValueError(f"Unsupported HTTP method: {method}")

            resp.raise_for_status()
            return resp.json() if resp.text else {}
        except requests.exceptions.RequestException as e:
            error_response = {'error': str(e), 'status': getattr(e.response, 'status_code', None)}
            # Try to get detailed validation errors from response
            if e.response is not None:
                try:
                    error_data = e.response.json()
                    error_response['details'] = error_data
                except:
                    error_response['response_text'] = e.response.text[:500]
            return error_response

    def get_lists(self) -> List[Dict]:
        """Fetch available contact lists."""
        result = self._make_request("GET", "/contacts/lists?limit=50")
        return result.get('lists', [])

    def upload_file(self, file_path: str, name: Optional[str] = None) -> Dict:
        """Upload a file to Brevo for use in campaigns."""
        url = f"{self.base_url}/smtp/files"

        try:
            filename = name or os.path.basename(file_path)

            with open(file_path, 'rb') as f:
                files = {'file': (filename, f)}
                headers = {"api-key": self.api_key}

                resp = requests.post(
                    url,
                    headers=headers,
                    files=files,
                    timeout=60
                )
                resp.raise_for_status()
                result = resp.json()

                return {
                    'success': True,
                    'file_id': result.get('fileId'),
                    'name': result.get('name', filename)
                }
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def upload_image_base64(self, image_data: bytes, name: str) -> Dict:
        """Upload image as base64 to Brevo."""
        url = f"{self.base_url}/smtp/files"

        try:
            # Brevo accepts base64 encoded images
            base64_data = base64.b64encode(image_data).decode('utf-8')

            data = {
                "name": name,
                "base64Content": base64_data
            }

            resp = requests.post(
                url,
                headers=self.headers,
                json=data,
                timeout=60
            )
            resp.raise_for_status()
            result = resp.json()

            return {
                'success': True,
                'file_id': result.get('fileId'),
                'name': result.get('name', name)
            }
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def create_campaign(self,
                       name: str,
                       subject: str,
                       html_content: str,
                       list_ids: Optional[List[int]] = None,
                       scheduled_at: Optional[str] = None) -> Dict:
        """
        Create a new email campaign.

        Args:
            name: Internal campaign name
            subject: Email subject line
            html_content: HTML email content
            list_ids: List of contact list IDs (defaults to configured list)
            scheduled_at: ISO 8601 datetime for scheduled send (optional)

        Returns:
            API response with campaign details
        """
        # Get default list if none provided
        if not list_ids:
            lists = self.get_lists()
            if lists:
                list_ids = [lists[0]['id']]

        data = {
            "name": name,
            "subject": subject,
            "type": "classic",
            "htmlContent": html_content,
            "sender": {
                "email": self.sender_email,
                "name": self.sender_name
            },
            "recipients": {
                "listIds": list_ids
            }
        }

        if scheduled_at:
            data['scheduledAt'] = scheduled_at

        return self._make_request("POST", "/emailCampaigns", data)

    def generate_html_from_mededelingen(self,
                                        service_date: datetime,
                                        predikant: str,
                                        mededelingen_data: Dict[str, Any],
                                        liturgie_data: Optional[Dict] = None,
                                        takenrooster_entry: Optional[Dict] = None,
                                        is_ole: bool = True,
                                        theme: str = "",
                                        bible_verse: str = "",
                                        youtube_link: str = "",
                                        liturgie_url: str = "",
                                        collecte_url: str = "",
                                        qr_image_url: str = "",
                                        ole_location: str = "",
                                        ole_time: str = "10:00") -> str:
        """
        Generate HTML email content matching the MailerLite OLE template structure.

        Args:
            service_date: The service date
            predikant: Predikant name
            mededelingen_data: Data from DropboxExcelReader.get_mededelingen()
            liturgie_data: Optional liturgie details
            takenrooster_entry: Optional takenrooster entry with OvD, etc.
            is_ole: Whether this is an Online Landelijke Eredienst
            theme: Sermon theme
            bible_verse: Bible verse reference
            youtube_link: YouTube livestream link
            liturgie_url: Liturgy document URL
            collecte_url: Collection payment URL
            qr_image_url: QR code image URL for collection
            ole_location: OLE location code (AM, DH, TB, etc.)
            ole_time: Service time (10:00, 10:30, etc.)

        Returns:
            Complete HTML email content
        """
        months = ['januari', 'februari', 'maart', 'april', 'mei', 'juni',
                  'juli', 'augustus', 'september', 'oktober', 'november', 'december']
        date_str = f"{service_date.day} {months[service_date.month - 1]} {service_date.year}"
        short_date = service_date.strftime('%d-%m-%Y')

        # Determine service type with location
        service_type = "OLE" if is_ole else "GKIN Amstelveen"

        # Map location codes to full names
        LOCATION_MAP = {
            'AM': 'in Amstelveen',
            'DH': 'vanuit de Marcuskerk in Den Haag',
            'TB': 'vanuit de Pauluskerk te Tilburg'
        }
        location_display = f" ({ole_location})" if ole_location else ""
        location_body = LOCATION_MAP.get(ole_location, ole_location) if ole_location else ""

        # Fix time format - remove 'u' suffix if present, then add ' uur'
        time_clean = ole_time.replace('u', '').replace('U', '') if ole_time else "10:00"
        time_display = time_clean

        # Generate the HTML (same structure as MailerLite version)
        html = f"""<!doctype html>
<html lang="nl" dir="ltr">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1, user-scalable=yes">
    <meta name="format-detection" content="telephone=no, date=no, address=no, email=no, url=no">
    <meta name="x-apple-disable-message-reformatting">
    <title>GKIN ({service_type}){location_display}: Online Landelijke Eredienst Zondag {date_str}, {time_clean}u</title>
</head>
<body style="margin: 0; padding: 0; background-color: #ffffff;">
    <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#ffffff">
        <tr><td align="center" valign="top">
            <table width="640" border="0" cellspacing="0" cellpadding="0" class="mobile-shell" style="background-color: #ffffff;">
                <tr><td>
                    <!-- Header with GKIN logo -->
                    <table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color: #000000;">
                        <tr><td style="padding: 25px; text-align: center;">
                            <h1 style="color: #ffffff; font-family: Arial, sans-serif; margin: 0;">GKIN Amstelveen</h1>
                        </td></tr>
                    </table>

                    <!-- Main Content -->
                    <table width="100%" border="0" cellspacing="0" cellpadding="0" style="padding: 40px 50px;">
                        <tr><td>
                            <p style="font-family: Arial, sans-serif; color: #515856; font-size: 16px; line-height: 1.5; margin-bottom: 15px;">Beste broeders en zusters,<br></p>
                            <p style="font-family: Arial, sans-serif; color: #515856; font-size: 16px; line-height: 1.5; margin-bottom: 15px;">Op {date_str} zal {predikant} voorgaan in de Online Landelijke Eredienst (OLE) van GKIN {location_body}, aanvang {time_display} uur.</p>
"""

        if theme:
            html += f"""                            <p style="font-family: Arial, sans-serif; color: #515856; font-size: 16px; line-height: 1.5; margin-bottom: 15px;">Het thema van de dienst is: <strong>"{theme}"</strong>"""
            if bible_verse:
                html += f""" genomen uit {bible_verse}."""
            html += """</p>
"""

        html += f"""                            <p style="font-family: Arial, sans-serif; color: #515856; font-size: 16px; line-height: 1.5; margin-bottom: 25px;">De dienst wordt live uitgezonden. De liturgie van de dienst kunt u hieronder vinden. Door op de link te klikken kunt u het bestand bekijken en downloaden. Via de eveneens hieronder vermelde link kunt u de dienst online volgen.</p>

                            <!-- Buttons -->
                            <table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-bottom: 25px;">
                                <tr>
                                    <td style="padding: 10px;">
                                        <a href="{liturgie_url}" target="_blank" style="display: inline-block; background-color: #000000; color: #ffffff; padding: 15px 30px; text-decoration: none; border-radius: 6px; font-family: Arial, sans-serif; font-size: 14px;">Liturgie ({ole_location or 'OLE'})</a>
                                    </td>
                                    <td style="padding: 10px;">
                                        <a href="{youtube_link}" target="_blank" style="display: inline-block; background-color: #000000; color: #ffffff; padding: 15px 30px; text-decoration: none; border-radius: 6px; font-family: Arial, sans-serif; font-size: 14px;">Webvideo ({ole_location or 'OLE'})</a>
                                    </td>
                                </tr>
                            </table>

                            <!-- Collecte Section -->
                            <table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color: #f8f9fa; padding: 25px; border-radius: 8px; margin-bottom: 25px;">
                                <tr><td>
                                    <h3 style="font-family: Arial, sans-serif; color: #000000; margin-bottom: 15px;">Collecte</h3>
                                    <p style="font-family: Arial, sans-serif; color: #515856; font-size: 14px; line-height: 1.5; margin-bottom: 10px;">De collecte kan via de onderstaande QR-code of via deze link:</p>
                                    <p style="font-family: Arial, sans-serif; margin-bottom: 15px;"><a href="{collecte_url}" target="_blank" style="color: #000000; text-decoration: underline;">{collecte_url}</a></p>
"""

        if qr_image_url:
            # If it's a local path, convert to base64
            if qr_image_url.startswith('/uploads/'):
                try:
                    # Build full path
                    upload_dir = os.environ.get('UPLOAD_DIR', '/app/uploads')
                    filename = qr_image_url.replace('/uploads/', '')
                    full_path = os.path.join(upload_dir, filename)
                    
                    if os.path.exists(full_path):
                        with open(full_path, 'rb') as f:
                            image_data = f.read()
                            base64_data = base64.b64encode(image_data).decode('utf-8')
                            # Determine mime type from extension
                            ext = os.path.splitext(filename)[1].lower()
                            mime_type = 'image/png' if ext == '.png' else 'image/jpeg' if ext in ['.jpg', '.jpeg'] else 'image/png'
                            qr_image_url = f"data:{mime_type};base64,{base64_data}"
                except Exception as e:
                    # If conversion fails, keep original URL
                    pass
            
            html += f"""                                    <img src="{qr_image_url}" alt="QR Code voor collecte" style="max-width: 200px; height: auto; border-radius: 8px;">
"""

        html += """                                </td></tr>
                            </table>

                            <p style="font-family: Arial, sans-serif; color: #515856; font-size: 14px; line-height: 1.5; margin-bottom: 20px;">De mededelingenbladen voor deze dienst zijn in de bijlage te vinden.</p>

                            <p style="font-family: Arial, sans-serif; color: #515856; font-size: 16px; line-height: 1.5; margin-bottom: 10px;">Namens de kerkenraad,<br><br>GKIN Amstelveen</p>
                        </td></tr>
                    </table>

                    <!-- Footer -->
                    <table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color: #f8f9fa; padding: 25px; text-align: center;">
                        <tr><td>
                            <p style="font-family: Arial, sans-serif; color: #6c757d; font-size: 12px;">
                                GKIN Amstelveen | E-mail: kerkenraad@gkin.nl | Website: www.gkin.nl
                            </p>
                        </td></tr>
                    </table>
                </td></tr>
            </table>
        </td></tr>
    </table>
</body>
</html>"""

        return html
