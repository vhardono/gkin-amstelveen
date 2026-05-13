"""
MailerLite Campaign Automation for GKIN Amstelveen
Creates email campaigns based on mededelingen and liturgie data.
"""

import os
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Any
import requests


class MailerLiteFileManager:
    """Handles file uploads to MailerLite for use in campaigns."""
    
    def __init__(self, api_key: Optional[str] = None):
        self.api_key = api_key or os.environ.get('MAILERLITE_API_KEY')
        self.base_url = "https://connect.mailerlite.com/api"
        self.headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Accept": "application/json"
        }
    
    def upload_file(self, file_path: str, name: Optional[str] = None) -> Dict:
        """
        Upload a file to MailerLite file manager.
        
        Args:
            file_path: Local path to file
            name: Optional custom name for the file
        
        Returns:
            Dict with 'id', 'url', 'name' of uploaded file
        """
        url = f"{self.base_url}/files"
        
        try:
            filename = name or os.path.basename(file_path)
            
            with open(file_path, 'rb') as f:
                files = {'file': (filename, f)}
                data = {'name': filename}
                
                resp = requests.post(
                    url,
                    headers={"Authorization": f"Bearer {self.api_key}"},
                    files=files,
                    data=data,
                    timeout=60
                )
                resp.raise_for_status()
                result = resp.json()
                
                return {
                    'success': True,
                    'id': result.get('data', {}).get('id'),
                    'url': result.get('data', {}).get('url'),
                    'name': result.get('data', {}).get('name', filename)
                }
        except Exception as e:
            return {'success': False, 'error': str(e)}
    
    def upload_bytes(self, content: bytes, name: str) -> Dict:
        """Upload file content directly from bytes."""
        url = f"{self.base_url}/files"
        
        try:
            files = {'file': (name, content)}
            data = {'name': name}
            
            resp = requests.post(
                url,
                headers={"Authorization": f"Bearer {self.api_key}"},
                files=files,
                data=data,
                timeout=60
            )
            resp.raise_for_status()
            result = resp.json()
            
            return {
                'success': True,
                'id': result.get('data', {}).get('id'),
                'url': result.get('data', {}).get('url'),
                'name': result.get('data', {}).get('name', name)
            }
        except Exception as e:
            return {'success': False, 'error': str(e)}
    
    def list_files(self) -> List[Dict]:
        """List all uploaded files in MailerLite."""
        url = f"{self.base_url}/files"
        try:
            resp = requests.get(
                url,
                headers={"Authorization": f"Bearer {self.api_key}", "Accept": "application/json"},
                timeout=30
            )
            resp.raise_for_status()
            result = resp.json()
            return result.get('data', [])
        except Exception as e:
            return []


class MailerLiteCampaignGenerator:
    """Handles MailerLite API interactions for campaign creation."""

    def __init__(self, api_key: Optional[str] = None, group_id: Optional[str] = None):
        self.api_key = api_key or os.environ.get('MAILERLITE_API_KEY')
        self.group_id = group_id or os.environ.get('MAILERLITE_GROUP_ID')
        self.base_url = "https://connect.mailerlite.com/api"
        self.headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
            "Accept": "application/json"
        }

    def _make_request(self, method: str, endpoint: str, data: Optional[Dict] = None) -> Dict:
        """Make authenticated request to MailerLite API."""
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
            return {'error': str(e), 'status': getattr(e.response, 'status_code', None)}

    def get_groups(self) -> List[Dict]:
        """Fetch available subscriber groups."""
        result = self._make_request("GET", "/groups")
        return result.get('data', [])

    def create_campaign(self, 
                       name: str,
                       subject: str,
                       html_content: str,
                       group_ids: Optional[List[str]] = None,
                       send_time: Optional[str] = None) -> Dict:
        """
        Create a new regular campaign.
        
        Args:
            name: Internal campaign name
            subject: Email subject line
            html_content: HTML email content
            group_ids: List of subscriber group IDs (defaults to configured group)
            send_time: ISO 8601 datetime for scheduled send (optional)
        
        Returns:
            API response with campaign details
        """
        groups = group_ids or ([self.group_id] if self.group_id else [])
        
        data = {
            "name": name,
            "subject": subject,
            "type": "regular",
            "emails": [{
                "subject": subject,
                "content": html_content,
                "from": {
                    "email": "kerkenraad@gkin.nl",
                    "name": "GKIN Amstelveen"
                }
            }],
            "groups": groups
        }
        
        if send_time:
            data['send_time'] = send_time
            
        return self._make_request("POST", "/campaigns", data)

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
        Generate HTML email content matching MailerLite OLE template structure.
        
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
            Complete HTML email content matching MailerLite format
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
        
        html = f"""<!doctype html>
<html lang="nl" dir="ltr">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1, user-scalable=yes">
    <meta name="format-detection" content="telephone=no, date=no, address=no, email=no, url=no">
    <meta name="x-apple-disable-message-reformatting">
    <title>GKIN ({service_type}){location_display}: Online Landelijke Eredienst Zondag {date_str}, {time_clean}u</title>
    <style type="text/css">
        html, body {{ margin: 0 !important; padding: 0 !important; width: 100% !important; height: 100% !important; }}
        body {{ -webkit-font-smoothing: antialiased; -moz-osx-font-smoothing: grayscale; text-rendering: optimizeLegibility; font-family: 'Inter', sans-serif; background-color: #F4F7FA; }}
        .document {{ margin: 0 !important; padding: 0 !important; width: 100% !important; }}
        img {{ border: 0; outline: none; text-decoration: none; -ms-interpolation-mode: bicubic; max-width: 100%; }}
        table {{ border-collapse: collapse; }}
        table, td {{ mso-table-lspace: 0pt; mso-table-rspace: 0pt; }}
        body, table, td, a {{ -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%; }}
        h1, h2, h3, h4, h5, p {{ margin: 0; word-break: break-word; }}
        a[x-apple-data-detectors] {{ color: inherit !important; text-decoration: none !important; font-size: inherit !important; font-family: inherit !important; font-weight: inherit !important; line-height: inherit !important; }}
        div[style*="margin: 16px 0;"] {{ margin: 0 !important; }}
        @media all and (max-width: 639px) {{
            .wrapper {{ width: 100% !important; }}
            .container {{ width: 100% !important; min-width: 100% !important; padding: 0 !important; }}
            .row {{ padding-left: 20px !important; padding-right: 20px !important; }}
            .col {{ display: block !important; width: 100% !important; }}
            .mobile-center {{ text-align: center !important; float: none !important; }}
            .mobile-hide {{ display: none !important; }}
            .ml-btn {{ width: 100% !important; max-width: 100% !important; }}
            .ml-btn-container {{ width: 100% !important; max-width: 100% !important; }}
            .mlContentTable {{ width: 100% !important; min-width: 10% !important; margin: 0 !important; float: none !important; }}
        }}
    </style>
</head>
<body style="margin: 0 !important; padding: 0 !important; background-color: #F4F7FA;">
    <div class="document" role="article" aria-roledescription="email" lang="nl" dir="ltr" style="background-color: #F4F7FA; line-height: 100%; font-size: medium; font-size: max(16px, 1rem);">
        <table width="100%" align="center" cellspacing="0" cellpadding="0" border="0">
            <tr>
                <td background="" bgcolor="#F4F7FA" align="center" valign="top" style="padding: 0 8px;">
                    <table class="container" align="center" width="640" cellpadding="0" cellspacing="0" border="0" style="max-width: 640px; margin-bottom: 0;">
                        <tr><td align="center"></td></tr>
                    </table>
                    <table width="640" class="wrapper" align="center" border="0" cellpadding="0" cellspacing="0" style="max-width: 640px; border: 1px solid #EAECED; border-radius: 8px; border-collapse: separate !important; overflow: hidden;">
                        <tr><td align="center">
                            <!-- Logo Section -->
                            <table class="ml-default" width="100%" border="0" cellspacing="0" cellpadding="0">
                                <tr><td>
                                    <table class="container ml-4 ml-default-border" width="640" bgcolor="#ffffff" align="center" border="0" cellspacing="0" cellpadding="0" style="color: #515856; width: 640px; min-width: 640px; margin-bottom: 0;">
                                        <tr><td class="ml-default-border container" height="20" style="line-height: 20px; min-width: 640px;"></td></tr>
                                        <tr><td class="row" style="padding: 0 50px;">
                                            <table align="center" width="100%" border="0" cellspacing="0" cellpadding="0">
                                                <tr>
                                                    <td class="col mobile-center" align="left" width="250">
                                                        <img src="https://storage.mlcdn.com/account_image/2074556/7TfhoufS1TCjs6j9QTvhhbioAu2JKwaBHdg693RM.png" border="0" alt="GKIN Logo" width="59" style="max-width: 59px; display: inline-block;">
                                                    </td>
                                                    <td class="col" width="40" height="20" style="line-height: 20px;"></td>
                                                    <td class="col mobile-center" align="right" style="text-align: right;">
                                                        <h3 style="font-family: 'Inter', sans-serif; color: #000000; font-size: 18px; line-height: 125%; font-weight: bold; margin-bottom: 0;">GKIN Amstelveen</h3>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td></tr>
                                    </table>
                                </td></tr>
                            </table>
                            <!-- Divider -->
                            <table class="ml-default" width="100%" border="0" cellspacing="0" cellpadding="0">
                                <tr><td>
                                    <table class="container ml-6 ml-default-border" width="640" bgcolor="#ffffff" align="center" border="0" cellspacing="0" cellpadding="0" style="width: 640px; min-width: 640px; margin-bottom: 0;">
                                        <tr><td class="ml-default-border container" height="20" style="line-height: 20px; min-width: 640px;"></td></tr>
                                        <tr><td class="row" style="padding: 0 50px;" align="center">
                                            <table role="presentation" cellpadding="0" cellspacing="0" border="0" align="center" width="100%">
                                                <tr><td style="border-top: 1px solid #EAECED;"></td></tr>
                                            </table>
                                        </td></tr>
                                        <tr><td height="20" style="line-height: 20px;"></td></tr>
                                    </table>
                                </td></tr>
                            </table>
                            <!-- Main Content -->
                            <table class="ml-default" width="100%" border="0" cellspacing="0" cellpadding="0">
                                <tr><td>
                                    <table class="container ml-8 ml-default-border" width="640" bgcolor="#ffffff" align="center" border="0" cellspacing="0" cellpadding="0" style="color: #515856; width: 640px; min-width: 640px; margin-bottom: 0;">
                                        <tr><td class="row" style="padding: 0 50px;">
                                            <table align="center" width="100%" border="0" cellspacing="0" cellpadding="0">
                                                <tr><td>
                                                    <p style="font-family: 'Inter', sans-serif; color: #515856; font-size: 16px; line-height: 137%; margin-top: 0; margin-bottom: 10px;">Beste broeders en zusters,<br></p>
                                                    <p style="font-family: 'Inter', sans-serif; color: #515856; font-size: 16px; line-height: 137%; margin-top: 0; margin-bottom: 10px;">Op {date_str} zal {predikant} voorgaan in de Online Landelijke Eredienst (OLE) van GKIN {location_body}, aanvang {time_display} uur.</p>
"""
        
        if theme:
            html += f"""                                                    <p style="font-family: 'Inter', sans-serif; color: #515856; font-size: 16px; line-height: 137%; margin-top: 0; margin-bottom: 10px;">Het thema van de dienst is: <strong>"{theme}"</strong>"""
            if bible_verse:
                html += f""" genomen uit {bible_verse}."""
            html += """<br><br></p>"""
        
        html += f"""                                                    <p style="font-family: 'Inter', sans-serif; color: #515856; font-size: 16px; line-height: 137%; margin-top: 0; margin-bottom: 0;">De dienst wordt live uitgezonden. De liturgie van de dienst kunt u hieronder vinden. Door op de link te klikken kunt u het bestand bekijken en downloaden. Via de eveneens hieronder vermelde link kunt u de dienst online volgen.</p>
                                                </td></tr>
                                            </table>
                                        </td></tr>
                                    </table>
                                </td></tr>
                            </table>
                            <!-- Buttons Section -->
                            <table class="ml-default" width="100%" border="0" cellspacing="0" cellpadding="0">
                                <tr><td>
                                    <table class="container ml-10 ml-default-border" width="640" bgcolor="#ffffff" align="center" border="0" cellspacing="0" cellpadding="0" style="width: 640px; min-width: 640px; margin-bottom: 0;">
                                        <tr><td class="ml-default-border container" height="20" style="line-height: 20px; min-width: 640px;"></td></tr>
                                        <tr><td class="row" style="padding: 0 50px;">
                                            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                                <tr>
                                                    <td class="col" width="250" valign="top">
                                                        <div class="ml-11" style="text-align: full;">
                                                            <table class="ml-btn-container" cellpadding="0" cellspacing="0" border="0" align="full" width="100%">
                                                                <tr><td align="center" valign="middle">
                                                                    <div>
                                                                        <table class="ml-btn ml-btn-secondary ml-btn-filled" border="0" cellpadding="0" cellspacing="0" role="presentation" style="border-collapse: separate; width: 250px;">
                                                                            <tr>
                                                                                <th align="center" valign="middle" style="background-color: #000000; border-radius: 6px; -webkit-font-smoothing: auto; word-break: break-all;">
                                                                                    <a href="{liturgie_url}" target="_blank" style="display: block; padding: 10px 25px; font-family: 'Inter', sans-serif; font-size: 14px; color: #ffffff; letter-spacing: 0.025em; text-decoration: none; border-radius: 6px; line-height: 16px; word-break: break-word; min-width: 70px; font-weight: normal;">
                                                                                        <div>Liturgie</div>
                                                                                        <div>({ole_location or 'OLE'})</div>
                                                                                    </a>
                                                                                </th>
                                                                            </tr>
                                                                        </table>
                                                                    </div>
                                                                </td></tr>
                                                            </table>
                                                        </div>
                                                    </td>
                                                    <td class="col" width="40" height="20" style="line-height: 20px;"></td>
                                                    <td class="col" width="250" valign="top">
                                                        <div class="ml-12" style="text-align: full;">
                                                            <table class="ml-btn-container" cellpadding="0" cellspacing="0" border="0" align="full" width="100%">
                                                                <tr><td align="center" valign="middle">
                                                                    <div>
                                                                        <table class="ml-btn ml-btn-secondary ml-btn-filled" border="0" cellpadding="0" cellspacing="0" role="presentation" style="border-collapse: separate; width: 250px;">
                                                                            <tr>
                                                                                <th align="center" valign="middle" style="background-color: #000000; border-radius: 6px; -webkit-font-smoothing: auto; word-break: break-all;">
                                                                                    <a href="{youtube_link}" target="_blank" style="display: block; padding: 10px 25px; font-family: 'Inter', sans-serif; font-size: 14px; color: #ffffff; letter-spacing: 0.025em; text-decoration: none; border-radius: 6px; line-height: 16px; word-break: break-word; min-width: 70px; font-weight: normal;">
                                                                                        <div>Webvideo</div>
                                                                                        <div>({ole_location or 'OLE'})</div>
                                                                                    </a>
                                                                                </th>
                                                                            </tr>
                                                                        </table>
                                                                    </div>
                                                                </td></tr>
                                                            </table>
                                                        </div>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td></tr>
                                        <tr><td height="10" style="line-height: 10px;"></td></tr>
                                    </table>
                                </td></tr>
                            </table>
                            <!-- Collection Section -->
                            <table class="ml-default" width="100%" border="0" cellspacing="0" cellpadding="0">
                                <tr><td>
                                    <table class="container ml-14 ml-default-border" width="640" bgcolor="#ffffff" align="center" border="0" cellspacing="0" cellpadding="0" style="color: #515856; width: 640px; min-width: 640px; margin-bottom: 0;">
                                        <tr><td class="ml-default-border container" height="20" style="line-height: 20px; min-width: 640px;"></td></tr>
                                        <tr><td class="row" style="padding: 0 50px;">
                                            <p style="font-family: 'Inter', sans-serif; color: #515856; font-size: 16px; line-height: 125%; margin-top: 0; margin-bottom: 10px;">In deze zondagsdienst wordt er 1 keer gecollecteerd.<br><br>De collecte is bestemd voor Landelijke kas ({service_type}). U kunt dit overmaken via: <a href="{collecte_url}" target="_blank" style="color: #2CB191; font-weight: normal; text-decoration: underline;">{collecte_url[:50]}...</a></p>
                                            <p style="font-family: 'Inter', sans-serif; color: #515856; font-size: 16px; line-height: 125%; margin-top: 0; margin-bottom: 0;">of door overmaking aan GEREJA KRISTEN INDONESIA NEDERLAND, IBAN: NL19 INGB 0002 6182 90 o.v.v. Collecte {service_type} {short_date}.</p>
                                        </td></tr>
                                    </table>
                                </td></tr>
                            </table>
                            <!-- QR Code Section -->
                            <table class="ml-default" width="100%" border="0" cellspacing="0" cellpadding="0">
                                <tr><td>
                                    <table class="container ml-16 ml-default-border" width="640" bgcolor="#ffffff" align="center" border="0" cellspacing="0" cellpadding="0" style="color: #515856; width: 640px; min-width: 640px; margin-bottom: 0;">
                                        <tr><td class="ml-default-border container" height="40" style="line-height: 40px; min-width: 640px;"></td></tr>
                                        <tr><td class="row" style="padding: 0 50px;">
                                            <table class="three-columns-layout" width="100%" border="0" cellspacing="0" cellpadding="0">
                                                <tr>
                                                    <td class="col" valign="top" width="160"></td>
                                                    <td class="col" width="30" height="30" style="line-height: 30px;"></td>
                                                    <td class="col" valign="top" width="160">
                                                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                                            <tr><td align="left">
                                                                <img src="{qr_image_url}" loading="lazy" border="0" alt="QR Code" width="160" style="display: block;">
                                                            </td></tr>
                                                        </table>
                                                    </td>
                                                    <td class="col" width="30" height="30" style="line-height: 30px;"></td>
                                                    <td class="col" valign="top" width="160"></td>
                                                </tr>
                                            </table>
                                        </td></tr>
                                        <tr><td height="40" style="line-height: 40px;"></td></tr>
                                    </table>
                                </td></tr>
                            </table>
                            <!-- Closing Section -->
                            <table class="ml-default" width="100%" border="0" cellspacing="0" cellpadding="0">
                                <tr><td>
                                    <table class="container ml-23 ml-default-border" width="640" bgcolor="#ffffff" align="center" border="0" cellspacing="0" cellpadding="0" style="color: #515856; width: 640px; min-width: 640px; margin-bottom: 0;">
                                        <tr><td class="row" style="padding: 0 50px;">
                                            <p style="font-family: 'Inter', sans-serif; color: #515856; font-size: 16px; line-height: 125%; margin-top: 0; margin-bottom: 10px;"><br></p>
                                            <p style="font-family: 'Inter', sans-serif; color: #515856; font-size: 16px; line-height: 125%; margin-top: 0; margin-bottom: 10px;">Wij wensen u allen een gezegende dienst toe.<br></p>
                                            <p style="font-family: 'Inter', sans-serif; color: #515856; font-size: 16px; line-height: 125%; margin-top: 0; margin-bottom: 10px;">Met broederlijke groet in Christus,</p>
                                            <p style="font-family: 'Inter', sans-serif; color: #515856; font-size: 16px; line-height: 125%; margin-top: 0; margin-bottom: 10px;"><br></p>
                                            <p style="font-family: 'Inter', sans-serif; color: #515856; font-size: 16px; line-height: 125%; margin-top: 0; margin-bottom: 10px;">Namens de landelijke kerkenraad GKIN,<br></p>
                                            <p style="font-family: 'Inter', sans-serif; color: #515856; font-size: 16px; line-height: 125%; margin-top: 0; margin-bottom: 0;">Vega Hardono, Regiosecretaris (AM)<br></p>
                                        </td></tr>
                                        <tr><td height="20" style="line-height: 20px;"></td></tr>
                                    </table>
                                </td></tr>
                            </table>
                            <!-- Divider -->
                            <table class="ml-default" width="100%" border="0" cellspacing="0" cellpadding="0">
                                <tr><td>
                                    <table class="container ml-25 ml-default-border" width="640" bgcolor="#ffffff" align="center" border="0" cellspacing="0" cellpadding="0" style="width: 640px; min-width: 640px; margin-bottom: 0;">
                                        <tr><td class="ml-default-border container" height="20" style="line-height: 20px; min-width: 640px;"></td></tr>
                                        <tr><td>
                                            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                                <tr><td class="row" style="padding: 0 50px;" align="center">
                                                    <table role="presentation" cellpadding="0" cellspacing="0" border="0" align="center" width="100%">
                                                        <tr><td style="border-top: 1px solid #EAECED;"></td></tr>
                                                    </table>
                                                </td></tr>
                                            </table>
                                        </td></tr>
                                        <tr><td height="20" style="line-height: 20px;"></td></tr>
                                    </table>
                                </td></tr>
                            </table>
                            <!-- Footer -->
                            <table class="ml-default" width="100%" border="0" cellspacing="0" cellpadding="0">
                                <tr><td>
                                    <table class="container ml-27 ml-default-border" width="640" bgcolor="#ffffff" align="center" border="0" cellspacing="0" cellpadding="0" style="width: 640px; min-width: 640px; margin-bottom: 0;">
                                        <tr><td class="row" style="padding: 0 50px;">
                                            <table align="center" width="100%" cellpadding="0" cellspacing="0" border="0">
                                                <tr>
                                                    <td class="col" align="left" width="250" valign="top" style="text-align: left !important;">
                                                        <p style="font-family: 'Inter', sans-serif; color: #515856; font-size: 14px; line-height: 150%; margin-bottom: 6px;"><strong>GKIN Amstelveen</strong></p>
                                                        <p style="font-family: 'Inter', sans-serif; color: #515856; font-size: 14px; line-height: 150%; margin-bottom: 0;">Bouwerij 52<br>1185XX Amstelveen</p>
                                                    </td>
                                                    <td class="col" width="40" height="30" style="line-height: 30px;"></td>
                                                    <td class="col" align="left" width="250" valign="top" style="text-align: left !important;">
                                                        <p style="font-family: 'Inter', sans-serif; color: #515856; font-size: 14px; line-height: 150%; margin-bottom: 6px;">Wilt u deze e-mails niet meer ontvangen? U kunt zich hier:</p>
                                                        <table width="100%" cellpadding="0" cellspacing="0" border="0">
                                                            <tr><td height="8" style="line-height: 8px;"></td></tr>
                                                            <tr><td align="left">
                                                                <p style="font-family: 'Inter', sans-serif; color: #515856; font-size: 14px; line-height: 150%; margin-bottom: 0;">
                                                                    <a href="{{unsubscribe}}" style="color: #515856; text-decoration: underline;">uitschrijven</a>
                                                                </p>
                                                            </td></tr>
                                                        </table>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td></tr>
                                        <tr><td height="20" style="line-height: 20px;"></td></tr>
                                    </table>
                                </td></tr>
                            </table>
                        </td></tr>
                    </table>
                </td></tr>
        </table>
    </div>
</body>
</html>"""
        return html

    def generate_campaign_preview(self,
                                  service_date: datetime,
                                  predikant: str,
                                  mededelingen_data: Dict[str, Any],
                                  takenrooster_entry: Optional[Dict] = None) -> Dict:
        """
        Generate preview data for a campaign without creating it.
        
        Returns:
            Dict with subject, html_content, and metadata
        """
        months = ['januari', 'februari', 'maart', 'april', 'mei', 'juni',
                  'juli', 'augustus', 'september', 'oktober', 'november', 'december']
        date_str = f"{service_date.day} {months[service_date.month - 1]} {service_date.year}"
        
        # Generate campaign name
        short_date = service_date.strftime('%y%m%d')
        campaign_name = f"GKIN Mededelingen {short_date}"
        
        # Generate subject line
        subject = f"Mededelingen GKIN Amstelveen – {date_str}"
        
        # Generate HTML content
        html_content = self.generate_html_from_mededelingen(
            service_date=service_date,
            predikant=predikant,
            mededelingen_data=mededelingen_data,
            takenrooster_entry=takenrooster_entry
        )
        
        return {
            'name': campaign_name,
            'subject': subject,
            'html_content': html_content,
            'service_date': date_str,
            'predikant': predikant,
            'ovd': takenrooster_entry.get('ovd', '') if takenrooster_entry else ''
        }

    def get_templates(self) -> List[Dict]:
        """Fetch available templates from MailerLite."""
        result = self._make_request("GET", "/templates")
        return result.get('data', [])
    
    def get_template(self, template_id: str) -> Dict:
        """Fetch a specific template by ID."""
        return self._make_request("GET", f"/templates/{template_id}")
    
    def create_campaign_from_template(self,
                                     name: str,
                                     subject: str,
                                     template_id: str,
                                     template_variables: Dict[str, str],
                                     attachment_urls: Optional[List[str]] = None,
                                     group_ids: Optional[List[str]] = None,
                                     send_time: Optional[str] = None) -> Dict:
        """
        Create campaign using a MailerLite template with variable substitution.
        
        Args:
            name: Campaign name
            subject: Email subject
            template_id: MailerLite template ID
            template_variables: Dict of {{placeholder}} -> value replacements
            attachment_urls: List of uploaded file URLs to include as links
            group_ids: Subscriber groups
            send_time: Schedule time
        
        Returns:
            API response with campaign details
        """
        # Get template content
        template = self.get_template(template_id)
        if 'error' in template:
            return template
        
        # Extract template HTML
        html_content = template.get('data', {}).get('html', '')
        
        # Replace template variables
        for key, value in template_variables.items():
            placeholder = f"{{{{{key}}}}}"
            html_content = html_content.replace(placeholder, str(value))
        
        # Add attachment links if provided
        if attachment_urls:
            attachment_section = '<div style="margin-top: 30px; padding: 15px; background: #f5f5f5; border-radius: 8px;"><h3>Bijlagen</h3><ul>'
            for url in attachment_urls:
                filename = url.split('/')[-1]
                attachment_section += f'<li><a href="{url}" target="_blank" style="color: #3A7C22;">{filename}</a></li>'
            attachment_section += '</ul></div>'
            
            # Insert before closing </body> tag or at end
            if '</body>' in html_content:
                html_content = html_content.replace('</body>', f'{attachment_section}</body>')
            else:
                html_content += attachment_section
        
        # Create campaign with template HTML
        return self.create_campaign(
            name=name,
            subject=subject,
            html_content=html_content,
            group_ids=group_ids,
            send_time=send_time
        )
    
    def create_campaign_with_attachments(self,
                                        name: str,
                                        subject: str,
                                        html_content: str,
                                        attachments: List[Dict],
                                        group_ids: Optional[List[str]] = None,
                                        send_time: Optional[str] = None) -> Dict:
        """
        Create campaign with file attachments uploaded to MailerLite.
        
        Args:
            attachments: List of dicts with 'path' (local file) or 'content' (bytes) + 'name'
        """
        from mailerlite_campaign import MailerLiteFileManager
        
        file_manager = MailerLiteFileManager(self.api_key)
        attachment_urls = []
        
        # Upload all attachments
        for att in attachments:
            if 'path' in att:
                result = file_manager.upload_file(att['path'], att.get('name'))
            elif 'content' in att:
                result = file_manager.upload_bytes(att['content'], att.get('name', 'attachment'))
            else:
                continue
            
            if result.get('success'):
                attachment_urls.append(result['url'])
        
        # Add attachment links to HTML
        if attachment_urls:
            attachment_html = '<div style="margin: 30px 0; padding: 20px; background: #f8f8f8; border-left: 4px solid #3A7C22;">'
            attachment_html += '<h3 style="margin-top: 0; color: #3A7C22;">📎 Bijlagen</h3><ul style="margin: 10px 0;">'
            for url in attachment_urls:
                filename = url.split('/')[-1]
                attachment_html += f'<li style="margin: 5px 0;"><a href="{url}" style="color: #3A7C22; text-decoration: none;">{filename}</a></li>'
            attachment_html += '</ul></div>'
            
            # Insert before closing body tag
            if '</body>' in html_content:
                html_content = html_content.replace('</body>', f'{attachment_html}</body>')
            else:
                html_content += attachment_html
        
        return self.create_campaign(
            name=name,
            subject=subject,
            html_content=html_content,
            group_ids=group_ids,
            send_time=send_time
        )
    
    def test_connection(self) -> Dict:
        """Test API connection and return status."""
        result = self._make_request("GET", "/groups")
        if 'error' in result:
            return {'success': False, 'error': result['error']}
        return {'success': True, 'groups_count': len(result.get('data', []))}
