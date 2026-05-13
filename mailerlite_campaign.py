"""
MailerLite Campaign Automation for GKIN Amstelveen
Creates email campaigns based on mededelingen and liturgie data.
"""

import os
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Any
import requests


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
                                        takenrooster_entry: Optional[Dict] = None) -> str:
        """
        Generate HTML email content from mededelingen and liturgie data.
        
        Args:
            service_date: The service date
            predikant: Predikant name
            mededelingen_data: Data from DropboxExcelReader.get_mededelingen()
            liturgie_data: Optional liturgie details
            takenrooster_entry: Optional takenrooster entry with OvD, etc.
        
        Returns:
            Complete HTML email content
        """
        months = ['januari', 'februari', 'maart', 'april', 'mei', 'juni',
                  'juli', 'augustus', 'september', 'oktober', 'november', 'december']
        date_str = f"{service_date.day} {months[service_date.month - 1]} {service_date.year}"
        
        html = f"""<!DOCTYPE html>
<html lang="nl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>GKIN Amstelveen Mededelingen</title>
    <style>
        body {{ font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 600px; margin: 0 auto; padding: 20px; }}
        .header {{ background-color: #3A7C22; color: white; padding: 20px; text-align: center; border-radius: 8px 8px 0 0; }}
        .header h1 {{ margin: 0; font-size: 24px; }}
        .header p {{ margin: 5px 0 0 0; opacity: 0.9; }}
        .content {{ background: white; padding: 25px; border: 1px solid #e0e0e0; }}
        .section {{ margin-bottom: 25px; }}
        .section h2 {{ color: #3A7C22; font-size: 18px; border-bottom: 2px solid #3A7C22; padding-bottom: 8px; margin-bottom: 15px; }}
        .section p {{ margin: 10px 0; }}
        .liturgie-item {{ display: flex; justify-content: space-between; padding: 8px 0; border-bottom: 1px solid #f0f0f0; }}
        .liturgie-item:last-child {{ border-bottom: none; }}
        .label {{ font-weight: bold; color: #555; }}
        .value {{ color: #333; }}
        .footer {{ background: #f5f5f5; padding: 20px; text-align: center; font-size: 12px; color: #666; border-radius: 0 0 8px 8px; }}
        .cta-button {{ display: inline-block; background-color: #3A7C22; color: white; padding: 12px 30px; text-decoration: none; border-radius: 5px; margin: 15px 0; }}
        .divider {{ border: none; height: 1px; background: #e0e0e0; margin: 20px 0; }}
    </style>
</head>
<body>
    <div class="header">
        <h1>GKIN Amstelveen</h1>
        <p>Mededelingen – {date_str}</p>
    </div>
    
    <div class="content">
        <div class="section">
            <h2>🙏 Dienst van zondag {date_str}</h2>
            <div class="liturgie-item">
                <span class="label">Voorganger:</span>
                <span class="value">{predikant or 'T.B.A.'}</span>
            </div>
"""
        
        if takenrooster_entry:
            ovd = takenrooster_entry.get('ovd', '')
            if ovd:
                html += f"""
            <div class="liturgie-item">
                <span class="label">Ouderling van Dienst:</span>
                <span class="value">{ovd}</span>
            </div>"""
            
            beamer = takenrooster_entry.get('beamer', '')
            if beamer:
                html += f"""
            <div class="liturgie-item">
                <span class="label">Beamer:</span>
                <span class="value">{beamer}</span>
            </div>"""
            
            voorzangers = takenrooster_entry.get('voorzangers', '')
            if voorzangers:
                html += f"""
            <div class="liturgie-item">
                <span class="label">Voorzangers:</span>
                <span class="value">{voorzangers}</span>
            </div>"""
        
        html += """
        </div>
"""
        
        # Add liturgie sections if available
        if liturgie_data:
            html += """
        <div class="section">
            <h2>🎵 Liturgie</h2>
"""
            for item in liturgie_data.get('items', []):
                title = item.get('title', '')
                type_ = item.get('type', '')
                if title:
                    html += f"""
            <div class="liturgie-item">
                <span class="label">{type_}:</span>
                <span class="value">{title}</span>
            </div>"""
            html += """
        </div>
"""
        
        # Add mededelingen sections
        landelijke = mededelingen_data.get('landelijke_nl', '')
        if landelijke:
            html += f"""
        <div class="section">
            <h2>📢 Landelijke Mededelingen</h2>
            <p>{landelijke.replace(chr(10), '<br>')}</p>
        </div>
"""
        
        regionale = mededelingen_data.get('regionale_nl', '')
        if regionale:
            html += f"""
        <div class="section">
            <h2>📢 Regionale Mededelingen</h2>
            <p>{regionale.replace(chr(10), '<br>')}</p>
        </div>
"""
        
        html += f"""
        <hr class="divider">
        
        <div class="section" style="text-align: center;">
            <p style="font-size: 14px; color: #666;">
                Volledige mededelingen en liturgie zijn beschikbaar op onze website.<br>
                <a href="https://gkin.nl" class="cta-button">Bezoek gkin.nl</a>
            </p>
        </div>
    </div>
    
    <div class="footer">
        <p>GKIN Amstelveen | <a href="mailto:kerkenraad@gkin.nl" style="color: #3A7C22;">kerkenraad@gkin.nl</a></p>
        <p style="margin-top: 10px;">
            <a href="{{unsubscribe}}" style="color: #666;">Uitschrijven</a> | 
            <a href="{{update_profile}}" style="color: #666;">Profiel bijwerken</a>
        </p>
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

    def test_connection(self) -> Dict:
        """Test API connection and return status."""
        result = self._make_request("GET", "/groups")
        if 'error' in result:
            return {'success': False, 'error': result['error']}
        return {'success': True, 'groups_count': len(result.get('data', []))}
