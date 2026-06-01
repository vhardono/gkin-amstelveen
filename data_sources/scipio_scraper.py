"""
Scipio Scraper
Authenticated scraper for birthday list from Scipio website
"""

import requests
from bs4 import BeautifulSoup
from typing import Dict, List, Any
from datetime import datetime, timedelta
from urllib.parse import urljoin
from config import Config

class ScipioScraper:
    def __init__(self):
        """Initialize the scraper with session"""
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        })
        self.base_url = Config.SCIPIO_BASE_URL
        self.logged_in = False
    
    def login(self) -> bool:
        """Login to Scipio system"""
        if not Config.SCIPIO_USERNAME or not Config.SCIPIO_PASSWORD:
            raise ValueError("Scipio credentials not configured")
        
        try:
            # First, get the login page to find the form
            login_url = urljoin(self.base_url, '/login')
            response = self.session.get(login_url)
            response.raise_for_status()
            
            soup = BeautifulSoup(response.content, 'html.parser')
            
            # Find login form details
            login_form = soup.find('form') or soup.find('form', {'action': True})
            
            # Extract form action and any hidden fields
            form_action = login_form.get('action') if login_form else '/login'
            form_action = urljoin(self.base_url, form_action)
            
            # Get all hidden inputs
            form_data = {}
            hidden_inputs = soup.find_all('input', {'type': 'hidden'})
            for hidden in hidden_inputs:
                name = hidden.get('name')
                value = hidden.get('value', '')
                if name:
                    form_data[name] = value
            
            # Add username, password, and pin with correct field names
            form_data['ctl00$Main$txtLogin'] = Config.SCIPIO_USERNAME
            form_data['ctl00$Main$txtPassword'] = Config.SCIPIO_PASSWORD
            if Config.SCIPIO_PIN:
                form_data['ctl00$Main$txtPIN'] = Config.SCIPIO_PIN
            
            # Submit login form with proper ASP.NET event targeting
            form_data['__EVENTTARGET'] = 'ctl00$Main$btnMyLogin'  # Trigger login button
            form_data['ctl00$Main$btnMyLogin.x'] = '1'  # Image button click coordinate
            form_data['ctl00$Main$btnMyLogin.y'] = '1'  # Image button click coordinate
            
            login_response = self.session.post(form_action, data=form_data)
            login_response.raise_for_status()
            
            # Update session with login response cookies
            if login_response.cookies:
                self.session.cookies.update(login_response.cookies)
            
            # Check if login was successful
            if self._is_login_successful(login_response):
                self.logged_in = True
                print("Successfully logged into Scipio")
                
                # Ensure session cookies are properly set
                print(f"Session cookies after login: {len(self.session.cookies)}")
                return True
            else:
                print("Login failed - check credentials")
                return False
                
        except requests.RequestException as e:
            print(f"Error during login: {e}")
            return False
        except Exception as e:
            print(f"Unexpected error during login: {e}")
            return False
    
    def _is_login_successful(self, response: requests.Response) -> bool:
        """Check if login was successful"""
        # Check for common indicators of successful login
        content = response.text.lower()
        
        # Look for logout link or user info
        if 'logout' in content or 'uitloggen' in content:
            return True
        
        # Check if we're redirected away from login page
        if 'login' not in response.url.lower() and response.status_code == 200:
            return True
        
        # Check for error messages
        error_indicators = ['error', 'fout', 'incorrect', 'verkeerd', 'invalid']
        if any(indicator in content for indicator in error_indicators):
            return False
        
        return True
    
    def get_birthday_list(self, mededelingen_date: datetime = None) -> Dict[str, Any]:
        """Get birthday list from Scipio for a 14-day window around the mededelingen date.

        The window is: mededelingen_date - 6 days to mededelingen_date + 7 days.
        E.g. for 10 May: searches 4 May to 17 May.

        Args:
            mededelingen_date: The selected mededelingen/bulletin date.
                               Defaults to next Sunday.
        """
        if not self.logged_in:
            if not self.login():
                return {'error': 'Failed to login to Scipio', 'birthdays': []}

        # Default to next Sunday
        if mededelingen_date is None:
            today = datetime.now().date()
            days_until_sunday = (6 - today.weekday()) % 7
            if days_until_sunday == 0:
                days_until_sunday = 7
            mededelingen_date = datetime.combine(
                today + timedelta(days=days_until_sunday), datetime.min.time()
            )

        # Calculate date range: selected date -6 to +7
        start_date = mededelingen_date - timedelta(days=6)
        end_date = mededelingen_date + timedelta(days=7)
        start_ddmm = start_date.strftime('%d%m')
        end_ddmm = end_date.strftime('%d%m')
        print(f"Birthday period: {start_date.strftime('%d-%m')} to {end_date.strftime('%d-%m')} "
              f"(mededelingen date: {mededelingen_date.strftime('%d-%m-%Y')})")

        try:
            url = urljoin(self.base_url, '/ledenadministratie/overzichten/birthday.aspx')
            print(f"Fetching birthday form: {url}")
            response = self.session.get(url, timeout=15)
            response.raise_for_status()

            soup = BeautifulSoup(response.content, 'html.parser')
            birthdays = self._submit_birthday_form(soup, url, start_ddmm, end_ddmm)

            print(f"Found {len(birthdays)} birthdays (status=actief)")

            # Split into two-column layout: first half left, second half right
            import math
            half = math.ceil(len(birthdays) / 2)
            left_col = birthdays[:half]
            right_col = birthdays[half:]

            birthday_table = []
            for i in range(half):
                left = left_col[i]['entry'] if i < len(left_col) else ''
                right = right_col[i]['entry'] if i < len(right_col) else ''
                birthday_table.append({'left': left, 'right': right})

            return {
                'source': url,
                'retrieved_at': datetime.now().isoformat(),
                'birthdays': birthdays,
                'birthday_table': birthday_table,
            }

        except Exception as e:
            print(f"Error fetching birthday list: {e}")
            return {'error': str(e), 'birthdays': []}
    
    def _submit_birthday_form(self, soup: BeautifulSoup, base_url: str,
                               start_ddmm: str, end_ddmm: str) -> List[Dict[str, Any]]:
        """Submit the Scipio birthday form and parse the result table."""
        try:
            form = soup.find('form')
            if not form:
                return []

            # Collect all hidden / text form fields
            form_data = {}
            for inp in form.find_all('input'):
                name = inp.get('name')
                if not name:
                    continue
                input_type = inp.get('type', '')
                if input_type == 'image':
                    continue
                if input_type in ('checkbox', 'radio'):
                    if inp.has_attr('checked'):
                        form_data[name] = inp.get('value', 'on')
                else:
                    form_data[name] = inp.get('value', '')

            # Collect select fields
            for sel in form.find_all('select'):
                name = sel.get('name')
                if name:
                    selected = sel.find('option', selected=True)
                    form_data[name] = (selected.get('value') if selected
                                       else sel.find('option').get('value', ''))

            # Use period mode with the requested date range
            form_data['ctl00$Main$optMaand'] = 'optVerjMaand2'
            form_data['ctl00$Main$txtDDMM1'] = start_ddmm
            form_data['ctl00$Main$txtDDMM2'] = end_ddmm

            # Image-button click coordinates (required by ASP.NET)
            form_data['ctl00$Main$btnExecute.x'] = '1'
            form_data['ctl00$Main$btnExecute.y'] = '1'

            submit_url = urljoin(base_url, form.get('action', ''))
            response = self.session.post(submit_url, data=form_data, timeout=30)
            response.raise_for_status()

            # Parse the result table
            result_soup = BeautifulSoup(response.content, 'html.parser')
            return self._parse_birthday_result_table(result_soup)

        except Exception as e:
            print(f"Error submitting birthday form: {e}")
            return []

    def _lookup_roepnaam_achternaam(self, regnr: str) -> tuple:
        """Look up Roepnaam and Achternaam from the persoonskaart via search redirect.

        Returns (roepnaam, achternaam) or (None, None) on failure.
        """
        try:
            pk_url = urljoin(
                self.base_url,
                f'/ledenadministratie/search.aspx?regnr={regnr}&query=&target=persoonskaart'
            )
            response = self.session.get(pk_url, timeout=15)
            if response.status_code != 200:
                return (None, None)

            soup = BeautifulSoup(response.content, 'html.parser')
            roepnaam_el = soup.find('input', {'id': lambda x: x and 'txtRoepnaam' in x})
            achternaam_el = soup.find('input', {'id': lambda x: x and 'txtAchternaam' in x})

            roepnaam = roepnaam_el.get('value', '').strip() if roepnaam_el else None
            achternaam = achternaam_el.get('value', '').strip() if achternaam_el else None
            return (roepnaam, achternaam)
        except Exception:
            return (None, None)

    def _parse_birthday_result_table(self, soup: BeautifulSoup) -> List[Dict[str, Any]]:
        """Parse the birthday result table returned by Scipio.

        Expected columns: #, Regnr, Wijk/sectie, Naam, G, Adres,
                          Postcode, Plaatsnaam, Geb.datum, Leeftijd, Status
        Only rows with Status == 'actief' are included.
        Looks up Roepnaam + Achternaam from persoonskaart for each person.
        """
        birthdays = []
        for table in soup.find_all('table'):
            rows = table.find_all('tr')
            if len(rows) < 2:
                continue
            header_cells = [c.get_text(strip=True).lower() for c in rows[0].find_all(['th', 'td'])]
            if 'naam' not in header_cells:
                continue

            # Map column names to indices
            col = {name: i for i, name in enumerate(header_cells)}

            for row in rows[1:]:
                cells = [c.get_text(strip=True) for c in row.find_all(['td', 'th'])]
                if len(cells) < len(header_cells):
                    continue

                status = cells[col['status']].lower() if 'status' in col else ''
                if status != 'actief':
                    continue

                regnr = cells[col.get('regnr', 1)]
                fallback_name = cells[col.get('naam', 3)]
                gender = cells[col.get('g', 4)]
                birth_date_raw = cells[col.get('geb.datum', 8)]
                salutation = 'zr.' if gender == 'V' else 'br.'

                # Look up Roepnaam + Achternaam from persoonskaart
                roepnaam, achternaam = self._lookup_roepnaam_achternaam(regnr)
                if roepnaam and achternaam:
                    display_name = f'{roepnaam} {achternaam}'
                else:
                    display_name = fallback_name

                # Format birth date as dd-mm (NEVER include year for privacy)
                dd_mm = birth_date_raw
                try:
                    parts = birth_date_raw.split()
                    if len(parts) >= 2:
                        day = int(parts[0])
                        month_map = {
                            'januari': 1, 'februari': 2, 'maart': 3, 'april': 4,
                            'mei': 5, 'juni': 6, 'juli': 7, 'augustus': 8,
                            'september': 9, 'oktober': 10, 'november': 11, 'december': 12,
                            'jan': 1, 'feb': 2, 'mrt': 3, 'mar': 3, 'apr': 4,
                            'mei': 5, 'jun': 6, 'jul': 7, 'aug': 8,
                            'sep': 9, 'sept': 9, 'okt': 10, 'nov': 11, 'dec': 12
                        }
                        month_num = month_map.get(parts[1].lower(), 0)
                        if month_num:
                            dd_mm = f'{day:02d}-{month_num:02d}'
                        else:
                            # Try parsing as dd-mm-yyyy or dd/mm/yyyy format
                            import re
                            date_match = re.match(r'(\d{1,2})[/-](\d{1,2})', birth_date_raw)
                            if date_match:
                                dd_mm = f'{int(date_match.group(1)):02d}-{int(date_match.group(2)):02d}'
                except Exception:
                    # If all parsing fails, try regex to extract just day-month
                    import re
                    try:
                        # Match patterns like "01 jun 1994" or "01-06-1994" or "01/06/1994"
                        patterns = [
                            r'(\d{1,2})\s+([a-z]{3,})',  # 01 jun
                            r'(\d{1,2})[/-](\d{1,2})'    # 01-06 or 01/06
                        ]
                        for pattern in patterns:
                            match = re.search(pattern, birth_date_raw, re.IGNORECASE)
                            if match:
                                day = int(match.group(1))
                                month = match.group(2).lower()
                                if month.isdigit():
                                    month_num = int(month)
                                else:
                                    month_num = month_map.get(month, 0)
                                if month_num:
                                    dd_mm = f'{day:02d}-{month_num:02d}'
                                    break
                    except Exception:
                        pass

                birthdays.append({
                    'entry': f'{dd_mm}: {salutation} {display_name}',
                })

        return birthdays
    
    def _search_for_birthdays(self, url: str) -> List[Dict[str, Any]]:
        """Search for birthday information on a page"""
        try:
            response = self.session.get(url)
            response.raise_for_status()
            
            soup = BeautifulSoup(response.content, 'html.parser')
            
            birthdays = []
            # Look for text containing birthday-related keywords
            birthday_keywords = ['verjaardag', 'jarig', 'geboren', 'birthday']
            
            for element in soup.find_all(text=True):
                text = element.strip()
                if any(keyword in text.lower() for keyword in birthday_keywords):
                    birthday = self._extract_birthday_from_text(text)
                    if birthday:
                        birthdays.append(birthday)
            
            return birthdays
            
        except requests.RequestException:
            return []
    
    def _parse_table_birthdays(self, soup: BeautifulSoup) -> List[Dict[str, Any]]:
        """Parse birthdays from HTML table"""
        birthdays = []
        
        tables = soup.find_all('table')
        for table in tables:
            # Check if table contains birthday-related headers
            headers = table.find_all(['th', 'td'])
            has_birthday_header = any(
                'verjaardag' in th.get_text().lower() or 
                'geboortedatum' in th.get_text().lower() or
                'birthday' in th.get_text().lower()
                for th in headers if th.name == 'th'
            )
            
            if has_birthday_header:
                rows = table.find_all('tr')[1:]  # Skip header row
                for row in rows:
                    cells = row.find_all(['td', 'th'])
                    if len(cells) >= 2:
                        birthday = self._extract_birthday_from_cells(cells)
                        if birthday:
                            birthdays.append(birthday)
        
        return birthdays
    
    def _parse_list_birthdays(self, soup: BeautifulSoup) -> List[Dict[str, Any]]:
        """Parse birthdays from list format"""
        birthdays = []
        
        lists = soup.find_all(['ul', 'ol'])
        for lst in lists:
            items = lst.find_all('li')
            for item in items:
                text = item.get_text().strip()
                if self._is_birthday_entry(text):
                    birthday = self._extract_birthday_from_text(text)
                    if birthday:
                        birthdays.append(birthday)
        
        return birthdays
    
    def _parse_card_birthdays(self, soup: BeautifulSoup) -> List[Dict[str, Any]]:
        """Parse birthdays from card/div format"""
        birthdays = []
        
        # Look for divs that might contain birthday information
        divs = soup.find_all('div', class_=lambda x: x and 'birthday' in x.lower())
        if not divs:
            divs = soup.find_all('div', class_=lambda x: x and 'verjaardag' in x.lower())
        
        for div in divs:
            text = div.get_text().strip()
            if self._is_birthday_entry(text):
                birthday = self._extract_birthday_from_text(text)
                if birthday:
                    birthdays.append(birthday)
        
        return birthdays
    
    def _is_birthday_entry(self, text: str) -> bool:
        """Check if text looks like a birthday entry"""
        birthday_indicators = ['verjaardag', 'jarig', 'geboren', 'birthday', 'leeftijd']
        date_indicators = ['/', '-', 'jan', 'feb', 'mrt', 'apr', 'mei', 'jun', 'jul', 'aug', 'sep', 'okt', 'nov', 'dec']
        
        text_lower = text.lower()
        has_birthday = any(indicator in text_lower for indicator in birthday_indicators)
        has_date = any(pattern in text for pattern in date_indicators)
        
        return has_birthday and len(text.split()) >= 2
    
    def _extract_birthday_from_cells(self, cells: List) -> Dict[str, Any]:
        """Extract birthday from table cells"""
        if len(cells) < 2:
            return None
        
        name = cells[0].get_text().strip()
        date_info = cells[1].get_text().strip()
        age = cells[2].get_text().strip() if len(cells) > 2 else ''
        
        return {
            'name': self._clean_name(name),
            'birth_date': self._clean_date(date_info),
            'age': self._clean_age(age),
            'additional_info': ' '.join(cell.get_text().strip() for cell in cells[3:]) if len(cells) > 3 else ''
        }
    
    def _extract_birthday_from_text(self, text: str) -> Dict[str, Any]:
        """Extract birthday from plain text"""
        # This is a simplified parser - adjust based on actual format
        parts = text.split()
        if len(parts) >= 2:
            name = parts[0]
            date_info = ' '.join(parts[1:])
            
            return {
                'name': self._clean_name(name),
                'birth_date': self._clean_date(date_info),
                'age': self._extract_age_from_text(date_info),
                'additional_info': text
            }
        
        return None
    
    def _extract_age_from_text(self, text: str) -> str:
        """Extract age from text"""
        import re
        
        # Look for age patterns like "25 jaar", "25j", etc.
        age_patterns = [r'(\d+)\s*jaar', r'(\d+)\s*j', r'leeftijd\s*[:\-]?\s*(\d+)']
        
        for pattern in age_patterns:
            match = re.search(pattern, text.lower())
            if match:
                return match.group(1)
        
        return ''
    
    def _clean_name(self, name_text: str) -> str:
        """Clean person's name"""
        return name_text.strip().title()
    
    def _clean_date(self, date_text: str) -> str:
        """Clean and standardize date format"""
        return date_text.strip()
    
    def _clean_age(self, age_text: str) -> str:
        """Clean age information"""
        import re
        
        # Extract numbers from age text
        numbers = re.findall(r'\d+', age_text)
        return numbers[0] if numbers else ''
