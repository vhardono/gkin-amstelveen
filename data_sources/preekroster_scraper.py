"""
Preekroster Scraper
Scrapes preaching roster from church website
"""

import requests
from bs4 import BeautifulSoup
from typing import Dict, List, Any
from datetime import datetime, timedelta
from config import Config

class PreekrosterScraper:
    def __init__(self):
        """Initialize the scraper with session"""
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        })
        self.url = Config.PREEKROSTER_URL
    
    def get_preekroster(self, mededelingen_date: datetime = None) -> Dict[str, Any]:
        """Get preekroster from website, filtered to 4 weeks from mededelingen_date.

        Args:
            mededelingen_date: The selected mededelingen/bulletin date.
                               Only entries from this date onward (up to 4 weeks) are included.
                               Defaults to next Sunday.
        """
        if mededelingen_date is None:
            today = datetime.now().date()
            days_until_sunday = (6 - today.weekday()) % 7
            if days_until_sunday == 0:
                days_until_sunday = 7
            mededelingen_date = datetime.combine(
                today + timedelta(days=days_until_sunday), datetime.min.time()
            )

        start_date = mededelingen_date.date() if hasattr(mededelingen_date, 'date') and callable(mededelingen_date.date) else mededelingen_date
        end_date = start_date + timedelta(weeks=4)
        print(f"Preekroster period: {start_date.strftime('%d-%m-%Y')} to {end_date.strftime('%d-%m-%Y')}")

        try:
            response = self.session.get(self.url)
            response.raise_for_status()
            
            soup = BeautifulSoup(response.content, 'html.parser')

            # Parse region->time mapping from the regio info table (first table)
            region_times = self._parse_region_times(soup)

            # Try different parsing methods
            roster = self._parse_table_roster(soup)
            if not roster:
                roster = self._parse_list_roster(soup)
            if not roster:
                roster = self._parse_div_roster(soup)
            
            # Create two separate tables as requested - OLE table only
            am_table = []
            ole_table = []
            processed_dates = set()  # Track processed dates to avoid duplicates
            
            for entry in roster:
                # Filter by date range
                entry_real_date = self._parse_dutch_date(entry.get('date', ''))
                if entry_real_date is None:
                    continue
                if entry_real_date <= start_date or entry_real_date > end_date:
                    continue

                region = entry.get('region', '')
                formatted_entry = {
                    'date': self._format_date(entry.get('date', '')),
                    'language': entry.get('language', ''),
                    'predikant': entry.get('speaker', ''),
                    'time': region_times.get(region, region_times.get('AM', '10.30u'))
                }
                
                # Check if this is an OLE entry (has OLE in additional_info)
                additional_info = entry.get('additional_info', '').upper()
                is_ole = 'OLE' in additional_info
                
                entry_date = self._format_date(entry.get('date', ''))
                
                # Add all AM entries (regardless of additional_info)
                if region == 'AM' and entry_date not in processed_dates:
                    am_entry = dict(formatted_entry)
                    am_entry['time'] = region_times.get('AM', '10.30u')
                    am_table.append(am_entry)
                    processed_dates.add(entry_date)

                # Add to OLE table for ANY region marked OLE (including AM itself).
                # On dates where AM is the streaming source, the OLE row reflects that.
                if region and is_ole:
                    ole_date_exists = any(ole_entry.get('date') == entry_date for ole_entry in ole_table)
                    if not ole_date_exists:
                        ole_entry = dict(formatted_entry)
                        ole_entry['regio'] = region
                        ole_entry['time'] = region_times.get(region, formatted_entry['time'])
                        ole_table.append(ole_entry)
            
            return {
                'source': self.url,
                'retrieved_at': datetime.now().isoformat(),
                'am_table': am_table,
                'ole_table': ole_table,
                'roster': roster  # Keep original for compatibility
            }
            
        except requests.RequestException as e:
            print(f"Error fetching preekroster: {e}")
            return {'error': str(e), 'am_table': [], 'tevens_online_table': [], 'roster': []}
        except Exception as e:
            print(f"Error parsing preekroster: {e}")
            return {'error': str(e), 'roster': []}
    
    def _parse_dutch_date(self, date_text: str):
        """Parse a Dutch date string like 'Zondag 10 mei 2026' into a date object.

        Also handles multiline strings like 'Donderdag 14 mei 2026\\nHemelvaart'.
        Returns None on failure.
        """
        import re
        from datetime import date as date_cls
        if not date_text:
            return None

        main_line = date_text.splitlines()[0].strip()
        dutch_month_to_num = {
            'januari': 1, 'februari': 2, 'maart': 3, 'april': 4,
            'mei': 5, 'juni': 6, 'juli': 7, 'augustus': 8,
            'september': 9, 'oktober': 10, 'november': 11, 'december': 12
        }
        try:
            parts = main_line.split()
            # Expect: [Weekday, day, month, year]
            if len(parts) >= 4:
                day = int(parts[1])
                month = dutch_month_to_num.get(parts[2].lower(), 0)
                year = int(parts[3])
                if month:
                    return date_cls(year, month, day)
        except Exception:
            pass
        return None

    def _format_date(self, date_text: str) -> str:
        """Format date as 'dd mmm' in Dutch, appending any extra label like '(Hemelvaart)'."""
        if not date_text:
            return ''

        # Date strings can include a newline-separated extra label, e.g.
        # 'Donderdag 14 mei 2026\nHemelvaart' or 'Zondag 31 mei 2026\n5de zondag'.
        lines = [ln.strip() for ln in date_text.splitlines() if ln.strip()]
        main_line = lines[0] if lines else date_text
        extra = ' '.join(lines[1:]).strip() if len(lines) > 1 else ''

        # Dutch month abbreviations (lowercase, as commonly used in NL)
        dutch_months = {
            'jan': 'jan', 'feb': 'feb', 'maa': 'mrt', 'mar': 'mrt', 'mrt': 'mrt',
            'apr': 'apr', 'mei': 'mei', 'jun': 'jun', 'jul': 'jul', 'aug': 'aug',
            'sep': 'sep', 'okt': 'okt', 'nov': 'nov', 'dec': 'dec'
        }

        try:
            parts = main_line.split()
            # Expect: [Weekday, day, month, year]
            if len(parts) >= 3:
                day = parts[1]
                month_raw = parts[2].lower()
                month_nl = dutch_months.get(month_raw[:3], month_raw[:3])
                formatted = f"{day} {month_nl}"
                if extra:
                    formatted += f" ({extra})"
                return formatted
        except Exception:
            pass

        return date_text

    def _parse_region_times(self, soup: BeautifulSoup) -> Dict[str, str]:
        """Parse the regio/adres/aanvangstijd table to map region code -> time string.

        Returns times formatted like '10.30u' / '13.30u'.
        """
        region_times: Dict[str, str] = {}
        for table in soup.find_all('table'):
            header_cells = table.find_all('tr')[0].find_all(['th', 'td']) if table.find_all('tr') else []
            header_texts = [c.get_text(strip=True).lower() for c in header_cells]
            if any('aanvangstijd' in h for h in header_texts) and any('regio' in h for h in header_texts):
                for row in table.find_all('tr')[1:]:
                    cells = [c.get_text(strip=True) for c in row.find_all(['td', 'th'])]
                    if len(cells) >= 3:
                        regio = cells[0].strip()
                        time_raw = cells[-1].strip()
                        region_times[regio] = self._format_time(time_raw)
                break
        return region_times

    def _format_time(self, time_text: str) -> str:
        """Convert '10:30 uur' or '13:30' style strings to '10.30u'."""
        import re
        m = re.search(r'(\d{1,2})[:.](\d{2})', time_text)
        if m:
            return f"{int(m.group(1))}.{m.group(2)}u"
        return time_text.strip()
    
    def _derive_time_from_am_table(self, entry: Dict, full_roster: List[Dict]) -> str:
        """Derive time from AM table entry"""
        # Find corresponding AM entry to get time
        for am_entry in full_roster:
            if (am_entry.get('region') == 'AM' and 
                am_entry.get('date') == entry.get('date') and
                am_entry.get('speaker') == entry.get('speaker')):
                return '10.30u'  # Use same hardcoded time as AM table
        
        return '10.30u'  # Default fallback
    
    def _parse_table_roster(self, soup: BeautifulSoup) -> List[Dict[str, Any]]:
        """Parse roster from HTML table format"""
        roster = []
        
        # Look for tables with common roster-related class names or headers
        tables = soup.find_all('table')
        for table in tables:
            headers = table.find_all(['th', 'td'])
            has_date_header = any('datum' in th.get_text().lower() or 'date' in th.get_text().lower() 
                                 for th in headers if th.name == 'th')
            
            # Also check for headers with 'Zondag' (Sunday) which indicates main roster table
            has_sunday_header = any('zondag' in th.get_text().lower() 
                                   for th in headers)
            
            # Check for 'Predikant' header which is main roster table
            has_predikant_header = any('predikant' in th.get_text().lower() 
                                      for th in headers)
            
                        
            if has_date_header or has_sunday_header or has_predikant_header or 'rooster' in table.get('class', []):
                rows = table.find_all('tr')
                
                # Find the header row with 'Predikant' to determine column positions
                header_row = None
                date_row = None
                for row in rows:
                    headers = row.find_all(['th', 'td'])
                    header_texts = [h.get_text().strip().lower() for h in headers]
                    if 'predikant' in header_texts:
                        header_row = row
                        break
                    # Also capture the date row (usually first row with date)
                    if not date_row and any('zondag' in text.lower() or 'maandag' in text.lower() or 'dinsdag' in text.lower() or 'woensdag' in text.lower() or 'donderdag' in text.lower() or 'vrijdag' in text.lower() or 'zaterdag' in text.lower() for text in header_texts):
                        date_row = row
                
                                
                if header_row:
                    # Get column indices
                    header_cells = header_row.find_all(['th', 'td'])
                    predikant_idx = None
                    bijzonderheden_idx = None
                    taal_idx = None
                    
                    for i, cell in enumerate(header_cells):
                        cell_text = cell.get_text().strip().lower()
                        if 'predikant' in cell_text:
                            predikant_idx = i
                        elif 'bijzonderheden' in cell_text:
                            bijzonderheden_idx = i
                        elif 'taal' in cell_text:
                            taal_idx = i
                    
                    # Get date from date row
                    date_text = ''
                    if date_row:
                        date_cells = date_row.find_all(['th', 'td'])
                        if date_cells:
                            date_text = date_cells[0].get_text().strip()
                    
                    # Process data rows (skip date row and header row)
                    data_rows = [row for row in rows if row != header_row and row != date_row]
                    for row in data_rows:
                        cells = row.find_all(['td', 'th'])
                        if len(cells) >= 2:
                            roster_entry = self._extract_gkin_roster_entry(cells, predikant_idx, bijzonderheden_idx, taal_idx, date_text)
                            if roster_entry:
                                roster.append(roster_entry)
        
        return roster
    
    def _parse_list_roster(self, soup: BeautifulSoup) -> List[Dict[str, Any]]:
        """Parse roster from list format (ul/ol)"""
        roster = []
        
        # Look for lists that might contain roster information
        lists = soup.find_all(['ul', 'ol'])
        for lst in lists:
            items = lst.find_all('li')
            for item in items:
                text = item.get_text().strip()
                if self._is_roster_entry(text):
                    entry = self._parse_text_roster_entry(text)
                    if entry:
                        roster.append(entry)
        
        return roster
    
    def _parse_div_roster(self, soup: BeautifulSoup) -> List[Dict[str, Any]]:
        """Parse roster from div-based layout"""
        roster = []
        
        # Look for divs with common roster-related content
        divs = soup.find_all('div')
        for div in divs:
            text = div.get_text().strip()
            if self._is_roster_entry(text):
                entry = self._parse_text_roster_entry(text)
                if entry:
                    roster.append(entry)
        
        return roster
    
    def _is_roster_entry(self, text: str) -> bool:
        """Check if text looks like a roster entry"""
        # Look for patterns that suggest a roster entry
        indicators = ['prediker', 'spreker', 'voorganger', 'dominee', 'pastoor']
        date_patterns = ['/', '-', 'jan', 'feb', 'mrt', 'apr', 'mei', 'jun', 'jul', 'aug', 'sep', 'okt', 'nov', 'dec']
        
        text_lower = text.lower()
        has_indicator = any(indicator in text_lower for indicator in indicators)
        has_date = any(pattern in text for pattern in date_patterns)
        
        return has_indicator and has_date
    
    def _extract_gkin_roster_entry(self, cells: List, predikant_idx: int, bijzonderheden_idx: int, taal_idx: int, date_text: str) -> Dict[str, Any]:
        """Extract GKIN roster entry from table cells"""
        if len(cells) < 2:
            return None
        
        # Extract region and preacher based on column positions
        region = cells[0].get_text().strip() if len(cells) > 0 else ''
        predikant = cells[predikant_idx].get_text().strip() if predikant_idx is not None and predikant_idx < len(cells) else ''
        taal = cells[taal_idx].get_text().strip() if taal_idx is not None and taal_idx < len(cells) else ''
        bijzonderheden = cells[bijzonderheden_idx].get_text().strip() if bijzonderheden_idx is not None and bijzonderheden_idx < len(cells) else ''
        
        return {
            'region': region,
            'speaker': self._clean_name(predikant),
            'language': taal,
            'service_type': 'Ochtenddienst',  # Default for GKIN
            'additional_info': bijzonderheden,
            'date': date_text  # Use the date from the date row
        }
    
    def _extract_roster_entry(self, cells: List) -> Dict[str, Any]:
        """Extract roster entry from table cells"""
        if len(cells) < 2:
            return None
        
        # Try to identify date and speaker/preacher
        date_text = cells[0].get_text().strip()
        speaker_text = cells[1].get_text().strip() if len(cells) > 1 else ''
        service_text = cells[2].get_text().strip() if len(cells) > 2 else ''
        
        return {
            'date': self._clean_date(date_text),
            'speaker': self._clean_name(speaker_text),
            'service_type': self._clean_service_type(service_text),
            'additional_info': ' '.join(cell.get_text().strip() for cell in cells[3:]) if len(cells) > 3 else ''
        }
    
    def _parse_text_roster_entry(self, text: str) -> Dict[str, Any]:
        """Parse roster entry from plain text"""
        # This is a simplified parser - you may need to adjust based on actual format
        parts = text.split('-')
        if len(parts) >= 2:
            date_part = parts[0].strip()
            rest = '-'.join(parts[1:]).strip()
            
            # Try to extract speaker name
            speaker = ''
            for keyword in ['prediker:', 'spreker:', 'voorganger:', 'dominee:']:
                if keyword in rest.lower():
                    speaker_parts = rest.lower().split(keyword)
                    if len(speaker_parts) > 1:
                        speaker = speaker_parts[1].split(',')[0].strip()
                        break
            
            return {
                'date': self._clean_date(date_part),
                'speaker': self._clean_name(speaker),
                'service_type': self._clean_service_type(rest),
                'additional_info': rest
            }
        
        return None
    
    def _clean_date(self, date_text: str) -> str:
        """Clean and standardize date format"""
        # Basic date cleaning - you may need to enhance this
        return date_text.strip()
    
    def _clean_name(self, name_text: str) -> str:
        """Clean speaker/preacher name"""
        return name_text.strip().title()
    
    def _clean_service_type(self, service_text: str) -> str:
        """Clean service type information"""
        service_text = service_text.lower()
        if 'ochtend' in service_text:
            return 'Ochtenddienst'
        elif 'middag' in service_text:
            return 'Middagdienst'
        elif 'avond' in service_text:
            return 'Avonddienst'
        return service_text.strip().title()
