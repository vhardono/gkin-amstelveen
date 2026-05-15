"""
GKIN OLE Website Scraper
Scrapes https://gkin.org/main/index.php/nl/online-landelijke-eredienst
to extract service details for a given target date.

Fields extracted per article:
  - date (datetime)
  - predikant
  - location (city name)
  - time
  - thema
  - bible_verse
  - youtube_link
  - liturgie_url  (direct PDF link on gkin.org)
  - collecte_url  (ING/OLE payment link)
"""

import re
import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
from typing import Dict, Any, Optional, List

OLE_INDEX_URL = "https://gkin.org/main/index.php/nl/online-landelijke-eredienst"

NL_MONTHS = {
    'januari': 1, 'februari': 2, 'maart': 3, 'april': 4,
    'mei': 5, 'juni': 6, 'juli': 7, 'augustus': 8,
    'september': 9, 'oktober': 10, 'november': 11, 'december': 12,
}

LOCATION_MAP = {
    'tilburg': ('TB', 'Pauluskerk te Tilburg'),
    'amsterdam': ('AM', 'Amsterdam'),
    'amstelveen': ('AM', 'Amstelveen'),
    'rotterdam': ('RT', 'Rotterdam'),
    'den haag': ('DH', 'Den Haag'),
    'groningen': ('GN', 'Groningen'),
    'utrecht': ('UT', 'Utrecht'),
    'eindhoven': ('EH', 'Eindhoven'),
}


class GKINOLEScraper:
    def __init__(self, timeout: int = 15):
        self.timeout = timeout
        self.session = requests.Session()
        self.session.headers.update({'User-Agent': 'Mozilla/5.0 (compatible; GKINBot/1.0)'})

    def _get(self, url: str) -> Optional[BeautifulSoup]:
        try:
            resp = self.session.get(url, timeout=self.timeout)
            resp.raise_for_status()
            return BeautifulSoup(resp.text, 'html.parser')
        except Exception as e:
            print(f"[GKINScraper] GET error {url}: {e}")
            return None

    def _parse_date_from_text(self, text: str) -> Optional[datetime]:
        """Parse Dutch date like '14 mei 2026' or 'donderdag 14 mei 2026'."""
        m = re.search(
            r'(\d{1,2})\s+(januari|februari|maart|april|mei|juni|juli|augustus|'
            r'september|oktober|november|december)\s+(\d{4})',
            text, re.IGNORECASE
        )
        if m:
            day, month_nl, year = int(m.group(1)), m.group(2).lower(), int(m.group(3))
            month = NL_MONTHS.get(month_nl)
            if month:
                return datetime(year, month, day)
        return None

    def _get_article_links(self, pages: int = 2) -> List[Dict[str, str]]:
        """Return list of {title, url} from the OLE index (first N pages)."""
        links = []
        for page in range(pages):
            url = OLE_INDEX_URL if page == 0 else f"{OLE_INDEX_URL}?start={page * 20}"
            soup = self._get(url)
            if not soup:
                continue
            for a in soup.find_all('a', href=True):
                href = a['href']
                if '/online-landelijke-eredienst/' in href and re.search(r'/\d+-', href):
                    full = href if href.startswith('http') else f"https://gkin.org{href}"
                    title = a.get_text(strip=True)
                    if title and full not in [l['url'] for l in links]:
                        links.append({'title': title, 'url': full})
        return links

    def _parse_article(self, url: str) -> Dict[str, Any]:
        """Scrape a single OLE article page and return extracted fields."""
        result: Dict[str, Any] = {
            'url': url, 'date': None, 'predikant': '', 'location': '',
            'location_code': '', 'time': '', 'thema': '', 'bible_verse': '',
            'youtube_link': '', 'liturgie_url': '', 'collecte_url': '',
        }
        soup = self._get(url)
        if not soup:
            return result

        # Get main article text
        article = soup.find('div', class_='item-page') or soup.find('article') or soup.find('main') or soup
        text = article.get_text(' ', strip=True)
        text = re.sub(r'\s+', ' ', text)

        # --- Date ---
        result['date'] = self._parse_date_from_text(text)

        # --- Predikant ---
        pred_m = re.search(
            r'\b(ds\.|zr\.|br\.)\s+([A-Z][^\.,]+?)(?=\s+voorgaan|\s+zal\s|\s+in\s+de\s+OLE)',
            text, re.IGNORECASE
        )
        if pred_m:
            result['predikant'] = pred_m.group(0).strip().rstrip(',')

        # --- Location ---
        loc_m = re.search(r'vanuit\s+(?:de\s+)?([^\.,]+?)(?:\s+te\s+([A-Za-z\s]+?))?(?:,|\.|aanvang)', text, re.IGNORECASE)
        if loc_m:
            loc_raw = (loc_m.group(2) or loc_m.group(1)).strip().lower()
            for key, (code, full) in LOCATION_MAP.items():
                if key in loc_raw:
                    result['location_code'] = code
                    result['location'] = full
                    break
            if not result['location']:
                result['location'] = (loc_m.group(2) or loc_m.group(1)).strip()

        # --- Time ---
        time_m = re.search(r'aanvang\s+(\d{1,2}[:.]\d{2})\s*uur', text, re.IGNORECASE)
        if time_m:
            result['time'] = time_m.group(1).replace('.', ':') + 'u'

        # --- Thema ---
        thema_m = re.search(
            r'thema van de dienst is\s*[:\-]?\s*["\u201c]?(.+?)(?=["\u201d]?\s*genomen uit)',
            text, re.IGNORECASE | re.DOTALL
        )
        if thema_m:
            result['thema'] = re.sub(r'\s+', ' ', thema_m.group(1)).strip().strip('\u201c\u201d"\'')
        else:
            # fallback: title of the article often contains the thema
            h_tag = soup.find(['h1', 'h2'])
            if h_tag:
                title_text = h_tag.get_text(strip=True)
                # strip predikant part after ' - '
                parts = title_text.split(' - ')
                if len(parts) > 1:
                    result['thema'] = parts[0].strip()

        # --- Bible verse ---
        bible_m = re.search(
            r'genomen uit\s+(.+?)(?=\s+De dienst|\s+De liturgie|\s+In deze|\s{3,}|$)',
            text, re.IGNORECASE | re.DOTALL
        )
        if bible_m:
            result['bible_verse'] = re.sub(r'\s+', ' ', bible_m.group(1)).strip()

        # --- YouTube link --- must have a non-empty path segment after the domain/path
        yt_m = re.search(r'https?://(?:www\.)?(?:youtube\.com/(?:live|watch)/[A-Za-z0-9_\-]{5,}|youtu\.be/[A-Za-z0-9_\-]{5,})[^\s<>"\']*', text)
        if yt_m:
            result['youtube_link'] = yt_m.group(0).rstrip('.,)')

        # --- Liturgie URL (direct link on gkin.org) ---
        for a in article.find_all('a', href=True):
            href = a['href']
            if 'liturgie' in href.lower() or href.lower().endswith('.pdf'):
                result['liturgie_url'] = href if href.startswith('http') else f"https://gkin.org{href}"
                break

        # --- Collecte URL (ING / OLE payment link) ---
        for a in article.find_all('a', href=True):
            href = a['href']
            if 'ing.nl/payreq' in href or 'tikkie.me' in href or 'ing.nl' in href:
                result['collecte_url'] = href
                break

        # --- QR image (download and encode as base64 data URI) ---
        for img in article.find_all('img', src=True):
            src = img['src']
            src_lower = src.lower()
            if 'qr' in src_lower or 'collecte' in src_lower or 'betaal' in src_lower:
                full_src = src if src.startswith('http') else f'https://gkin.org{src}'
                try:
                    r = self.session.get(full_src, timeout=10)
                    r.raise_for_status()
                    import base64 as _b64
                    ext = full_src.rsplit('.', 1)[-1].lower().split('?')[0]
                    mime = {'jpg': 'image/jpeg', 'jpeg': 'image/jpeg', 'gif': 'image/gif'}.get(ext, 'image/png')
                    b64 = _b64.b64encode(r.content).decode('utf-8')
                    result['qr_image_b64'] = f'data:{mime};base64,{b64}'
                    result['qr_image_url'] = full_src
                    print(f'[GKINScraper] QR image fetched: {full_src} ({len(r.content)} bytes)')
                except Exception as e:
                    print(f'[GKINScraper] QR image fetch error: {e}')
                break

        return result

    def fetch_for_date(self, target_date: datetime, look_ahead_days: int = 7) -> Dict[str, Any]:
        """
        Find the OLE article closest to target_date (within look_ahead_days).
        Returns parsed article dict or empty result with 'not_found' key.
        """
        print(f"[GKINScraper] Looking for OLE article for {target_date.strftime('%d-%m-%Y')}")
        links = self._get_article_links(pages=2)
        print(f"[GKINScraper] Found {len(links)} article links on index")

        # Score each link: prefer those whose title contains the target date month/day
        target_day = target_date.day
        target_month_nl = [k for k, v in NL_MONTHS.items() if v == target_date.month]
        target_month_str = target_month_nl[0] if target_month_nl else ''

        # First pass: find by scraping article date
        best: Optional[Dict[str, Any]] = None
        best_delta = timedelta(days=999)

        for link in links[:10]:  # check first 10 (most recent)
            article = self._parse_article(link['url'])
            art_date = article.get('date')
            if art_date:
                delta = abs(art_date - target_date)
                if delta <= timedelta(days=look_ahead_days) and delta < best_delta:
                    best = article
                    best_delta = delta
                    print(f"[GKINScraper] Candidate: {link['title']!r} date={art_date.strftime('%d-%m-%Y')} delta={delta.days}d")
                    if delta.days == 0:
                        break  # exact match

        if best:
            print(f"[GKINScraper] Best match: {best['url']} (delta {best_delta.days}d)")
            return best

        print(f"[GKINScraper] No article found within {look_ahead_days} days of {target_date.strftime('%d-%m-%Y')}")
        return {'not_found': True}
