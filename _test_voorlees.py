import sys
sys.path.insert(0, '/Users/vega/CascadeProjects/windsurf-project')

from voorlees_generator import VoorleesGenerator
from bulletin_generator import TEMPLATE_PATH
from docx import Document
from datetime import datetime

dt = datetime(2026, 5, 10)

meded = {
    'regionale_nl': 'Overlijdensbericht\nOp zondag 3 mei 2026 is op 81-jarige leeftijd in Tangerang, Indonesia overleden, mevr. Meike Charlotte Rosalin Soukotta-Frederik.\nZij was de moeder van zr. Elizabeth Soukotta en schoonmoeder van br. Bandung Nasserie van GKIN Amstelveen.\nDe gemeente en de kerkenraad bidden de familie in rouw Gods troost en nabijheid toe.\n\nOverlijdensbericht\nOp woensdag 29 april 2026 is op 60-jarige leeftijd in Utrecht overleden, zr. Siang Lan Goei Grace van GKIN Amstelveen.\nOp woensdag 29 april 2026 is op 77-jarige leeftijd in Balige, Indonesie overleden, mevr. Loide br Manurung, moeder van zr. Lita Napitupulu van GKIN Amstelveen.\nDe gemeente en de kerkenraad bidden de familie in rouw Gods troost en nabijheid toe.\n\nNieuwe leden in de seniorencommissie\nOp zondag 17 mei 2026 zullen br. Bandung Nasserie als bestuurslid van de seniorencommissie worden geinstalleerd tijdens de eredienst.\n\nPinksterviering 2026\nWij nodigen u met grote vreugde uit om de Pinksterdienst 2026 bij te wonen. De dienst wordt geleid door ds. Marla Winckler-Huliselan op zondag 24 mei 2026, beginnend om 10.30 uur in Bouwerij 52, Amstelveen.',
    'regionale_id': 'Berita Duka\nPada Minggu 3 Mei 2026, telah meninggal dunia dalam usia 81 tahun di Tangerang, Indonesia, Ibu Meike Charlotte Rosalin Soukotta-Frederik.\nBeliau adalah ibu dari sdri. Elizabeth Soukotta dan ibu mertua dari sdr. Bandung Nasserie dari GKIN Amstelveen.\nMajelis dan Jemaat mengucapkan turut berduka cita.\n\nBerita Duka\nPada Rabu 29 April 2026, telah meninggal dunia dalam usia 60 tahun di Utrecht, sdri. Siang Lan Goei Grace dari GKIN Amstelveen.\nPada Rabu 29 April 2026, telah meninggal dunia dalam usia 77 tahun di Balige, Indonesia, Ibu Loide br Manurung, ibu dari sdri. Lita Napitupulu dari GKIN Amstelveen.\nMajelis dan Jemaat mengucapkan turut berduka cita.\n\nAnggota baru komisi lansia\nPada Minggu 17 Mei 2026, sdr. Bandung Nasserie akan dilantik sebagai pengurus komisi lansia dalam ibadah.\n\nPerayaan Pentakosta 2026\nKami mengundang Anda dengan sukacita untuk menghadiri Ibadah Pentakosta 2026. Ibadah dipimpin oleh Pdt. Marla Winckler-Huliselan pada Minggu 24 Mei 2026 pukul 10.30 di Bouwerij 52, Amstelveen.',
    'landelijke_nl': 'Geloofsbelijdenis\nOp zondag 17 mei 2026 zal br. Bert Martin Rene Aper en zr Diah Sadiah de openbare Geloofsbelijdenis afleggen in de eredienst in de Marcuskerk te Den Haag onder leiding van ds. Stanley Tjahjadi.\nGaarne uw voorbede voor deze bijzondere gelegenheid.\n\nGKIN regionale bijeenkomsten over basispastoraat\nDe komende maanden biedt GKIN regionale bijeenkomsten over basispastoraat aan. De bijeenkomst in regio Amstelveen vindt plaats op 21 juni, na de eredienst waarin ds. S.M. Winckler-Huliselan voorgaat.',
    'landelijke_id': 'Pengakuan Percaya/Sidi\nPada Minggu 17 Mei 2026 di Marcuskerk, Den Haag, sdr. Bert Martin Rene Aper dan sdri Diah Sadiah akan mengaku percaya dalam kebaktian yang dipimpin oleh Pdt. Stanley Tjahjadi.\nMohon dukungan dan doa dari jemaat.\n\nPertemuan regional GKIN tentang pelayanan pastoral\nDalam beberapa bulan ke depan, GKIN menawarkan pertemuan regional tentang pelayanan pastoral dasar. Pertemuan di wilayah Amstelveen akan diadakan pada 21 Juni, setelah ibadah yang dipimpin oleh Pdt. S.M. Winckler-Huliselan.',
}

doc = Document(TEMPLATE_PATH)
welkom = []
in_welkom = False
for p in doc.paragraphs:
    if p.style.name.startswith('Heading') and 'welkomstwoord' in p.text.lower():
        in_welkom = True
        continue
    if in_welkom:
        if p.style.name.startswith('Heading'):
            break
        txt = p.text.strip()
        if txt:
            if 'Vandaag,' in txt:
                txt = 'Vandaag, zondag 10 mei 2026, gaat voor Ds. C. de Jonge. De ouderling van dienst is zr. Joyce Tearalangi. Als u vragen heeft, kunt u de ouderling van dienst aanspreken.'
            welkom.append(txt)

gen = VoorleesGenerator()
path = gen.generate(
    dt,
    {'predikant': 'Ds. C. de Jonge', 'ovd': 'zr. Joyce Tearalangi',
     '1eo': 'br. Hamra Simatupang', 'beamer': 'br. Chin Kie'},
    meded,
    {
        'dankoffer_url': 'https://tikkie.me/pay/GKINAM/in48FLYsgNfpWGVFeKpnxJ',
        'dankoffer_qr': 'output/_uploads/dankoffer_jft2wkmx.png',
        'ole_url': '', 'ole_qr': '',
        'collecte_contant': '', 'collecte_bonnen': '', 'collecte_bank': '',
        'collecte_tikkie': '', 'collecte_ole': '',
        'bezoekers_volwassenen': '', 'bezoekers_kinderen': '',
        'opbrengst_entries': [{
            'date_label': 'GKIN Amstelveen 3 mei 2026',
            'collecte_contant': '181,35', 'collecte_bonnen': '175,00',
            'collecte_bank': '39,00', 'collecte_tikkie': '415,28',
            'collecte_ole': '271,00',
            'bezoekers_volwassenen': '116', 'bezoekers_kinderen': '15',
        }],
    },
    welkom_paragraphs=welkom
)
print('Done:', path)
