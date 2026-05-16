const GKIN_I18N = {
  nl: {
    /* --- Common --- */
    'nav.subtitle':           'Kies een tool',
    'nav.back_dashboard':     'Terug naar dashboard',
    'nav.back_select':        'Terug naar template keuze',
    'common.open':            'Openen',
    'common.loading':         'Laden...',
    'common.save':            'Opslaan',
    'common.cancel':          'Annuleren',
    'common.date':            'Datum',
    'common.time':            'Tijd',
    'common.name':            'Naam',
    'common.location':        'Locatie',

    /* --- Login --- */
    'login.subtitle':         'Beveiligde toegang',
    'login.title':            'Inloggen vereist',
    'login.desc':             'Voer het wachtwoord in om verder te gaan.',
    'login.password':         'Wachtwoord',
    'login.btn':              'Inloggen',

    /* --- Home --- */
    'home.preek.title':       'Preekbevestiging',
    'home.preek.desc':        'Stuur preekbevestiging e-mail naar de predikant',
    'home.liturgie.title':    'Liturgie',
    'home.liturgie.desc':     'Genereer liturgie documenten en beamer PPT',
    'home.meded.title':       'Mededelingen',
    'home.meded.desc':        'Genereer het wekelijkse mededelingen bulletin',
    'home.mailing.title':     'Mailing List',
    'home.mailing.desc':      'Maak een e-mailcampagne via Sender.net',

    /* --- Campaign select --- */
    'cs.subtitle':            'Mailing List',
    'cs.heading':             'Kies een e-mail template',
    'cs.subheading':          'Selecteer het type campagne dat u wilt aanmaken.',
    'cs.ole.title':           'OLE Informatie',
    'cs.ole.desc':            'Te verzenden vóór de dienst. Bevat predikant, thema, bijbeltekst, liturgie en YouTube link.',
    'cs.pm.title':            'Preek & Mededelingen Informatie',
    'cs.pm.desc':             'Te verzenden na de dienst. Bevat preek opname, mededelingen en collecte informatie.',
    'cs.pm.soon':             'Binnenkort',

    /* --- Date sections --- */
    'date.section_heading':   'Eredienst datum',
    'date.label':             'Datum',
    'date.placeholder':       '— Kies een datum —',

    /* --- Campaign (OLE) --- */
    'camp.subtitle':          'Sender Campaign Generator',
    'camp.info':              'Genereer een Sender e-mailcampagne gebaseerd op mededelingen en liturgie gegevens.',
    'camp.date_section':      'Datum Selectie',
    'camp.date_placeholder':  '— Kies een datum —',
    'camp.date_label':        'Selecteer een datum',
    'camp.fetch_ole':         'Ophalen OLE gegevens',
    'camp.section_details':   'Campagne details',
    'camp.subject':           'E-mail onderwerp',
    'camp.preview_text':      'E-mail preview tekst',
    'camp.predikant':         'Predikant',
    'camp.thema':             'Thema',
    'camp.bijbel':            'Bijbeltekst',
    'camp.liturgie_url':      'Liturgie URL',
    'camp.collecte_url':      'OLE Collecte URL',
    'camp.qr_code':           'OLE QR Code',
    'camp.auto_fetched_note': 'Automatisch opgehaald van gkin.org · handmatig overschrijven mag.',
    'camp.fetched_from_gkin':  'Opgehaald van gkin.org',
    'camp.no_data_found':     'Geen data gevonden op gkin.org',
    'camp.section_list':      'Verzendlijst',
    'camp.list_placeholder':  '— Kies een lijst —',
    'camp.schedule':          'Inplannen',
    'camp.schedule_time':     'Verzendtijd',
    'camp.preview':           'Voorbeeld bekijken',
    'camp.create':            'Campagne aanmaken',
    'camp.campaign_name':     'Campagnenaam',
    'camp.ole_settings':      'OLE Template Instellingen',
    'camp.fetching_gkin':     'Gegevens ophalen van gkin.org...',
    'camp.liturgie_fetched':  'Liturgie opgehaald van gkin.org',
    'camp.liturgie_not_found':'Nog niet gepubliceerd op gkin.org – vul URL handmatig in',
    'camp.qr_fetched':        'QR opgehaald van gkin.org',
    'camp.missing_fields':    'Ontbrekende velden – vul handmatig in:',

    /* --- PM Campaign --- */
    'pm.info':                'Te verzenden na de dienst. Bevat links naar de preek opname, mededelingen en YouTube webvideo.',
    'pm.section_content':     'Inhoud',
    'pm.predikant_am':        'Predikant AM',
    'pm.meded_url':           'Mededelingen URL (PDF)',
    'pm.meded_url_hint':      'Link naar het gegenereerde Mededelingen PDF (bv. via Dropbox of Google Drive).',
    'pm.preek_am_url':        'Preek AM URL (PDF)',
    'pm.preek_am_url_hint':   'Link naar het Preek PDF voor de AM dienst.',
    'pm.ole_location':        'OLE Locatie',
    'pm.ole_predikant':       'OLE Predikant',
    'pm.preek_ole_url':       'Preek {loc} URL (PDF)',
    'pm.youtube':             'YouTube Link (OLE webvideo)',
    'pm.recipients':          'Ontvangers groep(en)',
    'pm.no_selection':        'Geen selectie = eerste groep wordt automatisch gebruikt',
    'pm.fetching':            'Rooster ophalen...',
    'pm.ole_not_found':       'OLE niet gevonden in rooster – vul handmatig in',
    'pm.ole_missing_manual':  'Niet gevonden op gkin.org – vul handmatig in',
    'pm.preek_ole_url_label': 'Preek OLE URL (PDF)',
    'pm.voorbeeld_heading':   'Voorbeeld',
    'pm.created':             'Campaign aangemaakt',
    'pm.scheduled':           'Ingepland voor',
    'pm.schedule_error':      'Inplannen mislukt',
    'pm.error':               'Fout',

    /* --- Mededelingen --- */
    'med.subtitle':           'Mededelingen Generator',
    'med.info':               'Genereer het wekelijkse mededelingen bulletin. Vul de velden in en klik op Genereer.',
    'med.date_label':         'Selecteer datum',
    'med.sec_overdenking':    'Overdenking',
    'med.predikant':          'Predikant',
    'med.thema':              'Thema',
    'med.schriftlezing':      'Schriftlezing',
    'med.content':            'Inhoud overdenking',
    'med.sec_collecte':       'Collecte opbrengsten',
    'med.collecte_week':      '(vorige week)',
    'med.load_email':         'Laden uit e-mail',
    'med.contant':            '1. Contant',
    'med.bonnen':             '2. Collectebonnen',
    'med.bank':               '3. Bankoverschrijving',
    'med.tikkie':             '4. Tikkie',
    'med.ole_collecte':       'OLE Collecte',
    'med.bez_vol':            'Bezoekers volwassenen',
    'med.bez_kind':           'Bezoekers kinderen',
    'med.overige':            'Overige ontvangsten',
    'med.sec_collecte_info':  'Collecte informatie – URLs & QR codes',
    'med.dankoffer':          'Dankoffer (GKIN Amstelveen)',
    'med.ole_section':        'Online Landelijke Eredienst (OLE)',
    'med.betaal_url':         'Betaalverzoek URL',
    'med.qr_upload':          'QR code afbeelding',
    'med.upload_btn':         'Uploaden',
    'med.outlook_login':      'Inloggen bij Outlook',
    'med.sec_activiteiten':   'Activiteiten kalender',
    'med.load_act':           'Laden uit mededelingen',
    'med.add_act':            'Rij toevoegen',
    'med.act_datum':          'Datum',
    'med.act_tijd':           'Tijd',
    'med.act_activiteit':     'Activiteit',
    'med.act_olv':            'O.l.v.',
    'med.act_locatie':        'Loc.',
    'med.sec_meded':          'Mededelingen tekst',
    'med.load_meded':         'Laden uit Dropbox',
    'med.regionale':          'Regionale Mededelingen',
    'med.landelijke':         'Landelijke Mededelingen',
    'med.sec_liederen':       'Liederen',
    'med.load_liederen':      'Laden uit e-mail',
    'med.sec_bdays':          'Verjaardagen',
    'med.generate':           'Genereer bulletin',
    'med.generate_voorlees':  'Genereer voorleesblad',

    /* --- WA-OLE --- */
    'wa.section_heading':     'WhatsApp berichten',
    'wa.ole_title':           'WA-OLE informatie',
    'wa.ole_hint':            'Weblinks OLE · stuur uiterlijk zaterdag naar Kerkenraad GKIN AM',
    'wa.am_title':            'WA-Dienst informatie',
    'wa.am_hint':             'Tikkie Dankoffer · stuur uiterlijk zaterdag naar Informatie GKIN Amstelveen',
    'wa.am_modal_title':      'WA-Dienst informatie',
    'wa.am_reminder':         '⏰ Stuur dit bericht uiterlijk zaterdag naar de <strong>Informatie GKIN Amstelveen</strong> WhatsApp groep. Voeg Liturgie en Mededelingen toe als bijlage.',
    'wa.section_title':       'WA-OLE informatie',
    'wa.download_liturgie':   'Download Liturgie',
    'wa.make_message':        'Maak WA bericht',
    'wa.hint':                'WhatsApp bericht voor OLE weblinks. OLE locatie & tijd worden opgehaald uit het preekroster. Liturgie wordt opgehaald van gkin.org.',
    'wa.reminder':            '⏰ Stuur dit bericht uiterlijk zaterdag naar de <strong>Kerkenraad GKIN AM</strong> WhatsApp groep.',
    'wa.modal_title':         'WA-OLE informatie',
    'wa.loading':             'OLE preekroster laden...',
    'wa.copy':                'Kopiëren',
    'wa.copied':              '✓ Gekopieerd!',
    'wa.no_date':             'Selecteer eerst een datum.',
    'wa.no_liturgie':         'Geen liturgie URL gevonden voor deze datum op gkin.org.',
    'wa.liturgie_error':      'Fout bij ophalen liturgie: ',

    /* --- Day / month names --- */
    'days':    ['maandag','dinsdag','woensdag','donderdag','vrijdag','zaterdag','zondag'],
    'months':  ['januari','februari','maart','april','mei','juni','juli','augustus','september','oktober','november','december'],
    'day.past': '(afgelopen)',

    /* --- Common --- */
    'common.back':            'Terug',
    'common.reset':           'Opnieuw beginnen',
    'common.footer_data':     'Data bronnen: Dropbox Takenrooster',
    'common.optional':        '(optioneel)',

    /* --- Outlook --- */
    'med.outlook_label':      'Outlook e-mail',
    'med.outlook_connected':  'Verbonden',
    'med.outlook_not_conn':   'Niet verbonden',
    'med.outlook_login':      'Inloggen bij Outlook',
    'med.ovd_content_label':  'Inhoud overdenking',
    'med.ovd_content_hint':   "(alinea's scheiden met een lege regel)",
    'med.docs_section':       'Te genereren documenten',
    'med.doc_gedrukt':        'Gedrukte versie',
    'med.doc_voorlees':       'Voorlees versie',
    'med.footer':             'Data bronnen: Dropbox Takenrooster & Mededelingen · Scipio Verjaardagen · GKIN Preekrooster',

    /* --- Preekbevestiging --- */
    'pb.subtitle':            'Preekbevestiging',
    'pb.info':                'Stuur een preekbevestiging e-mail naar de predikant.',
    'pb.date_label':          'Selecteer datum',
    'pb.predikant':           'Predikant',
    'pb.email':               'E-mailadres predikant',
    'pb.location':            'Locatie dienst',
    'pb.time':                'Aanvangstijd',
    'pb.thema':               'Thema',
    'pb.bijbel':              'Bijbeltekst',
    'pb.sec_ontvangers':      'Ontvangers',
    'pb.to':                  'Aan (To)',
    'pb.cc_hint':             'OvD, 1eO en Beamer worden automatisch toegevoegd op basis van datum.',
    'pb.sec_email':           'E-mail inhoud',
    'pb.subject_label':       'Onderwerp',
    'pb.body_label':          'Inhoud',
    'pb.download_label':      'Download Preekbevestiging.docx bijlage voor gastpredikant',
    'pb.reset':               'Opnieuw',
    'pb.send':                'Maak e-mail concept',
    'pb.footer':              'Data bronnen: Dropbox Takenrooster · GKIN Preekrooster',

    /* --- Liturgie --- */
    'lit.subtitle':           'Liturgie Generator',
    'lit.info':               'Upload het wekelijkse Excel-bestand en optioneel de Preek.docx. Selecteer welke bestanden u wilt genereren.',
    'lit.sec_main':           'Main Liturgy file.xlsx',
    'lit.auto_update':        'Automatisch bijwerken met laatste gegevens',
    'lit.auto_detail_1':      'Takenrooster (Voorganger, OvD, etc.)',
    'lit.auto_detail_2':      'Tikkie QR link',
    'lit.auto_detail_3':      'Dankoffer vers',
    'lit.preview_btn':        'Wijzigingen bekijken & bijwerken',
    'lit.modal_title':        'Wijzigingen bekijken',
    'lit.modal_date':         'Datum:',
    'lit.modal_not_found':    'Niet gevonden:',
    'lit.modal_col_field':    'Veld',
    'lit.modal_col_current':  'Huidig',
    'lit.modal_col_new':      'Nieuw',
    'lit.modal_warnings':     'Waarschuwingen:',
    'lit.modal_cancel':       'Annuleren',
    'lit.modal_apply':        'Toepassen',
    'lit.use_current':        'Gebruik huidige week bestand',
    'lit.use_current_hint':   'Main Liturgy file.xlsx uit werkomap (Dropbox)',
    'lit.upload_own':         'Upload eigen bestand',
    'lit.upload_own_hint':    'Selecteer een lokaal .xlsx bestand',
    'lit.sec_preek':          'Preek.docx',
    'lit.preek_use_current':  'Gebruik huidige week bestand',
    'lit.preek_use_hint':     'Preek.docx uit werkomap (Dropbox)',
    'lit.preek_upload':       'Upload eigen bestand',
    'lit.preek_upload_hint':  'Selecteer een lokaal .docx bestand',
    'lit.sec_output':         'Te genereren bestanden',
    'lit.litA_hint':          'Liturgie (gemeente)',
    'lit.litB_hint':          'Muziekteam',
    'lit.litP_hint':          'Beamer presentatie',
    'lit.generate':           'Genereer',
  },

  id: {
    /* --- Common --- */
    'nav.subtitle':           'Pilih alat',
    'nav.back_dashboard':     'Kembali ke dasbor',
    'nav.back_select':        'Kembali ke pilihan template',
    'common.open':            'Buka',
    'common.loading':         'Memuat...',
    'common.save':            'Simpan',
    'common.cancel':          'Batal',
    'common.date':            'Tanggal',
    'common.time':            'Waktu',
    'common.name':            'Nama',
    'common.location':        'Lokasi',

    /* --- Login --- */
    'login.subtitle':         'Akses terbatas',
    'login.title':            'Login diperlukan',
    'login.desc':             'Masukkan kata sandi untuk melanjutkan.',
    'login.password':         'Kata sandi',
    'login.btn':              'Masuk',

    /* --- Home --- */
    'home.preek.title':       'Konfirmasi Khotbah',
    'home.preek.desc':        'Kirim e-mail konfirmasi khotbah ke pendeta',
    'home.liturgie.title':    'Liturgi',
    'home.liturgie.desc':     'Buat dokumen liturgi dan presentasi beamer',
    'home.meded.title':       'Pengumuman',
    'home.meded.desc':        'Buat buletin pengumuman mingguan',
    'home.mailing.title':     'Mailing List',
    'home.mailing.desc':      'Buat kampanye e-mail via Sender.net',

    /* --- Campaign select --- */
    'cs.subtitle':            'Mailing List',
    'cs.heading':             'Pilih template e-mail',
    'cs.subheading':          'Pilih jenis kampanye yang ingin Anda buat.',
    'cs.ole.title':           'Informasi OLE',
    'cs.ole.desc':            'Dikirim sebelum ibadah. Berisi pendeta, tema, ayat, liturgi dan link YouTube.',
    'cs.pm.title':            'Informasi Khotbah & Mededelingen',
    'cs.pm.desc':             'Dikirim setelah ibadah. Berisi rekaman khotbah, mededelingen dan informasi kolekte.',
    'cs.pm.soon':             'Segera hadir',

    /* --- Date sections --- */
    'date.section_heading':   'Tanggal ibadah',
    'date.label':             'Tanggal',
    'date.placeholder':       '— Pilih tanggal —',

    /* --- Campaign (OLE) --- */
    'camp.subtitle':          'Generator Kampanye Sender',
    'camp.info':              'Buat kampanye e-mail Sender berdasarkan data mededelingen dan liturgi.',
    'camp.date_section':      'Pilih Tanggal',
    'camp.date_label':        'Pilih tanggal',
    'camp.date_placeholder':  '— Pilih tanggal —',
    'camp.fetch_ole':         'Ambil data OLE',
    'camp.section_details':   'Detail kampanye',
    'camp.subject':           'Subjek e-mail',
    'camp.preview_text':      'Teks preview e-mail',
    'camp.predikant':         'Pendeta',
    'camp.thema':             'Tema',
    'camp.bijbel':            'Ayat Alkitab',
    'camp.liturgie_url':      'URL Liturgi',
    'camp.collecte_url':      'OLE URL Kolekte',
    'camp.qr_code':           'OLE Kode QR',
    'camp.auto_fetched_note': 'Diambil otomatis dari gkin.org · boleh diubah manual.',
    'camp.fetched_from_gkin':  'Diambil dari gkin.org',
    'camp.no_data_found':     'Tidak ada data ditemukan di gkin.org',
    'camp.section_list':      'Daftar kirim',
    'camp.list_placeholder':  '— Pilih daftar —',
    'camp.schedule':          'Jadwalkan',
    'camp.schedule_time':     'Waktu kirim',
    'camp.preview':           'Lihat pratinjau',
    'camp.create':            'Buat kampanye',
    'camp.campaign_name':     'Nama kampanye',
    'camp.ole_settings':      'Pengaturan Template OLE',
    'camp.fetching_gkin':     'Mengambil data dari gkin.org...',
    'camp.liturgie_fetched':  'Liturgi diambil dari gkin.org',
    'camp.liturgie_not_found':'Belum dipublikasikan di gkin.org – isi URL secara manual',
    'camp.qr_fetched':        'QR diambil dari gkin.org',
    'camp.missing_fields':    'Kolom belum diisi – isi secara manual:',

    /* --- PM Campaign --- */
    'pm.info':                'Dikirim setelah ibadah. Berisi link rekaman khotbah, mededelingen dan video YouTube.',
    'pm.section_content':     'Konten',
    'pm.predikant_am':        'Pendeta AM',
    'pm.meded_url':           'URL Mededelingen (PDF)',
    'pm.meded_url_hint':      'Link ke PDF Mededelingen yang dibuat (mis. via Dropbox atau Google Drive).',
    'pm.preek_am_url':        'URL Khotbah AM (PDF)',
    'pm.preek_am_url_hint':   'Link ke PDF khotbah ibadah AM.',
    'pm.ole_location':        'Lokasi OLE',
    'pm.ole_predikant':       'Pendeta OLE',
    'pm.preek_ole_url':       'URL Khotbah {loc} (PDF)',
    'pm.youtube':             'Link YouTube (video OLE)',
    'pm.recipients':          'Grup penerima',
    'pm.no_selection':        'Tidak ada pilihan = grup pertama digunakan otomatis',
    'pm.fetching':            'Mengambil jadwal...',
    'pm.ole_not_found':       'OLE tidak ditemukan di jadwal – isi manual',
    'pm.ole_missing_manual':  'Tidak ditemukan di gkin.org – isi secara manual',
    'pm.preek_ole_url_label': 'URL Khotbah OLE (PDF)',
    'pm.voorbeeld_heading':   'Pratinjau',
    'pm.created':             'Kampanye dibuat',
    'pm.scheduled':           'Dijadwalkan pada',
    'pm.schedule_error':      'Penjadwalan gagal',
    'pm.error':               'Kesalahan',

    /* --- Mededelingen --- */
    'med.subtitle':           'Generator Pengumuman',
    'med.info':               'Buat buletin pengumuman mingguan. Isi kolom dan klik Buat.',
    'med.date_label':         'Pilih tanggal',
    'med.sec_overdenking':    'Renungan',
    'med.predikant':          'Pendeta',
    'med.thema':              'Tema',
    'med.schriftlezing':      'Bacaan Alkitab',
    'med.content':            'Isi renungan',
    'med.sec_collecte':       'Hasil kolekte',
    'med.collecte_week':      '(minggu lalu)',
    'med.load_email':         'Muat dari e-mail',
    'med.contant':            '1. Tunai',
    'med.bonnen':             '2. Bon kolekte',
    'med.bank':               '3. Transfer bank',
    'med.tikkie':             '4. Tikkie',
    'med.ole_collecte':       'Kolekte OLE',
    'med.bez_vol':            'Pengunjung dewasa',
    'med.bez_kind':           'Pengunjung anak',
    'med.overige':            'Penerimaan lainnya',
    'med.sec_collecte_info':  'Informasi kolekte – URL & QR',
    'med.dankoffer':          'Persembahan syukur (GKIN Amstelveen)',
    'med.ole_section':        'Ibadah Landelijke Online (OLE)',
    'med.betaal_url':         'URL pembayaran',
    'med.qr_upload':          'Gambar QR code',
    'med.upload_btn':         'Unggah',
    'med.outlook_login':      'Masuk ke Outlook',
    'med.sec_activiteiten':   'Kalender kegiatan',
    'med.load_act':           'Muat dari mededelingen',
    'med.add_act':            'Tambah baris',
    'med.act_datum':          'Tanggal',
    'med.act_tijd':           'Waktu',
    'med.act_activiteit':     'Kegiatan',
    'med.act_olv':            'O.l.v.',
    'med.act_locatie':        'Lok.',
    'med.sec_meded':          'Teks mededelingen',
    'med.load_meded':         'Muat dari Dropbox',
    'med.regionale':          'Mededelingen Regional',
    'med.landelijke':         'Mededelingen Nasional',
    'med.sec_liederen':       'Nyanyian',
    'med.load_liederen':      'Muat dari e-mail',
    'med.sec_bdays':          'Ulang tahun',
    'med.generate':           'Buat buletin',
    'med.generate_voorlees':  'Buat lembar baca',

    /* --- WA-OLE --- */
    'wa.section_heading':     'Pesan WhatsApp',
    'wa.ole_title':           'Info WA-OLE',
    'wa.ole_hint':            'Tautan OLE · kirim paling lambat Sabtu ke Kerkenraad GKIN AM',
    'wa.am_title':            'Info WA-Dienst',
    'wa.am_hint':             'Tikkie Dankoffer · kirim paling lambat Sabtu ke Informasi GKIN Amstelveen',
    'wa.am_modal_title':      'Info WA-Dienst',
    'wa.am_reminder':         '⏰ Kirim pesan ini paling lambat hari Sabtu ke grup WhatsApp <strong>Informasi GKIN Amstelveen</strong>. Lampirkan Liturgi dan Mededelingen.',
    'wa.section_title':       'Info WA-OLE',
    'wa.download_liturgie':   'Unduh Liturgi',
    'wa.make_message':        'Buat pesan WA',
    'wa.hint':                'Pesan WhatsApp untuk tautan OLE. Lokasi & waktu OLE diambil dari jadwal. Liturgi diambil dari gkin.org.',
    'wa.reminder':            '⏰ Kirim pesan ini paling lambat hari Sabtu ke grup WhatsApp <strong>Kerkenraad GKIN AM</strong>.',
    'wa.modal_title':         'Info WA-OLE',
    'wa.loading':             'Memuat jadwal OLE...',
    'wa.copy':                'Salin',
    'wa.copied':              '✓ Tersalin!',
    'wa.no_date':             'Pilih tanggal terlebih dahulu.',
    'wa.no_liturgie':         'URL liturgi tidak ditemukan untuk tanggal ini di gkin.org.',
    'wa.liturgie_error':      'Gagal mengambil liturgi: ',

    /* --- Day / month names --- */
    'days':    ['Senin','Selasa','Rabu','Kamis','Jumat','Sabtu','Minggu'],
    'months':  ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'],
    'day.past': '(lalu)',

    /* --- Common --- */
    'common.back':            'Kembali',
    'common.reset':           'Mulai ulang',
    'common.footer_data':     'Sumber data: Dropbox Takenrooster',
    'common.optional':        '(opsional)',

    /* --- Outlook --- */
    'med.outlook_label':      'E-mail Outlook',
    'med.outlook_connected':  'Terhubung',
    'med.outlook_not_conn':   'Tidak terhubung',
    'med.outlook_login':      'Masuk Outlook',
    'med.ovd_content_label':  'Isi renungan',
    'med.ovd_content_hint':   '(pisahkan alinea dengan baris kosong)',
    'med.docs_section':       'Dokumen yang akan dibuat',
    'med.doc_gedrukt':        'Versi cetak',
    'med.doc_voorlees':       'Versi baca',
    'med.footer':             'Sumber data: Dropbox Takenrooster & Mededelingen · Scipio Verjaardagen · GKIN Preekrooster',

    /* --- Preekbevestiging --- */
    'pb.subtitle':            'Konfirmasi Khotbah',
    'pb.info':                'Kirim e-mail konfirmasi khotbah ke pendeta.',
    'pb.date_label':          'Pilih tanggal',
    'pb.predikant':           'Pendeta',
    'pb.email':               'Alamat e-mail pendeta',
    'pb.location':            'Lokasi ibadah',
    'pb.time':                'Waktu mulai',
    'pb.thema':               'Tema',
    'pb.bijbel':              'Ayat Alkitab',
    'pb.sec_ontvangers':      'Penerima',
    'pb.to':                  'Kepada (To)',
    'pb.cc_hint':             'OvD, 1eO dan Beamer ditambahkan otomatis berdasarkan tanggal.',
    'pb.sec_email':           'Isi e-mail',
    'pb.subject_label':       'Subjek',
    'pb.body_label':          'Isi',
    'pb.download_label':      'Unduh lampiran Preekbevestiging.docx untuk pendeta tamu',
    'pb.reset':               'Mulai ulang',
    'pb.send':                'Buat konsep e-mail',
    'pb.footer':              'Sumber data: Dropbox Takenrooster · GKIN Preekrooster',

    /* --- Liturgie --- */
    'lit.subtitle':           'Generator Liturgi',
    'lit.info':               'Unggah file Excel mingguan dan opsional Preek.docx. Pilih file yang ingin dibuat.',
    'lit.sec_main':           'Main Liturgy file.xlsx',
    'lit.auto_update':        'Perbarui otomatis dengan data terbaru',
    'lit.auto_detail_1':      'Takenrooster (Voorganger, OvD, dll.)',
    'lit.auto_detail_2':      'Link QR Tikkie',
    'lit.auto_detail_3':      'Ayat Dankoffer',
    'lit.preview_btn':        'Lihat & terapkan perubahan',
    'lit.modal_title':        'Lihat perubahan',
    'lit.modal_date':         'Tanggal:',
    'lit.modal_not_found':    'Tidak ditemukan:',
    'lit.modal_col_field':    'Kolom',
    'lit.modal_col_current':  'Saat ini',
    'lit.modal_col_new':      'Baru',
    'lit.modal_warnings':     'Peringatan:',
    'lit.modal_cancel':       'Batal',
    'lit.modal_apply':        'Terapkan',
    'lit.use_current':        'Gunakan file minggu ini',
    'lit.use_current_hint':   'Main Liturgy file.xlsx dari folder kerja (Dropbox)',
    'lit.upload_own':         'Unggah file sendiri',
    'lit.upload_own_hint':    'Pilih file .xlsx lokal',
    'lit.sec_preek':          'Preek.docx',
    'lit.preek_use_current':  'Gunakan file minggu ini',
    'lit.preek_use_hint':     'Preek.docx dari folder kerja (Dropbox)',
    'lit.preek_upload':       'Unggah file sendiri',
    'lit.preek_upload_hint':  'Pilih file .docx lokal',
    'lit.sec_output':         'File yang akan dibuat',
    'lit.litA_hint':          'Liturgi (jemaat)',
    'lit.litB_hint':          'Tim musik',
    'lit.litP_hint':          'Presentasi beamer',
    'lit.generate':           'Buat',
  },

  en: {
    /* --- Common --- */
    'nav.subtitle':           'Choose a tool',
    'nav.back_dashboard':     'Back to dashboard',
    'nav.back_select':        'Back to template selection',
    'common.open':            'Open',
    'common.loading':         'Loading...',
    'common.save':            'Save',
    'common.cancel':          'Cancel',
    'common.date':            'Date',
    'common.time':            'Time',
    'common.name':            'Name',
    'common.location':        'Location',

    /* --- Login --- */
    'login.subtitle':         'Secure access',
    'login.title':            'Login required',
    'login.desc':             'Enter the password to continue.',
    'login.password':         'Password',
    'login.btn':              'Log in',

    /* --- Home --- */
    'home.preek.title':       'Sermon Confirmation',
    'home.preek.desc':        'Send sermon confirmation email to the minister',
    'home.liturgie.title':    'Liturgy',
    'home.liturgie.desc':     'Generate liturgy documents and beamer PPT',
    'home.meded.title':       'Announcements',
    'home.meded.desc':        'Generate the weekly announcements bulletin',
    'home.mailing.title':     'Mailing List',
    'home.mailing.desc':      'Create an email campaign via Sender.net',

    /* --- Campaign select --- */
    'cs.subtitle':            'Mailing List',
    'cs.heading':             'Choose an email template',
    'cs.subheading':          'Select the type of campaign you want to create.',
    'cs.ole.title':           'OLE Information',
    'cs.ole.desc':            'Sent before the service. Contains minister, theme, bible verse, liturgy and YouTube link.',
    'cs.pm.title':            'Sermon & Announcements Information',
    'cs.pm.desc':             'Sent after the service. Contains sermon recording, announcements and collection information.',
    'cs.pm.soon':             'Coming soon',

    /* --- Date sections --- */
    'date.section_heading':   'Service date',
    'date.label':             'Date',
    'date.placeholder':       '— Choose a date —',

    /* --- Campaign (OLE) --- */
    'camp.subtitle':          'Sender Campaign Generator',
    'camp.info':              'Generate a Sender email campaign based on announcements and liturgy data.',
    'camp.date_section':      'Date Selection',
    'camp.date_label':        'Select a date',
    'camp.date_placeholder':  '— Choose a date —',
    'camp.fetch_ole':         'Fetch OLE data',
    'camp.section_details':   'Campaign details',
    'camp.subject':           'Email subject',
    'camp.preview_text':      'Email preview text',
    'camp.predikant':         'Minister',
    'camp.thema':             'Theme',
    'camp.bijbel':            'Bible verse',
    'camp.liturgie_url':      'Liturgy URL',
    'camp.collecte_url':      'OLE Collecte URL',
    'camp.qr_code':           'OLE QR Code',
    'camp.auto_fetched_note': 'Automatically fetched from gkin.org · manual override allowed.',
    'camp.fetched_from_gkin':  'Fetched from gkin.org',
    'camp.no_data_found':     'No data found on gkin.org',
    'camp.section_list':      'Mailing list',
    'camp.list_placeholder':  '— Choose a list —',
    'camp.schedule':          'Schedule',
    'camp.schedule_time':     'Send time',
    'camp.preview':           'Preview',
    'camp.create':            'Create campaign',
    'camp.campaign_name':     'Campaign name',
    'camp.ole_settings':      'OLE Template Settings',
    'camp.fetching_gkin':     'Fetching data from gkin.org...',
    'camp.liturgie_fetched':  'Liturgy fetched from gkin.org',
    'camp.liturgie_not_found':'Not yet published on gkin.org – fill in URL manually',
    'camp.qr_fetched':        'QR fetched from gkin.org',
    'camp.missing_fields':    'Missing fields – fill in manually:',

    /* --- PM Campaign --- */
    'pm.info':                'To be sent after the service. Contains links to the sermon recording, announcements and YouTube webvideo.',
    'pm.section_content':     'Content',
    'pm.predikant_am':        'Preacher AM',
    'pm.meded_url':           'Announcements URL (PDF)',
    'pm.meded_url_hint':      'Link to the generated Announcements PDF (e.g. via Dropbox or Google Drive).',
    'pm.preek_am_url':        'Sermon AM URL (PDF)',
    'pm.preek_am_url_hint':   'Link to the Sermon PDF for the AM service.',
    'pm.ole_location':        'OLE Location',
    'pm.ole_predikant':       'OLE Preacher',
    'pm.preek_ole_url':       'Sermon {loc} URL (PDF)',
    'pm.youtube':             'YouTube Link (OLE webvideo)',
    'pm.recipients':          'Recipient group(s)',
    'pm.no_selection':        'No selection = first group used automatically',
    'pm.fetching':            'Fetching roster...',
    'pm.ole_not_found':       'OLE not found in roster – fill in manually',
    'pm.ole_missing_manual':  'Not found on gkin.org – fill in manually',
    'pm.preek_ole_url_label': 'Sermon OLE URL (PDF)',
    'pm.voorbeeld_heading':   'Preview',
    'pm.created':             'Campaign created',
    'pm.scheduled':           'Scheduled for',
    'pm.schedule_error':      'Scheduling failed',
    'pm.error':               'Error',

    /* --- Mededelingen --- */
    'med.subtitle':           'Announcements Generator',
    'med.info':               'Generate the weekly announcements bulletin. Fill in the fields and click Generate.',
    'med.date_label':         'Select date',
    'med.sec_overdenking':    'Sermon notes',
    'med.predikant':          'Minister',
    'med.thema':              'Theme',
    'med.schriftlezing':      'Bible reading',
    'med.content':            'Sermon content',
    'med.sec_collecte':       'Collection proceeds',
    'med.collecte_week':      '(previous week)',
    'med.load_email':         'Load from email',
    'med.contant':            '1. Cash',
    'med.bonnen':             '2. Collection vouchers',
    'med.bank':               '3. Bank transfer',
    'med.tikkie':             '4. Tikkie',
    'med.ole_collecte':       'OLE Collection',
    'med.bez_vol':            'Adult visitors',
    'med.bez_kind':           'Child visitors',
    'med.overige':            'Other receipts',
    'med.sec_collecte_info':  'Collection info – URLs & QR codes',
    'med.dankoffer':          'Thanksgiving offering (GKIN Amstelveen)',
    'med.ole_section':        'Online National Service (OLE)',
    'med.betaal_url':         'Payment URL',
    'med.qr_upload':          'QR code image',
    'med.upload_btn':         'Upload',
    'med.outlook_login':      'Sign in to Outlook',
    'med.sec_activiteiten':   'Activities calendar',
    'med.load_act':           'Load from announcements',
    'med.add_act':            'Add row',
    'med.act_datum':          'Date',
    'med.act_tijd':           'Time',
    'med.act_activiteit':     'Activity',
    'med.act_olv':            'Led by',
    'med.act_locatie':        'Loc.',
    'med.sec_meded':          'Announcements text',
    'med.load_meded':         'Load from Dropbox',
    'med.regionale':          'Regional Announcements',
    'med.landelijke':         'National Announcements',
    'med.sec_liederen':       'Songs',
    'med.load_liederen':      'Load from email',
    'med.sec_bdays':          'Birthdays',
    'med.generate':           'Generate bulletin',
    'med.generate_voorlees':  'Generate reading sheet',

    /* --- WA-OLE --- */
    'wa.section_heading':     'WhatsApp messages',
    'wa.ole_title':           'WA-OLE information',
    'wa.ole_hint':            'OLE web links · send by Saturday to Kerkenraad GKIN AM',
    'wa.am_title':            'WA-Dienst information',
    'wa.am_hint':             'Tikkie Dankoffer · send by Saturday to Informatie GKIN Amstelveen',
    'wa.am_modal_title':      'WA-Dienst information',
    'wa.am_reminder':         '⏰ Send this message by Saturday at the latest to the <strong>Informatie GKIN Amstelveen</strong> WhatsApp group. Attach Liturgy and Mededelingen.',
    'wa.section_title':       'WA-OLE information',
    'wa.download_liturgie':   'Download Liturgy',
    'wa.make_message':        'Create WA message',
    'wa.hint':                'WhatsApp message for OLE web links. OLE location & time are fetched from the preekroster. Liturgy is fetched from gkin.org.',
    'wa.reminder':            '⏰ Send this message by Saturday at the latest to the <strong>Kerkenraad GKIN AM</strong> WhatsApp group.',
    'wa.modal_title':         'WA-OLE information',
    'wa.loading':             'Loading OLE roster...',
    'wa.copy':                'Copy',
    'wa.copied':              '✓ Copied!',
    'wa.no_date':             'Please select a date first.',
    'wa.no_liturgie':         'No liturgy URL found for this date on gkin.org.',
    'wa.liturgie_error':      'Error fetching liturgy: ',

    /* --- Day / month names --- */
    'days':    ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'],
    'months':  ['January','February','March','April','May','June','July','August','September','October','November','December'],
    'day.past': '(past)',

    /* --- Common --- */
    'common.back':            'Back',
    'common.reset':           'Start over',
    'common.footer_data':     'Data sources: Dropbox Takenrooster',
    'common.optional':        '(optional)',

    /* --- Outlook --- */
    'med.outlook_label':      'Outlook email',
    'med.outlook_connected':  'Connected',
    'med.outlook_not_conn':   'Not connected',
    'med.outlook_login':      'Log in to Outlook',
    'med.ovd_content_label':  'Sermon content',
    'med.ovd_content_hint':   '(separate paragraphs with a blank line)',
    'med.docs_section':       'Documents to generate',
    'med.doc_gedrukt':        'Printed version',
    'med.doc_voorlees':       'Reading version',
    'med.footer':             'Data sources: Dropbox Takenrooster & Mededelingen · Scipio Verjaardagen · GKIN Preekrooster',

    /* --- Preekbevestiging --- */
    'pb.subtitle':            'Sermon Confirmation',
    'pb.info':                'Send a sermon confirmation email to the minister.',
    'pb.date_label':          'Select date',
    'pb.predikant':           'Minister',
    'pb.email':               'Minister email address',
    'pb.location':            'Service location',
    'pb.time':                'Start time',
    'pb.thema':               'Theme',
    'pb.bijbel':              'Bible verse',
    'pb.sec_ontvangers':      'Recipients',
    'pb.to':                  'To',
    'pb.cc_hint':             'OvD, 1eO and Beamer are added automatically based on the date.',
    'pb.sec_email':           'Email content',
    'pb.subject_label':       'Subject',
    'pb.body_label':          'Body',
    'pb.download_label':      'Download Preekbevestiging.docx attachment for guest minister',
    'pb.reset':               'Reset',
    'pb.send':                'Create email draft',
    'pb.footer':              'Data sources: Dropbox Takenrooster · GKIN Preekrooster',

    /* --- Liturgie --- */
    'lit.subtitle':           'Liturgy Generator',
    'lit.info':               'Upload the weekly Excel file and optionally Preek.docx. Select which files to generate.',
    'lit.sec_main':           'Main Liturgy file.xlsx',
    'lit.auto_update':        'Auto-update with latest data',
    'lit.auto_detail_1':      'Takenrooster (Minister, OvD, etc.)',
    'lit.auto_detail_2':      'Tikkie QR link',
    'lit.auto_detail_3':      'Dankoffer verse',
    'lit.preview_btn':        'Preview & apply changes',
    'lit.modal_title':        'Preview changes',
    'lit.modal_date':         'Date:',
    'lit.modal_not_found':    'Not found:',
    'lit.modal_col_field':    'Field',
    'lit.modal_col_current':  'Current',
    'lit.modal_col_new':      'New',
    'lit.modal_warnings':     'Warnings:',
    'lit.modal_cancel':       'Cancel',
    'lit.modal_apply':        'Apply',
    'lit.use_current':        'Use current week file',
    'lit.use_current_hint':   'Main Liturgy file.xlsx from working folder (Dropbox)',
    'lit.upload_own':         'Upload own file',
    'lit.upload_own_hint':    'Select a local .xlsx file',
    'lit.sec_preek':          'Preek.docx',
    'lit.preek_use_current':  'Use current week file',
    'lit.preek_use_hint':     'Preek.docx from working folder (Dropbox)',
    'lit.preek_upload':       'Upload own file',
    'lit.preek_upload_hint':  'Select a local .docx file',
    'lit.sec_output':         'Files to generate',
    'lit.litA_hint':          'Liturgy (congregation)',
    'lit.litB_hint':          'Music team',
    'lit.litP_hint':          'Beamer presentation',
    'lit.generate':           'Generate',
  },
};

window.GKIN_I18N = GKIN_I18N;

(function () {
  const STORAGE_KEY = 'gkin_lang';
  const DEFAULT_LANG = 'nl';

  function getLang() {
    return localStorage.getItem(STORAGE_KEY) || DEFAULT_LANG;
  }

  function applyLang(lang) {
    const dict = GKIN_I18N[lang] || GKIN_I18N[DEFAULT_LANG];
    document.querySelectorAll('[data-i18n]').forEach(el => {
      const key = el.getAttribute('data-i18n');
      if (dict[key] !== undefined) {
        if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') {
          el.placeholder = dict[key];
        } else {
          el.textContent = dict[key];
        }
      }
    });
    document.querySelectorAll('[data-i18n-html]').forEach(el => {
      const key = el.getAttribute('data-i18n-html');
      if (dict[key] !== undefined) el.innerHTML = dict[key];
    });

    // Rebuild date <option> labels using day/month arrays stored in the option's dataset
    const days   = dict['days']   || GKIN_I18N.nl['days'];
    const months = dict['months'] || GKIN_I18N.nl['months'];
    const pastLbl = dict['day.past'] || '';
    document.querySelectorAll('select[id="date"] option[data-day-idx], select[id="dateSelect"] option[data-day-idx]').forEach(opt => {
      const di = parseInt(opt.dataset.dayIdx, 10);
      const mi = parseInt(opt.dataset.monthIdx, 10);
      const dn = parseInt(opt.dataset.dayNum, 10);
      const yr = opt.dataset.year;
      const suffix = opt.dataset.suffix || '';
      // For campaign page the suffix may contain "(afgelopen)" — replace with translated version
      let translatedSuffix = suffix.replace('(afgelopen)', pastLbl).replace('(lalu)', pastLbl).replace('(past)', pastLbl);
      opt.textContent = `${days[di]} ${dn} ${months[mi]} ${yr}${translatedSuffix}`;
    });

    // Re-apply outlook login button text based on connected state
    const loginLbl = document.getElementById('emailLoginText');
    if (loginLbl) {
      const isConnected = loginLbl.dataset.connected === '1';
      loginLbl.textContent = isConnected
        ? '✓ ' + (dict['med.outlook_connected'] || 'Verbonden')
        : (dict['med.outlook_login'] || 'Inloggen bij Outlook');
    }
    // Re-apply outlook status pill
    const pill = document.getElementById('outlookStatus');
    if (pill && pill.dataset.connected === '1') {
      pill.textContent = dict['med.outlook_connected'] || 'Verbonden';
    }

    // Update switcher button highlights
    document.querySelectorAll('.lang-btn').forEach(btn => {
      btn.classList.toggle('bg-white', btn.dataset.lang === lang);
      btn.classList.toggle('text-gkin-green', btn.dataset.lang === lang);
      btn.classList.toggle('font-semibold', btn.dataset.lang === lang);
      btn.classList.toggle('text-green-100', btn.dataset.lang !== lang);
    });
    localStorage.setItem(STORAGE_KEY, lang);
    window.GKIN_APPLY_LANG = applyLang;

    // Global translation function for inline JS use
    window.t = function(key) {
      const lang = getLang();
      const dict = GKIN_I18N[lang] || GKIN_I18N[DEFAULT_LANG];
      return dict[key] || key;
    };
  }

  function initSwitcher() {
    document.querySelectorAll('.lang-btn').forEach(btn => {
      btn.addEventListener('click', () => applyLang(btn.dataset.lang));
    });
    applyLang(getLang());
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initSwitcher);
  } else {
    initSwitcher();
  }
})();
