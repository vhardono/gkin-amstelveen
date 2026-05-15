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
    'camp.predikant':         'Predikant',
    'camp.thema':             'Thema',
    'camp.bijbel':            'Bijbeltekst',
    'camp.liturgie_url':      'Liturgie URL',
    'camp.collecte_url':      'Collecte URL',
    'camp.section_list':      'Verzendlijst',
    'camp.list_placeholder':  '— Kies een lijst —',
    'camp.schedule':          'Inplannen',
    'camp.schedule_time':     'Verzendtijd',
    'camp.preview':           'Voorbeeld bekijken',
    'camp.create':            'Campagne aanmaken',

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
    'pb.send':                'Verstuur bevestiging',

    /* --- Liturgie --- */
    'lit.subtitle':           'Liturgie Generator',
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
    'home.meded.title':       'Mededelingen',
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
    'camp.predikant':         'Pendeta',
    'camp.thema':             'Tema',
    'camp.bijbel':            'Ayat Alkitab',
    'camp.liturgie_url':      'URL Liturgi',
    'camp.collecte_url':      'URL Kolekte',
    'camp.section_list':      'Daftar kirim',
    'camp.list_placeholder':  '— Pilih daftar —',
    'camp.schedule':          'Jadwalkan',
    'camp.schedule_time':     'Waktu kirim',
    'camp.preview':           'Lihat pratinjau',
    'camp.create':            'Buat kampanye',

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
    'pb.send':                'Kirim konfirmasi',

    /* --- Liturgie --- */
    'lit.subtitle':           'Generator Liturgi',
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
    'camp.predikant':         'Minister',
    'camp.thema':             'Theme',
    'camp.bijbel':            'Bible verse',
    'camp.liturgie_url':      'Liturgy URL',
    'camp.collecte_url':      'Collection URL',
    'camp.section_list':      'Mailing list',
    'camp.list_placeholder':  '— Choose a list —',
    'camp.schedule':          'Schedule',
    'camp.schedule_time':     'Send time',
    'camp.preview':           'Preview',
    'camp.create':            'Create campaign',

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
    'pb.send':                'Send confirmation',

    /* --- Liturgie --- */
    'lit.subtitle':           'Liturgy Generator',
    'lit.generate':           'Generate',
  },
};

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
        } else if (el.tagName === 'OPTION') {
          el.textContent = dict[key];
        } else {
          el.textContent = dict[key];
        }
      }
    });
    // Update switcher button highlights
    document.querySelectorAll('.lang-btn').forEach(btn => {
      btn.classList.toggle('bg-white', btn.dataset.lang === lang);
      btn.classList.toggle('text-gkin-green', btn.dataset.lang === lang);
      btn.classList.toggle('font-semibold', btn.dataset.lang === lang);
      btn.classList.toggle('text-green-100', btn.dataset.lang !== lang);
    });
    localStorage.setItem(STORAGE_KEY, lang);
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
