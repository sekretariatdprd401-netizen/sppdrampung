/**
 * ============================================================================
 * SIAP SPPD & RAMPUNG — CODE.GS (SATU FILE UTAMA BACKEND)
 *
 * File ini adalah GABUNGAN otomatis dari modul-modul pada folder backend-gas/:
 *   Config.gs, Utils.gs, Database.gs, Auth.gs, User.gs, SPPD.gs, Rampung.gs, Program.gs, StandarHarga.gs, Pejabat.gs, Pengaturan.gs, DasarHukum.gs, Print.gs, Logo.gs, WebApp.gs, Api.gs.
 *
 * Google Apps Script memperlakukan semua file .gs sebagai satu namespace global,
 * sehingga versi gabungan ini berperilaku identik dengan versi multi-file.
 *
 * Gunakan SALAH SATU saja (jangan keduanya dalam satu project):
 *   Opsi A : satu file Code.gs ini, ATAU
 *   Opsi B : 16 file .gs terpisah dari folder backend-gas/.
 *
 * DATABASE DIMULAI DALAM KONDISI KOSONG (tanpa data dummy/contoh):
 *   - setupDatabase() hanya membuat Sheet + header + akun admin awal.
 *   - Semua data transaksi (SPPD, Rampung, Program, Standar Harga, dll.)
 *     masuk hanya melalui input pengguna di aplikasi.
 *   - Untuk membersihkan data percobaan: jalankan resetData({ confirm: "RESET" })
 *     dari editor Apps Script (header & struktur tetap dipertahankan).
 *
 * Setup: jalankan setupDatabase() lalu deploy sebagai Web App.
 * ============================================================================
 */
/* ==========================================================================
 * BEGIN FILE ASLI: Config.gs
 * ========================================================================== */
/**
 * ============================================================================
 * SIAP SPPD & RAMPUNG — KONFIGURASI TERPUSAT (BACKEND GOOGLE APPS SCRIPT)
 * File     : Config.gs
 * Runtime  : Google Apps Script V8 (tanpa Node.js / npm / server eksternal)
 * Database : Google Sheets (SpreadsheetApp)
 *
 * Isi file ini:
 *   1. Identitas aplikasi & konfigurasi Spreadsheet ID (satu titik, terpusat).
 *   2. Definisi SELURUH Sheet + Header (dibuat otomatis oleh setupDatabase()).
 *
 * CATATAN PENTING:
 *  - Nama Header di bawah diambil dari field yang BENAR-BENAR dipakai
 *    frontend (index.html): form SPPD, form Rampung, rincian biaya pegawai,
 *    komponen biaya, Program, Standar Harga, Kop, Pejabat, Dasar Hukum,
 *    User, Instansi, dan Cetak Config.
 *  - setupDatabase() hanya MENAMBAH Sheet/Header yang belum ada. Data lama,
 *    Sheet lama, dan Header lama TIDAK PERNAH dihapus atau diganti nama.
 *  - Kolom `extraJson` adalah cadangan otomatis: field apa pun yang belum
 *    punya kolom sendiri disimpan utuh di sana (anti kehilangan data bila
 *    struktur aplikasi berkembang di masa depan).
 * ============================================================================
 */

/** Identitas aplikasi & versi skema database. */
var APP_NAME = 'SIAP SPPD & RAMPUNG';
var APP_VERSION = '1.0.0';
var SCHEMA_VERSION = '2026.09.16';

/** Nama properti (PropertiesService) yang dipakai backend. */
var PROP_SPREADSHEET_ID = 'SPREADSHEET_ID';
var PROP_LAST_SETUP = 'LAST_SETUP_AT';
var PROP_APP_VERSION = 'APP_VERSION';
var PROP_INSTALLED_AT = 'INSTALLED_AT';
var PROP_LOGO_FILE_ID = 'INSTANSI_LOGO_FILE_ID';
var PROP_SESSION_PREFIX = 'SESS_';

/**
 * OPSI 1 (paling sederhana): tulis Spreadsheet ID di bawah ini bila ingin
 * mengunci database ke Spreadsheet tertentu. Contoh:
 *   var SPREADSHEET_ID_DEFAULT = '1AbCdEfGhIjKlMnOpQrStUvWxYz1234567890';
 * Biarkan kosong ('') untuk mode otomatis:
 *   - Script terikat (container-bound)  → memakai Spreadsheet induknya.
 *   - Script berdiri sendiri (standalone) dan belum diatur → sistem MEMBUAT
 *     Spreadsheet baru bernama "DATABASE SIAP SPPD & RAMPUNG" + menyimpan ID
 *     di PropertiesService. Tidak ada pembuatan Sheet/Header manual.
 * OPSI 2: jalankan setSpreadsheetId('ID_SPREADSHEET') dari editor Apps Script.
 */
var SPREADSHEET_ID_DEFAULT = '';

/** Nama Spreadsheet yang dibuat otomatis bila database belum ada. */
var SPREADSHEET_AUTO_NAME = 'DATABASE SIAP SPPD & RAMPUNG';

/** Folder Drive untuk berkas besar (mis. logo instansi). */
var DRIVE_FOLDER_NAME = 'SIAP SPPD - BERKAS';

/** Batas panjang aman menyimpan teks dalam satu sel Google Sheets. */
var CELL_TEXT_LIMIT = 45000;

/** Masa berlaku token sesi login (detik) — 8 jam. */
var SESSION_TTL_SECONDS = 8 * 60 * 60;

/* ============================================================================
 * DEFINISI SHEET & HEADER
 * ----------------------------------------------------------------------------
 * kind:
 *   'table'     = banyak baris data (punya keyField = ID unik / primary key)
 *   'child'     = baris anak yang terhubung ke tabel induk (parentKeyField)
 *   'singleton' = hanya 1 baris pengaturan
 *   'keyvalue'  = tabel kunci/nilai
 * jsonFields  = kolom yang isinya JSON (array/objek)
 * extraJson   = kolom cadangan untuk field yang belum punya kolom sendiri
 *
 * ID unik SELALU dibuat backend dan TIDAK bergantung nomor baris, sehingga
 * urutan baris boleh berubah dan baris boleh dihapus.
 * ========================================================================== */
function sheetDefs_() {
  var defs = {};
  addDefsSppd_(defs);
  addDefsRampung_(defs);
  addDefsMaster_(defs);
  addDefsPengaturan_(defs);
  return defs;
}

/** Urutan pembuatan Sheet saat setupDatabase() (tabel induk lebih dulu). */
function sheetOrder_() {
  return [
    'SPPD', 'SPPD_PELAKSANA',
    'RAMPUNG', 'RAMPUNG_RINCIAN', 'RAMPUNG_KOMPONEN',
    'PROGRAM', 'PROGRAM_REKENING',
    'STANDAR_HARGA', 'KOP', 'PEJABAT', 'DASAR_HUKUM', 'USER',
    'INSTANSI', 'CETAK_CONFIG', 'SISTEM'
  ];
}

/**
 * Alias nama Sheet: bila pengguna/GAS sudah punya Sheet dengan penamaan lain,
 * backend tetap mengenalinya (data TIDAK dipindahkan, hanya dibaca).
 */
function sheetAliases_() {
  return {
    'STANDAR HARGA': 'STANDAR_HARGA',
    STANDARHARGA: 'STANDAR_HARGA',
    'DASAR HUKUM': 'DASAR_HUKUM',
    DASARHUKUM: 'DASAR_HUKUM',
    PEJABAT_PENANDATANGAN: 'PEJABAT'
  };
}
/* ------------------------- SHEET: SPPD & PELAKSANA ------------------------ */
function addDefsSppd_(d) {
  d.SPPD = {
    name: 'SPPD',
    label: 'Data Surat Perjalanan Dinas (SPT/SPPD)',
    kind: 'table',
    keyField: 'idSPPD',
    idPrefix: 'sppd',
    extraJson: 'extraJson',
    jsonFields: ['kegiatanJson'],
    headers: [
      'idSPPD', 'noSpt', 'noSppd', 'jenisSuratTugas', 'tglSpt', 'tglSppd',
      'maksud', 'tempatBerangkat', 'kotaTujuan', 'instansiTujuan',
      'transport', 'noTransport', 'ketTransport',
      'ttdSpt', 'ttdSpd', 'pptk',
      'hotel', 'alamatHotel', 'checkin', 'checkout', 'jumlahMalam', 'noKamar',
      'tglBerangkat', 'jamBerangkat', 'tglKembali', 'jamKembali', 'lamaHari',
      'biayaUangHarian', 'biayaRepresentasi', 'biayaTransport', 'biayaTaksi',
      'biayaPenginapan', 'biayaSewaKendaraan', 'biayaLainnya', 'biayaPpn', 'biayaPph',
      'total',
      'dokSpt', 'dokSppd', 'dokTiket', 'dokBoarding', 'dokHotel', 'dokTransport', 'dokPendukung',
      'program', 'kegiatan', 'subkegiatan', 'rekening', 'kegiatanJson',
      'status', 'createdAt', 'updatedAt', 'extraJson', 'createdBy'
    ]
  };
  d.SPPD_PELAKSANA = {
    name: 'SPPD_PELAKSANA',
    label: 'Pelaksana / pegawai pada setiap SPPD',
    kind: 'child',
    keyField: 'idPelaksana',
    idPrefix: 'plk',
    parentKeyField: 'idSPPD',
    parentSheet: 'SPPD',
    extraJson: 'extraJson',
    headers: [
      'idPelaksana', 'idSPPD', 'noSppd', 'nama', 'nip', 'jabatan', 'golongan', 'unit',
      'urutan', 'createdAt', 'updatedAt', 'extraJson'
    ]
  };
}

/* ------------------------ SHEET: RAMPUNG & RINCIAN ----------------------- */
function addDefsRampung_(d) {
  d.RAMPUNG = {
    name: 'RAMPUNG',
    label: 'Pertanggungjawaban (Rampung) Perjalanan Dinas',
    kind: 'table',
    keyField: 'idRampung',
    idPrefix: 'rmp',
    extraJson: 'extraJson',
    jsonFields: ['kegiatanJson', 'pejabatJson'],
    headers: [
      'idRampung', 'idSPPD', 'noSppd', 'noSpt', 'nama', 'nip', 'jabatan',
      'tujuan', 'maksud', 'tglBerangkat', 'tglKembali', 'lamaHari',
      'totalSPPD', 'totalRealisasi', 'selisih', 'totalRincian', 'jumlahRincian',
      'reaUangHarian', 'reaTransport', 'reaTaksi', 'reaPenginapan',
      'reaSewaKendaraan', 'reaRepresentasi', 'reaLainnya',
      'buktiTiket', 'buktiBoarding', 'buktiHotel', 'buktiTransport', 'buktiLainnya',
      'kegSubkegiatanId', 'kegKodeOrganisasi', 'kegKodeProgram', 'kegProgram',
      'kegKodeKegiatan', 'kegKegiatan', 'kegKodeSubkegiatan', 'kegSubkegiatan',
      'kegRekening', 'kegRekeningUraian', 'kegKeperluan', 'kegSumberDana',
      'kegTahunAnggaran', 'kegiatanJson',
      'pejPA', 'pejPPTK', 'pejBendahara', 'pejabatJson',
      'tahunAnggaran',
      'status', 'createdAt', 'updatedAt', 'extraJson', 'createdBy'
    ]
  };
  d.RAMPUNG_RINCIAN = {
    name: 'RAMPUNG_RINCIAN',
    label: 'Rincian biaya per pegawai pada setiap Rampung',
    kind: 'child',
    keyField: 'idRincian',
    idPrefix: 'rci',
    parentKeyField: 'idRampung',
    parentSheet: 'RAMPUNG',
    extraJson: 'extraJson',
    headers: [
      'idRincian', 'idRampung', 'noSppd', 'urutan', 'nama', 'nip', 'jabatan', 'golongan', 'unit',
      'noSpd', 'tglSpd', 'jumlahHari', 'uangHarian', 'transport', 'taksi', 'penginapan',
      'sewaKendaraan', 'representasi', 'lainnya', 'total',
      'createdAt', 'updatedAt', 'extraJson'
    ]
  };
  d.RAMPUNG_KOMPONEN = {
    name: 'RAMPUNG_KOMPONEN',
    label: 'Komponen biaya pada setiap rincian pegawai Rampung',
    kind: 'child',
    keyField: 'idKomponen',
    idPrefix: 'kmp',
    parentKeyField: 'idRincian',
    parentSheet: 'RAMPUNG_RINCIAN',
    extraJson: 'extraJson',
    headers: [
      'idKomponen', 'idRincian', 'idRampung', 'urutan', 'deskripsi', 'harga', 'qty',
      'satuan', 'persen', 'rill', 'subtotal', 'createdAt', 'updatedAt', 'extraJson'
    ]
  };
}

/* ------------------ SHEET: PROGRAM, STANDAR HARGA, MASTER ----------------- */
function addDefsMaster_(d) {
  d.PROGRAM = {
    name: 'PROGRAM',
    label: 'Master Program / Kegiatan / Subkegiatan',
    kind: 'table',
    keyField: 'idProgram',
    idPrefix: 'prg',
    extraJson: 'extraJson',
    headers: [
      'idProgram', 'kodeOrganisasi', 'kodeProgram', 'kode', 'nama',
      'kodeKegiatan', 'kegiatan', 'kodeSubkegiatan', 'subkegiatan',
      'rekening', 'uraianRekening', 'pagu', 'status',
      'createdAt', 'updatedAt', 'extraJson'
    ]
  };
  d.PROGRAM_REKENING = {
    name: 'PROGRAM_REKENING',
    label: 'Daftar Kode Rekening pada setiap Program',
    kind: 'child',
    keyField: 'idProgramRekening',
    idPrefix: 'prk',
    parentKeyField: 'idProgram',
    parentSheet: 'PROGRAM',
    extraJson: 'extraJson',
    headers: [
      'idProgramRekening', 'idProgram', 'urutan', 'kode', 'uraian',
      'createdAt', 'updatedAt', 'extraJson'
    ]
  };
  d.STANDAR_HARGA = {
    name: 'STANDAR_HARGA',
    label: 'Standar harga per daerah tujuan',
    kind: 'table',
    keyField: 'idStandarHarga',
    idPrefix: 'sh',
    extraJson: 'extraJson',
    headers: [
      'idStandarHarga', 'daerahTujuan', 'uangHarian', 'uangHarianDiklat', 'uangRepresentasi',
      'penginapan', 'uangTransport', 'taksiSetempat', 'taksiTanjungBill', 'biayaKontribusi',
      'status', 'createdAt', 'updatedAt', 'extraJson'
    ]
  };
  d.KOP = {
    name: 'KOP',
    label: 'Pengaturan Kop dokumen (SPD / Kwitansi)',
    kind: 'table',
    keyField: 'idKop',
    idPrefix: 'kop',
    extraJson: 'extraJson',
    headers: [
      'idKop', 'jenisSuratTugas', 'namaInstansi', 'namaUnitKerja', 'alamat', 'telepon',
      'laman', 'kota', 'default', 'notaDinas', 'logo',
      'createdAt', 'updatedAt', 'extraJson'
    ]
  };
  d.PEJABAT = {
    name: 'PEJABAT',
    label: 'Pejabat penandatangan (PA / PPTK / Bendahara / SPT / SPD)',
    kind: 'table',
    keyField: 'idPejabat',
    idPrefix: 'pj',
    extraJson: 'extraJson',
    headers: [
      'idPejabat', 'nama', 'jabatan', 'pangkat', 'nip', 'unitKerja', 'kotaKecamatan',
      'peran', 'status', 'createdAt', 'updatedAt', 'extraJson'
    ]
  };
  d.DASAR_HUKUM = {
    name: 'DASAR_HUKUM',
    label: 'Dasar hukum dokumen SPPD / Rampung',
    kind: 'table',
    keyField: 'idDasarHukum',
    idPrefix: 'dh',
    extraJson: 'extraJson',
    headers: [
      'idDasarHukum', 'no', 'dasarHukum', 'createdAt', 'updatedAt', 'extraJson'
    ]
  };
  d.USER = {
    name: 'USER',
    label: 'Pengguna aplikasi (Admin & Pengguna)',
    kind: 'table',
    keyField: 'idUser',
    idPrefix: 'usr',
    extraJson: 'extraJson',
    headers: [
      'idUser', 'nama', 'username', 'nip', 'role', 'status', 'password', 'access',
      'lastLogin', 'createdAt', 'updatedAt', 'extraJson'
    ]
  };
}
/* --------------- SHEET: PENGATURAN (INSTANSI, CETAK, SISTEM) -------------- */
function addDefsPengaturan_(d) {
  d.INSTANSI = {
    name: 'INSTANSI',
    label: 'Identitas instansi aplikasi (1 baris)',
    kind: 'singleton',
    keyField: 'kunci',
    extraJson: 'extraJson',
    headers: [
      'kunci', 'nama', 'organisasi', 'alamat', 'kabkota', 'provinsi',
      'telepon', 'laman', 'logo', 'namaKepala', 'nipKepala', 'jabatanKepala',
      'updatedAt', 'extraJson'
    ]
  };
  d.CETAK_CONFIG = {
    name: 'CETAK_CONFIG',
    label: 'Konfigurasi cetak dokumen (1 baris)',
    kind: 'singleton',
    keyField: 'kunci',
    extraJson: 'extraJson',
    headers: [
      'kunci', 'kertas', 'orientasi', 'margin', 'font', 'skala', 'header', 'footer',
      'logo', 'ttd', 'jarakTtd', 'posisiTtd', 'kop', 'updatedAt', 'extraJson'
    ]
  };
  d.SISTEM = {
    name: 'SISTEM',
    label: 'Metadata sistem (versi skema, dsb.)',
    kind: 'keyvalue',
    keyField: 'kunci',
    headers: ['kunci', 'nilai', 'keterangan', 'updatedAt']
  };
}

/** Kunci baris tunggal untuk Sheet INSTANSI & CETAK_CONFIG. */
var SINGLETON_KEY_INSTANSI = 'INSTANSI';
var SINGLETON_KEY_CETAK = 'CETAK_CONFIG';

/**
 * Nilai default 1 baris pengaturan.
 * PENTING: nilai ini HANYA dipakai saat baris pengaturan belum ada.
 * Bila baris sudah ada, isi lama dipertahankan seluruhnya (tidak ditimpa).
 */
function defaultInstansi_() {
  /* KOSONG — tidak ada data contoh. Nilai diisi pengguna melalui menu
     Pengaturan → Instansi/Kop. Baris singleton tetap dibuat agar struktur tersedia. */
  return {
    nama: '',
    organisasi: '',
    alamat: '',
    kabkota: '',
    provinsi: '',
    telepon: '',
    laman: '',
    logo: '',
    namaKepala: '',
    nipKepala: '',
    jabatanKepala: ''
  };
}

function defaultCetakConfig_() {
  return {
    kertas: 'A4', orientasi: 'potrait', margin: 25, font: 12, skala: 100,
    header: true, footer: true, logo: true, ttd: true, jarakTtd: 2,
    posisiTtd: 'kanan', kop: true
  };
}

/** Akun Admin awal — dibuat HANYA bila Sheet USER masih kosong. */
function defaultAdminUser_() {
  return {
    idUser: 'usr-admin',
    nama: 'Administrator',
    username: 'admin',
    nip: '',
    role: 'admin',
    status: 'Aktif',
    password: 'admin',
    access: ['all'],
    lastLogin: 'Belum pernah login'
  };
}

/* ==========================================================================
 * END FILE ASLI: Config.gs
 * ========================================================================== */

/* ==========================================================================
 * BEGIN FILE ASLI: Utils.gs
 * ========================================================================== */
/**
 * ============================================================================
 * SIAP SPPD & RAMPUNG — UTILITAS BERSAMA
 * File   : Utils.gs
 * Fungsi : format respons konsisten, ID unik, konversi nilai sel, validasi,
 *          dan fungsi hitung yang SAMA dengan frontend (index.html).
 * ============================================================================
 */

/* ------------------------------ RESPONS API ------------------------------- */

/** Respons sukses standar: { success:true, data, message }. */
function ok_(data, message) {
  return {
    success: true,
    data: (data === undefined ? null : data),
    message: message || 'Berhasil.'
  };
}

/** Respons gagal standar: { success:false, message, error }. */
function fail_(message, error) {
  return {
    success: false,
    data: null,
    message: message || 'Terjadi kesalahan.',
    error: (error === undefined || error === null) ? null : String(error && error.message ? error.message : error)
  };
}

/** Jalankan fungsi, selalu kembalikan respons standar (tidak pernah throw). */
function safeRun_(messageOk, fn) {
  try {
    var data = fn();
    return ok_(data, messageOk || 'Berhasil.');
  } catch (err) {
    logError_('safeRun_', err);
    return fail_((err && err.message) ? err.message : 'Terjadi kesalahan.', err);
  }
}

/** Melempar Error dengan pesan jelas (dipakai untuk validasi). */
function throwError_(msg) {
  throw new Error(msg);
}

/* ------------------------------- ID & WAKTU ------------------------------- */

/** ID unik buatan backend (tidak bergantung nomor baris Google Sheets). */
function newId_(prefix) {
  var p = prefix || 'id';
  return p + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, 12);
}

/** Waktu sekarang (ISO 8601) untuk kolom createdAt / updatedAt. */
function nowIso_() {
  return new Date().toISOString();
}

function scriptTimeZone_() {
  var tz = 'Asia/Jakarta';
  try { tz = Session.getScriptTimeZone() || tz; } catch (e) { }
  return tz;
}

/** Stempel waktu untuk kolom `lastLogin` — padanan nowStamp() frontend. */
function lastLoginStamp_() {
  return Utilities.formatDate(new Date(), scriptTimeZone_(), 'yyyy-MM-dd HH:mm');
}

/** Bila sel berisi Date, ubah menjadi teks tanggal/waktu yang konsisten. */
function dateCellToString_(d) {
  var tz = scriptTimeZone_();
  var jam = Utilities.formatDate(d, tz, 'HH:mm:ss');
  if (jam === '00:00:00') return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  return Utilities.formatDate(d, tz, 'yyyy-MM-dd HH:mm:ss');
}

/* ----------------------------- KONVERSI NILAI ----------------------------- */

/** Ubah nilai apa pun menjadi nilai aman untuk sel Google Sheets. */
function cellToString_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return dateCellToString_(v);
  if (typeof v === 'object') return jsonStringify_(v);
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  return String(v);
}

/** Nilai sel → string (selalu string, '' bila kosong). */
function toStr_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return dateCellToString_(v);
  return String(v);
}

/** Nilai sel → angka (0 bila bukan angka). Mendukung format Indonesia. */
function toNum_(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  var s = String(v).replace(/[^\d,.\-]/g, '');
  if (s.indexOf(',') > -1) s = s.replace(/\./g, '').replace(',', '.');
  var n = Number(s);
  return isNaN(n) ? 0 : n;
}

/** Nilai sel → bilangan bulat.
 *  BUG YANG DIPERBAIKI: sebelumnya `!n` membuat nilai 0 yang SAH selalu
 *  diganti menjadi `def` (mis. jumlahHari/jumlahMalam = 0 menjadi 1/def).
 *  Sekarang `def` hanya dipakai bila nilai bukan angka sama sekali. */
function toInt_(v, def) {
  if (v === null || v === undefined || v === '') return (def === undefined) ? 0 : def;
  if (typeof v === 'number') return isFinite(v) ? Math.round(v) : ((def === undefined) ? 0 : def);
  var s = String(v).replace(/[^\d,.\-]/g, '');
  if (s === '' || s === '-' || s === '.' || s === '-.') return (def === undefined) ? 0 : def;
  var n = toNum_(v);
  if (!isFinite(n)) return (def === undefined) ? 0 : def;
  return Math.round(n);
}

/** Nilai sel → boolean (menerima TRUE/FALSE/1/0/ya/tidak/on). */
function toBool_(v) {
  if (v === true || v === false) return v;
  if (v === null || v === undefined || v === '') return false;
  var s = String(v).trim().toLowerCase();
  return (s === 'true' || s === '1' || s === 'ya' || s === 'yes' || s === 'on');
}

/** Nilai terisi? (angka 0 dianggap terisi). */
function isFilled_(v) {
  return !(v === null || v === undefined || String(v).trim() === '');
}
/* --------------------------------- JSON ---------------------------------- */

function jsonStringify_(v) {
  try { return JSON.stringify(v); } catch (e) { return ''; }
}

function jsonParse_(txt, fallback) {
  if (txt === null || txt === undefined || txt === '') return fallback;
  if (typeof txt === 'object') return txt;
  try {
    var v = JSON.parse(String(txt));
    return (v === null || v === undefined) ? fallback : v;
  } catch (e) { return fallback; }
}

/** Salinan dalam (deep copy) sederhana. */
function clone_(v) {
  if (v === null || v === undefined) return v;
  return jsonParse_(jsonStringify_(v), v);
}

/* --------------------------------- LIST ---------------------------------- */

/** Array → teks dipisah '|' (kolom peran / access). */
function joinList_(arr) {
  if (!Array.isArray(arr)) return '';
  var out = [];
  for (var i = 0; i < arr.length; i++) {
    var s = toStr_(arr[i]).trim();
    if (s) out.push(s);
  }
  return out.join(' | ');
}

/** Teks '|' atau ',' → array (tahan terhadap format data lama). */
function splitList_(txt) {
  var arr = txt;
  if (arr === null || arr === undefined) return [];
  if (typeof arr === 'string') {
    var parsed = jsonParse_(arr, null);
    if (Array.isArray(parsed)) arr = parsed;
    else arr = String(arr).split(/[;,]+/);
  }
  if (!Array.isArray(arr)) arr = [arr];
  var out = [];
  for (var i = 0; i < arr.length; i++) {
    var v = toStr_(arr[i]).trim();
    if (v && out.indexOf(v) === -1) out.push(v);
  }
  return out;
}

/* ------------------------------- VALIDASI -------------------------------- */

/** Validasi field wajib. spec = [{field, label}]. */
function requireFields_(obj, spec, contextLabel) {
  var o = obj || {};
  var missing = [];
  for (var i = 0; i < spec.length; i++) {
    if (!isFilled_(o[spec[i].field])) missing.push(spec[i].label || spec[i].field);
  }
  if (missing.length) {
    throwError_((contextLabel ? contextLabel + ': ' : '') +
      'data wajib belum lengkap — ' + missing.join(', ') + '.');
  }
}

/** Validasi ID wajib (untuk update / delete). */
function requireId_(id, label) {
  if (!isFilled_(id)) throwError_((label || 'ID') + ' wajib diisi.');
  return String(id).trim();
}

/** Lempar error bila data tidak ditemukan. */
function assertFound_(obj, label) {
  if (!obj) throwError_((label || 'Data') + ' tidak ditemukan.');
  return obj;
}

/* --------------------------- HITUNG (SAMA FRONTEND) ---------------------- */

/** Total biaya SPPD — identik dengan calcBiaya() pada frontend. */
function calcBiaya_(b) {
  b = b || {};
  return toNum_(b.uangHarian) + toNum_(b.representasi) + toNum_(b.transport) +
    toNum_(b.taksi) + toNum_(b.penginapan) + toNum_(b.sewaKendaraan) +
    toNum_(b.lainnya) + toNum_(b.ppn) - toNum_(b.pph);
}

/** Total realisasi Rampung — identik dengan calcRealisasi() pada frontend. */
function calcRealisasi_(r) {
  r = r || {};
  return toNum_(r.uangHarian) + toNum_(r.transport) + toNum_(r.taksi) +
    toNum_(r.penginapan) + toNum_(r.sewaKendaraan) + toNum_(r.representasi) +
    toNum_(r.lainnya);
}

/** Lamanya hari perjalanan — identik dengan hitungLamaHari() pada frontend. */
function hitungLamaHari_(tglBerangkat, tglKembali) {
  if (!tglBerangkat || !tglKembali) return 1;
  var a = new Date(String(tglBerangkat) + 'T00:00:00');
  var b = new Date(String(tglKembali) + 'T00:00:00');
  var diff = Math.round((b.getTime() - a.getTime()) / 86400000);
  return diff >= 0 ? diff + 1 : 1;
}

/** Subtotal komponen — identik dengan rmpItemSubtotal() pada frontend. */
function rmpItemSubtotal_(it) {
  it = it || {};
  var p = (it.persen === undefined || it.persen === null || it.persen === '') ? 100 : toNum_(it.persen);
  return Math.round(toNum_(it.harga) * toNum_(it.qty) * (p / 100));
}

/* -------------------------------- HEADER -------------------------------- */

/** Peta header → indeks kolom (0-based) berdasarkan NAMA Header asli. */
function headerIndexMap_(headers) {
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    var h = toStr_(headers[i]).trim();
    if (h && map[h] === undefined) map[h] = i;
  }
  return map;
}

/** Urutkan objek dari yang paling baru (createdAt desc, tie-break baris sheet). */
function sortByCreatedDesc_(arr) {
  var list = arr || [];
  list.sort(function (a, b) {
    var x = toStr_(a.createdAt), y = toStr_(b.createdAt);
    if (x === y) return toNum_(b.__row) - toNum_(a.__row);
    return x < y ? 1 : -1;
  });
  return list;
}

/** Stempel createdAt (saat dibuat) & updatedAt (selalu) untuk penulisan. */
function stampForWrite_(obj, isCreate) {
  var o = obj || {};
  if (isCreate && !isFilled_(o.createdAt)) o.createdAt = nowIso_();
  o.updatedAt = nowIso_();
  return o;
}

/** Escape HTML — identik dengan escapeHtml() pada frontend. */
function escapeHtml_(str) {
  return String(str == null ? '' : str)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/** Terbilang — identik dengan terbilang() pada frontend. */
function terbilang_(n) {
  var satuan = ['', 'satu', 'dua', 'tiga', 'empat', 'lima', 'enam', 'tujuh',
    'delapan', 'sembilan', 'sepuluh', 'sebelas'];
  n = Math.abs(Math.round(Number(n) || 0));
  if (n < 12) return satuan[n] || '';
  if (n < 20) return satuan[n - 10] + ' belas';
  if (n < 100) return satuan[Math.floor(n / 10)] + ' puluh' + (n % 10 ? ' ' + satuan[n % 10] : '');
  if (n < 200) return 'seratus' + (n % 100 ? ' ' + terbilang_(n % 100) : '');
  if (n < 1000) return satuan[Math.floor(n / 100)] + ' ratus' + (n % 100 ? ' ' + terbilang_(n % 100) : '');
  return String(n);
}

/**
 * Normalisasi payload pembaruan agar mendukung dua gaya pemanggilan:
 *   updateXxx(token, data)          → { id: data.id, ...data }
 *   updateXxx(token, id, data)      → { id: id, ...data }
 */
function normalizeUpdatePayload_(data, dataAlt, idField, label) {
  var payload;
  if (typeof data === 'string' || typeof data === 'number') {
    payload = Object.assign({}, dataAlt || {});
    payload[idField] = String(data);
  } else {
    payload = Object.assign({}, data || {});
  }
  if (!isFilled_(payload[idField]) && isFilled_(payload.id)) payload[idField] = payload.id;
  payload[idField] = requireId_(payload[idField], label || 'ID');
  payload.id = payload[idField];
  return payload;
}

/* ------------------------------- DAFTAR --------------------------------- */

/** Normalisasi daftar hak akses user menu frontend. */
function normalizeUserAccess_(access) {
  var arr = splitList_(access);
  var allow = ['all', 'sppd', 'rampung', 'program', 'standar-harga', 'pengaturan', 'user'];
  var out = [];
  for (var i = 0; i < arr.length; i++) {
    var v = toStr_(arr[i]).trim();
    if (allow.indexOf(v) > -1 && out.indexOf(v) === -1) out.push(v);
  }
  return out;
}

/**
 * KUNCI TULIS RE-ENTRANT
 * README menjanjikan "Operasi tulis memakai LockService", tetapi praktiknya
 * hanya resetData() yang mengunci — semua penulisan CRUD (append/update/
 * delete) berjalan tanpa kunci sehingga dua pengguna yang menyimpan bersamaan
 * dapat saling menimpa baris (getLastRow lalu setValues bukan operasi atomik).
 * Flag eksekusi membuat kunci bersarang aman: satu operasi komposit
 * (mis. simpan Rampung + rincian + komponen) cukup memegang SATU kunci.
 */
var GAS_LOCK_HELD_ = false;

function withScriptLock_(fn) {
  var acquired = false;
  if (!GAS_LOCK_HELD_) {
    try {
      var lock = LockService.getScriptLock();
      lock.waitLock(30000);
      GAS_LOCK_HELD_ = true;
      acquired = true;
    } catch (e) {
      /* waitLock gagal (timeout / runtime tanpa LockService seperti editor):
         lanjut tanpa kunci — perilaku lama, tidak pernah menggagalkan request
         hanya karena layanan kunci tidak tersedia. */
    }
  }
  try { return fn(); }
  finally {
    if (acquired) {
      try { LockService.getScriptLock().releaseLock(); } catch (e2) { }
      GAS_LOCK_HELD_ = false;
    }
  }
}

/* --------------------------------- LOGGING ------------------------------- */

function logInfo_(tag, msg) {
  try { Logger.log('[' + tag + '] ' + msg); } catch (e) { }
}

function logError_(tag, err) {
  try { Logger.log('[' + tag + '] ERROR: ' + (err && err.stack ? err.stack : err)); } catch (e) { }
}


/* ------------------------- CACHE DATA REFERENSI -------------------------- */
/* Data referensi (pejabat, dasar hukum, kop, instansi, standar harga,
 * program) jarang berubah — disimpan sementara di CacheService (TTL pendek)
 * agar pembacaan berulang tidak menembak Google Sheets setiap saat.
 * Setiap operasi tulis memanggil refCacheBust_ sehingga data selalu segar.
 * Data transaksi (SPPD, Rampung, USER) TIDAK di-cache (harus real-time). */
var REF_CACHE_TTL_SECONDS_ = 120;

/** Ambil dari cache; bila kosong, jalankan producer() lalu simpan hasilnya. */
function refCacheGet_(key, producer) {
  var cache = null;
  try { cache = CacheService.getScriptCache(); } catch (e) { cache = null; }
  var cacheKey = 'ref_' + key;
  if (cache) {
    try {
      var raw = cache.get(cacheKey);
      if (raw) { var parsed = JSON.parse(raw); if (parsed && parsed.__refCache) return parsed.value; }
    } catch (e) { /* cache bermasalah -> hitung ulang secara normal */ }
  }
  var value = producer();
  if (cache) {
    try {
      var payload = JSON.stringify({ __refCache: true, value: value });
      /* Batas 1 entri CacheService ~100KB; bila lebih, lewati penyimpanan. */
      if (payload.length <= 98000) cache.put(cacheKey, payload, REF_CACHE_TTL_SECONDS_);
    } catch (e) { /* gagal simpan cache tidak boleh menggagalkan request */ }
  }
  return value;
}

/** Hapus entri cache suatu modul referensi (dipanggil setelah tulis/ubah/hapus). */
function refCacheBust_(key) {
  try { CacheService.getScriptCache().remove('ref_' + key); } catch (e) { }
}

/** Seluruh kunci cache data referensi yang dipakai aplikasi. */
var REF_CACHE_KEYS_ = ['pejabat', 'dasarhukum', 'kop', 'instansi', 'cetakconfig', 'standarharga', 'program'];

/**
 * Hapus SEMUA cache data referensi sekaligus.
 * WAJIB dipakai setiap kali database berubah secara massal (resetData,
 * setSpreadsheetId) — kalau hanya sebagian kunci yang dibersihkan, sisa kunci
 * akan menyajikan data LAMA hingga TTL (120 detik) habis.
 */
function refCacheBustAll_() {
  for (var i = 0; i < REF_CACHE_KEYS_.length; i++) refCacheBust_(REF_CACHE_KEYS_[i]);
}

/* ==========================================================================
 * END FILE ASLI: Utils.gs
 * ========================================================================== */

/* ==========================================================================
 * BEGIN FILE ASLI: Database.gs
 * ========================================================================== */
/**
 * ============================================================================
 * SIAP SPPD & RAMPUNG — ENGINE DATABASE GOOGLE SHEETS
 * File   : Database.gs
 * Isi    : konfigurasi Spreadsheet, setupDatabase() otomatis (Sheet + Header),
 *          pembacaan/penulisan baris, dan helper tabel anak.
 *
 * JAMINAN KEAMANAN DATA:
 *  - Sheet yang sudah ada TIDAK PERNAH dihapus atau diganti nama.
 *  - Header yang sudah ada TIDAK PERNAH dihapus/diubah; hanya kolom baru yang
 *    DITAMBAHKAN di ujung kanan.
 *  - Baris data lama TIDAK PERNAH dikosongkan oleh setupDatabase().
 *  - Tidak ada reset database otomatis.
 * ============================================================================
 */

/* --------------------------- SPREADSHEET / CONFIG ------------------------- */

/** Baca Spreadsheet ID terpusat (PropertiesService → konstanta → otomatis). */
function getSpreadsheetId_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(PROP_SPREADSHEET_ID);
  if (isFilled_(id)) return String(id).trim();
  if (isFilled_(SPREADSHEET_ID_DEFAULT)) {
    props.setProperty(PROP_SPREADSHEET_ID, String(SPREADSHEET_ID_DEFAULT).trim());
    return String(SPREADSHEET_ID_DEFAULT).trim();
  }
  /* Script terikat (container-bound) → pakai Spreadsheet induknya. */
  var active = null;
  try { active = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { active = null; }
  if (active) {
    props.setProperty(PROP_SPREADSHEET_ID, active.getId());
    return active.getId();
  }
  /* Belum ada database → buat Spreadsheet baru otomatis. */
  var ss = SpreadsheetApp.create(SPREADSHEET_AUTO_NAME);
  props.setProperty(PROP_SPREADSHEET_ID, ss.getId());
  logInfo_('getSpreadsheetId_', 'Spreadsheet baru dibuat: ' + ss.getId());
  return ss.getId();
}

/** Simpan Spreadsheet ID ke konfigurasi terpusat. */
function setSpreadsheetId(id) {
    memoClearAll_(); /* ID berganti → referensi lama tidak berlaku. */
    refCacheBustAll_(); /* cache referensi milik Spreadsheet lama harus dibuang. */
  var sid = requireId_(id, 'Spreadsheet ID');
  var ss;
  try { ss = SpreadsheetApp.openById(sid); } catch (e) {
    return fail_('Spreadsheet ID tidak dapat dibuka. Pastikan ID benar dan akun memiliki akses.', e);
  }
  PropertiesService.getScriptProperties().setProperty(PROP_SPREADSHEET_ID, sid);
  return ok_({ spreadsheetId: sid, name: ss.getName(), url: ss.getUrl() },
    'Spreadsheet ID berhasil disimpan. Jalankan setupDatabase() untuk menyiapkan Sheet & Header.');
}

/** Informasi konfigurasi database (tidak mengubah apa pun). */
function getDatabaseConfig() {
  return safeRun_('Konfigurasi database dibaca.', function () {
    var props = PropertiesService.getScriptProperties();
    var sid = props.getProperty(PROP_SPREADSHEET_ID) || '';
    var url = '';
    var name = '';
    if (sid) {
      try { var ss = SpreadsheetApp.openById(sid); name = ss.getName(); url = ss.getUrl(); }
      catch (e) { name = '(tidak dapat dibuka)'; }
    }
    return {
      appName: APP_NAME, appVersion: APP_VERSION, schemaVersion: SCHEMA_VERSION,
      spreadsheetId: sid, spreadsheetName: name, spreadsheetUrl: url,
      lastSetupAt: props.getProperty(PROP_LAST_SETUP) || '',
      installedAt: props.getProperty(PROP_INSTALLED_AT) || ''
    };
  });
}

/* ---------------------- MEMOISASI PER-EKSEKUSI --------------------------- */
/* Referensi Spreadsheet/Sheet/Header disimpan selama SATU eksekusi request,
 * sehingga openById/getSheetByName/pembacaan header tidak diulang untuk
 * tiap operasi dalam proses yang sama. Memori dibersihkan otomatis karena
 * setiap google.script.run memulai eksekusi baru. */
var MEMO_SS_ = null;
var MEMO_SS_ID_ = null;
var MEMO_SHEETS_ = {};
var MEMO_HEADERS_ = {};

/** Bersihkan seluruh memo (dipakai setup/reset/penggantian ID Spreadsheet). */
function memoClearAll_() {
  MEMO_SS_ = null; MEMO_SS_ID_ = null; MEMO_SHEETS_ = {}; MEMO_HEADERS_ = {};
}

/** Hapus memo satu Sheet (setelah sheet dibuat/dihapus/header diubah). */
function memoClearSheet_(name) {
  if (name && MEMO_SHEETS_[name]) delete MEMO_SHEETS_[name];
  if (name && MEMO_HEADERS_[name]) delete MEMO_HEADERS_[name];
}

/* --------------------- REVISI DATA (VALIDASI CACHE FRONTEND) --------------
 * Setiap penulisan data menaikkan "revisi" Sheet yang bersangkutan. Frontend
 * memakai revisi ini untuk MEMASTIKAN data yang di-cache masih sama dengan
 * server sebelum memutuskan untuk tidak memuat ulang. Hasilnya:
 *   - data tidak dimuat ulang bila memang tidak ada perubahan;
 *   - data yang berubah tetap langsung terambil (tidak ada cache basi).
 *
 * Disimpan di Script Properties (BUKAN di Sheet) supaya:
 *   - pembacaannya tidak perlu membuka Google Sheets sama sekali, dan
 *   - tidak ada penulisan Sheet tambahan pada setiap operasi CRUD.
 * Nilainya bilangan bulat yang selalu naik (monotonik).
 * ------------------------------------------------------------------------ */
var REV_PREFIX_ = 'REV_';

/** Naikkan revisi satu Sheet. Aman dipanggil dari jalur tulis mana pun. */
function bumpDataRev_(sheetName) {
  var nm = toStr_(sheetName).trim();
  if (!nm) return;
  try {
    var props = PropertiesService.getScriptProperties();
    var key = REV_PREFIX_ + nm;
    var cur = Number(props.getProperty(key)) || 0;
    var next = Date.now();
    if (next <= cur) next = cur + 1;
    props.setProperty(key, String(next));
  } catch (e) {
    /* Revisi hanya dipakai untuk optimasi cache — kegagalan di sini TIDAK
       boleh menggagalkan operasi tulis yang sedang berjalan. */
  }
}

/** Naikkan revisi beberapa Sheet sekaligus. */
function bumpDataRevs_(names) {
  var arr = names || [];
  for (var i = 0; i < arr.length; i++) bumpDataRev_(arr[i]);
}

/** Naikkan revisi SELURUH Sheet aplikasi (dipakai setelah pengosongan massal). */
function bumpAllDataRevs_() {
  var order = sheetOrder_();
  for (var i = 0; i < order.length; i++) {
    try { bumpDataRev_(getDef_(order[i]).name); } catch (e) { }
  }
}

/** Baca seluruh revisi yang tercatat (satu pembacaan, tanpa membuka Sheet). */
function readDataRevs_() {
  var out = {};
  try {
    var all = PropertiesService.getScriptProperties().getProperties();
    for (var k in all) {
      if (k.indexOf(REV_PREFIX_) !== 0) continue;
      var nm = k.slice(REV_PREFIX_.length);
      if (nm) out[nm] = all[k];
    }
  } catch (e) { }
  return out;
}

function getSpreadsheet_() {
  var sid = getSpreadsheetId_();
  if (MEMO_SS_ && MEMO_SS_ID_ === sid) return MEMO_SS_;
  try {
    MEMO_SS_ = SpreadsheetApp.openById(sid);
    MEMO_SS_ID_ = sid;
    MEMO_SHEETS_ = {}; MEMO_HEADERS_ = {};
    return MEMO_SS_;
  } catch (e) {
    throwError_('Spreadsheet database tidak dapat dibuka (ID: ' + sid + '). Periksa konfigurasi SPREADSHEET_ID.');
  }
}

function getDef_(key) {
  var defs = sheetDefs_();
  var def = defs[key] || defs[String(key).toUpperCase()];
  if (!def) throwError_('Definisi Sheet "' + key + '" tidak dikenal.');
  return def;
}
/* ------------------------------ SETUP DATABASE --------------------------- */

/**
 * SETUP DATABASE OTOMATIS (idempotent & aman dijalankan berkali-kali):
 *   cek Spreadsheet → cek seluruh Sheet → buat Sheet yang belum ada →
 *   cek Header → tambah Header yang belum ada → database siap dipakai.
 * Data lama tidak pernah disentuh.
 *
 * Cara pakai:
 *   - Dari editor Apps Script: pilih fungsi `setupDatabase` → Run.
 *   - Dari frontend/API       : apiCall({ token, method:'setupDatabase' }).
 */
function setupDatabase(opt) {
  return safeRun_('Database siap digunakan.', function () {
    var options = opt || {};
    if (isFilled_(options.spreadsheetId)) {
      PropertiesService.getScriptProperties()
        .setProperty(PROP_SPREADSHEET_ID, String(options.spreadsheetId).trim());
    }
    var ss = getSpreadsheet_();
    var defs = sheetDefs_();
    var order = sheetOrder_();

    var createdSheets = [];
    var existingSheets = [];
    var addedHeaders = [];

    for (var i = 0; i < order.length; i++) {
      var def = getDef_(order[i]);
      var res = ensureSheet_(ss, def);
      if (res.created) createdSheets.push(def.name); else existingSheets.push(def.name);
      if (res.addedHeaders.length) addedHeaders.push(def.name + ' → ' + res.addedHeaders.join(', '));
    }

    /* Baris pengaturan tunggal (HANYA dibuat bila belum ada). */
    var instansiRowCreated = ensureSingletonRow_(ss, defs.INSTANSI, SINGLETON_KEY_INSTANSI, defaultInstansi_());
    var cetakRowCreated = ensureSingletonRow_(ss, defs.CETAK_CONFIG, SINGLETON_KEY_CETAK, defaultCetakConfig_());

    /* Admin awal (HANYA bila Sheet USER masih kosong). */
    var adminCreated = ensureAdminUser_(ss, defs.USER);

    /* Metadata sistem. */
    setSettingIfEmpty_(ss, 'schemaVersion', SCHEMA_VERSION, 'Versi skema database');
    setSetting_(ss, 'appVersion', APP_VERSION, 'Versi aplikasi backend');
    setSetting_(ss, 'lastSetupAt', nowIso_(), 'Waktu setup terakhir');

    removeEmptyDefaultSheet_(ss, defs);

    /* Pastikan setiap Sheet punya revisi awal, sehingga validasi cache
       frontend langsung aktif sejak instalasi pertama (bukan menunggu
       penulisan pertama). */
    bumpAllDataRevs_();

    var props = PropertiesService.getScriptProperties();
    props.setProperty(PROP_LAST_SETUP, nowIso_());
    props.setProperty(PROP_APP_VERSION, APP_VERSION);
    if (!props.getProperty(PROP_INSTALLED_AT)) props.setProperty(PROP_INSTALLED_AT, nowIso_());

    return {
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName(),
      spreadsheetUrl: ss.getUrl(),
      schemaVersion: SCHEMA_VERSION,
      sheetCount: order.length,
      createdSheets: createdSheets,
      existingSheets: existingSheets,
      addedHeaders: addedHeaders,
      instansiRowCreated: instansiRowCreated,
      cetakConfigRowCreated: cetakRowCreated,
      adminUserCreated: adminCreated,
      message: 'Setup selesai. Sheet & Header siap; data lama tetap utuh.'
    };
  });
}

/** Cari Sheet berdasarkan nama resmi ATAU alias yang didukung. */
function findSheetByName_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (sh) return sh;
  var aliases = sheetAliases_();
  if (aliases[name]) {
    sh = ss.getSheetByName(aliases[name]);
    if (sh) return sh;
  }
  for (var k in aliases) {
    if (aliases[k] === name) {
      sh = ss.getSheetByName(k);
      if (sh) return sh;
    }
  }
  return null;
}

/** Pastikan Sheet ada (buat bila belum ada) + Header lengkap. */
function ensureSheet_(ss, def) {
  var sh = findSheetByName_(ss, def.name);
  var created = false;
  if (!sh) {
    sh = ss.insertSheet(def.name);
    created = true;
  }
  var added = ensureHeaders_(sh, def.headers);
  /* Header bisa saja baru ditulis/dilengkapi di sini — buang memo header sheet ini. */
  try { memoClearSheet_(sh.getName()); } catch (e2) { }
  if (created) {
    try {
      sh.getRange(1, 1, 1, def.headers.length).setFontWeight('bold').setBackground('#e7f0f8');
      sh.setFrozenRows(1);
    } catch (e) { logError_('ensureSheet_', e); }
  }
  /* Format Plain Text untuk kolom Kode Program agar 1.02.01.2.01
     tidak otomatis menjadi tanggal. Hanya PROGRAM/kodeProgram & kode. */
  try {
    if (def.name === 'PROGRAM') {
      var __hdrs = getSheetHeaders_(sh);
      var __cols = [];
      var __i1 = __hdrs.indexOf('kodeProgram');
      if (__i1 !== -1) __cols.push(__i1 + 1);
      var __i2 = __hdrs.indexOf('kode');
      if (__i2 !== -1 && __i2 !== __i1) __cols.push(__i2 + 1);
      for (var __ci = 0; __ci < __cols.length; __ci++) {
        sh.getRange(2, __cols[__ci], Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat('@');
      }
    }
  } catch (e) { /* format awal gagal tidak boleh menggagalkan setup */ }
  return { sheet: sh, created: created, addedHeaders: added };
}

/**
 * Pastikan seluruh Header tersedia.
 * - Header lama dipertahankan apa adanya (nama, urutan, posisi).
 * - Header yang belum ada DITAMBAHKAN di ujung kanan (tanpa duplikat).
 * - Tidak menyentuh baris data.
 */
function ensureHeaders_(sh, headers) {
  var lastCol = sh.getLastColumn();
  var row = [];
  if (lastCol > 0) row = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var existing = [];
  for (var i = 0; i < row.length; i++) {
    var h = toStr_(row[i]).trim();
    if (h) existing.push(h);
  }
  var added = [];
  for (var j = 0; j < headers.length; j++) {
    if (existing.indexOf(headers[j]) === -1) added.push(headers[j]);
  }
  if (added.length) {
    var startCol = Math.max(row.length, existing.length) + 1;
    sh.getRange(1, startCol, 1, added.length).setValues([added]);
  }
  return added;
}

/**
 * Hapus Sheet bawaan Google ("Sheet1") HANYA bila benar-benar kosong dan
 * bukan salah satu Sheet aplikasi. Sheet berisi data TIDAK PERNAH dihapus.
 */
function removeEmptyDefaultSheet_(ss, defs) {
  var names = ['Sheet1', 'Sheet 1'];
  for (var i = 0; i < names.length; i++) {
    var sh = null;
    try { sh = ss.getSheetByName(names[i]); } catch (e) { sh = null; }
    if (!sh) continue;
    var isAppSheet = false;
    for (var k in defs) { if (defs[k].name === sh.getName()) isAppSheet = true; }
    if (isAppSheet) continue;
    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    var isEmpty = true;
    if (lastRow > 0 && lastCol > 0) {
      var vals = sh.getRange(1, 1, lastRow, lastCol).getValues();
      for (var r = 0; r < vals.length && isEmpty; r++) {
        for (var c = 0; c < vals[r].length; c++) {
          if (isFilled_(vals[r][c])) { isEmpty = false; break; }
        }
      }
    }
    if (!isEmpty) continue;
    if (ss.getSheets().length <= 1) continue;
    try { ss.deleteSheet(sh); logInfo_('setupDatabase', 'Sheet bawaan kosong dihapus: ' + names[i]); } catch (e) { }
    memoClearAll_();
  }
}

/* ------------------------------- RESET DATA ------------------------------- */

/** Sheet yang baris datanya dikosongkan oleh resetData() (header baris 1 tetap). */
var RESETTABLE_SHEETS_ = ['SPPD', 'SPPD_PELAKSANA', 'RAMPUNG', 'RAMPUNG_RINCIAN', 'RAMPUNG_KOMPONEN',
  'PROGRAM', 'PROGRAM_REKENING', 'STANDAR_HARGA', 'KOP', 'PEJABAT', 'DASAR_HUKUM', 'USER'];

/**
 * RESET DATA — kosongkan seluruh baris data percobaan pada Sheet transaksi &
 * master. Header (baris 1) dan struktur kolom TIDAK disentuh. Dipertahankan:
 *   - CETAK_CONFIG (konfigurasi teknis cetak);
 *   - SISTEM (metadata aplikasi);
 *   - INSTANSI ditulis ulang menjadi KOSONG (diisi pengguna via Pengaturan);
 *   - akun Admin dibuat ulang otomatis (wajib untuk login).
 * TIDAK dijalankan otomatis oleh doGet/setupDatabase/aplikasi — hanya manual.
 *
 * Cara pakai (dari editor Apps Script):
 *   resetData({ confirm: 'RESET' })
 */
function resetData(opt) {
  opt = opt || {};
  if (String(opt.confirm || '').trim().toUpperCase() !== 'RESET') {
    return fail_('Reset dibatalkan. Konfirmasi wajib: panggil resetData({ confirm: "RESET" }).', null);
  }
  return safeRun_('Reset selesai. Header & struktur tetap ada; data transaksi = 0.', function () {
    /* withScriptLock_ menangani runtime yang tidak menyediakan LockService
       (di editor, waitLock dapat melempar) — reset tidak boleh gagal karenanya. */
    return withScriptLock_(function () {
    /* Penulisan massal: mulai dari memo bersih agar pembacaan berikutnya
       benar-benar melihat kondisi setelah reset (bukan metadata lama). */
    memoClearAll_();
    var ss = getSpreadsheet_();
      var defs = sheetDefs_();
      var cleared = [];

      /* 1) Kosongkan baris data (mulai baris 2) pada semua Sheet yang ditentukan. */
      for (var i = 0; i < RESETTABLE_SHEETS_.length; i++) {
        var key = RESETTABLE_SHEETS_[i];
        var def = defs[key];
        if (!def) continue;
        var res = ensureSheet_(ss, def);
        var sh = res.sheet;
        var lastRow = sh.getLastRow();
        var lastCol = sh.getLastColumn();
        if (lastRow > 1 && lastCol > 0) {
          sh.getRange(2, 1, lastRow - 1, lastCol).clearContent();
          cleared.push(def.name);
        }
      }

      /* 2) INSTANSI: tulis ulang baris singleton menjadi nilai kosong. */
      var defIns = defs.INSTANSI;
      var shIns = ensureSheet_(ss, defIns).sheet;
      var headersIns = getSheetHeaders_(shIns);
      var rowIns = objToRowArray_(defIns, headersIns, defaultInstansi_(), Math.max(headersIns.length, 1));
      if (headersIns.length) shIns.getRange(2, 1, 1, rowIns.length).setValues([rowIns]);
      var instansiReset = true;

      /* Seluruh isi Sheet dikosongkan langsung (clearContent), jadi revisi
         dinaikkan untuk SEMUA Sheet agar cache frontend tidak menyajikan
         data hasil reset yang lama. */
      bumpAllDataRevs_();

      /* 3) CETAK_CONFIG & SISTEM dipertahankan apa adanya (konfigurasi sistem). */

      /* 4) Buat ulang akun Admin (Sheet USER baru saja dikosongkan). */
      var adminCreated = ensureAdminUser_(ss, defs.USER);
      /* Semua data referensi berubah -> bersihkan SELURUH cache referensi,
         bukan hanya instansi/cetakconfig (sisanya akan basi sampai TTL habis). */
      refCacheBustAll_();

      var props = PropertiesService.getScriptProperties();
      props.setProperty(PROP_LAST_SETUP, nowIso_());
      logInfo_('resetData', 'Sheet dikosongkan: ' + cleared.join(', '));

      return {
        clearedSheets: cleared,
        instansiReset: instansiReset,
        adminUserCreated: adminCreated,
        preservedSheets: ['INSTANSI (dikosongkan)', 'CETAK_CONFIG', 'SISTEM'],
        message: 'Reset selesai. Header tetap ada; data transaksi = 0; akun admin dibuat ulang.'
      };
    });
  });
}

/* --------------------------- PENGATURAN 1 BARIS -------------------------- */

/** Buat baris pengaturan bila belum ada (mengembalikan true bila dibuat). */
function ensureSingletonRow_(ss, def, key, defaults) {
  var sh = ensureSheet_(ss, def).sheet;
  var rows = readObjects_(sh, def);
  var found = null;
  for (var i = 0; i < rows.length; i++) {
    if (toStr_(rows[i][def.keyField]).trim() === key) { found = rows[i]; break; }
  }
  var changed = false;
  if (!found) {
    var obj = clone_(defaults);
    obj[def.keyField] = key;
    obj.updatedAt = nowIso_();
    appendObject_(sh, def, obj);
    changed = true;
  } else {
    /* Hanya MELENGKAPI field yang belum ada — nilai lama tidak diubah. */
    var patch = {};
    for (var f in defaults) {
      if (found[f] === undefined || found[f] === null || found[f] === '') {
        var dv = defaults[f];
        if (dv !== '' && dv !== null && dv !== undefined) patch[f] = dv;
      }
    }
    if (Object.keys(patch).length) { updateObject_(sh, def, key, patch); changed = true; }
  }
  return changed;
}

/** Tabel kunci/nilai (Sheet SISTEM). */
function getSetting_(ss, kunci) {
  var def = getDef_('SISTEM');
  var sh = findSheetByName_(ss, def.name);
  if (!sh) return '';
  var rows = readObjects_(sh, def);
  for (var i = 0; i < rows.length; i++) {
    if (toStr_(rows[i].kunci).trim() === kunci) return toStr_(rows[i].nilai);
  }
  return '';
}

function setSetting_(ss, kunci, nilai, keterangan) {
  var def = getDef_('SISTEM');
  var sh = ensureSheet_(ss, def).sheet;
  var rows = readObjects_(sh, def);
  var obj = { kunci: kunci, nilai: nilai, keterangan: keterangan || '', updatedAt: nowIso_() };
  for (var i = 0; i < rows.length; i++) {
    if (toStr_(rows[i].kunci).trim() === kunci) {
      return updateObject_(sh, def, kunci, obj);
    }
  }
  return appendObject_(sh, def, obj);
}

function setSettingIfEmpty_(ss, kunci, nilai, keterangan) {
  if (isFilled_(getSetting_(ss, kunci))) return false;
  setSetting_(ss, kunci, nilai, keterangan);
  return true;
}

/** Buat akun Admin awal HANYA bila Sheet USER belum punya baris data. */
function ensureAdminUser_(ss, def) {
  var sh = ensureSheet_(ss, def).sheet;
  var rows = readObjects_(sh, def);
  if (rows.length > 0) return false;
  var admin = defaultAdminUser_();
  admin.createdAt = nowIso_();
  admin.updatedAt = nowIso_();
  appendObject_(sh, def, admin);
  return true;
}

/* ------------------------------ AKSES BARIS ------------------------------ */

/** Buka Sheet untuk definisi tertentu (otomatis dibuat bila belum ada). */
function openSheet_(def) {
  var ss = getSpreadsheet_();
  if (MEMO_SHEETS_[def.name]) return MEMO_SHEETS_[def.name];
  var sh = ensureSheet_(ss, def).sheet;
  MEMO_SHEETS_[def.name] = sh;
  return sh;
}

/** Nama Header pada baris 1 (sepanjang kolom yang terpakai). */
function getSheetHeaders_(sh) {
  var name = sh.getName();
  if (MEMO_HEADERS_[name]) return MEMO_HEADERS_[name];
  var lastCol = sh.getLastColumn();
  if (lastCol <= 0) return [];
  var row = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var out = [];
  for (var i = 0; i < row.length; i++) out.push(toStr_(row[i]).trim());
  MEMO_HEADERS_[name] = out;
  return out;
}

/** True bila objek punya minimal satu nilai terisi (selain __row). */
function hasAnyValue_(obj) {
  for (var k in obj) {
    if (k === '__row') continue;
    if (isFilled_(obj[k])) return true;
  }
  return false;
}

/** Satu baris Sheet → objek mentah berdasarkan NAMA header. */
function rowToObject_(def, headers, rowValues) {
  var obj = {};
  var extra = {};
  var extraField = def.extraJson || 'extraJson';
  for (var c = 0; c < headers.length; c++) {
    var h = headers[c];
    if (!h) continue;
    var raw = rowValues[c];
    if (h === extraField) {
      var ex = jsonParse_(raw, null);
      if (ex && typeof ex === 'object') extra = ex;
      continue;
    }
    obj[h] = raw;
  }
  /* Field cadangan hanya ditambahkan bila belum punya kolom sendiri. */
  for (var k in extra) { if (obj[k] === undefined) obj[k] = extra[k]; }
  return obj;
}

/** Objek → array nilai sel sesuai urutan Header (panjang = width). */
function objToRowArray_(def, headers, obj, width) {
  obj = obj || {};
  var extraField = def.extraJson || null;
  var extra = {};
  for (var k in obj) {
    if (k === '__row') continue;
    if (headers.indexOf(k) === -1 && k !== extraField) extra[k] = obj[k];
  }
  var hasExtra = false;
  for (var kk in extra) { hasExtra = true; break; }
  var row = [];
  for (var c = 0; c < width; c++) {
    var h = headers[c];
    if (!h) { row.push(''); continue; }
    if (h === extraField) { row.push(hasExtra ? jsonStringify_(extra) : ''); continue; }
    if (obj[h] === undefined) { row.push(''); continue; }
    row.push(cellToString_(obj[h]));
  }
  return row;
}

/** Baca baris Sheet sebagai objek (+ __row = nomor baris Sheet).
 *  OPT: filterField/filterValue — filter di level baris SEBELUM konversi objek
 *  (menghemat waktu untuk dataset besar). */
function readObjects_(sh, def, filterField, filterValue) {
  var headers = getSheetHeaders_(sh);
  var out = [];
  if (!headers.length) return out;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return out;
  var filterCol = -1;
  var filterTarget = '';
  if (filterField && filterValue !== undefined) {
    filterCol = headers.indexOf(filterField);
    filterTarget = toStr_(filterValue).trim();
  }
  var values = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  for (var i = 0; i < values.length; i++) {
    /* OPT: skip baris yang tidak cocok tanpa konversi objek. */
    if (filterCol >= 0 && toStr_(values[i][filterCol]).trim() !== filterTarget) continue;
    var obj = rowToObject_(def, headers, values[i]);
    if (!hasAnyValue_(obj)) continue;      /* lewati baris kosong */
    obj.__row = i + 2;
    out.push(obj);
  }
  return out;
}

/** Cari nomor baris Sheet berdasarkan nilai kolom ID unik (keyField). */
function findRowIndex_(sh, def, idValue) {
  var headers = getSheetHeaders_(sh);
  var map = headerIndexMap_(headers);
  var col = map[def.keyField];
  if (col === undefined) {
    throwError_('Kolom "' + def.keyField + '" tidak ditemukan pada Sheet ' + def.name + '.');
  }
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return -1;
  var values = sh.getRange(2, col + 1, lastRow - 1, 1).getValues();
  var target = toStr_(idValue).trim();
  for (var i = 0; i < values.length; i++) {
    if (toStr_(values[i][0]).trim() === target) return i + 2;
  }
  return -1;
}

/** Pastikan kolom Kode Program pada Sheet PROGRAM diformat Plain Text (@)
 *  agar nilai seperti 1.02.01.2.01 tidak diubah Google Sheets menjadi tanggal.
 *  Hanya kolom Kode Program (kodeProgram & kode) yang diformat — kolom lain tidak disentuh. */
function ensureKodeProgramTextFormat_(sh, def, row, numRows) {
  if (!sh || !def || def.name !== 'PROGRAM') return;
  var r = row || 2;
  var n = numRows || 1;
  try {
    var headers = getSheetHeaders_(sh);
    var cols = [];
    var idx1 = headers.indexOf('kodeProgram');
    if (idx1 !== -1) cols.push(idx1 + 1);
    var idx2 = headers.indexOf('kode');
    if (idx2 !== -1 && idx2 !== idx1) cols.push(idx2 + 1);
    for (var i = 0; i < cols.length; i++) {
      sh.getRange(r, cols[i], n, 1).setNumberFormat('@');
    }
  } catch (e) { /* format gagal tidak boleh menggagalkan tulis */ }
}

/** Tambah satu baris data (selalu di bawah data terakhir). */
function appendObject_(sh, def, obj) {
  return withScriptLock_(function () {
    var headers = getSheetHeaders_(sh);
    var width = headers.length;
    var row = objToRowArray_(def, headers, obj, width);
    var target = Math.max(sh.getLastRow(), 1) + 1;
    ensureKodeProgramTextFormat_(sh, def, target, 1);
    sh.getRange(target, 1, 1, width).setValues([row]);
    /* Isi Sheet berubah -> revisi naik, agar cache frontend tidak menyajikan
       data lama. Dipusatkan di sini supaya TIDAK ADA jalur tulis yang terlewat. */
    bumpDataRev_(def && def.name);
    return obj;
  });
}

/** Tambah BANYAK baris dalam SATU panggilan setValues (batch).
 *  Hasil AKHIR identik dengan memanggil appendObject_() berulang (baris
 *  tertulis berurutan tepat di bawah data terakhir), tetapi hanya memakai
 *  1x getLastRow + 1x setValues + 1x bump revisi — jauh lebih cepat untuk
 *  rincian/komponen Rampung dan pelaksana SPPD yang berjumlah banyak. */
function appendObjects_(sh, def, list) {
  var arr = list || [];
  if (!arr.length) return 0;
  return withScriptLock_(function () {
    var headers = getSheetHeaders_(sh);
    var width = headers.length;
    if (!width) return 0;
    var rows = [];
    for (var i = 0; i < arr.length; i++) {
      rows.push(objToRowArray_(def, headers, arr[i], width));
    }
    var target = Math.max(sh.getLastRow(), 1) + 1;
    ensureKodeProgramTextFormat_(sh, def, target, rows.length);
    sh.getRange(target, 1, rows.length, width).setValues(rows);
    bumpDataRev_(def && def.name);
    return arr.length;
  });
}

/**
 * Ubah satu baris data berdasarkan ID unik.
 * `patch` hanya menimpa field yang dikirim; field lain (termasuk kolom
 * tambahan buatan pengguna) tetap utuh.
 */
function updateObject_(sh, def, idValue, patch) {
  return withScriptLock_(function () {
    var r = findRowIndex_(sh, def, idValue);
    if (r < 0) {
      throwError_(def.label + ' dengan ID "' + idValue + '" tidak ditemukan.');
    }
    var headers = getSheetHeaders_(sh);
    var width = headers.length;
    var current = sh.getRange(r, 1, 1, width).getValues()[0];
    var obj = rowToObject_(def, headers, current);
    for (var k in (patch || {})) {
      if (patch[k] !== undefined) obj[k] = patch[k];
    }
    var row = objToRowArray_(def, headers, obj, width);
    /* Kolom tanpa header (gap) tidak boleh dikosongkan. */
    for (var c = 0; c < width; c++) {
      if (!headers[c]) row[c] = current[c];
    }
    ensureKodeProgramTextFormat_(sh, def, r, 1);
    sh.getRange(r, 1, 1, width).setValues([row]);
    bumpDataRev_(def && def.name);
    return obj;
  });
}
/** Ambil satu objek mentah berdasarkan ID unik (null bila tidak ada). */
function getObjectById_(sh, def, idValue) {
  var r = findRowIndex_(sh, def, idValue);
  if (r < 0) return null;
  var headers = getSheetHeaders_(sh);
  var row = sh.getRange(r, 1, 1, headers.length).getValues()[0];
  var obj = rowToObject_(def, headers, row);
  obj.__row = r;
  return obj;
}

/** Hapus SATU baris data berdasarkan ID unik. */
function deleteObject_(sh, def, idValue) {
  return withScriptLock_(function () {
    var r = findRowIndex_(sh, def, idValue);
    if (r < 0) {
      throwError_(def.label + ' dengan ID "' + idValue + '" tidak ditemukan.');
    }
    sh.deleteRow(r);
    bumpDataRev_(def && def.name);
    return true;
  });
}

/* ------------------------------ TABEL ANAK ------------------------------- */

/** Baris anak yang terhubung ke satu induk (parentKeyField = parentId).
 *  OPT: filter di level sheet SEBELUM konversi objek. */
function readChildObjects_(sh, def, parentId) {
  return readObjects_(sh, def, def.parentKeyField, parentId);
}

/** Hapus seluruh baris anak milik satu induk (dari bawah ke atas). */
function deleteChildRows_(sh, def, parentId) {
  return withScriptLock_(function () {
    var rows = readChildObjects_(sh, def, parentId);
    if (!rows.length) return 0;
    var idx = [];
    for (var i = 0; i < rows.length; i++) idx.push(rows[i].__row);
    idx.sort(function (a, b) { return b - a; });
    for (var j = 0; j < idx.length; j++) sh.deleteRow(idx[j]);
    bumpDataRev_(def && def.name);
    return idx.length;
  });
}

/** Ganti seluruh baris anak milik satu induk dengan daftar baru.
 *  Penulisan memakai SATU batch setValues (lihat appendObjects_) — hasil
 *  akhir identik dengan penulisan satu per satu. */
function replaceChildObjects_(sh, def, parentId, list) {
  return withScriptLock_(function () {
    deleteChildRows_(sh, def, parentId);
    var arr = list || [];
    var batch = [];
    for (var i = 0; i < arr.length; i++) {
      var obj = clone_(arr[i]);
      obj[def.parentKeyField] = parentId;
      if (!isFilled_(obj[def.keyField])) obj[def.keyField] = newId_(def.idPrefix);
      if (obj.createdAt === undefined) obj.createdAt = nowIso_();
      obj.updatedAt = nowIso_();
      batch.push(obj);
    }
    appendObjects_(sh, def, batch);
    return arr.length;
  });
}

/* ----------------------- BARIS ANAK BERDASAR FIELD ------------------------ */

/**
 * Hapus seluruh baris yang nilai field-nya cocok (umumnya kolom idRampung /
 * idRincian pada sheet anak). Dipakai Rampung.gs.
 * OPT: gunakan filter di readObjects_() untuk skip baris tidak cocok.
 */
function deleteChildRowsByField_(sh, fieldName, fieldValue) {
  return withScriptLock_(function () {
    var def = getDefForSheet_(sh);
    var rows = readObjects_(sh, def, fieldName, fieldValue);
    if (!rows.length) return 0;
    var idx = [];
    for (var i = 0; i < rows.length; i++) idx.push(rows[i].__row);
    idx.sort(function (a, b) { return b - a; });
    for (var j = 0; j < idx.length; j++) sh.deleteRow(idx[j]);
    try { bumpDataRev_(sh.getName()); } catch (e) { }
    return idx.length;
  });
}

/** Cari definisi Sheet dari objek Sheet (untuk deleteChildRowsByField_). */
function getDefForSheet_(sh) {
  var defs = sheetDefs_();
  var nm = '';
  try { nm = sh.getName(); } catch (e) { nm = ''; }
  var aliases = sheetAliases_();
  if (aliases[nm]) nm = aliases[nm];
  for (var k in defs) {
    if (defs[k].name === nm) return defs[k];
    if (k === nm) return defs[k];
  }
  return { name: nm, keyField: 'id', extraJson: 'extraJson' };
}

/** Jumlah baris data pada sebuah Sheet. */
function countRows_(sh) {
  return Math.max(sh.getLastRow() - 1, 0);
}

/* ------------------------- PEMERIKSAAN KESEHATAN DB --------------------- */

/**
 * periksaDatabase() — laporan kondisi database (tidak mengubah data).
 * Berguna untuk memastikan seluruh Sheet & Header siap sebelum dipakai.
 */
function periksaDatabase() {
  return safeRun_('Pemeriksaan database selesai.', function () {
    var ss = getSpreadsheet_();
    var defs = sheetDefs_();
    var order = sheetOrder_();
    var report = [];
    var masalah = [];

    for (var i = 0; i < order.length; i++) {
      var def = getDef_(order[i]);
      var sh = findSheetByName_(ss, def.name);
      if (!sh) {
        masalah.push('Sheet ' + def.name + ' belum ada.');
        report.push({ sheet: def.name, ada: false, baris: 0, header: [] });
        continue;
      }
      var headers = getSheetHeaders_(sh);
      var missing = [];
      for (var j = 0; j < def.headers.length; j++) {
        if (headers.indexOf(def.headers[j]) === -1) missing.push(def.headers[j]);
      }
      if (missing.length) masalah.push('Sheet ' + def.name + ' kurang header: ' + missing.join(', '));
      report.push({ sheet: def.name, ada: true, baris: countRows_(sh), kolom: headers.length, header: headers, headerKurang: missing });
    }

    return {
      spreadsheetId: ss.getId(),
      spreadsheetUrl: ss.getUrl(),
      schemaVersion: SCHEMA_VERSION,
      siap: masalah.length === 0,
      masalah: masalah,
      sheets: report
    };
  });
}

/* ==========================================================================
 * END FILE ASLI: Database.gs
 * ========================================================================== */

/* ==========================================================================
 * BEGIN FILE ASLI: Auth.gs
 * ========================================================================== */
/**
 * ============================================================================
 * SIAP SPPD & RAMPUNG — AUTENTIKASI & OTORISASI (SISI SERVER)
 * File   : Auth.gs
 *
 * Aturan login DIPERTAHANKAN sama dengan frontend:
 *   1. Admin    : username `admin` / password `admin`.
 *   2. Pengguna : username = NIP / password = NIP (USER.nip, status Aktif).
 *   3. Tambahan : username & password persis untuk akun ber-role `admin`
 *                 (akun hasil pembuatan Admin pada menu User).
 *
 * Sesi TIDAK dipercaya dari browser: setiap permintaan membawa token yang
 * dibuat & diverifikasi di server (CacheService, berlaku 8 jam).
 * ============================================================================
 */

/* ------------------------------- SESSION STORE --------------------------- */

function sessionStorePut_(token, sess) {
  var raw = jsonStringify_(sess);
  var key = PROP_SESSION_PREFIX + token;
  /* BUG YANG DIPERBAIKI: CacheService menolak TTL > 3600 detik pada akun
     konsumen (Google Workspace pribadi), sehingga put() selalu melempar dan
     jalur cepat cache mati — setiap request jatuh ke cadangan Script
     Properties yang baru terpakai penuh bila banyak pengguna login.
     TTL cache ditekan ke maks 1 jam; masa berlaku 8 jam tetap dijaga oleh
     sess.expiresAt (cadangan Properties mengikat selama 8 jam penuh). */
  var cacheTtl = Math.max(60, Math.min(SESSION_TTL_SECONDS, 3600));
  try {
    CacheService.getScriptCache().put(key, raw, cacheTtl);
  } catch (e) { logError_('sessionStorePut_', e); }
  /* Cadangan bila CacheService bermasalah (kedaluwarsa otomatis saat dibaca). */
  try {
    PropertiesService.getScriptProperties().setProperty(key, raw);
  } catch (e2) { }
  /* BUG YANG DIPERBAIKI: token sesi yang tidak pernah di-logout menumpuk
     permanen di Script Properties (pembersihan lama hanya terjadi saat token
     dibaca). Store Properties punya batas jumlah kunci & ukuran; sesak oleh
     sesi mati akan mengganggu penyimpanan konfigurasi lain. Sapu sesi
     kedaluwarsa secara berkala (maks. sekali per 15 menit). */
  sessionSweepExpired_();
  return true;
}

var SESS_SWEEP_INTERVAL_MS_ = 15 * 60 * 1000;
/* CATATAN: nama marker sengaja TIDAK berawalan SESS_ agar tidak ikut
   terhapus oleh sweep itu sendiri (yang menyapu semua kunci SESS_*). */
var SESS_SWEEP_AT_PROP_ = 'SESSION_SWEEP_AT';

/** Hapus properti sesi (SESS_*) yang expiresAt-nya sudah lewat. */
function sessionSweepExpired_() {
  var props = null;
  try { props = PropertiesService.getScriptProperties(); } catch (e) { return; }
  if (!props) return;
  var now = Date.now();
  try {
    var lastRun = parseInt(props.getProperty(SESS_SWEEP_AT_PROP_) || '0', 10);
    if (lastRun && (now - lastRun) < SESS_SWEEP_INTERVAL_MS_) return;
    props.setProperty(SESS_SWEEP_AT_PROP_, String(now));
  } catch (e2) { return; }
  try {
    var all = props.getProperties();
    var hapus = [];
    for (var key in all) {
      if (key.indexOf(PROP_SESSION_PREFIX) !== 0) continue;
      var s = jsonParse_(all[key], null);
      var exp = s && s.expiresAt ? new Date(s.expiresAt).getTime() : 0;
      if (!exp || exp < now) hapus.push(key);
    }
    if (hapus.length) props.deleteProperties(hapus);
  } catch (e3) { logError_('sessionSweepExpired_', e3); }
}

function sessionStoreGet_(token) {
  if (!isFilled_(token)) return null;
  var key = PROP_SESSION_PREFIX + String(token).trim();
  var raw = '';
  try { raw = CacheService.getScriptCache().get(key) || ''; } catch (e) { raw = ''; }
  if (!raw) {
    try { raw = PropertiesService.getScriptProperties().getProperty(key) || ''; } catch (e2) { raw = ''; }
  }
  if (!raw) return null;
  var sess = jsonParse_(raw, null);
  if (!sess) return null;
  if (sess.expiresAt && new Date(sess.expiresAt).getTime() < new Date().getTime()) {
    sessionStoreDrop_(token);
    return null;
  }
  return sess;
}

function sessionStoreDrop_(token) {
  if (!isFilled_(token)) return false;
  var key = PROP_SESSION_PREFIX + String(token).trim();
  try { CacheService.getScriptCache().remove(key); } catch (e) { }
  try { PropertiesService.getScriptProperties().deleteProperty(key); } catch (e2) { }
  return true;
}

/* --------------------------------- LOGIN -------------------------------- */

/**
 * login(username, password) — juga menerima objek { username, password }.
 * Mengembalikan respons standar { success, data:{session}, message }.
 */
function login(username, password) {
  var payload = username;
  if (payload && typeof payload === 'object') {
    password = payload.password;
    username = payload.username;
  }
  var u = toStr_(username).trim();
  var p = toStr_(password);
  if (!u || !p) return fail_('Username dan password wajib diisi.');

  var users = listUsersRaw_();
  var rec = null;
  var i;

  /* 1) Admin utama — selalu `admin` / `admin`. */
  if (u === 'admin' && p === 'admin') {
    for (i = 0; i < users.length; i++) {
      if (toStr_(users[i].username).toLowerCase() === 'admin') { rec = users[i]; break; }
    }
    if (!rec) { rec = defaultAdminUser_(); rec.idUser = 'usr-admin'; }
    return buildLoginSuccess_(rec, 'admin', ['all']);
  }

  /* 2) Pengguna — username = NIP, password = NIP. */
  for (i = 0; i < users.length; i++) {
    if (toStr_(users[i].nip).trim() === u) { rec = users[i]; break; }
  }
  if (rec) {
    if (toStr_(rec.status || 'Aktif').trim() !== 'Aktif') {
      return fail_('Akun Anda tidak aktif. Hubungi Administrator.');
    }
    var nip = toStr_(rec.nip).trim();
    var pw = toStr_(rec.password);
    if (p !== nip && p !== pw) return fail_('Username atau password salah.');
    return buildLoginSuccess_(rec, toStr_(rec.role || 'pengguna'), splitList_(rec.access));
  }

  /* 3) Akun Admin buatan Administrator (username & password persis). */
  for (i = 0; i < users.length; i++) {
    var r2 = users[i];
    if (toStr_(r2.role) !== 'admin') continue;
    if (toStr_(r2.username).trim() !== u) continue;
    if (!isFilled_(r2.password) || toStr_(r2.password) !== p) continue;
    if (toStr_(r2.status || 'Aktif').trim() !== 'Aktif') {
      return fail_('Akun Anda tidak aktif. Hubungi Administrator.');
    }
    return buildLoginSuccess_(r2, 'admin', splitList_(r2.access));
  }

  return fail_('Username atau password salah.');
}

/** Buat token sesi + perbarui lastLogin, lalu kembalikan respons sukses. */
function buildLoginSuccess_(rec, role, access) {
  var token = newId_('tkn');
  var acc = (Array.isArray(access) && access.length) ? access : (role === 'admin' ? ['all'] : []);
  var sess = {
    token: token,
    userId: toStr_(rec.idUser || rec.id || ''),
    nama: toStr_(rec.nama || (role === 'admin' ? 'Administrator' : '')),
    username: toStr_(rec.username || ''),
    nip: toStr_(rec.nip || ''),
    role: role,
    access: acc,
    loginAt: nowIso_(),
    expiresAt: new Date(new Date().getTime() + SESSION_TTL_SECONDS * 1000).toISOString()
  };
  sessionStorePut_(token, sess);
  try {
    if (isFilled_(sess.userId)) {
      updateUser_(null, { idUser: sess.userId, lastLogin: lastLoginStamp_() }, true);
    }
  } catch (e) { logError_('buildLoginSuccess_/lastLogin', e); }
  var publik = clone_(sess);
  delete publik.expiresAt;
  return ok_({ session: publik }, 'Login berhasil.');
}
/** Akhiri sesi (token dihapus di server). */
function logout(token) {
  sessionStoreDrop_(token);
  return ok_(true, 'Logout berhasil.');
}

/** Info sesi aktif berdasarkan token. */
function getSession(token) {
  var sess = sessionStoreGet_(token);
  if (!sess) return fail_('Sesi tidak valid atau sudah berakhir. Silakan login kembali.');
  var publik = clone_(sess);
  delete publik.expiresAt;
  return ok_(publik, 'Sesi aktif.');
}

/* --------------------------- VALIDASI & OTORISASI ------------------------ */

/** Ambil sesi valid atau lempar error (dipakai setiap permintaan API). */
function requireSession_(token) {
  var sess = sessionStoreGet_(token);
  if (!sess) throwError_('Sesi tidak valid atau sudah berakhir. Silakan login kembali.');
  if (toStr_(sess.role) !== 'admin') {
    /* Pastikan akun masih ada & masih Aktif (hak akses bisa berubah). */
    var rec = findUserById_(sess.userId);
    if (rec) {
      if (toStr_(rec.status || 'Aktif').trim() !== 'Aktif') {
        sessionStoreDrop_(token);
        throwError_('Akun Anda tidak aktif. Hubungi Administrator.');
      }
      sess.access = splitList_(rec.access);
      sess.role = toStr_(rec.role || sess.role);
    }
  }
  return sess;
}

function requireAdmin_(token) {
  var sess = requireSession_(token);
  if (toStr_(sess.role) !== 'admin') {
    throwError_('Hanya Administrator yang diizinkan mengakses menu ini.');
  }
  return sess;
}

/** Boleh MENULIS modul ini? (mengikuti daftar `access` milik user). */
function canWriteModule_(sess, moduleName) {
  if (!sess) return false;
  if (toStr_(sess.role) === 'admin') return true;
  var acc = Array.isArray(sess.access) ? sess.access : [];
  if (acc.indexOf('all') > -1) return true;
  return acc.indexOf(moduleName) > -1;
}

/**
 * Boleh MEMBACA modul ini?
 * Membaca data referensi (SPPD, Rampung, Program, Standar Harga, Pengaturan)
 * diizinkan untuk semua sesi yang sah, karena Dashboard, pemilih pejabat/kop,
 * dan seluruh proses cetak memerlukannya. Menu tetap dibatasi di frontend
 * (canAccess) dan setiap PENULISAN tetap dibatasi di server (requireWrite_).
 * Daftar USER hanya boleh dibaca Administrator.
 */
function canReadModule_(sess, moduleName) {
  if (!sess) return false;
  if (moduleName === 'user') return toStr_(sess.role) === 'admin';
  return true;
}

/** Pastikan sesi memiliki hak tulis modul terkait. */
function requireWrite_(sess, moduleName) {
  if (!canWriteModule_(sess, moduleName)) {
    throwError_('Anda tidak memiliki hak akses untuk mengubah data ' + moduleName + '.');
  }
  return true;
}

/* ---------------- KEPEMILIKAN BARIS (ROW-LEVEL ACCESS) -------------------
 * Aturan: Admin melihat/mengelola SEMUA data. Pengguna hanya data miliknya
 * sendiri (kolom `createdBy` = userId pembuat). Baris lama yang `createdBy`-
 * nya kosong dianggap data lawas dan tetap terlihat semua pihak — tidak
 * ditetapkan ke siapa pun dan tidak diubah.
 * Identitas diambil dari SESI SERVER (tidak dipercaya dari browser).
 * ------------------------------------------------------------------------ */

/** True bila sesi milik Administrator. */
function isAdmin_(sess) {
  return !!sess && toStr_(sess.role) === 'admin';
}

/** Identitas pemilik untuk sesi ini (stabil: userId → username → nip). */
function sessionOwnerId_(sess) {
  if (!sess) return '';
  var id = toStr_(sess.userId).trim();
  if (id) return id;
  id = toStr_(sess.username).trim();
  if (id) return id;
  return toStr_(sess.nip).trim();
}

/** Pemilik yang tercatat pada satu baris data (kolom createdBy). */
function recordOwnerId_(rec) {
  if (!rec) return '';
  return toStr_(rec.createdBy || rec.createdby || '').trim();
}

/** Boleh sesi ini melihat/mengelola baris ini? */
function canAccessRecord_(sess, rec) {
  if (isAdmin_(sess)) return true;
  var owner = recordOwnerId_(rec);
  if (!owner) return true;
  return owner === sessionOwnerId_(sess);
}

/** Pastikan sesi boleh mengelola baris ini, atau lempar error. */
function requireRecordAccess_(sess, rec, label) {
  if (!canAccessRecord_(sess, rec)) {
    throwError_((label || 'Data') + ' bukan milik Anda. Anda hanya dapat mengelola data yang Anda buat sendiri.');
  }
  return true;
}

/** Saring daftar agar hanya berisi baris milik sesi ini (Admin = semua).
 *  Selalu kembalikan array baru; daftar asli tidak diubah. */
function filterOwnedList_(sess, list) {
  if (isAdmin_(sess) || !Array.isArray(list)) return list || [];
  var me = sessionOwnerId_(sess);
  var out = [];
  for (var i = 0; i < list.length; i++) {
    var owner = recordOwnerId_(list[i]);
    if (!owner || owner === me) out.push(list[i]);
  }
  return out;
}

/* ==========================================================================
 * END FILE ASLI: Auth.gs
 * ========================================================================== */

/* ==========================================================================
 * BEGIN FILE ASLI: User.gs
 * ========================================================================== */
/**
 * ============================================================================
 * SIAP SPPD & RAMPUNG - MODUL USER / LOGIN (SISI SERVER)
 * File   : User.gs
 * Struktur frontend: { id,nama,username,nip,role,status,password,access,
 *   lastLogin,createdAt,updatedAt }
 * Aturan login DIPERTAHANKAN: admin/admin, pengguna NIP/NIP, akun admin persis.
 * ============================================================================
 */

function userSheet_() {
  var def = getDef_('USER');
  return { def: def, sh: openSheet_(def) };
}

function userFromRecord_(rec) {
  rec = rec || {};
  return {
    id: toStr_(rec.idUser),
    nama: toStr_(rec.nama),
    username: toStr_(rec.username),
    nip: toStr_(rec.nip),
    role: isFilled_(rec.role) ? toStr_(rec.role) : 'pengguna',
    status: isFilled_(rec.status) ? toStr_(rec.status) : 'Aktif',
    password: isFilled_(rec.password) ? toStr_(rec.password) : '',
    access: normalizeUserAccess_(rec.access),
    lastLogin: toStr_(rec.lastLogin) || 'Belum pernah login',
    createdAt: toStr_(rec.createdAt),
    updatedAt: toStr_(rec.updatedAt)
  };
}

function userPublic_(obj) {
  var o = clone_(obj) || {};
  delete o.password;
  return o;
}

function userToRecord_(data) {
  var d = clone_(data) || {};
  return {
    idUser: toStr_(d.id || d.idUser),
    nama: toStr_(d.nama),
    username: toStr_(d.username),
    nip: toStr_(d.nip),
    role: isFilled_(d.role) ? toStr_(d.role) : 'pengguna',
    status: isFilled_(d.status) ? toStr_(d.status) : 'Aktif',
    password: isFilled_(d.password) ? toStr_(d.password) : '',
    access: jsonStringify_(normalizeUserAccess_(d.access)),
    lastLogin: isFilled_(d.lastLogin) ? toStr_(d.lastLogin) : 'Belum pernah login',
    createdAt: toStr_(d.createdAt),
    updatedAt: toStr_(d.updatedAt)
  };
}

/** Seluruh User MENTAH (dipakai Auth.gs - termasuk password). */
function listUsersRaw_() {
  var s = userSheet_();
  return readObjects_(s.sh, s.def);
}

function fetchUsersFull_() {
  var s = userSheet_();
  var rows = readObjects_(s.sh, s.def);
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var obj = userFromRecord_(rows[i]);
    obj.__row = rows[i].__row;
    out.push(obj);
  }
  out = sortByCreatedDesc_(out);
  for (var j = 0; j < out.length; j++) delete out[j].__row;
  return out;
}

/** Daftar User PUBLIK tanpa password (dipakai API getUsers). */
function fetchUsersPublic_() {
  var list = fetchUsersFull_();
  var out = [];
  for (var i = 0; i < list.length; i++) out.push(userPublic_(list[i]));
  return out;
}

function findUserById_(id) {
  var key = toStr_(id).trim();
  if (!key) return null;
  var s = userSheet_();
  return getObjectById_(s.sh, s.def, key);
}
/** Simpan (buat/ubah) satu User. buatBaru=true -> username unik + password wajib. */
function saveUserRecord_(data, buatBaru, strict) {
  var d = clone_(data) || {};
  var s = userSheet_();
  var id = toStr_(d.id || d.idUser);
  var lama = id ? getObjectById_(s.sh, s.def, id) : null;
  if (!lama) id = id || newId_('usr');
  d.id = id;
  d.idUser = id;
  if (strict !== false) {
    requireFields_(d, [
      { field: 'nama', label: 'Nama User' },
      { field: 'username', label: 'Username' },
      { field: 'role', label: 'Role' },
      { field: 'status', label: 'Status' }
    ], 'User');
    var role = toStr_(d.role);
    if (role !== 'admin' && role !== 'pengguna') {
      throwError_('User: Role harus "admin" atau "pengguna".');
    }
    if (role === 'pengguna' && !isFilled_(d.nip)) {
      throwError_('User: NIP wajib diisi agar pengguna dapat login.');
    }
  }
  if (buatBaru && !isFilled_(d.password)) {
    throwError_('User: Password wajib diisi (min. 6 karakter).');
  }
  var semua = readObjects_(s.sh, s.def);
  for (var i = 0; i < semua.length; i++) {
    if (toStr_(semua[i].idUser).trim() === id) continue;
    if (isFilled_(d.username) &&
        toStr_(semua[i].username).trim().toLowerCase() ===
        toStr_(d.username).trim().toLowerCase()) {
      throwError_('User: Username "' + d.username + '" sudah dipakai.');
    }
  }
  if (!lama) {
    if (!isFilled_(d.lastLogin)) d.lastLogin = 'Belum pernah login';
  } else if (!isFilled_(d.password)) {
    d.password = toStr_(lama.password);
  }
  if (!isFilled_(d.access)) d.access = normalizeUserAccess_(lama ? lama.access : []);
  var gabung = {};
  if (lama) {
    for (var k in lama) { if (k !== '__row') gabung[k] = lama[k]; }
    gabung.createdAt = toStr_(lama.createdAt);
  }
  for (var key in d) { if (d[key] !== undefined) gabung[key] = d[key]; }
  gabung.id = id;
  gabung.idUser = id;
  stampForWrite_(gabung, !lama);
  var rec = userToRecord_(gabung);
  rec.idUser = id;
  if (!lama) {
    if (!isFilled_(rec.createdAt)) rec.createdAt = nowIso_();
    appendObject_(s.sh, s.def, rec);
  } else {
    if (!isFilled_(rec.createdAt)) rec.createdAt = toStr_(lama.createdAt);
    updateObject_(s.sh, s.def, id, rec);
  }
  return fetchUserById_(id);
}

/** Satu User publik berdasarkan ID (null bila tidak ada). */
function fetchUserById_(id) {
  var s = userSheet_();
  var rec = getObjectById_(s.sh, s.def, id);
  return rec ? userPublic_(userFromRecord_(rec)) : null;
}

function removeUser_(id) {
  var key = requireId_(id, 'ID User');
  var s = userSheet_();
  assertFound_(getObjectById_(s.sh, s.def, key), 'User');
  deleteObject_(s.sh, s.def, key);
  return true;
}

/** Perbarui field kecil tanpa validasi penuh (mis. lastLogin). silent=true tak melempar. */
function updateUser_(token, data, silent) {
  var d = clone_(data) || {};
  var id = toStr_(d.id || d.idUser || d.idUsr);
  if (!id) {
    if (silent) return null;
    throwError_('ID User tidak ditemukan.');
  }
  var s = userSheet_();
  var lama = getObjectById_(s.sh, s.def, id);
  if (!lama) {
    if (silent) return null;
    throwError_('User dengan ID "' + id + '" tidak ditemukan.');
  }
  var patch = {};
  for (var k in d) {
    if (k === 'id' || k === 'idUser' || k === '__row') continue;
    if (d[k] !== undefined) patch[k] = d[k];
  }
  patch.updatedAt = nowIso_();
  updateObject_(s.sh, s.def, id, patch);
  return fetchUserById_(id);
}

/* ---- API PUBLIK (dipakai Api.gs) ---- */
function getUsers(token) {
  return safeRun_('Data User berhasil dibaca.', function () {
    requireAdmin_(token);
    return fetchUsersPublic_();
  });
}

function createUser(token, data) {
  return safeRun_('Data User berhasil disimpan.', function () {
    requireAdmin_(token);
    return saveUserRecord_(data, true, true);
  });
}

function saveUser(token, data) {
  return createUser(token, data);
}

function updateUserPub(token, data, dataAlt) {
  return safeRun_('Data User berhasil diperbarui.', function () {
    requireAdmin_(token);
    var payload = normalizeUpdatePayload_(data, dataAlt, 'id', 'ID User');
    return saveUserRecord_(payload, false, true);
  });
}

function deleteUser(token, id) {
  return safeRun_('Data User berhasil dihapus.', function () {
    requireAdmin_(token);
    return removeUser_(id);
  });
}

/* ==========================================================================
 * END FILE ASLI: User.gs
 * ========================================================================== */

/* ==========================================================================
 * BEGIN FILE ASLI: SPPD.gs
 * ========================================================================== */
/**
 * ============================================================================
 * SIAP SPPD & RAMPUNG — MODUL SPPD
 * File   : SPPD.gs
 *
 * Struktur data mengikuti field nyata frontend (index.html):
 *   { id, noSpt, noSppd, jenisSuratTugas, tglSpt, tglSppd, maksud,
 *     tempatBerangkat, kotaTujuan, instansiTujuan, transport, noTransport,
 *     ketTransport, ttdSpt, ttdSpd, pptk, hotel, alamatHotel, checkin, checkout,
 *     jumlahMalam, noKamar, pelaksana:[{nama,nip,jabatan,golongan,unit}],
 *     tglBerangkat, jamBerangkat, tglKembali, jamKembali, lamaHari,
 *     biaya:{uangHarian,representasi,transport,taksi,penginapan,sewaKendaraan,
 *            lainnya,ppn,pph}, total,
 *     dokumen:{spt,sppd,tiket,boarding,hotel,transport,pendukung},
 *     program, kegiatan, subkegiatan, rekening, status, createdAt }
 *
 * Relasi: SPPD (induk) → SPPD_PELAKSANA (anak, idSPPD), dan dipakai modul
 *         Rampung melalui idSPPD (lihat Rampung.gs).
 * ============================================================================
 */

/* -------------------------------- MAPPING -------------------------------- */

function sppdBiayaMap_() {
  return {
    biayaUangHarian: 'uangHarian', biayaRepresentasi: 'representasi',
    biayaTransport: 'transport', biayaTaksi: 'taksi', biayaPenginapan: 'penginapan',
    biayaSewaKendaraan: 'sewaKendaraan', biayaLainnya: 'lainnya',
    biayaPpn: 'ppn', biayaPph: 'pph'
  };
}

function sppdDokumenMap_() {
  return {
    dokSpt: 'spt', dokSppd: 'sppd', dokTiket: 'tiket', dokBoarding: 'boarding',
    dokHotel: 'hotel', dokTransport: 'transport', dokPendukung: 'pendukung'
  };
}

function sppdDefaultBiaya_() {
  return {
    uangHarian: 0, representasi: 0, transport: 0, taksi: 0,
    penginapan: 0, sewaKendaraan: 0, lainnya: 0, ppn: 0, pph: 0
  };
}

function sppdDefaultDokumen_() {
  return {
    spt: null, sppd: null, tiket: null, boarding: null,
    hotel: null, transport: null, pendukung: null
  };
}

/** Teks → null bila kosong (padanan dokumen:{...:null} pada frontend). */
function nullableText_(v) {
  var s = toStr_(v);
  return s === '' ? null : s;
}

/** Objek SPPD frontend → baris sheet (flat). */
function sppdToRecord_(data) {
  var d = data || {};
  var rec = {};
  var scalars = [
    'noSpt', 'noSppd', 'jenisSuratTugas', 'tglSpt', 'tglSppd', 'maksud',
    'tempatBerangkat', 'kotaTujuan', 'instansiTujuan', 'transport', 'noTransport',
    'ketTransport', 'ttdSpt', 'ttdSpd', 'pptk', 'hotel', 'alamatHotel', 'checkin',
    'checkout', 'noKamar', 'tglBerangkat', 'jamBerangkat', 'tglKembali', 'jamKembali',
    'program', 'subkegiatan', 'rekening', 'status', 'createdAt', 'updatedAt'
  ];
  for (var i = 0; i < scalars.length; i++) {
    if (d[scalars[i]] !== undefined) rec[scalars[i]] = d[scalars[i]];
  }
  rec.idSPPD = toStr_(d.id || d.idSPPD || '');
  if (d.createdBy !== undefined) rec.createdBy = toStr_(d.createdBy);
  if (d.jumlahMalam !== undefined) rec.jumlahMalam = toInt_(d.jumlahMalam, 0);
  if (d.lamaHari !== undefined) rec.lamaHari = toInt_(d.lamaHari, 1);

  /* Hanya menulis kolom biaya/dokumen bila datanya memang dikirim, agar
     pembaruan sebagian (partial update) tidak menghapus nilai lama. */
  if (d.biaya !== undefined) {
    var biaya = d.biaya || {};
    var bMap = sppdBiayaMap_();
    for (var col in bMap) rec[col] = toNum_(biaya[bMap[col]]);
    rec.total = calcBiaya_(biaya);
  }
  if (d.dokumen !== undefined) {
    var dok = d.dokumen || {};
    var dMap = sppdDokumenMap_();
    for (var dcol in dMap) rec[dcol] = nullableText_(dok[dMap[dcol]]);
  }

  /* `kegiatan` pada SPPD bisa string (form SPPD) atau objek (data lama). */
  if (d.kegiatan !== undefined && d.kegiatan !== null) {
    if (typeof d.kegiatan === 'object') {
      rec.kegiatanJson = jsonStringify_(d.kegiatan);
      rec.kegiatan = toStr_(d.kegiatan.kegiatan || '');
    } else {
      rec.kegiatan = toStr_(d.kegiatan);
    }
  }
  return rec;
}

/** Baris sheet → objek SPPD frontend (nested, siap dipakai UI & cetak). */
function sppdFromRecord_(rec) {
  rec = rec || {};
  var obj = {
    id: toStr_(rec.idSPPD),
    noSpt: toStr_(rec.noSpt),
    noSppd: toStr_(rec.noSppd),
    jenisSuratTugas: toStr_(rec.jenisSuratTugas),
    tglSpt: toStr_(rec.tglSpt),
    tglSppd: toStr_(rec.tglSppd),
    maksud: toStr_(rec.maksud),
    tempatBerangkat: toStr_(rec.tempatBerangkat),
    kotaTujuan: toStr_(rec.kotaTujuan),
    instansiTujuan: toStr_(rec.instansiTujuan),
    transport: toStr_(rec.transport) || 'kendaraan-dinas',
    noTransport: toStr_(rec.noTransport),
    ketTransport: toStr_(rec.ketTransport),
    ttdSpt: toStr_(rec.ttdSpt),
    ttdSpd: toStr_(rec.ttdSpd),
    pptk: toStr_(rec.pptk),
    hotel: toStr_(rec.hotel),
    alamatHotel: toStr_(rec.alamatHotel),
    checkin: toStr_(rec.checkin),
    checkout: toStr_(rec.checkout),
    jumlahMalam: toNum_(rec.jumlahMalam),
    noKamar: toStr_(rec.noKamar),
    tglBerangkat: toStr_(rec.tglBerangkat),
    jamBerangkat: toStr_(rec.jamBerangkat),
    tglKembali: toStr_(rec.tglKembali),
    jamKembali: toStr_(rec.jamKembali),
    lamaHari: toInt_(rec.lamaHari, 1),
    pelaksana: [],
    biaya: sppdDefaultBiaya_(),
    total: 0,
    dokumen: sppdDefaultDokumen_(),
    program: toStr_(rec.program),
    kegiatan: '',
    subkegiatan: toStr_(rec.subkegiatan),
    rekening: toStr_(rec.rekening),
    createdBy: toStr_(rec.createdBy),
    createdAt: toStr_(rec.createdAt),
    updatedAt: toStr_(rec.updatedAt)
  };

  /* biaya — semua kunci selalu ada (sama dengan makeSPPD() frontend). */
  var bMap = sppdBiayaMap_();
  for (var col in bMap) {
    if (rec[col] !== undefined && rec[col] !== '') obj.biaya[bMap[col]] = toNum_(rec[col]);
  }
  obj.total = (rec.total === undefined || rec.total === '') ? calcBiaya_(obj.biaya) : toNum_(rec.total);

  /* dokumen — semua kunci selalu ada; kosong = null. */
  var dMap = sppdDokumenMap_();
  for (var dcol in dMap) obj.dokumen[dMap[dcol]] = nullableText_(rec[dcol]);

  /* kegiatan — objek data lama (kegiatanJson) dipakai bila ada. */
  var kegObj = jsonParse_(rec.kegiatanJson, null);
  obj.kegiatan = (kegObj && typeof kegObj === 'object') ? kegObj : toStr_(rec.kegiatan);

  /* status hanya diisi bila pernah diset (badge tampil apa adanya). */
  if (isFilled_(rec.status)) obj.status = toStr_(rec.status);

  return obj;
}

/** Objek pelaksana frontend → baris sheet anak. */
function pelaksanaToRecord_(p, urutan, idSPPD, noSppd) {
  var o = p || {};
  return {
    idPelaksana: toStr_(o.idPelaksana || ''),
    idSPPD: toStr_(idSPPD),
    noSppd: toStr_(noSppd),
    nama: toStr_(o.nama),
    nip: toStr_(o.nip),
    jabatan: toStr_(o.jabatan),
    golongan: toStr_(o.golongan),
    unit: toStr_(o.unit),
    urutan: urutan
  };
}

function pelaksanaFromRecord_(rec) {
  return {
    nama: toStr_(rec.nama),
    nip: toStr_(rec.nip),
    jabatan: toStr_(rec.jabatan),
    golongan: toStr_(rec.golongan),
    unit: toStr_(rec.unit)
  };
}

/* ------------------------------ PEMBACAAN ------------------------------- */

/** Helper bersama: Sheet + definisi SPPD / SPPD_PELAKSANA. */
function sppdSheet_() { return { def: getDef_('SPPD'), sh: openSheet_(getDef_('SPPD')) }; }
function sppdPelaksanaSheet_() {
  var def = getDef_('SPPD_PELAKSANA');
  return { def: def, sh: openSheet_(def) };
}

function pelaksanaDariSheet_(idSPPD) {
  var s = sppdPelaksanaSheet_();
  var kids = readChildObjects_(s.sh, s.def, idSPPD);
  kids.sort(function (a, b) { return toNum_(a.urutan) - toNum_(b.urutan); });
  return kids.map(pelaksanaFromRecord_);
}

/** Daftar SPPD lengkap (termasuk pelaksana) — tanpa pemeriksaan hak akses. */
function fetchSPPDList_() {
  var s = sppdSheet_();
  var rows = readObjects_(s.sh, s.def);
  var sp = sppdPelaksanaSheet_();
  var pelRows = readObjects_(sp.sh, sp.def);
  var byParent = {};
  for (var i = 0; i < pelRows.length; i++) {
    var pid = toStr_(pelRows[i].idSPPD).trim();
    if (!byParent[pid]) byParent[pid] = [];
    byParent[pid].push(pelRows[i]);
  }
  var out = [];
  for (var j = 0; j < rows.length; j++) {
    var obj = sppdFromRecord_(rows[j]);
    obj.__row = rows[j].__row;
    var kids = byParent[obj.id] || [];
    kids.sort(function (a, b) { return toNum_(a.urutan) - toNum_(b.urutan); });
    obj.pelaksana = kids.map(pelaksanaFromRecord_);
    out.push(obj);
  }
  out = sortByCreatedDesc_(out);
  for (var k = 0; k < out.length; k++) delete out[k].__row;
  return out;
}

/** Satu SPPD lengkap berdasarkan ID (null bila tidak ada). */
function fetchSPPDById_(id) {
  var s = sppdSheet_();
  var rec = getObjectById_(s.sh, s.def, id);
  if (!rec) return null;
  var obj = sppdFromRecord_(rec);
  obj.pelaksana = pelaksanaDariSheet_(obj.id);
  return obj;
}

/* ------------------------------- PENULISAN ------------------------------- */

/** Field wajib SPPD — sama dengan validasi pada form SPPD frontend. */
function sppdRequiredFields_() {
  return [
    { field: 'jenisSuratTugas', label: 'Jenis Surat Tugas' },
    { field: 'noSpt', label: 'Nomor SPT' },
    { field: 'noSppd', label: 'Nomor SPD' },
    { field: 'kotaTujuan', label: 'Daerah/Tempat Tujuan' },
    { field: 'transport', label: 'Alat Angkutan' },
    { field: 'tglBerangkat', label: 'Tanggal Pergi' },
    { field: 'tglKembali', label: 'Tanggal Pulang' },
    { field: 'maksud', label: 'Untuk / Maksud' },
    { field: 'ttdSpt', label: 'Penandatangan SPT' },
    { field: 'ttdSpd', label: 'Penandatangan SPD' },
    { field: 'pptk', label: 'PPTK' }
  ];
}

/**
 * Simpan (buat/ubah) satu SPPD beserta pelaksananya.
 * `strict === false` dipakai oleh proses migrasi data lama (tanpa validasi).
 */
function saveSPPDRecord_(data, isCreate, strict) {
  var d = clone_(data) || {};
  if (isCreate) {
    if (!isFilled_(d.id)) d.id = newId_('sppd');
  } else {
    d.id = requireId_(d.id || d.idSPPD, 'ID SPPD');
  }

  if (strict !== false) {
    requireFields_(d, sppdRequiredFields_(), 'SPPD');
    var pelCheck = Array.isArray(d.pelaksana) ? d.pelaksana : [];
    var adaNama = false;
    for (var q = 0; q < pelCheck.length; q++) {
      if (pelCheck[q] && isFilled_(pelCheck[q].nama)) adaNama = true;
    }
    if (!adaNama) throwError_('SPPD: minimal satu pelaksana (dengan nama) wajib diisi.');
  }
  if (isFilled_(d.lamaHari) === false || toNum_(d.lamaHari) <= 0) {
    d.lamaHari = hitungLamaHari_(d.tglBerangkat, d.tglKembali);
  }

  var s = sppdSheet_();
  stampForWrite_(d, isCreate);

  if (isCreate) {
    var recNew = sppdToRecord_(d);
    recNew.idSPPD = d.id;
    if (!isFilled_(recNew.createdAt)) recNew.createdAt = nowIso_();
    appendObject_(s.sh, s.def, recNew);
  } else {
    var exists = getObjectById_(s.sh, s.def, d.id);
    assertFound_(exists, 'SPPD');
    var rec = sppdToRecord_(d);
    rec.idSPPD = d.id;
    if (!isFilled_(rec.createdAt)) rec.createdAt = toStr_(exists.createdAt);
    updateObject_(s.sh, s.def, d.id, rec);
  }

  /* Pelaksana disimpan di tabel anak (relasi idSPPD). */
  var sp = sppdPelaksanaSheet_();
  var list = [];
  var pel = Array.isArray(d.pelaksana) ? d.pelaksana : [];
  for (var i = 0; i < pel.length; i++) {
    if (!pel[i]) continue;
    list.push(pelaksanaToRecord_(pel[i], i + 1, d.id, d.noSppd));
  }
  replaceChildObjects_(sp.sh, sp.def, d.id, list);

  return fetchSPPDById_(d.id);
}

/** Hitung berapa Rampung yang masih mereferensikan SPPD ini (informasi).
 *  OPT: filter di level baris — baris tak cocok dilewati tanpa konversi objek. */
function countRampungBySppd_(idSPPD) {
  var def = getDef_('RAMPUNG');
  var sh = openSheet_(def);
  return readObjects_(sh, def, 'idSPPD', idSPPD).length;
}

/** Hapus SPPD beserta pelaksananya (Rampung terkait TIDAK dihapus). */
function removeSPPD_(id) {
  var key = requireId_(id, 'ID SPPD');
  var s = sppdSheet_();
  var rec = getObjectById_(s.sh, s.def, key);
  assertFound_(rec, 'SPPD');
  var sp = sppdPelaksanaSheet_();
  deleteChildRows_(sp.sh, sp.def, key);
  deleteObject_(s.sh, s.def, key);
  return true;
}
/* ------------------------- API PUBLIK (dipakai Api.gs) ------------------ */

/**
 * getSPPD(token)
 * Mengembalikan seluruh data SPPD (termasuk pelaksana) siap dipakai frontend.
 */
function getSPPD(token) {
  return safeRun_('Data SPPD berhasil dibaca.', function () {
    var sess = requireSession_(token);
    if (!canReadModule_(sess, 'sppd')) throwError_('Anda tidak memiliki hak akses melihat data SPPD.');
    return filterOwnedList_(sess, fetchSPPDList_());
  });
}

/** getSPPDById(token, id) — satu SPPD lengkap berdasarkan ID unik. */
function getSPPDById(token, id) {
  return safeRun_('Data SPPD berhasil dibaca.', function () {
    var sess = requireSession_(token);
    if (!canReadModule_(sess, 'sppd')) throwError_('Anda tidak memiliki hak akses melihat data SPPD.');
    var key = requireId_(id, 'ID SPPD');
    var obj = fetchSPPDById_(key);
    assertFound_(obj, 'SPPD');
    requireRecordAccess_(sess, obj, 'SPPD');
    return obj;
  });
}

/**
 * createSPPD(token, data) — menyimpan SPPD baru.
 * ID unik, total biaya, dan createdAt dibuat backend.
 */
function createSPPD(token, data) {
  return safeRun_('Data SPPD berhasil disimpan.', function () {
    var sess = requireSession_(token);
    requireWrite_(sess, 'sppd');
    /* Pemilik dicatat dari sesi server — nilai dari browser diabaikan. */
    var obj = Object.assign({}, data || {});
    obj.createdBy = sessionOwnerId_(sess);
    return saveSPPDRecord_(obj, true, true);
  });
}

/**
 * updateSPPD(token, data) — memperbarui SPPD.
 * Didukung juga bentuk updateSPPD(token, id, data).
 * Baris lama WAJIB ada (tidak pernah membuat data baru / menghapus data).
 */
function updateSPPD(token, data, dataAlt) {
  return safeRun_('Data SPPD berhasil diperbarui.', function () {
    var sess = requireSession_(token);
    requireWrite_(sess, 'sppd');
    var payload = normalizeUpdatePayload_(data, dataAlt, 'id', 'ID SPPD');
    var lama = fetchSPPDById_(payload.id);
    assertFound_(lama, 'SPPD');
    requireRecordAccess_(sess, lama, 'SPPD');
    /* Kepemilikan tidak boleh diubah dari browser — pakai yang tersimpan. */
    delete payload.createdBy;
    /* BUG YANG DIPERBAIKI: payload separuh (mis. hanya {id, status}) dulu
       diteruskan apa adanya. Karena `pelaksana` tidak dikirim, saveSPPDRecord_
       menganggapnya [] dan replaceChildObjects_ MENGHAPUS SEMUA baris
       SPPD_PELAKSANA milik SPPD ini. Field yang tidak dikirim kini dilengkapi
       dari data lama — sama seperti perlakuan updateRampung(). */
    for (var k in lama) {
      if (payload[k] === undefined) payload[k] = lama[k];
    }
    payload.id = lama.id;
    return saveSPPDRecord_(payload, false, true);
  });
}

/** deleteSPPD(token, id) — menghapus SPPD + pelaksananya. */
function deleteSPPD(token, id) {
  return safeRun_('Data SPPD berhasil dihapus.', function () {
    var sess = requireSession_(token);
    requireWrite_(sess, 'sppd');
    var key = requireId_(id, 'ID SPPD');
    var lama = fetchSPPDById_(key);
    assertFound_(lama, 'SPPD');
    requireRecordAccess_(sess, lama, 'SPPD');
    return removeSPPD_(key);
  });
}

/**
 * Info relasi SPPD → Rampung (dipakai halaman/fitur "Ambil Data dari SPPD").
 */
function getSPPDRelations(token, id) {
  return safeRun_('Relasi SPPD dibaca.', function () {
    var sess = requireSession_(token);
    var key = requireId_(id, 'ID SPPD');
    var sppd = fetchSPPDById_(key);
    assertFound_(sppd, 'SPPD');
    requireRecordAccess_(sess, sppd, 'SPPD');
    var rampung = filterOwnedList_(sess, fetchRampungList_());
    var terkait = [];
    for (var i = 0; i < rampung.length; i++) {
      if (toStr_(rampung[i].sppdId).trim() === key || toStr_(rampung[i].noSppd).trim() === toStr_(sppd.noSppd).trim()) {
        terkait.push({ idRampung: rampung[i].id, noSppd: rampung[i].noSppd, nama: rampung[i].nama, status: toStr_(rampung[i].status) });
      }
    }
    return { sppd: sppd, rampungTerkait: terkait, jumlahRampung: terkait.length };
  });
}

/* ==========================================================================
 * END FILE ASLI: SPPD.gs
 * ========================================================================== */

/* ==========================================================================
 * BEGIN FILE ASLI: Rampung.gs
 * ========================================================================== */
/**
 * ============================================================================
 * SIAP SPPD & RAMPUNG — MODUL RAMPUNG
 * File   : Rampung.gs
 *
 * Struktur data mengikuti field nyata frontend (index.html):
 *   { id, sppdId, noSppd, noSpt, nama, nip, jabatan, tujuan, maksud,
 *     tglBerangkat, tglKembali, totalSPPD, totalRincian, totalRealisasi,
 *     selisih, tahunAnggaran, status,
 *     kegiatan:{subkegiatanId,kodeOrganisasi,kodeProgram,program,kodeKegiatan,
 *               kegiatan,kodeSubkegiatan,subkegiatan,rekening,rekeningUraian,
 *               keperluan,sumberDana,tahunAnggaran},
 *     pejabat:{pa,pptk,bendahara},
 *     rincian:[{nama,nip,jabatan,golongan,unit,noSpd,tglSpd,jumlahHari,
 *               uangHarian,transport,taksi,penginapan,sewaKendaraan,
 *               representasi,lainnya,komponen:[{deskripsi,harga,qty,satuan,
 *               persen,rill,subtotal}],total}],
 *     realisasi:{...}, bukti:{...}, createdAt }
 *
 * Relasi: RAMPUNG.idSPPD → SPPD.idSPPD (Ambil Data dari SPPD),
 *         RAMPUNG_RINCIAN.idRampung → RAMPUNG,
 *         RAMPUNG_KOMPONEN.idRincian → RAMPUNG_RINCIAN.
 * ============================================================================
 */

/* -------------------------------- MAPPING -------------------------------- */

/** Kegiatan (objek) Rampung → kolom Sheet. */
function rampungKegiatanMap_() {
  return {
    kegSubkegiatanId: 'subkegiatanId',
    kegKodeOrganisasi: 'kodeOrganisasi',
    kegKodeProgram: 'kodeProgram',
    kegProgram: 'program',
    kegKodeKegiatan: 'kodeKegiatan',
    kegKegiatan: 'kegiatan',
    kegKodeSubkegiatan: 'kodeSubkegiatan',
    kegSubkegiatan: 'subkegiatan',
    kegRekening: 'rekening',
    kegRekeningUraian: 'rekeningUraian',
    kegKeperluan: 'keperluan',
    kegSumberDana: 'sumberDana',
    kegTahunAnggaran: 'tahunAnggaran'
  };
}

/** Realisasi (objek) Rampung → kolom Sheet. */
function rampungRealisasiMap_() {
  return {
    reaUangHarian: 'uangHarian', reaTransport: 'transport', reaTaksi: 'taksi',
    reaPenginapan: 'penginapan', reaSewaKendaraan: 'sewaKendaraan',
    reaRepresentasi: 'representasi', reaLainnya: 'lainnya'
  };
}

/** Bukti (objek) Rampung → kolom Sheet. */
function rampungBuktiMap_() {
  return {
    buktiTiket: 'tiket', buktiBoarding: 'boarding', buktiHotel: 'hotel',
    buktiTransport: 'transport', buktiLainnya: 'lainnya'
  };
}

function rampungDefaultRealisasi_() {
  return {
    uangHarian: 0, transport: 0, taksi: 0, penginapan: 0,
    sewaKendaraan: 0, representasi: 0, lainnya: 0
  };
}

function rampungDefaultBukti_() {
  return { tiket: null, boarding: null, hotel: null, transport: null, lainnya: null };
}

function rampungDefaultKegiatan_() {
  return {
    subkegiatanId: '', kodeOrganisasi: '', kodeProgram: '', program: '',
    kodeKegiatan: '', kegiatan: '', kodeSubkegiatan: '', subkegiatan: '',
    rekening: '', rekeningUraian: '', keperluan: '', sumberDana: '', tahunAnggaran: ''
  };
}

function rampungDefaultPejabat_() {
  return { pa: '', pptk: '', bendahara: '' };
}

/** Objek Rampung frontend → baris sheet (flat). */
function rampungToRecord_(data) {
  var d = data || {};
  var rec = {};
  var scalars = [
    'noSppd', 'noSpt', 'nama', 'nip', 'jabatan', 'tujuan', 'maksud',
    'tglBerangkat', 'tglKembali', 'lamaHari', 'tahunAnggaran', 'status', 'createdAt', 'updatedAt'
  ];
  for (var i = 0; i < scalars.length; i++) {
    if (d[scalars[i]] !== undefined) rec[scalars[i]] = d[scalars[i]];
  }
  rec.idRampung = toStr_(d.id || d.idRampung || '');
  rec.idSPPD = toStr_(d.sppdId || d.idSPPD || '');
  if (d.createdBy !== undefined) rec.createdBy = toStr_(d.createdBy);

  if (d.totalSPPD !== undefined) rec.totalSPPD = toNum_(d.totalSPPD);

  /* Kegiatan (objek) — kolom terbaca + JSON cadangan untuk field lain. */
  if (d.kegiatan !== undefined && d.kegiatan !== null) {
    var keg = (typeof d.kegiatan === 'object') ? d.kegiatan : {};
    var kMap = rampungKegiatanMap_();
    for (var kcol in kMap) {
      var v = keg[kMap[kcol]];
      rec[kcol] = (v === undefined || v === null) ? '' : v;
    }
    var sisa = {};
    var terpakai = {};
    for (var k2 in kMap) terpakai[kMap[k2]] = true;
    for (var kk in keg) { if (!terpakai[kk]) sisa[kk] = keg[kk]; }
    var adaSisa = false;
    for (var sk in sisa) { adaSisa = true; break; }
    rec.kegiatanJson = adaSisa ? jsonStringify_(sisa) : '';
  }

  /* Pejabat penandatangan. */
  if (d.pejabat !== undefined && d.pejabat !== null) {
    var pej = (typeof d.pejabat === 'object') ? d.pejabat : {};
    rec.pejPA = toStr_(pej.pa);
    rec.pejPPTK = toStr_(pej.pptk);
    rec.pejBendahara = toStr_(pej.bendahara);
    var sisaPej = {};
    for (var pk in pej) {
      if (pk !== 'pa' && pk !== 'pptk' && pk !== 'bendahara') sisaPej[pk] = pej[pk];
    }
    var adaPej = false;
    for (var sp in sisaPej) { adaPej = true; break; }
    rec.pejabatJson = adaPej ? jsonStringify_(sisaPej) : '';
  }

  /* Realisasi (obligasi hanya bila dikirim — sama dengan perilaku frontend). */
  if (d.realisasi !== undefined && d.realisasi !== null) {
    var rl = d.realisasi || {};
    var rMap = rampungRealisasiMap_();
    for (var rcol in rMap) rec[rcol] = toNum_(rl[rMap[rcol]]);
  }

  /* Bukti dokumen. */
  if (d.bukti !== undefined && d.bukti !== null) {
    var bk = d.bukti || {};
    var bMap = rampungBuktiMap_();
    for (var bcol in bMap) rec[bcol] = nullableText_(bk[bMap[bcol]]);
  }
  return rec;
}
/** Baris sheet → objek Rampung frontend (nested, siap UI & cetak). */
function rampungFromRecord_(rec) {
  rec = rec || {};
  var obj = {
    id: toStr_(rec.idRampung),
    noSppd: toStr_(rec.noSppd),
    nama: toStr_(rec.nama),
    nip: toStr_(rec.nip),
    jabatan: toStr_(rec.jabatan),
    tujuan: toStr_(rec.tujuan),
    maksud: toStr_(rec.maksud),
    tglBerangkat: toStr_(rec.tglBerangkat),
    tglKembali: toStr_(rec.tglKembali),
    /* BUG YANG DIPERBAIKI: kolom `lamaHari` ditulis ke Sheet tetapi tidak pernah
       dipetakan kembali ke objek frontend, sehingga field ini hilang saat data
       dibaca ulang (round-trip tulis -> baca tidak simetris). Akibatnya nilai
       yang sudah tersimpan tidak dapat dipakai lagi oleh proses berikutnya dan
       berbeda antara sesi yang baru menyimpan dan sesi yang baru memuat.
       Bila kolom kosong (baris lama), hitung ulang dari tanggal — sama seperti
       yang dilakukan sisi penulisan. */
    lamaHari: isFilled_(rec.lamaHari)
      ? toNum_(rec.lamaHari)
      : (hitungLamaHari_(toStr_(rec.tglBerangkat), toStr_(rec.tglKembali)) || 1),
    totalSPPD: toNum_(rec.totalSPPD),
    totalRincian: toNum_(rec.totalRincian),
    totalRealisasi: toNum_(rec.totalRealisasi),
    selisih: toNum_(rec.selisih),
    tahunAnggaran: toStr_(rec.tahunAnggaran),
    kegiatan: rampungDefaultKegiatan_(),
    pejabat: rampungDefaultPejabat_(),
    rincian: [],
    createdBy: toStr_(rec.createdBy),
    createdAt: toStr_(rec.createdAt),
    updatedAt: toStr_(rec.updatedAt)
  };
  if (isFilled_(rec.idSPPD)) obj.sppdId = toStr_(rec.idSPPD);
  if (isFilled_(rec.noSpt)) obj.noSpt = toStr_(rec.noSpt);
  obj.status = isFilled_(rec.status) ? toStr_(rec.status) : 'Draf';

  /* Kegiatan (objek) dari kolom keg* + sisa field dari kegiatanJson. */
  var kMap = rampungKegiatanMap_();
  for (var kcol in kMap) {
    if (isFilled_(rec[kcol])) obj.kegiatan[kMap[kcol]] = toStr_(rec[kcol]);
  }
  var sisaKeg = jsonParse_(rec.kegiatanJson, null);
  if (sisaKeg && typeof sisaKeg === 'object') {
    for (var sk in sisaKeg) obj.kegiatan[sk] = sisaKeg[sk];
  }
  if (!isFilled_(obj.tahunAnggaran) && isFilled_(obj.kegiatan.tahunAnggaran)) {
    obj.tahunAnggaran = obj.kegiatan.tahunAnggaran;
  }

  /* Pejabat penandatangan. */
  if (isFilled_(rec.pejPA)) obj.pejabat.pa = toStr_(rec.pejPA);
  if (isFilled_(rec.pejPPTK)) obj.pejabat.pptk = toStr_(rec.pejPPTK);
  if (isFilled_(rec.pejBendahara)) obj.pejabat.bendahara = toStr_(rec.pejBendahara);
  var sisaPej = jsonParse_(rec.pejabatJson, null);
  if (sisaPej && typeof sisaPej === 'object') {
    for (var sp in sisaPej) obj.pejabat[sp] = sisaPej[sp];
  }

  /* Realisasi — hanya bila ada isinya (menjaga perilaku frontend). */
  var rMap = rampungRealisasiMap_();
  var adaRealisasi = false;
  for (var rcol in rMap) { if (isFilled_(rec[rcol])) { adaRealisasi = true; break; } }
  if (adaRealisasi) {
    var rl = rampungDefaultRealisasi_();
    for (var rcol2 in rMap) rl[rMap[rcol2]] = toNum_(rec[rcol2]);
    obj.realisasi = rl;
    if (!isFilled_(rec.totalRealisasi)) obj.totalRealisasi = calcRealisasi_(rl);
  } else if (!isFilled_(rec.totalRealisasi)) {
    obj.totalRealisasi = obj.totalRincian;
  }
  if (!isFilled_(rec.selisih)) obj.selisih = obj.totalRealisasi - obj.totalSPPD;

  /* Bukti dokumen. */
  var bMap = rampungBuktiMap_();
  var adaBukti = false;
  for (var bcol in bMap) { if (isFilled_(rec[bcol])) { adaBukti = true; break; } }
  if (adaBukti) {
    var bk = rampungDefaultBukti_();
    for (var bcol2 in bMap) bk[bMap[bcol2]] = nullableText_(rec[bcol2]);
    obj.bukti = bk;
  }

  return obj;
}
/* --------------------------- RINCIAN & KOMPONEN -------------------------- */

/** Field opsional rincian: hanya dimunculkan bila ada isinya. */
function rincianOptionalFields_() {
  return ['golongan', 'unit', 'jumlahHari', 'uangHarian', 'transport', 'taksi',
    'penginapan', 'sewaKendaraan', 'representasi', 'lainnya'];
}

/**
 * Total satu rincian.
 * Bila `total` sudah ada → dipakai apa adanya (tidak dihitung ulang), sama
 * seperti frontend yang menyimpan `total` hasil perhitungan form.
 */
function rincianTotalOf_(p) {
  var o = p || {};
  if (o.total !== undefined && o.total !== null && o.total !== '') return toNum_(o.total);
  var sum = 0;
  var komp = Array.isArray(o.komponen) ? o.komponen : [];
  for (var i = 0; i < komp.length; i++) {
    var c = komp[i] || {};
    if (c.subtotal !== undefined && c.subtotal !== null && c.subtotal !== '') {
      sum += toNum_(c.subtotal);
    } else {
      var persen = (c.persen === undefined || c.persen === null || c.persen === '') ? 100 : toNum_(c.persen);
      sum += toNum_(c.harga) * toNum_(c.qty) * (persen / 100);
    }
  }
  return sum;
}

function rincianToRecord_(p, urutan, idRampung, noSppd) {
  var o = p || {};
  var rec = {
    idRincian: toStr_(o.idRincian || o.id || ''),
    idRampung: toStr_(idRampung),
    noSppd: toStr_(noSppd),
    urutan: urutan,
    nama: toStr_(o.nama),
    nip: toStr_(o.nip),
    jabatan: toStr_(o.jabatan),
    golongan: toStr_(o.golongan),
    unit: toStr_(o.unit),
    noSpd: toStr_(o.noSpd),
    tglSpd: toStr_(o.tglSpd),
    total: rincianTotalOf_(o)
  };
  if (o.jumlahHari !== undefined) rec.jumlahHari = toInt_(o.jumlahHari, 0);
  var numFields = ['uangHarian', 'transport', 'taksi', 'penginapan', 'sewaKendaraan', 'representasi', 'lainnya'];
  for (var i = 0; i < numFields.length; i++) {
    if (o[numFields[i]] !== undefined) rec[numFields[i]] = toNum_(o[numFields[i]]);
  }
  return rec;
}

function rincianFromRecord_(rec) {
  rec = rec || {};
  var obj = {
    id: toStr_(rec.idRincian),
    nama: toStr_(rec.nama),
    nip: toStr_(rec.nip),
    jabatan: toStr_(rec.jabatan),
    noSpd: toStr_(rec.noSpd),
    tglSpd: toStr_(rec.tglSpd),
    komponen: [],
    total: toNum_(rec.total)
  };
  var opts = rincianOptionalFields_();
  for (var i = 0; i < opts.length; i++) {
    if (isFilled_(rec[opts[i]])) {
      obj[opts[i]] = (opts[i] === 'jumlahHari') ? toInt_(rec[opts[i]], 0) : toNum_(rec[opts[i]]);
    }
  }
  return obj;
}

function komponenToRecord_(c, urutan, idRincian, idRampung) {
  var o = c || {};
  var persen = (o.persen === undefined || o.persen === null || o.persen === '') ? 100 : toNum_(o.persen);
  var subtotal = (o.subtotal === undefined || o.subtotal === null || o.subtotal === '')
    ? (toNum_(o.harga) * toNum_(o.qty) * (persen / 100))
    : toNum_(o.subtotal);
  return {
    idKomponen: toStr_(o.idKomponen || o.id || ''),
    idRincian: toStr_(idRincian),
    idRampung: toStr_(idRampung),
    urutan: urutan,
    deskripsi: toStr_(o.deskripsi),
    harga: toNum_(o.harga),
    qty: toNum_(o.qty),
    satuan: toStr_(o.satuan),
    persen: persen,
    rill: !!o.rill,
    subtotal: subtotal
  };
}

/* ------------------------------ PEMBACAAN ------------------------------- */

function rampungSheet_() { return { def: getDef_('RAMPUNG'), sh: openSheet_(getDef_('RAMPUNG')) }; }
function rincianSheet_() { return { def: getDef_('RAMPUNG_RINCIAN'), sh: openSheet_(getDef_('RAMPUNG_RINCIAN')) }; }
function komponenSheet_() { return { def: getDef_('RAMPUNG_KOMPONEN'), sh: openSheet_(getDef_('RAMPUNG_KOMPONEN')) }; }

function sortByUrutan_(arr) {
  (arr || []).sort(function (a, b) { return toNum_(a.urutan) - toNum_(b.urutan); });
  return arr;
}

/** Semua komponen, dikelompokkan berdasarkan idRincian. */
function komponenGrouped_() {
  var k = komponenSheet_();
  var rows = sortByUrutan_(readObjects_(k.sh, k.def));
  var byRincian = {};
  for (var i = 0; i < rows.length; i++) {
    var id = toStr_(rows[i].idRincian).trim();
    if (!id) continue;
    if (!byRincian[id]) byRincian[id] = [];
    byRincian[id].push(komponenFromRecord_(rows[i]));
  }
  return byRincian;
}

/** Semua rincian, dikelompokkan berdasarkan idRampung (termasuk komponen). */
function rincianGrouped_() {
  var r = rincianSheet_();
  var rows = sortByUrutan_(readObjects_(r.sh, r.def));
  var komp = komponenGrouped_();
  var byRampung = {};
  for (var i = 0; i < rows.length; i++) {
    var pid = toStr_(rows[i].idRampung).trim();
    if (!pid) continue;
    var obj = rincianFromRecord_(rows[i]);
    obj.komponen = komp[toStr_(rows[i].idRincian).trim()] || [];
    if (!byRampung[pid]) byRampung[pid] = [];
    byRampung[pid].push(obj);
  }
  return byRampung;
}

/** Daftar Rampung lengkap (dengan rincian + komponen) tanpa cek hak akses. */
function fetchRampungList_() {
  var s = rampungSheet_();
  var rows = readObjects_(s.sh, s.def);
  var rincianByParent = rincianGrouped_();
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var obj = rampungFromRecord_(rows[i]);
    obj.__row = rows[i].__row;
    obj.rincian = rincianByParent[obj.id] || [];
    out.push(obj);
  }
  out = sortByCreatedDesc_(out);
  for (var j = 0; j < out.length; j++) delete out[j].__row;
  return out;
}

/** Satu Rampung lengkap berdasarkan ID unik (null bila tidak ada). */
function fetchRampungById_(id) {
  var s = rampungSheet_();
  var rec = getObjectById_(s.sh, s.def, id);
  if (!rec) return null;
  var obj = rampungFromRecord_(rec);
  var r = rincianSheet_();
  var rows = sortByUrutan_(readChildObjects_(r.sh, r.def, obj.id));
  var komp = komponenGrouped_();
  obj.rincian = [];
  for (var i = 0; i < rows.length; i++) {
    var item = rincianFromRecord_(rows[i]);
    item.komponen = komp[toStr_(rows[i].idRincian).trim()] || [];
    obj.rincian.push(item);
  }
  return obj;
}

/** Cari Rampung berdasarkan idSPPD atau nomor SPPD (untuk cegah duplikat). */
function fetchRampungBySppd_(sppdId, noSppd) {
  var list = fetchRampungList_();
  var key = toStr_(sppdId).trim();
  var no = toStr_(noSppd).trim();
  for (var i = 0; i < list.length; i++) {
    if (key && toStr_(list[i].sppdId).trim() === key) return list[i];
  }
  for (var j = 0; j < list.length; j++) {
    if (no && toStr_(list[j].noSppd).trim() === no) return list[j];
  }
  return null;
}
function komponenFromRecord_(rec) {
  return {
    deskripsi: toStr_(rec.deskripsi),
    harga: toNum_(rec.harga),
    qty: toNum_(rec.qty),
    satuan: toStr_(rec.satuan),
    persen: isFilled_(rec.persen) ? toNum_(rec.persen) : 100,
    rill: toBool_(rec.rill),
    subtotal: toNum_(rec.subtotal)
  };
}
/* ------------------------------ PENULISAN ------------------------------- */

/**
 * Validasi Rampung sisi server.
 * Aturan ketat form (Subkegiatan, Rekening, Tahun Anggaran, Keperluan,
 * Sumber Dana, PA/PPTK/Bendahara) TETAP dijalankan frontend seperti semula.
 * Di server diperiksa hal yang tidak boleh kosong dalam kondisi apa pun,
 * agar alur "Ambil Data dari SPPD" dan data lama tetap dapat disimpan.
 */
function validateRampungDraft_(d, isCreate) {
  if (!d) throwError_('Data Rampung kosong.');
  var rincian = Array.isArray(d.rincian) ? d.rincian : [];
  if (!rincian.length) throwError_('Rampung: minimal satu Rincian Biaya Pegawai wajib diisi.');
  var adaNama = false;
  for (var i = 0; i < rincian.length; i++) {
    if (rincian[i] && isFilled_(rincian[i].nama)) adaNama = true;
    var komp = (rincian[i] && Array.isArray(rincian[i].komponen)) ? rincian[i].komponen : [];
    for (var j = 0; j < komp.length; j++) {
      if (komp[j] && !isFilled_(komp[j].deskripsi) && toNum_(komp[j].harga) > 0) {
        throwError_('Rampung: setiap komponen biaya wajib memiliki Deskripsi.');
      }
    }
  }
  if (!adaNama) throwError_('Rampung: setiap rincian wajib memiliki Nama Pegawai.');
  var keperluan = (d.kegiatan && d.kegiatan.keperluan) || d.maksud;
  if (!isFilled_(keperluan)) throwError_('Rampung: Keperluan / Judul Rincian wajib diisi.');
  if (!isFilled_(d.noSppd) && !(d.rincian[0] && isFilled_(d.rincian[0].noSpd))) {
    throwError_('Rampung: Nomor SPD wajib diisi.');
  }
}

/** Hitung total — MENGIKUTI rumus frontend (saveRampung/updateRampung). */
function computeRampungTotals_(d) {
  var list = Array.isArray(d.rincian) ? d.rincian : [];
  var totalRincian = 0;
  for (var i = 0; i < list.length; i++) totalRincian += rincianTotalOf_(list[i]);
  var hasRealisasi = (d.realisasi !== undefined && d.realisasi !== null);
  var totalRealisasi = hasRealisasi ? calcRealisasi_(d.realisasi) : totalRincian;
  var totalSPPD = toNum_(d.totalSPPD);
  return {
    totalRincian: totalRincian,
    totalRealisasi: totalRealisasi,
    selisih: totalRealisasi - totalSPPD,
    totalSPPD: totalSPPD,
    jumlahRincian: list.length
  };
}

/* deleteChildRowsByField_() versi kanonis ada di Database.gs (dipakai juga
   oleh modul ini). Definisi ganda di sini SENGAJA dihapus: duplikat membuat
   salah satu versi mati-surih (yang menang tergantung urutan muat berkas),
   sehingga perbaikan pada satu salinan tidak berpengaruh. */
/** Simpan ulang seluruh rincian + komponen milik satu Rampung. */
function saveRincianChildren_(idRampung, noSppd, rincianList) {
  var r = rincianSheet_();
  var k = komponenSheet_();

  /* createdAt lama dipertahankan bila ID-nya masih sama. */
  var lamaRincian = {};
  var rowsLama = readChildObjects_(r.sh, r.def, idRampung);
  for (var a = 0; a < rowsLama.length; a++) {
    lamaRincian[toStr_(rowsLama[a].idRincian).trim()] = toStr_(rowsLama[a].createdAt);
  }
  var lamaKomponen = {};
  /* OPT: hanya baris milik Rampung ini yang dibaca/dikonversi. */
  var kompLama = readObjects_(k.sh, k.def, 'idRampung', idRampung);
  for (var b = 0; b < kompLama.length; b++) {
    lamaKomponen[toStr_(kompLama[b].idKomponen).trim()] = toStr_(kompLama[b].createdAt);
  }

  deleteChildRowsByField_(k.sh, 'idRampung', idRampung);
  deleteChildRows_(r.sh, r.def, idRampung);

  /* OPT: rincian + komponen ditulis masing-masing dalam SATU batch setValues
     (lihat appendObjects_) — hasil akhir identik dengan penulisan satu per
     satu, tetapi jauh lebih sedikit panggilan ke Google Sheets. */
  var list = Array.isArray(rincianList) ? rincianList : [];
  var batchRincian = [];
  var batchKomponen = [];
  for (var i = 0; i < list.length; i++) {
    var rec = rincianToRecord_(list[i], i + 1, idRampung, noSppd);
    if (!isFilled_(rec.idRincian)) rec.idRincian = newId_('rci');
    rec.createdAt = lamaRincian[rec.idRincian] || nowIso_();
    rec.updatedAt = nowIso_();
    batchRincian.push(rec);

    var komp = (list[i] && Array.isArray(list[i].komponen)) ? list[i].komponen : [];
    for (var c = 0; c < komp.length; c++) {
      var krec = komponenToRecord_(komp[c], c + 1, rec.idRincian, idRampung);
      if (!isFilled_(krec.idKomponen)) krec.idKomponen = newId_('kmp');
      krec.createdAt = lamaKomponen[krec.idKomponen] || rec.createdAt;
      krec.updatedAt = nowIso_();
      batchKomponen.push(krec);
    }
  }
  appendObjects_(r.sh, r.def, batchRincian);
  appendObjects_(k.sh, k.def, batchKomponen);
  return list.length;
}
/**
 * Simpan (buat/ubah) satu Rampung beserta rincian & komponennya.
 * `strict === false` dipakai proses migrasi data lama.
 */
function saveRampungRecord_(data, isCreate, strict) {
  var d = clone_(data) || {};
  if (isCreate) {
    if (!isFilled_(d.id)) d.id = newId_('rmp');
  } else {
    d.id = requireId_(d.id || d.idRampung, 'ID Rampung');
  }
  if (strict !== false) validateRampungDraft_(d, isCreate);

  var s = rampungSheet_();
  var totals = computeRampungTotals_(d);
  d.totalRincian = totals.totalRincian;
  d.totalRealisasi = totals.totalRealisasi;
  d.selisih = totals.selisih;
  d.totalSPPD = totals.totalSPPD;
  stampForWrite_(d, isCreate);

  var rec = rampungToRecord_(d);
  rec.idRampung = d.id;
  rec.totalRincian = totals.totalRincian;
  rec.totalRealisasi = totals.totalRealisasi;
  rec.selisih = totals.selisih;
  rec.totalSPPD = totals.totalSPPD;
  rec.jumlahRincian = totals.jumlahRincian;
  if (!isFilled_(rec.lamaHari)) rec.lamaHari = hitungLamaHari_(d.tglBerangkat, d.tglKembali);

  if (isCreate) {
    if (!isFilled_(rec.createdAt)) rec.createdAt = nowIso_();
    appendObject_(s.sh, s.def, rec);
  } else {
    var exists = getObjectById_(s.sh, s.def, d.id);
    assertFound_(exists, 'Rampung');
    if (!isFilled_(rec.createdAt)) rec.createdAt = toStr_(exists.createdAt);
    updateObject_(s.sh, s.def, d.id, rec);
  }

  saveRincianChildren_(d.id, toStr_(d.noSppd), Array.isArray(d.rincian) ? d.rincian : []);
  return fetchRampungById_(d.id);
}

/** Hapus Rampung beserta rincian & komponennya. */
function removeRampung_(id) {
  var key = requireId_(id, 'ID Rampung');
  var s = rampungSheet_();
  var rec = getObjectById_(s.sh, s.def, key);
  assertFound_(rec, 'Rampung');
  var k = komponenSheet_();
  deleteChildRowsByField_(k.sh, 'idRampung', key);
  var r = rincianSheet_();
  deleteChildRows_(r.sh, r.def, key);
  deleteObject_(s.sh, s.def, key);
  return true;
}

/**
 * Verifikasi Rampung — siklus status sama dengan frontend:
 * Diajukan → Diverifikasi → Selesai → Diajukan.
 */
function verifikasiRampungRecord_(id) {
  var key = requireId_(id, 'ID Rampung');
  var s = rampungSheet_();
  var rec = getObjectById_(s.sh, s.def, key);
  assertFound_(rec, 'Rampung');
  var status = toStr_(rec.status);
  var next = (status === 'Diajukan') ? 'Diverifikasi'
    : (status === 'Diverifikasi') ? 'Selesai' : 'Diajukan';
  updateObject_(s.sh, s.def, key, { status: next, updatedAt: nowIso_() });
  return fetchRampungById_(key);
}

/* ----------------------- RELASI: SPPD → RAMPUNG ------------------------- */

/** Rentang tanggal naratif — port identik dari docTglRentang() frontend. */
function docTglRentang_(tglPergi, tglPulang, lamanya) {
  var bulanNama = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli',
    'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  if (!tglPergi || !tglPulang) return null;
  var d1 = new Date(String(tglPergi).length === 10 ? tglPergi + 'T00:00:00' : tglPergi);
  var d2 = new Date(String(tglPulang).length === 10 ? tglPulang + 'T00:00:00' : tglPulang);
  if (isNaN(d1.getTime()) || isNaN(d2.getTime())) return null;
  var hari = Number(lamanya) || Math.round((d2.getTime() - d1.getTime()) / 86400000) + 1;
  var lamaStr = hari + ' (' + terbilang_(hari) + ') hari';
  var tglStr = '';
  if (hari === 1) {
    tglStr = 'tanggal ' + d1.getDate() + ' ' + bulanNama[d1.getMonth()] + ' ' + d1.getFullYear();
  } else if (d1.getMonth() === d2.getMonth() && d1.getFullYear() === d2.getFullYear()) {
    tglStr = 'dari tanggal ' + d1.getDate() + ' s/d ' + d2.getDate() + ' ' +
      bulanNama[d1.getMonth()] + ' ' + d1.getFullYear();
  } else {
    tglStr = 'dari tanggal ' + d1.getDate() + ' ' + bulanNama[d1.getMonth()] + ' s/d ' +
      d2.getDate() + ' ' + bulanNama[d2.getMonth()] + ' ' + d2.getFullYear();
  }
  return { hari: hari, lamaStr: lamaStr, tglStr: tglStr };
}

/** Preposisi tujuan — port identik dari tujuanPreposisi() frontend.
 *  "di" untuk hotel/tempat acara, "ke" untuk instansi/lembaga (default). */
function tujuanPreposisi_(tujuan) {
  var dt = String(tujuan || '').trim().toLowerCase();
  if (!dt) return 'ke';
  if (/^(di|ke|dari|pada)\s/.test(dt)) return '';
  if (/(^|[^a-zà-ÿ])(hotel|resort|convention|conference|ballroom|aula|gedung|venue|wisma|losmen|homestay|villa|cottage|motel|penginapan|resto|restoran|caf[eé]|kafe|mall|plaza|stadi(?:on|um)|arena|auditorium|hall)(?![a-zà-ÿ])/.test(dt)) return 'di';
  return 'ke';
}

/** 'Untuk' extra — port identik dari docUntukExtra() frontend. */
function docUntukExtra_(tujuan, tglPergi, tglPulang, lamanya) {
  if (!tujuan || !tglPergi || !tglPulang || !lamanya) return '';
  var preposisi = tujuanPreposisi_(tujuan);
  var rg = docTglRentang_(tglPergi, tglPulang, lamanya);
  if (!rg) return '';
  return (preposisi ? ' ' + preposisi : '') + ' ' + escapeHtml_(String(tujuan).trim()) +
    ' selama ' + rg.lamaStr + ', ' + rg.tglStr;
}

/** Teks "Untuk" dari SPPD — SUMBER TUNGGAL, port dari sptUntukRaw() frontend. */
function sptUntukRaw_(r) {
  var d = r || {};
  return toStr_(d.maksud).trim() +
    docUntukExtra_(d.kotaTujuan, d.tglBerangkat, d.tglKembali, d.lamaHari);
}
/**
 * Bangun RANGKA data Rampung dari satu SPPD.
 * Port logika buildRampungFromSPPD() frontend, sehingga hasil "Ambil Data dari
 * SPPD" identik dan tetap memegang referensi idSPPD (tidak ada duplikasi data).
 */
function buildRampungDraftFromSPPD_(s) {
  var sppd = s || {};
  var pel = (Array.isArray(sppd.pelaksana) && sppd.pelaksana[0]) ? sppd.pelaksana[0] : {};
  var tglSpd = sppd.tglBerangkat || '';
  var programs = fetchProgramList_();
  var prog = null;
  var i;

  if (sppd.kegiatan && sppd.kegiatan.subkegiatanId) {
    for (i = 0; i < programs.length; i++) {
      if (toStr_(programs[i].id) === toStr_(sppd.kegiatan.subkegiatanId)) { prog = programs[i]; break; }
    }
  }
  if (!prog && sppd.kegiatan && sppd.kegiatan.subkegiatan) {
    var nm = toStr_(sppd.kegiatan.subkegiatan).toLowerCase();
    if (nm) {
      for (i = 0; i < programs.length; i++) {
        if (toStr_(programs[i].subkegiatan).toLowerCase() === nm) { prog = programs[i]; break; }
      }
    }
  }

  var keperluan = sptUntukRaw_(sppd);
  var kegiatan = prog ? {
    subkegiatanId: prog.id,
    kodeOrganisasi: toStr_(prog.kodeOrganisasi),
    kodeProgram: toStr_(prog.kodeProgram),
    program: toStr_(prog.program),
    kodeKegiatan: toStr_(prog.kodeKegiatan),
    kegiatan: toStr_(prog.kegiatan),
    kodeSubkegiatan: toStr_(prog.kodeSubkegiatan),
    subkegiatan: toStr_(prog.subkegiatan),
    rekening: toStr_(prog.rekening),
    rekeningUraian: toStr_(prog.rekeningUraian),
    keperluan: keperluan
  } : {
    subkegiatanId: '', kodeOrganisasi: '', kodeProgram: '', program: '',
    kodeKegiatan: '', kegiatan: '', kodeSubkegiatan: '',
    subkegiatan: (sppd.kegiatan && sppd.kegiatan.subkegiatan) || '',
    rekening: '', rekeningUraian: '', keperluan: keperluan
  };

  var biaya = sppd.biaya || {};
  var rincian = [];
  var pelaksana = Array.isArray(sppd.pelaksana) ? sppd.pelaksana : [];
  for (i = 0; i < pelaksana.length; i++) {
    var p = pelaksana[i] || {};
    rincian.push({
      noSpd: sppd.noSppd || '-',
      nama: p.nama || '-',
      nip: p.nip || '-',
      jabatan: p.jabatan || '',
      golongan: p.golongan || '',
      unit: p.unit || '',
      tglSpd: tglSpd,
      jumlahHari: sppd.lamaHari || 1,
      uangHarian: toNum_(biaya.uangHarian),
      transport: toNum_(biaya.transport),
      taksi: toNum_(biaya.taksi),
      penginapan: toNum_(biaya.penginapan),
      sewaKendaraan: toNum_(biaya.sewaKendaraan),
      representasi: toNum_(biaya.representasi),
      lainnya: toNum_(biaya.lainnya),
      komponen: []
    });
  }
  var totalRincian = 0;
  for (i = 0; i < rincian.length; i++) totalRincian += rincianTotalOf_(rincian[i]);

  var totalSPPD = (typeof sppd.total === 'number') ? sppd.total : calcBiaya_(biaya);
  var dok = sppd.dokumen;
  var bukti = dok ? {
    tiket: nullableText_(dok.tiket), boarding: nullableText_(dok.boarding),
    hotel: nullableText_(dok.hotel), transport: nullableText_(dok.transport),
    lainnya: nullableText_(dok.pendukung)
  } : {};

  return {
    id: null,
    sppdId: toStr_(sppd.id),
    noSppd: sppd.noSppd || '-',
    noSpt: sppd.noSpt || '',
    nama: pel.nama || '-',
    nip: pel.nip || '-',
    jabatan: pel.jabatan || '-',
    tujuan: sppd.kotaTujuan || '-',
    maksud: keperluan,
    tglBerangkat: tglSpd || '',
    tglKembali: sppd.tglKembali || '',
    totalSPPD: totalSPPD,
    kegiatan: kegiatan,
    pejabat: { pa: '', pptk: sppd.pptk || '', bendahara: '' },
    rincian: rincian,
    tahunAnggaran: (sppd.kegiatan && sppd.kegiatan.tahunAnggaran) || '',
    totalRincian: totalRincian,
    bukti: bukti,
    status: 'Draf'
  };
}

/**
 * Ambil RANGKA Rampung dari SPPD (belum disimpan) + status duplikat.
 * Dipakai oleh fitur "Ambil Data dari SPPD" agar relasi idSPPD selalu terjaga.
 */
function getRampungDraftFromSppd_(sppdId) {
  var key = requireId_(sppdId, 'ID SPPD');
  var sppd = fetchSPPDById_(key);
  assertFound_(sppd, 'SPPD');
  var existing = fetchRampungBySppd_(sppd.id, sppd.noSppd);
  return {
    sppd: sppd,
    draft: buildRampungDraftFromSPPD_(sppd),
    sudahAdaRampung: !!existing,
    idRampungAda: existing ? existing.id : '',
    rampungAda: existing || null
  };
}

/**
 * Buat Rampung dari SPPD (tanpa duplikat):
 * bila SPPD sudah punya Rampung → sertakan data yang ada, tidak membuat baru.
 */
function createRampungFromSppd_(sppdId, force, ownerId) {
  var info = getRampungDraftFromSppd_(sppdId);
  if (info.sudahAdaRampung && force !== true) {
    return {
      dibuat: false,
      alasan: 'SPPD ' + info.sppd.noSppd + ' sudah memiliki data Rampung.',
      idRampung: info.idRampungAda,
      rampung: info.rampungAda
    };
  }
  if (isFilled_(ownerId)) info.draft.createdBy = toStr_(ownerId);
  var saved = saveRampungRecord_(info.draft, true, false);
  return { dibuat: true, alasan: '', idRampung: saved.id, rampung: saved };
}
/* ------------------------- API PUBLIK (dipakai Api.gs) ------------------ */

/** getRampung(token) — data Rampung milik sesi + rincian + komponen. */
function getRampung(token) {
  return safeRun_('Data Rampung berhasil dibaca.', function () {
    var sess = requireSession_(token);
    if (!canReadModule_(sess, 'rampung')) throwError_('Anda tidak memiliki hak akses melihat data Rampung.');
    return filterOwnedList_(sess, fetchRampungList_());
  });
}

/** getRampungById(token, id) — satu Rampung lengkap. */
function getRampungById(token, id) {
  return safeRun_('Data Rampung berhasil dibaca.', function () {
    var sess = requireSession_(token);
    if (!canReadModule_(sess, 'rampung')) throwError_('Anda tidak memiliki hak akses melihat data Rampung.');
    var obj = fetchRampungById_(requireId_(id, 'ID Rampung'));
    assertFound_(obj, 'Rampung');
    requireRecordAccess_(sess, obj, 'Rampung');
    return obj;
  });
}

/** createRampung(token, data) — simpan Rampung baru + rincian + komponen. */
function createRampung(token, data) {
  return safeRun_('Data Rampung berhasil disimpan.', function () {
    var sess = requireSession_(token);
    requireWrite_(sess, 'rampung');
    /* Pemilik dicatat dari sesi server — nilai dari browser diabaikan. */
    var obj = Object.assign({}, data || {});
    obj.createdBy = sessionOwnerId_(sess);
    return saveRampungRecord_(obj, true, true);
  });
}

/**
 * updateRampung(token, data) — perbarui Rampung.
 * Didukung juga bentuk updateRampung(token, id, data).
 */
function updateRampung(token, data, dataAlt) {
  return safeRun_('Data Rampung berhasil diperbarui.', function () {
    var sess = requireSession_(token);
    requireWrite_(sess, 'rampung');
    var payload = normalizeUpdatePayload_(data, dataAlt, 'id', 'ID Rampung');
    var lama = fetchRampungById_(payload.id);
    assertFound_(lama, 'Rampung');
    requireRecordAccess_(sess, lama, 'Rampung');
    /* Kepemilikan tidak boleh diubah dari browser — pakai yang tersimpan. */
    delete payload.createdBy;
    /* Lengkapi field yang tidak dikirim dengan data lama (update sebagian). */
    for (var k in lama) {
      if (payload[k] === undefined) payload[k] = lama[k];
    }
    payload.id = lama.id;
    return saveRampungRecord_(payload, false, true);
  });
}

/** deleteRampung(token, id) — hapus Rampung + rincian + komponen. */
function deleteRampung(token, id) {
  return safeRun_('Data Rampung berhasil dihapus.', function () {
    var sess = requireSession_(token);
    requireWrite_(sess, 'rampung');
    var key = requireId_(id, 'ID Rampung');
    var lama = fetchRampungById_(key);
    assertFound_(lama, 'Rampung');
    requireRecordAccess_(sess, lama, 'Rampung');
    return removeRampung_(key);
  });
}

/** verifikasiRampung(token, id) — ubah status verifikasi. */
function verifikasiRampung(token, id) {
  return safeRun_('Status Rampung diperbarui.', function () {
    var sess = requireSession_(token);
    requireWrite_(sess, 'rampung');
    var key = requireId_(id, 'ID Rampung');
    var lama = fetchRampungById_(key);
    assertFound_(lama, 'Rampung');
    requireRecordAccess_(sess, lama, 'Rampung');
    return verifikasiRampungRecord_(key);
  });
}

/**
 * getRampungDraftFromSppd(token, sppdId)
 * Data SPPD + rangka Rampung yang siap dipakai fitur "Ambil Data dari SPPD".
 */
function getRampungDraftFromSppd(token, sppdId) {
  return safeRun_('Data SPPD untuk Rampung berhasil dibaca.', function () {
    var sess = requireSession_(token);
    var info = getRampungDraftFromSppd_(sppdId);
    requireRecordAccess_(sess, info.sppd, 'SPPD');
    return info;
  });
}

/**
 * createRampungFromSppd(token, sppdId) — buat Rampung dari SPPD tanpa duplikat.
 */
function createRampungFromSppd(token, sppdId) {
  return safeRun_('Data Rampung dibuat dari SPPD.', function () {
    var sess = requireSession_(token);
    requireWrite_(sess, 'rampung');
    var sppd = fetchSPPDById_(requireId_(sppdId, 'ID SPPD'));
    assertFound_(sppd, 'SPPD');
    requireRecordAccess_(sess, sppd, 'SPPD');
    var res = createRampungFromSppd_(sppdId, false, sessionOwnerId_(sess));
    return res;
  });
}

/* ==========================================================================
 * END FILE ASLI: Rampung.gs
 * ========================================================================== */

/* ==========================================================================
 * BEGIN FILE ASLI: Program.gs
 * ========================================================================== */
/**
 * ============================================================================
 * SIAP SPPD & RAMPUNG — MODUL PROGRAM / KEGIATAN / SUBKEGIATAN
 * File   : Program.gs
 *
 * Struktur data mengikuti frontend (normalizeProgram):
 *   { id, kodeOrganisasi, kodeProgram, kode, nama, kodeKegiatan, kegiatan,
 *     kodeSubkegiatan, subkegiatan, rekeningList:[{kode,uraian}], rekening,
 *     uraianRekening, pagu, status, createdAt }
 *
 * Catatan: field `pagu` sudah tidak dipakai form Program, tetapi NILAI LAMA
 * TETAP DIPERTAHANKAN (dibaca & ditulis kembali apa adanya).
 * ============================================================================
 */

var KODE_ORGANISASI_DEFAULT = '4.01.2.24.0.00.01.00';

/** Kode Program — port progKodeProgram() frontend. */
function progKodeProgram_(p) {
  p = p || {};
  return toStr_(p.kodeProgram || p.kode);
}

/** Kode Kegiatan — port progKodeKegiatan() frontend. */
function progKodeKegiatan_(p) {
  p = p || {};
  return toStr_(p.kodeKegiatan);
}

/** Daftar rekening — port progRekList() frontend. */
function progRekList_(p) {
  p = p || {};
  if (Array.isArray(p.rekeningList) && p.rekeningList.length) return p.rekeningList;
  if (isFilled_(p.rekening)) {
    return [{ kode: toStr_(p.rekening), uraian: toStr_(p.uraianRekening) || 'Belanja Perjalanan Dinas' }];
  }
  return [];
}

/** Normalisasi Program — port normalizeProgram() frontend. */
function normalizeProgram_(p) {
  var o = Object.assign({}, p || {});
  if (!o.kodeOrganisasi) o.kodeOrganisasi = KODE_ORGANISASI_DEFAULT;
  if (!o.kodeProgram) o.kodeProgram = toStr_(o.kode);
  if (!o.kode) o.kode = toStr_(o.kodeProgram);
  if (!o.kodeKegiatan) o.kodeKegiatan = '';
  if (!o.kodeSubkegiatan) o.kodeSubkegiatan = '';
  if (!Array.isArray(o.rekeningList) || !o.rekeningList.length) {
    if (isFilled_(o.rekening)) {
      o.rekeningList = [{ kode: toStr_(o.rekening), uraian: toStr_(o.uraianRekening) || 'Belanja Perjalanan Dinas' }];
    } else {
      o.rekeningList = [];
    }
  }
  if (!isFilled_(o.rekening) && o.rekeningList.length) o.rekening = toStr_(o.rekeningList[0].kode);
  if (!isFilled_(o.uraianRekening) && o.rekeningList.length) o.uraianRekening = toStr_(o.rekeningList[0].uraian);
  return o;
}

/* -------------------------------- MAPPING -------------------------------- */

function programToRecord_(data) {
  var d = normalizeProgram_(data);
  return {
    idProgram: toStr_(d.id || d.idProgram),
    kodeOrganisasi: toStr_(d.kodeOrganisasi),
    kodeProgram: toStr_(d.kodeProgram),
    kode: toStr_(d.kode),
    nama: toStr_(d.nama),
    kodeKegiatan: toStr_(d.kodeKegiatan),
    kegiatan: toStr_(d.kegiatan),
    kodeSubkegiatan: toStr_(d.kodeSubkegiatan),
    subkegiatan: toStr_(d.subkegiatan),
    rekening: toStr_(d.rekening),
    uraianRekening: toStr_(d.uraianRekening),
    pagu: (d.pagu === undefined ? '' : toNum_(d.pagu)),
    status: toStr_(d.status) || 'Aktif',
    createdAt: toStr_(d.createdAt),
    updatedAt: toStr_(d.updatedAt)
  };
}

function programFromRecord_(rec, rekeningList) {
  rec = rec || {};
  var obj = {
    id: toStr_(rec.idProgram),
    kodeOrganisasi: toStr_(rec.kodeOrganisasi),
    kodeProgram: toStr_(rec.kodeProgram),
    kode: toStr_(rec.kode),
    nama: toStr_(rec.nama),
    kodeKegiatan: toStr_(rec.kodeKegiatan),
    kegiatan: toStr_(rec.kegiatan),
    kodeSubkegiatan: toStr_(rec.kodeSubkegiatan),
    subkegiatan: toStr_(rec.subkegiatan),
    rekening: toStr_(rec.rekening),
    uraianRekening: toStr_(rec.uraianRekening),
    rekeningList: Array.isArray(rekeningList) ? rekeningList : [],
    status: toStr_(rec.status) || 'Aktif',
    createdAt: toStr_(rec.createdAt),
    updatedAt: toStr_(rec.updatedAt)
  };
  /* `pagu` lama tetap dibawa bila pernah ada. */
  if (isFilled_(rec.pagu)) obj.pagu = toNum_(rec.pagu);
  return normalizeProgram_(obj);
}

function rekeningToRecord_(r, urutan, idProgram) {
  var o = r || {};
  return {
    idProgramRekening: toStr_(o.idProgramRekening || o.id || ''),
    idProgram: toStr_(idProgram),
    urutan: urutan,
    kode: toStr_(o.kode),
    uraian: toStr_(o.uraian)
  };
}

function rekeningFromRecord_(rec) {
  return { kode: toStr_(rec.kode), uraian: toStr_(rec.uraian) };
}
/* ------------------------------ PEMBACAAN ------------------------------- */

function programSheets_() {
  return {
    def: getDef_('PROGRAM'), sh: openSheet_(getDef_('PROGRAM')),
    defRek: getDef_('PROGRAM_REKENING'), shRek: openSheet_(getDef_('PROGRAM_REKENING'))
  };
}

/** Seluruh Program (termasuk rekeningList) tanpa cek hak akses. */
/** Versi ber-cache (CacheService, TTL pendek): mempercepat pembacaan berulang.
 * Otomatis dibatalkan oleh refCacheBust_ setiap kali data program berubah. */
function fetchProgramList_() {
  return refCacheGet_('program', fetchProgramList_Uncached_);
}

/** Implementasi asli (tanpa cache) — perilaku identik dengan sebelumnya. */
function fetchProgramList_Uncached_() {
  var s = programSheets_();
  var rows = readObjects_(s.sh, s.def);
  var rekRows = sortByUrutan_(readObjects_(s.shRek, s.defRek));
  var byParent = {};
  for (var i = 0; i < rekRows.length; i++) {
    var pid = toStr_(rekRows[i].idProgram).trim();
    if (!pid) continue;
    if (!byParent[pid]) byParent[pid] = [];
    byParent[pid].push(rekeningFromRecord_(rekRows[i]));
  }
  var out = [];
  for (var j = 0; j < rows.length; j++) {
    var obj = programFromRecord_(rows[j], byParent[toStr_(rows[j].idProgram).trim()] || []);
    obj.__row = rows[j].__row;
    out.push(obj);
  }
  out = sortByCreatedDesc_(out);
  for (var k = 0; k < out.length; k++) delete out[k].__row;
  return out;
}

/** Satu Program berdasarkan ID (null bila tidak ada). */
function fetchProgramById_(id) {
  var s = programSheets_();
  var rec = getObjectById_(s.sh, s.def, id);
  if (!rec) return null;
  var rek = sortByUrutan_(readChildObjects_(s.shRek, s.defRek, id));
  return programFromRecord_(rec, rek.map(rekeningFromRecord_));
}

/* ------------------------------ PENULISAN ------------------------------- */

/** Simpan (buat/ubah) satu Program beserta daftar rekeningnya. */
function saveProgramRecord_(data, isCreate, strict) {
  var d = normalizeProgram_(clone_(data) || {});
  if (isCreate) {
    if (!isFilled_(d.id)) d.id = newId_('prg');
  } else {
    d.id = requireId_(d.id || d.idProgram, 'ID Program');
  }
  if (strict !== false) {
    requireFields_(d, [
      { field: 'kodeOrganisasi', label: 'Kode Organisasi' },
      { field: 'kodeProgram', label: 'Kode Program' },
      { field: 'nama', label: 'Program' },
      { field: 'kodeKegiatan', label: 'Kode Kegiatan' },
      { field: 'kegiatan', label: 'Kegiatan' },
      { field: 'kodeSubkegiatan', label: 'Kode Sub Kegiatan' },
      { field: 'subkegiatan', label: 'Sub Kegiatan' }
    ], 'Program');
    if (!d.rekeningList.length) throwError_('Program: minimal 1 Kode Rekening wajib diisi.');
  }

  var s = programSheets_();
  stampForWrite_(d, isCreate);
  var rec = programToRecord_(d);
  rec.idProgram = d.id;

  if (isCreate) {
    if (!isFilled_(rec.createdAt)) rec.createdAt = nowIso_();
    appendObject_(s.sh, s.def, rec);
  } else {
    var exists = getObjectById_(s.sh, s.def, d.id);
    assertFound_(exists, 'Program');
    if (!isFilled_(rec.createdAt)) rec.createdAt = toStr_(exists.createdAt);
    /* `pagu` lama dipertahankan bila data baru tidak memuatnya. */
    if (!isFilled_(rec.pagu) && isFilled_(exists.pagu)) rec.pagu = toNum_(exists.pagu);
    updateObject_(s.sh, s.def, d.id, rec);
  }

  var list = [];
  for (var i = 0; i < d.rekeningList.length; i++) {
    list.push(rekeningToRecord_(d.rekeningList[i], i + 1, d.id));
  }
  replaceChildObjects_(s.shRek, s.defRek, d.id, list);

  /* Cache referensi modul ini dibatalkan — data berubah. */
  refCacheBust_('program');
  return fetchProgramById_(d.id);
}

/** Hapus Program beserta daftar rekeningnya. */
function removeProgram_(id) {
  var key = requireId_(id, 'ID Program');
  var s = programSheets_();
  var rec = getObjectById_(s.sh, s.def, key);
  assertFound_(rec, 'Program');
  deleteChildRows_(s.shRek, s.defRek, key);
  deleteObject_(s.sh, s.def, key);
  /* Cache referensi modul ini dibatalkan — data berubah. */
  refCacheBust_('program');
  return true;
}

/* ------------------------- API PUBLIK (dipakai Api.gs) ------------------ */

function getProgram(token) {
  return safeRun_('Data Program berhasil dibaca.', function () {
    requireSession_(token);
    return fetchProgramList_();
  });
}

function getProgramById(token, id) {
  return safeRun_('Data Program berhasil dibaca.', function () {
    requireSession_(token);
    var obj = fetchProgramById_(requireId_(id, 'ID Program'));
    assertFound_(obj, 'Program');
    return obj;
  });
}

function createProgram(token, data) {
  return safeRun_('Data Program berhasil disimpan.', function () {
    var sess = requireSession_(token);
    requireWrite_(sess, 'program');
    return saveProgramRecord_(data, true, true);
  });
}

function updateProgram(token, data, dataAlt) {
  return safeRun_('Data Program berhasil diperbarui.', function () {
    var sess = requireSession_(token);
    requireWrite_(sess, 'program');
    var payload = normalizeUpdatePayload_(data, dataAlt, 'id', 'ID Program');
    return saveProgramRecord_(payload, false, true);
  });
}

function deleteProgram(token, id) {
  return safeRun_('Data Program berhasil dihapus.', function () {
    var sess = requireSession_(token);
    requireWrite_(sess, 'program');
    return removeProgram_(id);
  });
}

/* ==========================================================================
 * END FILE ASLI: Program.gs
 * ========================================================================== */

/* ==========================================================================
 * BEGIN FILE ASLI: StandarHarga.gs
 * ========================================================================== */
/**
 * ============================================================================
 * SIAP SPPD & RAMPUNG — MODUL STANDAR HARGA
 * File   : StandarHarga.gs
 *
 * Struktur data mengikuti frontend:
 *   { id, daerahTujuan, uangHarian, uangHarianDiklat, uangRepresentasi,
 *     penginapan, uangTransport, taksiSetempat, taksiTanjungBill,
 *     biayaKontribusi, status, createdAt }
 *
 * Catatan: field `tahun` sudah tidak dipakai frontend (sanitizeStandarHarga
 * menghapusnya), sehingga tidak disimpan / tidak dimunculkan kembali.
 * ============================================================================
 */

var STANDAR_HARGA_FIELDS = ['daerahTujuan', 'uangHarian', 'uangHarianDiklat',
  'uangRepresentasi', 'penginapan', 'uangTransport', 'taksiSetempat',
  'taksiTanjungBill', 'biayaKontribusi'];

/** Baris sheet → objek frontend. */
function standarHargaFromRecord_(rec) {
  rec = rec || {};
  var obj = { id: toStr_(rec.idStandarHarga), daerahTujuan: toStr_(rec.daerahTujuan) };
  for (var i = 0; i < STANDAR_HARGA_FIELDS.length; i++) {
    var f = STANDAR_HARGA_FIELDS[i];
    if (f === 'daerahTujuan') continue;
    obj[f] = toNum_(rec[f]);
  }
  obj.status = isFilled_(rec.status) ? toStr_(rec.status) : 'Aktif';
  obj.createdAt = toStr_(rec.createdAt);
  obj.updatedAt = toStr_(rec.updatedAt);
  return obj;
}

/** Objek frontend → baris sheet. */
function standarHargaToRecord_(data) {
  var d = clone_(data) || {};
  delete d.tahun;
  var rec = {
    idStandarHarga: toStr_(d.id || d.idStandarHarga),
    daerahTujuan: toStr_(d.daerahTujuan),
    status: toStr_(d.status) || 'Aktif',
    createdAt: toStr_(d.createdAt),
    updatedAt: toStr_(d.updatedAt)
  };
  for (var i = 0; i < STANDAR_HARGA_FIELDS.length; i++) {
    var f = STANDAR_HARGA_FIELDS[i];
    if (f === 'daerahTujuan') continue;
    rec[f] = toNum_(d[f]);
  }
  return rec;
}

function standarHargaSheet_() {
  return { def: getDef_('STANDAR_HARGA'), sh: openSheet_(getDef_('STANDAR_HARGA')) };
}

/** Seluruh Standar Harga (tanpa cek hak akses). */
/** Versi ber-cache (CacheService, TTL pendek): mempercepat pembacaan berulang.
 * Otomatis dibatalkan oleh refCacheBust_ setiap kali data standarharga berubah. */
function fetchStandarHargaList_() {
  return refCacheGet_('standarharga', fetchStandarHargaList_Uncached_);
}

/** Implementasi asli (tanpa cache) — perilaku identik dengan sebelumnya. */
function fetchStandarHargaList_Uncached_() {
  var s = standarHargaSheet_();
  var rows = readObjects_(s.sh, s.def);
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var obj = standarHargaFromRecord_(rows[i]);
    obj.__row = rows[i].__row;
    out.push(obj);
  }
  out = sortByCreatedDesc_(out);
  for (var j = 0; j < out.length; j++) delete out[j].__row;
  return out;
}

function fetchStandarHargaById_(id) {
  var s = standarHargaSheet_();
  var rec = getObjectById_(s.sh, s.def, id);
  return rec ? standarHargaFromRecord_(rec) : null;
}

/**
 * Simpan (buat/ubah) Standar Harga.
 * Saat mengubah, field yang tidak dikirim akan memakai nilai lama
 * (perilaku sama dengan updateStandarHarga() frontend).
 */
function saveStandarHargaRecord_(data, isCreate, strict) {
  var d = clone_(data) || {};
  if (isCreate) {
    if (!isFilled_(d.id)) d.id = newId_('sh');
  } else {
    d.id = requireId_(d.id || d.idStandarHarga, 'ID Standar Harga');
  }
  if (strict !== false) {
    requireFields_(d, [{ field: 'daerahTujuan', label: 'Daerah Tujuan' }], 'Standar Harga');
  }

  var s = standarHargaSheet_();
  var lama = isCreate ? null : getObjectById_(s.sh, s.def, d.id);
  if (!isCreate) assertFound_(lama, 'Standar Harga');

  var gabung = {};
  if (lama) {
    for (var k in lama) { if (k !== '__row') gabung[k] = lama[k]; }
    gabung.idStandarHarga = d.id;
    gabung.createdAt = toStr_(lama.createdAt);
  }
  for (var key in d) { if (d[key] !== undefined) gabung[key] = d[key]; }
  gabung.id = d.id;
  stampForWrite_(gabung, isCreate);

  var rec = standarHargaToRecord_(gabung);
  rec.idStandarHarga = d.id;
  if (isCreate) {
    if (!isFilled_(rec.createdAt)) rec.createdAt = nowIso_();
    appendObject_(s.sh, s.def, rec);
  } else {
    updateObject_(s.sh, s.def, d.id, rec);
  }
  /* Cache referensi modul ini dibatalkan — data berubah. */
  refCacheBust_('standarharga');
  return fetchStandarHargaById_(d.id);
}

function removeStandarHarga_(id) {
  var key = requireId_(id, 'ID Standar Harga');
  var s = standarHargaSheet_();
  assertFound_(getObjectById_(s.sh, s.def, key), 'Standar Harga');
  deleteObject_(s.sh, s.def, key);
  /* Cache referensi modul ini dibatalkan — data berubah. */
  refCacheBust_('standarharga');
  return true;
}

/* ------------------------- API PUBLIK (dipakai Api.gs) ------------------ */

function getStandarHarga(token) {
  return safeRun_('Data Standar Harga berhasil dibaca.', function () {
    requireSession_(token);
    return fetchStandarHargaList_();
  });
}

function getStandarHargaById(token, id) {
  return safeRun_('Data Standar Harga berhasil dibaca.', function () {
    requireSession_(token);
    var obj = fetchStandarHargaById_(requireId_(id, 'ID Standar Harga'));
    assertFound_(obj, 'Standar Harga');
    return obj;
  });
}

function createStandarHarga(token, data) {
  return safeRun_('Data Standar Harga berhasil disimpan.', function () {
    var sess = requireSession_(token);
    requireWrite_(sess, 'standar-harga');
    return saveStandarHargaRecord_(data, true, true);
  });
}

function updateStandarHarga(token, data, dataAlt) {
  return safeRun_('Data Standar Harga berhasil diperbarui.', function () {
    var sess = requireSession_(token);
    requireWrite_(sess, 'standar-harga');
    var payload = normalizeUpdatePayload_(data, dataAlt, 'id', 'ID Standar Harga');
    return saveStandarHargaRecord_(payload, false, true);
  });
}

function deleteStandarHarga(token, id) {
  return safeRun_('Data Standar Harga berhasil dihapus.', function () {
    var sess = requireSession_(token);
    requireWrite_(sess, 'standar-harga');
    return removeStandarHarga_(id);
  });
}

/* ==========================================================================
 * END FILE ASLI: StandarHarga.gs
 * ========================================================================== */

/* ==========================================================================
 * BEGIN FILE ASLI: Pejabat.gs
 * ========================================================================== */
/** SIAP SPPD & RAMPUNG — MODUL PEJABAT PENANDATANGAN (Pejabat.gs)
 * Struktur frontend: { id, nama, jabatan, pangkat, nip, peran:[...], status }
 * `peran` adalah ARRAY (disimpan JSON). `nip` + `status` WAJIB dipertahankan
 * karena dipakai dropdown pegawai & cetak (pejabatOptions/makeSig).
 */
function peranResmiPejabat_() {
  return ['Pengguna Anggaran', 'Kuasa Pengguna Anggaran', 'PPTK',
    'Bendahara', 'Penandatangan SPT', 'Penandatangan SPD'];
}
function normalizePeranPejabat_(peran) {
  var arr = [];
  if (Array.isArray(peran)) arr = peran;
  else if (isFilled_(peran)) {
    var parsed = jsonParse_(peran, null);
    if (Array.isArray(parsed)) arr = parsed;
    else arr = splitList_(peran);
  }
  var out = [];
  for (var i = 0; i < arr.length; i++) {
    var v = toStr_(arr[i]).trim();
    if (v && out.indexOf(v) === -1) out.push(v);
  }
  return out;
}
function pejabatFromRecord_(rec) {
  rec = rec || {};
  return {
    id: toStr_(rec.idPejabat), nama: toStr_(rec.nama),
    jabatan: toStr_(rec.jabatan), pangkat: toStr_(rec.pangkat),
    nip: toStr_(rec.nip),
    /* `unitKerja` & `kotaKecamatan` ada di Sheet + form Pejabat, tetapi
       sebelumnya TIDAK dibaca di sini maupun ditulis di pejabatToRecord_() —
       kedua isian pengguna hilang diam-diam dan selalu tampil kosong. */
    unitKerja: toStr_(rec.unitKerja), kotaKecamatan: toStr_(rec.kotaKecamatan),
    peran: normalizePeranPejabat_(rec.peran),
    status: isFilled_(rec.status) ? toStr_(rec.status) : 'Aktif',
    createdAt: toStr_(rec.createdAt), updatedAt: toStr_(rec.updatedAt)
  };
}
function pejabatToRecord_(data) {
  var d = clone_(data) || {};
  return {
    idPejabat: toStr_(d.id || d.idPejabat), nama: toStr_(d.nama),
    jabatan: toStr_(d.jabatan), pangkat: toStr_(d.pangkat),
    nip: toStr_(d.nip),
    /* Wajib dipetakan — lihat catatan pada pejabatFromRecord_(). */
    unitKerja: toStr_(d.unitKerja), kotaKecamatan: toStr_(d.kotaKecamatan),
    peran: jsonStringify_(normalizePeranPejabat_(d.peran)),
    status: toStr_(d.status) || 'Aktif',
    createdAt: toStr_(d.createdAt), updatedAt: toStr_(d.updatedAt)
  };
}
function pejabatSheet_() {
  var def = getDef_('PEJABAT');
  return { def: def, sh: openSheet_(def) };
}
/** Versi ber-cache (CacheService, TTL pendek): mempercepat pembacaan berulang.
 * Otomatis dibatalkan oleh refCacheBust_ setiap kali data pejabat berubah. */
function fetchPejabatList_() {
  return refCacheGet_('pejabat', fetchPejabatList_Uncached_);
}

/** Implementasi asli (tanpa cache) — perilaku identik dengan sebelumnya. */
function fetchPejabatList_Uncached_() {
  var s = pejabatSheet_();
  var rows = readObjects_(s.sh, s.def);
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var obj = pejabatFromRecord_(rows[i]);
    obj.__row = rows[i].__row;
    out.push(obj);
  }
  out = sortByCreatedDesc_(out);
  for (var j = 0; j < out.length; j++) delete out[j].__row;
  return out;
}
function fetchPejabatById_(id) {
  var s = pejabatSheet_();
  var rec = getObjectById_(s.sh, s.def, id);
  return rec ? pejabatFromRecord_(rec) : null;
}
/* Simpan Pejabat (upsert — sama seperti updatePejabat() frontend). */
function savePejabatRecord_(data, strict) {
  var d = clone_(data) || {};
  var s = pejabatSheet_();
  var id = toStr_(d.id || d.idPejabat);
  var lama = id ? getObjectById_(s.sh, s.def, id) : null;
  if (!lama) id = id ? id : newId_('pj');
  d.id = id; d.idPejabat = id;
  d.peran = normalizePeranPejabat_(d.peran);
  if (strict !== false) {
    requireFields_(d, [
      { field: 'nama', label: 'Nama Pejabat' },
      { field: 'nip', label: 'NIP Pejabat' }
    ], 'Pejabat');
  }
  var gabung = {};
  if (lama) { for (var k in lama) { if (k !== '__row') gabung[k] = lama[k]; } }
  for (var key in d) { if (d[key] !== undefined) gabung[key] = d[key]; }
  gabung.id = id; gabung.idPejabat = id;
  stampForWrite_(gabung, !lama);
  var rec = pejabatToRecord_(gabung);
  rec.idPejabat = id;
  if (!lama) {
    if (!isFilled_(rec.createdAt)) rec.createdAt = nowIso_();
    appendObject_(s.sh, s.def, rec);
  } else {
    if (!isFilled_(rec.createdAt)) rec.createdAt = toStr_(lama.createdAt);
    updateObject_(s.sh, s.def, id, rec);
  }
  /* Cache referensi modul ini dibatalkan — data berubah. */
  refCacheBust_('pejabat');
  return fetchPejabatById_(id);
}
function removePejabat_(id) {
  var key = requireId_(id, 'ID Pejabat');
  var s = pejabatSheet_();
  assertFound_(getObjectById_(s.sh, s.def, key), 'Pejabat');
  deleteObject_(s.sh, s.def, key);
  /* Cache referensi modul ini dibatalkan — data berubah. */
  refCacheBust_('pejabat');
  return true;
}
/* ---- API PUBLIK (dipakai Api.gs) ---- */
function getPejabat(token) {
  return safeRun_('Data Pejabat berhasil dibaca.', function () {
    requireSession_(token);
    return fetchPejabatList_();
  });
}
function getPejabatById(token, id) {
  return safeRun_('Data Pejabat berhasil dibaca.', function () {
    requireSession_(token);
    var obj = fetchPejabatById_(requireId_(id, 'ID Pejabat'));
    assertFound_(obj, 'Pejabat');
    return obj;
  });
}
function updatePejabat(token, data, dataAlt) {
  return safeRun_('Data Pejabat berhasil disimpan.', function () {
    var sess = requireSession_(token);
    requireWrite_(sess, 'pengaturan');
    var payload;
    if (typeof data === 'string' || typeof data === 'number') {
      payload = Object.assign({}, dataAlt || {});
      payload.id = String(data);
    } else { payload = Object.assign({}, data || {}); }
    if (!isFilled_(payload.id) && isFilled_(payload.idPejabat)) payload.id = payload.idPejabat;
    return savePejabatRecord_(payload, true);
  });
}
function deletePejabat(token, id) {
  return safeRun_('Data Pejabat berhasil dihapus.', function () {
    var sess = requireSession_(token);
    requireWrite_(sess, 'pengaturan');
    return removePejabat_(id);
  });
}

/* ==========================================================================
 * END FILE ASLI: Pejabat.gs
 * ========================================================================== */

/* ==========================================================================
 * BEGIN FILE ASLI: Pengaturan.gs
 * ========================================================================== */
/**
 * ============================================================================
 * SIAP SPPD & RAMPUNG â€” MODUL PENGATURAN
 * File   : Pengaturan.gs
 * Isi    : A. Pengaturan Kop, B. Identitas Instansi (INSTANSI),
 *          C. Konfigurasi Cetak (CETAK_CONFIG), D. Logo instansi (DriveApp).
 *
 * Struktur data Kop mengikuti frontend:
 *   { id, jenisSuratTugas, namaInstansi, namaUnitKerja, alamat, telepon,
 *     laman, default, notaDinas }
 *
 * B. Instansi     : identitas instansi dari frontend (index.html hal. Pengaturan).
 * C. Cetak Config : pengaturan cetak dari frontend `DB.cetakConfig`.
 * D. Logo         : logo instansi disimpan ke Google Drive (DriveApp) agar tidak
 *                   memenuhi sel Google Sheets.
 * ============================================================================
 */

/* --------------------------------- KOP ---------------------------------- */

function kopSheet_() { return { def: getDef_('KOP'), sh: openSheet_(getDef_('KOP')) }; }

function kopFromRecord_(rec) {
  rec = rec || {};
  var obj = {
    id: toStr_(rec.idKop),
    jenisSuratTugas: toStr_(rec.jenisSuratTugas),
    namaInstansi: toStr_(rec.namaInstansi),
    namaUnitKerja: toStr_(rec.namaUnitKerja),
    alamat: toStr_(rec.alamat),
    telepon: toStr_(rec.telepon),
    laman: toStr_(rec.laman),
    /* `kota` dipakai dokumen cetak pada bagian "Dikeluarkan di [Kota], tanggal".
       Sebelumnya kolom ini ada di Sheet + form, tetapi TIDAK dibaca di sini
       maupun ditulis di kopToRecord_() — isian pengguna hilang diam-diam. */
    kota: toStr_(rec.kota),
    default: toBool_(rec.default),
    notaDinas: toBool_(rec.notaDinas)
  };
  if (isFilled_(rec.logo)) obj.logo = toStr_(rec.logo);
  obj.createdAt = toStr_(rec.createdAt);
  obj.updatedAt = toStr_(rec.updatedAt);
  return obj;
}

function kopToRecord_(data) {
  var d = clone_(data) || {};
  return {
    idKop: toStr_(d.id || d.idKop),
    jenisSuratTugas: toStr_(d.jenisSuratTugas),
    namaInstansi: toStr_(d.namaInstansi),
    namaUnitKerja: toStr_(d.namaUnitKerja),
    alamat: toStr_(d.alamat),
    telepon: isFilled_(d.telepon) ? toStr_(d.telepon) : '-',
    laman: isFilled_(d.laman) ? toStr_(d.laman) : '-',
    /* Wajib dipetakan — lihat catatan pada kopFromRecord_(). */
    kota: toStr_(d.kota),
    default: !!d.default,
    notaDinas: !!d.notaDinas,
    logo: isFilled_(d.logo) ? toStr_(d.logo) : '',
    createdAt: toStr_(d.createdAt),
    updatedAt: toStr_(d.updatedAt)
  };
}

/** Seluruh Pengaturan Kop (tanpa cek hak akses). */
/** Versi ber-cache (CacheService, TTL pendek): mempercepat pembacaan berulang.
 * Otomatis dibatalkan oleh refCacheBust_ setiap kali data kop berubah. */
function fetchKopList_() {
  return refCacheGet_('kop', fetchKopList_Uncached_);
}

/** Implementasi asli (tanpa cache) — perilaku identik dengan sebelumnya. */
function fetchKopList_Uncached_() {
  var s = kopSheet_();
  var rows = readObjects_(s.sh, s.def);
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var obj = kopFromRecord_(rows[i]);
    obj.__row = rows[i].__row;
    out.push(obj);
  }
  out = sortByCreatedDesc_(out);
  for (var j = 0; j < out.length; j++) delete out[j].__row;
  return out;
}

function fetchKopById_(id) {
  var s = kopSheet_();
  var rec = getObjectById_(s.sh, s.def, id);
  return rec ? kopFromRecord_(rec) : null;
}

/**
 * Simpan Kop (membuat bila ID belum ada â€” sama seperti updateKop() frontend).
 * Bila `default` dicentang â†’ Kop lain otomatis dijadikan non-default.
 */
function saveKopRecord_(data, strict) {
  var d = clone_(data) || {};
  var s = kopSheet_();
  var id = toStr_(d.id || d.idKop);
  var lama = id ? getObjectById_(s.sh, s.def, id) : null;
  if (!lama) id = id ? id : newId_('kop');
  d.id = id;
  d.idKop = id;

  if (strict !== false) {
    requireFields_(d, [
      { field: 'jenisSuratTugas', label: 'Jenis Surat Tugas' },
      { field: 'namaInstansi', label: 'Nama Instansi' },
      { field: 'namaUnitKerja', label: 'Nama Unit Kerja' },
      { field: 'alamat', label: 'Alamat' }
    ], 'Pengaturan Kop');
  }

  var gabung = {};
  if (lama) {
    for (var k in lama) { if (k !== '__row') gabung[k] = lama[k]; }
  }
  for (var key in d) { if (d[key] !== undefined) gabung[key] = d[key]; }
  gabung.id = id; gabung.idKop = id;
  stampForWrite_(gabung, !lama);

  var rec = kopToRecord_(gabung);
  rec.idKop = id;
  if (!lama) {
    if (!isFilled_(rec.createdAt)) rec.createdAt = nowIso_();
    appendObject_(s.sh, s.def, rec);
  } else {
    if (!isFilled_(rec.createdAt)) rec.createdAt = toStr_(lama.createdAt);
    updateObject_(s.sh, s.def, id, rec);
  }

  /* Hanya satu Kop yang boleh menjadi Default. */
  if (rec.default) {
    var semua = readObjects_(s.sh, s.def);
    for (var i = 0; i < semua.length; i++) {
      var oid = toStr_(semua[i].idKop);
      if (oid === id) continue;
      if (toBool_(semua[i].default)) {
        updateObject_(s.sh, s.def, oid, { default: false, updatedAt: nowIso_() });
      }
    }
  }
  /* Cache referensi dibatalkan lebih dulu, BARU hasil dibaca ulang.
     BUG YANG DIPERBAIKI: dengan urutan lama, fetch di bawah bisa menyajikan
     entri lama yang masih tertinggal di cache (mis. put() gagal atau payload
     lewat batas 98KB), sehingga respons "berhasil disimpan" berisi data
     sebelum perubahan. */
  refCacheBust_('kop');
  return fetchKopById_(id);
}

function removeKop_(id) {
  var key = requireId_(id, 'ID Kop');
  var s = kopSheet_();
  assertFound_(getObjectById_(s.sh, s.def, key), 'Pengaturan Kop');
  deleteObject_(s.sh, s.def, key);
  /* Cache referensi modul ini dibatalkan — data berubah. */
  refCacheBust_('kop');
  return true;
}

/* ---- API PUBLIK KOP (dipakai Api.gs) ---- */

function getKop(token) {
  return safeRun_('Data Pengaturan Kop berhasil dibaca.', function () {
    requireSession_(token);
    return fetchKopList_();
  });
}

function getKopById(token, id) {
  return safeRun_('Data Pengaturan Kop berhasil dibaca.', function () {
    requireSession_(token);
    var obj = fetchKopById_(requireId_(id, 'ID Kop'));
    assertFound_(obj, 'Pengaturan Kop');
    return obj;
  });
}

function updateKop(token, data, dataAlt) {
  return safeRun_('Pengaturan Kop berhasil disimpan.', function () {
    var sess = requireSession_(token);
    requireWrite_(sess, 'pengaturan');
    var payload;
    if (typeof data === 'string' || typeof data === 'number') {
      payload = Object.assign({}, dataAlt || {});
      payload.id = String(data);
    } else { payload = Object.assign({}, data || {}); }
    if (!isFilled_(payload.id) && isFilled_(payload.idKop)) payload.id = payload.idKop;
    return saveKopRecord_(payload, true);
  });
}

function deleteKop(token, id) {
  return safeRun_('Pengaturan Kop berhasil dihapus.', function () {
    var sess = requireSession_(token);
    requireWrite_(sess, 'pengaturan');
    return removeKop_(id);
  });
}

/* --------------------------------- INSTANSI -------------------------------- */

function instansiSheet_() { return { def: getDef_('INSTANSI'), sh: openSheet_(getDef_('INSTANSI')) }; }

/** Versi ber-cache (CacheService, TTL pendek): mempercepat pembacaan berulang.
 * Otomatis dibatalkan oleh refCacheBust_ setiap kali data instansi berubah. */
function fetchInstansiRecord_() {
  return refCacheGet_('instansi', fetchInstansiRecord_Uncached_);
}

/** Implementasi asli (tanpa cache) — perilaku identik dengan sebelumnya. */
function fetchInstansiRecord_Uncached_() {
  var s = instansiSheet_();
  var rows = readObjects_(s.sh, s.def);
  for (var i = 0; i < rows.length; i++) {
    if (toStr_(rows[i].kunci).trim() === SINGLETON_KEY_INSTANSI) return rows[i];
  }
  return null;
}

function instansiFromRecord_(rec) {
  rec = rec || {};
  return {
    nama: toStr_(rec.nama), organisasi: toStr_(rec.organisasi),
    alamat: toStr_(rec.alamat), kabkota: toStr_(rec.kabkota),
    provinsi: toStr_(rec.provinsi), telepon: toStr_(rec.telepon),
    laman: toStr_(rec.laman),
    logo: isFilled_(rec.logo) ? toStr_(rec.logo) : '',
    namaKepala: toStr_(rec.namaKepala), nipKepala: toStr_(rec.nipKepala),
    jabatanKepala: toStr_(rec.jabatanKepala)
  };
}

function saveInstansiRecord_(data) {
  var s = instansiSheet_();
  var existing = fetchInstansiRecord_();
  var patch = Object.assign({}, data || {});
  delete patch.kunci; delete patch.extraJson;
  delete patch.createdAt; delete patch.updatedAt;
  if (!existing) {
    var obj = Object.assign({}, patch);
    obj.kunci = SINGLETON_KEY_INSTANSI;
    obj.updatedAt = nowIso_();
    appendObject_(s.sh, s.def, obj);
  } else {
    var merged = Object.assign({}, existing, patch);
    merged.kunci = SINGLETON_KEY_INSTANSI;
    merged.updatedAt = nowIso_();
    updateObject_(s.sh, s.def, SINGLETON_KEY_INSTANSI, merged);
  }
  /* Cache dibatalkan lebih dulu, baru hasil dibaca ulang — melihat catatan
     pada saveKopRecord_ (kembalikan data segar, bukan entri cache lama). */
  refCacheBust_('instansi');
  var rec = fetchInstansiRecord_();
  return rec ? instansiFromRecord_(rec) : instansiFromRecord_({});
}

/* ------------------------------- CETAK CONFIG -------------------------------- */

function cetakConfigSheet_() { return { def: getDef_('CETAK_CONFIG'), sh: openSheet_(getDef_('CETAK_CONFIG')) }; }

/** Versi ber-cache (CacheService, TTL pendek): mempercepat pembacaan berulang.
 * Otomatis dibatalkan oleh refCacheBust_ setiap kali data cetakconfig berubah. */
function fetchCetakConfigRecord_() {
  return refCacheGet_('cetakconfig', fetchCetakConfigRecord_Uncached_);
}

/** Implementasi asli (tanpa cache) — perilaku identik dengan sebelumnya. */
function fetchCetakConfigRecord_Uncached_() {
  var s = cetakConfigSheet_();
  var rows = readObjects_(s.sh, s.def);
  for (var i = 0; i < rows.length; i++) {
    if (toStr_(rows[i].kunci).trim() === SINGLETON_KEY_CETAK) return rows[i];
  }
  return null;
}

function cetakConfigFromRecord_(rec) {
  rec = rec || {};
  return {
    kertas: toStr_(rec.kertas) || 'A4',
    orientasi: toStr_(rec.orientasi) || 'potrait',
    margin: toNum_(rec.margin) || 25,
    font: toNum_(rec.font) || 12,
    skala: toNum_(rec.skala) || 100,
    header: toBool_(rec.header), footer: toBool_(rec.footer),
    logo: toBool_(rec.logo), ttd: toBool_(rec.ttd),

    jarakTtd: toNum_(rec.jarakTtd) || 2,
    posisiTtd: toStr_(rec.posisiTtd) || 'kanan',
    kop: toBool_(rec.kop)
  };
}

/* ------------------------- API PUBLIK (INSTANSI) ------------------------ */

function getInstansi(token) {
  return safeRun_('Data identitas instansi berhasil dibaca.', function () {
    requireSession_(token);
    var rec = fetchInstansiRecord_();
    if (!rec) return instansiFromRecord_({ nama: '', organisasi: '', alamat: '', kabkota: '', provinsi: '', telepon: '', laman: '', logo: '', namaKepala: '', nipKepala: '', jabatanKepala: '' });
    return instansiFromRecord_(rec);
  });
}

function updateInstansi(token, data) {
  return safeRun_('Data identitas instansi berhasil diperbarui.', function () {
    var sess = requireSession_(token);
    requireWrite_(sess, 'pengaturan');
    return saveInstansiRecord_(data);
  });
}

/* ------------------------- API PUBLIK (CETAK_CONFIG) ------------------- */

function getCetakConfig(token) {
  return safeRun_('Konfigurasi cetak berhasil dibaca.', function () {
    requireSession_(token);
    var rec = fetchCetakConfigRecord_();
    if (!rec) return cetakConfigFromRecord_({ kertas: 'A4', orientasi: 'potrait', margin: 25, font: 12, skala: 100, header: true, footer: true, logo: true, ttd: true, jarakTtd: 2, posisiTtd: 'kanan', kop: true });
    return cetakConfigFromRecord_(rec);
  });
}

function updateCetakConfig(token, data) {
  return safeRun_('Konfigurasi cetak berhasil diperbarui.', function () {
    var sess = requireSession_(token);
    requireWrite_(sess, 'pengaturan');
    return saveCetakConfigRecord_(data);
  });
}

function saveCetakConfigRecord_(data) {
  var s = cetakConfigSheet_();
  var existing = fetchCetakConfigRecord_();
  var patch = Object.assign({}, data || {});
  delete patch.kunci; delete patch.extraJson;
  delete patch.createdAt; delete patch.updatedAt;
  if (!existing) {
    var obj = Object.assign({}, patch);
    obj.kunci = SINGLETON_KEY_CETAK;
    obj.updatedAt = nowIso_();
    appendObject_(s.sh, s.def, obj);
  } else {
    var merged = Object.assign({}, existing, patch);
    merged.kunci = SINGLETON_KEY_CETAK;
    merged.updatedAt = nowIso_();
    updateObject_(s.sh, s.def, SINGLETON_KEY_CETAK, merged);
  }
  /* Cache dibatalkan lebih dulu, baru hasil dibaca ulang — melihat catatan
     pada saveKopRecord_ (kembalikan data segar, bukan entri cache lama). */
  refCacheBust_('cetakconfig');
  var rec = fetchCetakConfigRecord_();
  return rec ? cetakConfigFromRecord_(rec) : cetakConfigFromRecord_({});
}

/* ==========================================================================
 * END FILE ASLI: Pengaturan.gs
 * ========================================================================== */

/* ==========================================================================
 * BEGIN FILE ASLI: DasarHukum.gs
 * ========================================================================== */
/**
 * ============================================================================
 * SIAP SPPD & RAMPUNG — MODUL DASAR HUKUM
 * File   : DasarHukum.gs
 * Struktur frontend: { id, no, dasarHukum }
 *  - `no` hanya nomor urut tampilan (opsional); backend tidak memaksanya unik.
 * ============================================================================
 */

function dasarHukumSheet_() {
  var def = getDef_('DASAR_HUKUM');
  return { def: def, sh: openSheet_(def) };
}

function dasarHukumFromRecord_(rec) {
  rec = rec || {};
  var obj = {
    id: toStr_(rec.idDasarHukum),
    dasarHukum: toStr_(rec.dasarHukum),
    createdAt: toStr_(rec.createdAt),
    updatedAt: toStr_(rec.updatedAt)
  };
  if (isFilled_(rec.no)) obj.no = toStr_(rec.no);
  return obj;
}

function dasarHukumToRecord_(data) {
  var d = clone_(data) || {};
  return {
    idDasarHukum: toStr_(d.id || d.idDasarHukum),
    no: toStr_(d.no),
    dasarHukum: toStr_(d.dasarHukum),
    createdAt: toStr_(d.createdAt),
    updatedAt: toStr_(d.updatedAt)
  };
}

/** Versi ber-cache (CacheService, TTL pendek): mempercepat pembacaan berulang.
 * Otomatis dibatalkan oleh refCacheBust_ setiap kali data dasarhukum berubah. */
function fetchDasarHukumList_() {
  return refCacheGet_('dasarhukum', fetchDasarHukumList_Uncached_);
}

/** Implementasi asli (tanpa cache) — perilaku identik dengan sebelumnya. */
function fetchDasarHukumList_Uncached_() {
  var s = dasarHukumSheet_();
  var rows = readObjects_(s.sh, s.def);
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var obj = dasarHukumFromRecord_(rows[i]);
    obj.__row = rows[i].__row;
    out.push(obj);
  }
  out = sortByCreatedDesc_(out);
  for (var j = 0; j < out.length; j++) delete out[j].__row;
  return out;
}

function fetchDasarHukumById_(id) {
  var s = dasarHukumSheet_();
  var rec = getObjectById_(s.sh, s.def, id);
  return rec ? dasarHukumFromRecord_(rec) : null;
}

/* Simpan Dasar Hukum (upsert — sama seperti updateDasarHukum() frontend). */
function saveDasarHukumRecord_(data, strict) {
  var d = clone_(data) || {};
  var s = dasarHukumSheet_();
  var id = toStr_(d.id || d.idDasarHukum);
  var lama = id ? getObjectById_(s.sh, s.def, id) : null;
  if (!lama) id = id ? id : newId_('dh');
  d.id = id; d.idDasarHukum = id;
  if (strict !== false) {
    requireFields_(d, [{ field: 'dasarHukum', label: 'Dasar Hukum' }], 'Dasar Hukum');
  }
  var gabung = {};
  if (lama) { for (var k in lama) { if (k !== '__row') gabung[k] = lama[k]; } }
  for (var key in d) { if (d[key] !== undefined) gabung[key] = d[key]; }
  gabung.id = id; gabung.idDasarHukum = id;
  stampForWrite_(gabung, !lama);
  var rec = dasarHukumToRecord_(gabung);
  rec.idDasarHukum = id;
  if (!lama) {
    if (!isFilled_(rec.createdAt)) rec.createdAt = nowIso_();
    appendObject_(s.sh, s.def, rec);
  } else {
    if (!isFilled_(rec.createdAt)) rec.createdAt = toStr_(lama.createdAt);
    updateObject_(s.sh, s.def, id, rec);
  }
  /* Cache referensi modul ini dibatalkan — data berubah. */
  refCacheBust_('dasarhukum');
  return fetchDasarHukumById_(id);
}

function removeDasarHukum_(id) {
  var key = requireId_(id, 'ID Dasar Hukum');
  var s = dasarHukumSheet_();
  assertFound_(getObjectById_(s.sh, s.def, key), 'Dasar Hukum');
  deleteObject_(s.sh, s.def, key);
  /* Cache referensi modul ini dibatalkan — data berubah. */
  refCacheBust_('dasarhukum');
  return true;
}

/* ---- API PUBLIK (dipakai Api.gs) ---- */
function getDasarHukum(token) {
  return safeRun_('Data Dasar Hukum berhasil dibaca.', function () {
    requireSession_(token);
    return fetchDasarHukumList_();
  });
}

function getDasarHukumById(token, id) {
  return safeRun_('Data Dasar Hukum berhasil dibaca.', function () {
    requireSession_(token);
    var obj = fetchDasarHukumById_(requireId_(id, 'ID Dasar Hukum'));
    assertFound_(obj, 'Dasar Hukum');
    return obj;
  });
}

function updateDasarHukum(token, data, dataAlt) {
  return safeRun_('Data Dasar Hukum berhasil disimpan.', function () {
    var sess = requireSession_(token);
    requireWrite_(sess, 'pengaturan');
    var payload;
    if (typeof data === 'string' || typeof data === 'number') {
      payload = Object.assign({}, dataAlt || {});
      payload.id = String(data);
    } else { payload = Object.assign({}, data || {}); }
    if (!isFilled_(payload.id) && isFilled_(payload.idDasarHukum)) payload.id = payload.idDasarHukum;
    return saveDasarHukumRecord_(payload, true);
  });
}

function deleteDasarHukum(token, id) {
  return safeRun_('Data Dasar Hukum berhasil dihapus.', function () {
    var sess = requireSession_(token);
    requireWrite_(sess, 'pengaturan');
    return removeDasarHukum_(id);
  });
}

/* ==========================================================================
 * END FILE ASLI: DasarHukum.gs
 * ========================================================================== */

/* ==========================================================================
 * BEGIN FILE ASLI: Print.gs
 * ========================================================================== */
/**
 * SIAP SPPD & RAMPUNG — MODULE CETAK DOKUMEN
 * File   : Print.gs
 *
 * Menghasilkan HTML siap cetak untuk:
 *   SPPD  : SPT, SPD, Lampiran, Nota Dinas, SPD Per Orang, Lampiran Per Orang
 *   Rampung : Rampung 1, Kwitansi 1, Kwitansi 2, Kwitansi 3, Rill, Rekap
 *
 * Format HTML mengikuti logika migrasi dari Refrensi/Rampung.html
 * yang sudah dipakai frontend (index.html).
 *
 * UTILITAS: menggunakan fungsi dari Utils.gs (toStr_, toNum_, escapeHtml_,
 * terbilang_, calcBiaya_, calcRealisasi_, nowIso_) — TIDAK diduplikasi.
 */

/* ============================ FORMAT CETAK ============================ */

function fmtRupiah_(n) {
  n = toNum_(n);
  return 'Rp ' + n.toLocaleString('id-ID');
}

function fmtDate_(v) {
  if (!v) return '-';
  var d = new Date(String(v).length === 10 ? v + 'T00:00:00' : v);
  if (isNaN(d.getTime())) return String(v);
  var bulan = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
  return d.getDate() + ' ' + bulan[d.getMonth()] + ' ' + d.getFullYear();
}

/* ============================ TEMPLATE DASAR ============================ */

function printTemplate_(title, content) {
  var inst = getInstansiRecord_() || {};
  var logoHtml = isFilled_(inst.logo)
    ? '<img src="' + escapeHtml_(inst.logo) + '" style="width:92px;height:92px;object-fit:contain" />'
    : '<i class="fas fa-landmark" style="font-size:32px;color:#1f5c8b"></i>';
  return '<style>' +
    '@page{size:A4;margin:15mm 20mm}' +
    '#printArea{font-family:"Arial Narrow",Arial,sans-serif;font-size:11pt;line-height:1.15;margin:0;padding:0;color:#000}' +
    '#printArea table{border-collapse:collapse;width:100%}' +
    '#printArea td{vertical-align:top;padding:1px 0}' +
    '#printArea .pp-title{text-align:center;font-weight:bold;font-size:13pt;letter-spacing:4pt;margin-bottom:4px}' +
    '#printArea .pp-note{margin-top:12px;font-size:10pt;line-height:1.4}' +
    '#printArea .pp-sign{margin-top:30px;display:flex;justify-content:flex-end}' +
    '#printArea .pp-signbox{text-align:center;min-width:200px}' +
    '#printArea .pp-signbox .gap{height:50px}' +
    '#printArea .data-table{width:100%;border:1.2px solid #000;margin:8px 0}' +
    '#printArea .data-table th,#printArea .data-table td{border:1.2px solid #000;padding:4px 6px;font-size:10pt;text-align:left}' +
    '#printArea .data-table th{background:#f0f0f0;font-weight:600}' +
    /* BUG YANG DIPERBAIKI: sel uang/qty pada cetak Rampung & Kwitansi memakai
       class text-end/text-center yang tidak pernah didefinisikan di CSS cetak,
       sehingga angka tampil rata kiri. Kelas kini didefinisikan + digit tabular
       agar kolom uang sejajar vertikal. */
    '#printArea .text-end{text-align:right;font-variant-numeric:tabular-nums}' +
    '#printArea .text-center{text-align:center}' +
    '#printArea .top-section{width:100%;margin-bottom:0}' +
    '#printArea .top-section td{vertical-align:top;padding:1px 0}' +
    '#printArea .top-lbl{width:130px}' +
    '#printArea .top-sep{width:10px;text-align:center}' +
    '#printArea .grid{width:100%;border:1.2px solid #000;table-layout:fixed;margin-top:0}' +
    '#printArea .grid td{border:1.2px solid #000;vertical-align:top;padding:4px 6px;font-size:10pt}' +
    '#printArea .grid .sec{width:28px;text-align:center;font-weight:normal;vertical-align:top}' +
    '#printArea .grid .left-col{width:50%}#printArea .grid .right-col{width:50%}' +
    '#printArea .sig-area{min-height:55px;text-align:center;padding-top:5px}' +
    '#printArea .sig-dots{margin-top:35px;margin-bottom:2px}' +
    '#printArea .perhatian{margin-top:12px;font-size:9.5pt;line-height:1.3}' +
    '#printArea .field-row{margin-bottom:1px}' +
    '#printArea .field-row .fl{display:inline-block;width:100px}' +
    '#printArea .field-row .fs{display:inline-block;width:12px;text-align:center}' +
    '.page-break{page-break-after:always}' +
  '</style>' +
  '<div style="width:92px;text-align:center;float:left">' + logoHtml + '</div>' +
  '<div style="text-align:center;margin-left:100px">' +
    '<div style="font-size:14pt;font-weight:bold;letter-spacing:1.5pt;text-transform:uppercase">' + escapeHtml_(inst.nama || 'Nama Instansi') + '</div>' +
    '<div style="font-size:14pt;font-weight:bold;letter-spacing:1pt;text-transform:uppercase;white-space:nowrap">' + escapeHtml_(inst.organisasi || '') + '</div>' +
    '<div style="font-size:9pt;margin-top:2px">' + escapeHtml_(inst.alamat || '') + '</div>' +
    '<div style="font-size:9pt">' + escapeHtml_(inst.kabkota || '') + ', ' + escapeHtml_(inst.provinsi || '') + '</div>' +
  '</div>' +
  '<div style="border-top:3px solid #000;margin-top:5px;border-bottom:1px solid #000;padding-bottom:2px;margin-bottom:10px;clear:both"></div>' +
  content;
}

function tandaTangan_() {
  var inst = getInstansiRecord_() || {};
  var pej = findPejabatByNama_(inst.namaKepala, 'Pengguna Anggaran') ||
    findPejabatByRole_('Pengguna Anggaran') ||
    { nama: inst.namaKepala, nip: inst.nipKepala, jabatan: inst.jabatanKepala };
  var cfg = getCetakConfigRecord_() || {};
  var cls = cfg.posisiTtd === 'tengah' ? 'justify-content-center' : cfg.posisiTtd === 'kiri' ? '' : 'justify-content-end';
  var spacing = (toNum_(cfg.jarakTtd) || 2) * 16 + 'px';
  return '<div class="pp-sign ' + cls + '"><div class="pp-signbox">' +
    '<div>' + escapeHtml_(pej.jabatan || 'Pejabat') + ',</div>' +
    '<div class="gap" style="height:' + spacing + '"></div>' +
    '<div style="font-weight:700;text-decoration:underline">' + escapeHtml_(pej.nama || '-') + '</div>' +
    '<div>NIP. ' + escapeHtml_(pej.nip || '-') + '</div>' +
  '</div></div>';
}

/* ============================ IDENTITAS ============================ */

function getInstansiRecord_() {
  var cached = fetchInstansiRecord_();
  return cached || defaultInstansi_();
}

function getCetakConfigRecord_() {
  var cached = fetchCetakConfigRecord_();
  return cached || defaultCetakConfig_();
}

function getDefaultKop_() {
  var list = fetchKopList_();
  for (var i = 0; i < list.length; i++) {
    if (toBool_(list[i].default)) return list[i];
  }
  return list.length ? list[0] : {};
}

function findKopByJenis_(jenis) {
  var list = fetchKopList_();
  var j = toStr_(jenis).trim().toLowerCase();
  for (var i = 0; i < list.length; i++) {
    if (toStr_(list[i].jenisSuratTugas).trim().toLowerCase() === j) return list[i];
  }
  return null;
}

function findPejabatByNama_(nama, peran) {
  var list = fetchPejabatList_();
  var n = toStr_(nama || '').trim().toLowerCase();
  for (var i = 0; i < list.length; i++) {
    if (toStr_(list[i].status) !== 'Aktif') continue;
    if (n && toStr_(list[i].nama).trim().toLowerCase() === n) return list[i];
  }
  if (peran) {
    for (var j = 0; j < list.length; j++) {
      if (toStr_(list[j].status) !== 'Aktif') continue;
      var arr = normalizePeranPejabat_(list[j].peran);
      if (arr.indexOf(peran) > -1) return list[j];
    }
  }
  return null;
}

function findPejabatByRole_(peran) {
  return findPejabatByNama_('', peran);
}

function printIdentityConfig_(r, forSpt) {
  var inst = getInstansiRecord_() || {};
  var defaultKop = getDefaultKop_() || {};
  var m = {};
  if (forSpt && r) {
    var jenis = toStr_(r.jenisSuratTugas || '').trim();
    m = jenis ? (findKopByJenis_(jenis) || {}) : {};
  }
  var src = Object.keys(m).length ? m : defaultKop;
  return {
    namaInstansi: toStr_(src.namaInstansi) || toStr_(inst.nama) || toStr_(defaultKop.namaInstansi) || 'Nama Instansi',
    namaUnitKerja: toStr_(src.namaUnitKerja) || toStr_(inst.organisasi) || toStr_(defaultKop.namaUnitKerja) || '',
    alamat: toStr_(src.alamat) || toStr_(inst.alamat) || toStr_(defaultKop.alamat) || '',
    telp: (toStr_(src.telepon) && toStr_(src.telepon) !== '-' ? toStr_(src.telepon) : '') || toStr_(inst.telepon) || toStr_(defaultKop.telepon) || '',
    laman: (toStr_(src.laman) && toStr_(src.laman) !== '-' ? toStr_(src.laman) : '') || toStr_(inst.laman) || toStr_(defaultKop.laman) || '',
    kota: toStr_(src.kota) || toStr_(inst.kabkota) || '',
    tahun: new Date().getFullYear(),
    logoBase64: toStr_(inst.logo) || toStr_(src.logo) || '',
    dasarHukum: (fetchDasarHukumList_() || []).map(function (x) { return x.dasarHukum; }).filter(Boolean).join('\n')
  };
}

function transportLabel_(v) {
  var opts = {
    'pesawat': 'Pesawat Terbang', 'bus': 'Bus', 'kereta': 'Kereta Api',
    'kendaraan-dinas': 'Kendaraan Dinas', 'kapal': 'Kapal Laut'
  };
  return opts[toStr_(v)] || toStr_(v) || '-';
}

function sppdPakaiNotaDinas_(r) {
  var k = findKopByJenis_(toStr_(r.jenisSuratTugas || ''));
  return k ? toBool_(k.notaDinas) : false;
}

/* ============================ VALIDASI ============================ */

function validatePrintSPPD_(r, forPerOrang) {
  if (!r) return false;
  var missing = [];
  if (!toStr_(r.noSpt)) missing.push('Nomor SPT');
  if (!toStr_(r.noSppd)) missing.push('Nomor SPPD');
  if (!toStr_(r.tglSpt)) missing.push('Tanggal SPT');
  if (!toStr_(r.tglSppd)) missing.push('Tanggal SPPD');
  if (!toStr_(r.kotaTujuan)) missing.push('Daerah/Tempat Tujuan');
  if (!toStr_(r.transport)) missing.push('Alat Angkutan');
  if (!toStr_(r.tglBerangkat)) missing.push('Tanggal Pergi');
  if (!toStr_(r.tglKembali)) missing.push('Tanggal Pulang');
  if (!toStr_(r.maksud)) missing.push('Untuk / Maksud');
  if (!toStr_(r.ttdSpt)) missing.push('Penandatangan SPT');
  if (!toStr_(r.ttdSpd)) missing.push('Penandatangan SPD');
  if (!toStr_(r.pptk)) missing.push('PPTK');
  if (forPerOrang) {
    var pel = r.pelaksana || [];
    for (var i = 0; i < pel.length; i++) {
      if (!toStr_(pel[i].nama)) missing.push('Nama Pelaksana');
      if (!toStr_(pel[i].nip)) missing.push('NIP Pelaksana');
      if (!toStr_(pel[i].jabatan)) missing.push('Jabatan Pelaksana');
    }
  }
  return missing.length === 0;
}

/* ============================ BUILDERS SPPD ============================ */

function buildSppdHeader_(r, title) {
  var cfg = printIdentityConfig_(r, title === 'SPT');
  return '<div class="pp-title">' + escapeHtml_(title) + '</div>' +
    '<div style="font-size:12px;text-align:center;margin-bottom:14px">Nomor: ' + escapeHtml_(r.noSppd || '') + '</div>';
}

function buildSppdSPT_(r) {
  var cfg = printIdentityConfig_(r, true);
  var pel = r.pelaksana || [];
  var pelHTML = pel.map(function (p, i) {
    return '<tr><td style="width:30px">' + (i + 1) + '</td>' +
      '<td>' + escapeHtml_(p.nama || '-') + '</td>' +
      '<td>' + escapeHtml_(p.nip || '-') + '</td>' +
      '<td>' + escapeHtml_(p.jabatan || '-') + '</td>' +
      '<td>' + escapeHtml_(p.golongan || '-') + '</td></tr>';
  }).join('');
  var content = '<div class="pp-title">SURAT PERINTAH TUGAS</div>' +
    '<div style="font-size:12px;text-align:center;margin-bottom:14px">Nomor: ' + escapeHtml_(r.noSpt || '') + '</div>' +
    '<table class="top-section">' +
    '<tr><td class="top-lbl">Dasar</td><td class="top-sep">:</td><td>' + escapeHtml_(cfg.dasarHukum || '-') + '</td></tr>' +
    '<tr><td class="top-lbl">Nama Instansi</td><td class="top-sep">:</td><td>' + escapeHtml_(cfg.namaInstansi) + '</td></tr>' +
    '<tr><td class="top-lbl">Unit Kerja</td><td class="top-sep">:</td><td>' + escapeHtml_(cfg.namaUnitKerja) + '</td></tr>' +
    '</table>' +
    '<table class="grid"><tr><td class="sec">A</td><td colspan="3"><b>Yang diperintahkan</b></td></tr>' +
    '<tr><td></td><td class="left-col">Nama</td><td class="left-col">NIP</td><td>Jabatan</td></tr>' +
    pelHTML +
    '</table>' +
    '<table class="grid"><tr><td class="sec">B</td><td colspan="3"><b>Perintah Tugas</b></td></tr>' +
    '<tr><td></td><td class="left-col">Maksud</td><td class="left-col">Untuk</td><td>Tempat Tujuan</td></tr>' +
    '<tr><td></td><td class="left-col">' + escapeHtml_(r.maksud || '-') + '</td><td class="left-col">' + escapeHtml_(sptUntukRaw_(r) || '-') + '</td><td>' + escapeHtml_(r.kotaTujuan || '-') + '</td></tr>' +
    '<tr><td></td><td class="left-col">Tanggal Pergi</td><td class="left-col">' + fmtDate_(r.tglBerangkat) + '</td><td>Tanggal Pulang</td><td>' + fmtDate_(r.tglKembali) + '</td></tr>' +
    '<tr><td></td><td class="left-col">Alat Angkutan</td><td class="left-col">' + escapeHtml_(transportLabel_(r.transport)) + '</td><td></td><td></td></tr>' +
    '</table>' +
    '<div style="margin-top:12px">Demikian Surat Perintah Tugas ini dibuat untuk dilaksanakan dengan penuh tanggung jawab.</div>' +
    tandaTangan_();
  return printTemplate_('', content);
}

function buildSppdSPD_(r, pel) {
  /* BUG YANG DIPERBAIKI: pemanggil printSPPD format 'spd' hanya mengirim
     (r) — pelaksana jatuh ke {}. Karena identitas pelaksana berada di
     r.pelaksana, SPD non-perorang selalu mencetak "Nama: -, NIP: -". */
  pel = pel || (r && Array.isArray(r.pelaksana) && r.pelaksana[0]) || {};
  var cfg = printIdentityConfig_(r, false);
  var biaya = r.biaya || {};
  var total = calcBiaya_(biaya);
  var content = '<div class="pp-title">SURAT PERJALANAN DINAS</div>' +
    '<div style="font-size:12px;text-align:center;margin-bottom:14px">Nomor: ' + escapeHtml_(r.noSppd || '') + '</div>' +
    '<table class="grid"><tr><td class="sec">A</td><td colspan="3"><b>Dasar Perjalanan Dinas</b></td></tr>' +
    '<tr><td></td><td class="left-col">Nomor SPT</td><td class="left-col">' + escapeHtml_(r.noSpt || '-') + '</td><td>Tanggal</td><td>' + fmtDate_(r.tglSpt) + '</td></tr>' +
    '<tr><td class="sec">B</td><td colspan="3"><b>Pejabat Pembuat Komitmen</b></td></tr>' +
    '<tr><td></td><td class="left-col">Nama</td><td class="left-col">' + escapeHtml_(r.pptk || '-') + '</td><td></td><td></td></tr>' +
    '<tr><td class="sec">C</td><td colspan="3"><b>Pelaksana Perjalanan Dinas</b></td></tr>' +
    '<tr><td></td><td class="left-col">Nama</td><td class="left-col">' + escapeHtml_(pel.nama || '-') + '</td><td>NIP</td><td>' + escapeHtml_(pel.nip || '-') + '</td></tr>' +
    '<tr><td></td><td class="left-col">Jabatan</td><td class="left-col">' + escapeHtml_(pel.jabatan || '-') + '</td><td>Pangkat/Gol</td><td>' + escapeHtml_(pel.golongan || '-') + '</td></tr>' +
    '<tr><td class="sec">D</td><td colspan="3"><b>Informasi Perjalanan</b></td></tr>' +
    '<tr><td></td><td class="left-col">Maksud</td><td class="left-col" colspan="3">' + escapeHtml_(r.maksud || '-') + '</td></tr>' +
    '<tr><td></td><td class="left-col">Tujuan</td><td class="left-col" colspan="3">' + escapeHtml_(r.kotaTujuan || '-') + '</td></tr>' +
    '<tr><td></td><td class="left-col">Tanggal Pergi</td><td class="left-col">' + fmtDate_(r.tglBerangkat) + '</td><td>Tanggal Pulang</td><td>' + fmtDate_(r.tglKembali) + '</td></tr>' +
    '<tr><td></td><td class="left-col">Alat Angkutan</td><td class="left-col" colspan="3">' + escapeHtml_(transportLabel_(r.transport)) + ' ' + escapeHtml_(r.noTransport || '') + '</td></tr>' +
    '<tr><td class="sec">E</td><td colspan="3"><b>Rincian Biaya</b></td></tr>' +
    '<tr><td></td><td class="left-col">Uang Harian</td><td class="left-col text-end">' + fmtRupiah_(biaya.uangHarian) + '</td><td>Representasi</td><td class="text-end">' + fmtRupiah_(biaya.representasi) + '</td></tr>' +
    '<tr><td></td><td class="left-col">Transport</td><td class="left-col text-end">' + fmtRupiah_(biaya.transport) + '</td><td>Taksi</td><td class="text-end">' + fmtRupiah_(biaya.taksi) + '</td></tr>' +
    '<tr><td></td><td class="left-col">Penginapan</td><td class="left-col text-end">' + fmtRupiah_(biaya.penginapan) + '</td><td>Sewa Kendaraan</td><td class="text-end">' + fmtRupiah_(biaya.sewaKendaraan) + '</td></tr>' +
    '<tr><td></td><td class="left-col">Biaya Lainnya</td><td class="left-col text-end">' + fmtRupiah_(biaya.lainnya) + '</td><td>PPn</td><td class="text-end">' + fmtRupiah_(biaya.ppn) + '</td></tr>' +
    '<tr><td></td><td class="left-col">PPH</td><td class="left-col text-end">' + fmtRupiah_(biaya.pph) + '</td><td>Total</td><td class="text-end"><b>' + fmtRupiah_(total) + '</b></td></tr>' +
    '</table>' +
    '<div style="margin-top:12px">Demikian Surat Perjalanan Dinas ini dibuat untuk dipergunakan sebagaimana mestinya.</div>' +
    tandaTangan_();
  return printTemplate_('', content);
}

function buildSppdLampiran_(r, pel) {
  /* Sama seperti buildSppdSPD_: ambil pelaksana pertama bila tidak dikirim. */
  pel = pel || (r && Array.isArray(r.pelaksana) && r.pelaksana[0]) || {};
  var cfg = printIdentityConfig_(r, false);
  var biaya = r.biaya || {};
  var content = '<div class="pp-title">LAMPIRAN SURAT PERJALANAN DINAS</div>' +
    '<div style="font-size:12px;text-align:center;margin-bottom:14px">Nomor: ' + escapeHtml_(r.noSppd || '') + '</div>' +
    '<table class="grid"><tr><td class="sec">I</td><td colspan="3"><b>Identitas</b></td></tr>' +
    '<tr><td></td><td class="left-col">Nama</td><td class="left-col">' + escapeHtml_(pel.nama || '-') + '</td><td>NIP</td><td>' + escapeHtml_(pel.nip || '-') + '</td></tr>' +
    '<tr><td></td><td class="left-col">Jabatan</td><td class="left-col">' + escapeHtml_(pel.jabatan || '-') + '</td><td>Unit</td><td>' + escapeHtml_(pel.unit || '-') + '</td></tr>' +
    '<tr><td class="sec">II</td><td colspan="3"><b>Rincian Biaya</b></td></tr>' +
    '<tr><td></td><td class="left-col">Uang Harian</td><td class="left-col text-end">' + fmtRupiah_(biaya.uangHarian) + '</td><td>Representasi</td><td class="text-end">' + fmtRupiah_(biaya.representasi) + '</td></tr>' +
    '<tr><td></td><td class="left-col">Transport</td><td class="left-col text-end">' + fmtRupiah_(biaya.transport) + '</td><td>Taksi</td><td class="text-end">' + fmtRupiah_(biaya.taksi) + '</td></tr>' +
    '<tr><td></td><td class="left-col">Penginapan</td><td class="left-col text-end">' + fmtRupiah_(biaya.penginapan) + '</td><td>Sewa Kendaraan</td><td class="text-end">' + fmtRupiah_(biaya.sewaKendaraan) + '</td></tr>' +
    '<tr><td></td><td class="left-col">Biaya Lainnya</td><td class="left-col text-end">' + fmtRupiah_(biaya.lainnya) + '</td><td>PPn</td><td class="text-end">' + fmtRupiah_(biaya.ppn) + '</td></tr>' +
    '<tr><td></td><td class="left-col">PPH</td><td class="left-col text-end">' + fmtRupiah_(biaya.pph) + '</td><td>Total</td><td class="text-end"><b>' + fmtRupiah_(calcBiaya_(biaya)) + '</b></td></tr>' +
    '</table>' +
    '<div style="margin-top:12px">Lampiran ini merupakan bagian yang tidak terpisahkan dari Surat Perjalanan Dinas.</div>' +
    tandaTangan_();
  return printTemplate_('', content);
}

function buildSppdNotaDinas_(r) {
  var cfg = printIdentityConfig_(r, false);
  var content = '<div class="pp-title">NOTA DINAS</div>' +
    '<div style="font-size:12px;text-align:center;margin-bottom:14px">Nomor: ' + escapeHtml_(r.noSppd || '') + '</div>' +
    '<table class="grid"><tr><td class="left-col"><b>Kepada</b></td><td>' + escapeHtml_(r.ttdSpd || '-') + '</td></tr>' +
    '<tr><td><b>Dari</b></td><td>' + escapeHtml_(r.ttdSpt || '-') + '</td></tr>' +
    '<tr><td><b>Tanggal</b></td><td>' + fmtDate_(r.tglSpt) + '</td></tr>' +
    '<tr><td><b>Nomor</b></td><td>' + escapeHtml_(r.noSpt || '-') + '</td></tr></table>' +
    '<div style="margin-top:12px">Dengan ini disampaikan bahwa pegawai tersebut di bawah ini akan melaksanakan perjalanan dinas:</div>' +
    '<table class="grid"><tr><td class="left-col">Nama</td><td>' + escapeHtml_((r.pelaksana && r.pelaksana[0]) ? r.pelaksana[0].nama : '-') + '</td></tr>' +
    '<tr><td>NIP</td><td>' + escapeHtml_((r.pelaksana && r.pelaksana[0]) ? r.pelaksana[0].nip : '-') + '</td></tr>' +
    '<tr><td>Maksud</td><td>' + escapeHtml_(r.maksud || '-') + '</td></tr>' +
    '<tr><td>Tujuan</td><td>' + escapeHtml_(r.kotaTujuan || '-') + '</td></tr>' +
    '<tr><td>Tanggal</td><td>' + fmtDate_(r.tglBerangkat) + ' s/d ' + fmtDate_(r.tglKembali) + '</td></tr>' +
    '</table>' +
    '<div style="margin-top:12px">Demikian Nota Dinas ini disampaikan untuk menjadi perhatian dan dilaksanakan.</div>' +
    tandaTangan_();
  return printTemplate_('', content);
}

function buildSppdPerOrang_(r, builder) {
  var pelAll = r.pelaksana || [];
  if (pelAll.length <= 1) return builder(r, pelAll[0] || {});
  var out = [];
  for (var i = 0; i < pelAll.length; i++) {
    var doc = builder(r, pelAll[i]);
    out.push(i < pelAll.length - 1 ? '<div style="page-break-after:always">' + doc + '</div>' : doc);
  }
  return out.join('');
}

/* ============================ API PRINT SPPD ============================ */

function printSPPD(token, id, format) {
  return safeRun_('HTML cetak SPPD berhasil dibuat.', function () {
    var sess = requireSession_(token);
    var sppd = fetchSPPDById_(requireId_(id, 'ID SPPD'));
    assertFound_(sppd, 'SPPD');
    requireRecordAccess_(sess, sppd, 'SPPD');
    format = toStr_(format || 'spd').toLowerCase();
    if (format === 'spt') return { html: buildSppdSPT_(sppd), filename: 'SPT_' + sppd.noSpt.replace(/\//g, '_') + '.html' };
    if (format === 'spd') return { html: buildSppdSPD_(sppd), filename: 'SPD_' + sppd.noSppd.replace(/\//g, '_') + '.html' };
    if (format === 'lampiran') return { html: buildSppdLampiran_(sppd), filename: 'Lampiran_' + sppd.noSppd.replace(/\//g, '_') + '.html' };
    if (format === 'nota') {
      if (!sppdPakaiNotaDinas_(sppd)) throwError_('Nota Dinas tidak tersedia untuk Jenis Surat Tugas ini.');
      return { html: buildSppdNotaDinas_(sppd), filename: 'NotaDinas_' + sppd.noSppd.replace(/\//g, '_') + '.html' };
    }
    if (format === 'spd-perorang') return { html: buildSppdPerOrang_(sppd, buildSppdSPD_), filename: 'SPD_PerOrang_' + sppd.noSppd.replace(/\//g, '_') + '.html' };
    if (format === 'lampiran-perorang') return { html: buildSppdPerOrang_(sppd, buildSppdLampiran_), filename: 'Lampiran_PerOrang_' + sppd.noSppd.replace(/\//g, '_') + '.html' };
    throwError_('Format cetak SPPD tidak dikenal: ' + format);
  });
}

/* ============================ BUILDERS RAMPUNG ============================ */

function rampungIdentHTML_(r) {
  var k = r.kegiatan || {};
  var rows = '';
  if (k.program || k.kegiatan || k.subkegiatan || k.rekening) {
    rows += '<tr><td style="width:180px">Kode Organisasi</td><td>: ' + escapeHtml_(k.kodeOrganisasi || '-') + '</td></tr>';
    rows += '<tr><td>Program</td><td>: ' + escapeHtml_(k.program || '-') + '</td></tr>';
    rows += '<tr><td>Kegiatan</td><td>: ' + escapeHtml_(k.kodeKegiatan || '-') + ' — ' + escapeHtml_(k.kegiatan || '-') + '</td></tr>';
    rows += '<tr><td>Subkegiatan</td><td>: ' + escapeHtml_(k.kodeSubkegiatan || '-') + ' — ' + escapeHtml_(k.subkegiatan || '-') + '</td></tr>';
    rows += '<tr><td>Kode Rekening</td><td>: ' + escapeHtml_(k.rekening || '-') + '</td></tr>';
    rows += '<tr><td>Sumber Dana</td><td>: ' + escapeHtml_(k.sumberDana || '-') + '</td></tr>';
    rows += '<tr><td>Keperluan</td><td>: ' + escapeHtml_(k.keperluan || '-') + '</td></tr>';
  }
  rows += '<tr><td style="width:180px">Nama</td><td>: ' + escapeHtml_(r.nama) + '</td></tr>';
  rows += '<tr><td>NIP</td><td>: ' + escapeHtml_(r.nip) + '</td></tr>';
  rows += '<tr><td>Jabatan</td><td>: ' + escapeHtml_(r.jabatan) + '</td></tr>';
  rows += '<tr><td>Tujuan</td><td>: ' + escapeHtml_(r.tujuan || '-') + '</td></tr>';
  rows += '<tr><td>Maksud</td><td>: ' + escapeHtml_(r.maksud || '-') + '</td></tr>';
  rows += '<tr><td>Periode</td><td>: ' + fmtDate_(r.tglBerangkat) + ' — ' + fmtDate_(r.tglKembali) + '</td></tr>';
  return '<table style="margin-bottom:8px">' + rows + '</table>';
}

function rampungBiayaHTML_(r) {
  var html = '';
  if (r.rincian && r.rincian.length) {
    html = '<thead><tr><th>Komponen</th><th class="text-center">QTY</th><th class="text-end">Harga</th><th class="text-center">%</th><th class="text-end">Subtotal</th></tr></thead><tbody>';
    for (var i = 0; i < r.rincian.length; i++) {
      var p = r.rincian[i];
      html += '<tr style="background:#f5f7fa"><td colspan="5" style="font-weight:600">' + (i + 1) + '. ' + escapeHtml_(p.nama) + ' — NIP ' + escapeHtml_(p.nip) + ' · No. SPD ' + escapeHtml_(p.noSpd) + ' (' + fmtDate_(p.tglSpd) + ')</td></tr>';
      var komp = p.komponen || [];
      if (!komp.length) html += '<tr><td colspan="5">-</td></tr>';
      for (var j = 0; j < komp.length; j++) {
        var c = komp[j];
        html += '<tr><td>' + escapeHtml_(c.deskripsi || '-') + (c.rill ? ' <b>(Rill)</b>' : '') + '</td><td class="text-center">' + escapeHtml_(c.qty != null ? c.qty : '-') + ' ' + escapeHtml_(c.satuan || '') + '</td><td class="text-end">' + fmtRupiah_(c.harga) + '</td><td class="text-center">' + escapeHtml_(c.persen != null ? c.persen + '%' : '-') + '</td><td class="text-end">' + fmtRupiah_(c.subtotal) + '</td></tr>';
      }
      html += '<tr style="background:#eef6ff"><td colspan="4" style="font-weight:700">Subtotal — ' + escapeHtml_(p.nama) + '</td><td class="text-end" style="font-weight:700">' + fmtRupiah_(p.total) + '</td></tr>';
    }
    html += '</tbody>';
  } else {
    var rl = r.realisasi || {};
    html = '<thead><tr><th style="width:220px">Komponen</th><th style="width:120px" class="text-end">Nominal</th></tr></thead><tbody>';
    html += '<tr><td>Uang Harian</td><td class="text-end">' + fmtRupiah_(rl.uangHarian) + '</td></tr>';
    html += '<tr><td>Transport</td><td class="text-end">' + fmtRupiah_(rl.transport) + '</td></tr>';
    html += '<tr><td>Taksi</td><td class="text-end">' + fmtRupiah_(rl.taksi) + '</td></tr>';
    html += '<tr><td>Penginapan</td><td class="text-end">' + fmtRupiah_(rl.penginapan) + '</td></tr>';
    html += '<tr><td>Sewa Kendaraan</td><td class="text-end">' + fmtRupiah_(rl.sewaKendaraan) + '</td></tr>';
    html += '<tr><td>Representasi</td><td class="text-end">' + fmtRupiah_(rl.representasi) + '</td></tr>';
    html += '<tr><td>Biaya Lainnya</td><td class="text-end">' + fmtRupiah_(rl.lainnya) + '</td></tr></tbody>';
  }
  return html;
}

function buildRampungRampung1_(r) {
  var ident = rampungIdentHTML_(r);
  var biaya = rampungBiayaHTML_(r);
  var content = '<div class="pp-title">RAMPUNG PERJALANAN DINAS</div>' +
    '<div style="font-size:12px;text-align:center;margin-bottom:14px">Nomor SPPD: ' + escapeHtml_(r.noSppd) + '</div>' +
    ident +
    '<div style="font-weight:700;margin:12px 0 6px">' + ((r.rincian && r.rincian.length) ? 'Rincian Biaya Pegawai' : 'Rincian Realisasi Biaya') + '</div>' +
    '<table>' + biaya +
    '<tr style="background:#eef6ff"><td colspan="4" style="font-weight:700">Total Realisasi</td><td class="text-end" style="font-weight:700">' + fmtRupiah_(r.totalRealisasi) + '</td></tr>' +
    '</table>' +
    '<div class="pp-note">Berkas rampung telah diverifikasi dan dinyatakan <b>' + escapeHtml_(r.status) + '</b>.</div>' +
    tandaTangan_();
  return printTemplate_('', content);
}

function buildRampungKwitansi1_(r) {
  var content = '<div class="pp-title">KWITANSI PERJALANAN DINAS</div>' +
    '<div style="font-size:12px;text-align:center;margin-bottom:14px">Nomor SPPD: ' + escapeHtml_(r.noSppd) + '</div>' +
    '<table class="grid"><tr><td class="left-col">Telah terima dari</td><td>' + escapeHtml_(r.nama) + '</td></tr>' +
    '<tr><td>Jumlah</td><td>' + fmtRupiah_(r.totalRealisasi) + '</td></tr>' +
    '<tr><td>Untuk pembayaran</td><td>' + escapeHtml_(r.maksud || '-') + '</td></tr>' +
    '<tr><td>Periode</td><td>' + fmtDate_(r.tglBerangkat) + ' s/d ' + fmtDate_(r.tglKembali) + '</td></tr></table>' +
    '<div style="margin-top:12px">Demikian kwitansi ini dibuat dengan sebenarnya untuk dipergunakan sebagaimana mestinya.</div>' +
    tandaTangan_();
  return printTemplate_('', content);
}

function buildRampungKwitansi2_(r) {
  var content = '<div class="pp-title">KWITANSI PERJALANAN DINAS</div>' +
    '<div style="font-size:12px;text-align:center;margin-bottom:14px">Nomor: ' + escapeHtml_(r.noSppd) + '</div>' +
    '<table class="grid"><tr><td class="left-col">Yang menerima</td><td>' + escapeHtml_(r.nama) + '</td></tr>' +
    '<tr><td>NIP</td><td>' + escapeHtml_(r.nip) + '</td></tr>' +
    '<tr><td>Jumlah diterima</td><td>' + fmtRupiah_(r.totalRealisasi) + '</td></tr>' +
    '<tr><td>Keperluan</td><td>' + escapeHtml_(r.maksud || '-') + '</td></tr></table>' +
    '<div style="margin-top:12px">Kwitansi ini ditandatangani sebagai bukti pembayaran yang sah.</div>' +
    tandaTangan_();
  return printTemplate_('', content);
}

function buildRampungKwitansi3_(r) {
  var content = '<div class="pp-title">KWITANSI PERJALANAN DINAS (PER ORANG)</div>' +
    '<div style="font-size:12px;text-align:center;margin-bottom:14px">Nomor SPPD: ' + escapeHtml_(r.noSppd) + '</div>' +
    '<table class="grid"><tr><td class="left-col">Nama</td><td>' + escapeHtml_(r.nama) + '</td></tr>' +
    '<tr><td>NIP</td><td>' + escapeHtml_(r.nip) + '</td></tr>' +
    '<tr><td>Jabatan</td><td>' + escapeHtml_(r.jabatan) + '</td></tr>' +
    '<tr><td>Total diterima</td><td><b>' + fmtRupiah_(r.totalRealisasi) + '</b></td></tr></table>' +
    '<div style="margin-top:12px">Telah diterima dengan sebenarnya uang perjalanan dinas tersebut di atas.</div>' +
    tandaTangan_();
  return printTemplate_('', content);
}

function buildRampungRill_(r) {
  var content = '<div class="pp-title">BUKTI PENGELUARAN RILL</div>' +
    '<div style="font-size:12px;text-align:center;margin-bottom:14px">Nomor SPPD: ' + escapeHtml_(r.noSppd) + '</div>' +
    '<table class="grid"><tr><td class="left-col">Nama</td><td>' + escapeHtml_(r.nama) + '</td></tr>' +
    '<tr><td>NIP</td><td>' + escapeHtml_(r.nip) + '</td></tr>' +
    '<tr><td>Keperluan</td><td>' + escapeHtml_(r.maksud || '-') + '</td></tr></table>' +
    '<table class="data-table"><thead><tr><th>Deskripsi</th><th class="text-center">Qty</th><th class="text-end">Harga</th><th class="text-center">%</th><th class="text-end">Subtotal</th></tr></thead><tbody>' +
    rampungBiayaHTML_(r) +
    '</tbody></table>' +
    '<div style="margin-top:12px">Bukti pengeluaran rill dilampirkan sesuai ketentuan yang berlaku.</div>' +
    tandaTangan_();
  return printTemplate_('', content);
}

function buildRampungRekap_(r) {
  var content = '<div class="pp-title">REKAPITULASI PERJALANAN DINAS</div>' +
    '<div style="font-size:12px;text-align:center;margin-bottom:14px">Nomor SPPD: ' + escapeHtml_(r.noSppd) + '</div>' +
    '<table class="grid"><tr><td class="left-col">Nama</td><td>' + escapeHtml_(r.nama) + '</td></tr>' +
    '<tr><td>NIP</td><td>' + escapeHtml_(r.nip) + '</td></tr>' +
    '<tr><td>Tujuan</td><td>' + escapeHtml_(r.tujuan || '-') + '</td></tr>' +
    '<tr><td>Periode</td><td>' + fmtDate_(r.tglBerangkat) + ' s/d ' + fmtDate_(r.tglKembali) + '</td></tr>' +
    '<tr><td>Total SPPD</td><td>' + fmtRupiah_(r.totalSPPD) + '</td></tr>' +
    '<tr><td>Total Realisasi</td><td>' + fmtRupiah_(r.totalRealisasi) + '</td></tr>' +
    '<tr><td>Selisih</td><td>' + fmtRupiah_(r.selisih) + '</td></tr></table>' +
    '<div style="margin-top:12px">Rekapitulasi ini dibuat sebagai laporan pertanggungjawaban perjalanan dinas.</div>' +
    tandaTangan_();
  return printTemplate_('', content);
}

/* ============================ API PRINT RAMPUNG ============================ */

function printRampung(token, id, format) {
  return safeRun_('HTML cetak Rampung berhasil dibuat.', function () {
    var sess = requireSession_(token);
    var r = fetchRampungById_(requireId_(id, 'ID Rampung'));
    assertFound_(r, 'Rampung');
    requireRecordAccess_(sess, r, 'Rampung');
    format = toStr_(format || 'rampung').toLowerCase();
    if (format === 'rampung' || format === 'rampung1') return { html: buildRampungRampung1_(r), filename: 'Rampung1_' + r.noSppd.replace(/\//g, '_') + '.html' };
    if (format === 'kwitansi1') return { html: buildRampungKwitansi1_(r), filename: 'Kwitansi1_' + r.noSppd.replace(/\//g, '_') + '.html' };
    if (format === 'kwitansi2') return { html: buildRampungKwitansi2_(r), filename: 'Kwitansi2_' + r.noSppd.replace(/\//g, '_') + '.html' };
    if (format === 'kwitansi3') return { html: buildRampungKwitansi3_(r), filename: 'Kwitansi3_' + r.noSppd.replace(/\//g, '_') + '.html' };
    if (format === 'rill') return { html: buildRampungRill_(r), filename: 'Rill_' + r.noSppd.replace(/\//g, '_') + '.html' };
    if (format === 'rekap') return { html: buildRampungRekap_(r), filename: 'Rekap_' + r.noSppd.replace(/\//g, '_') + '.html' };
    throwError_('Format cetak Rampung tidak dikenal: ' + format);
  });
}

function printPreview(token, html) {
  return safeRun_('Preview HTML berhasil dibuat.', function () {
    requireSession_(token);
    return { html: toStr_(html) };
  });
}

/* ==========================================================================
 * END FILE ASLI: Print.gs
 * ========================================================================== */

/* ==========================================================================
 * BEGIN FILE ASLI: Logo.gs
 * ========================================================================== */
/**
 * SIAP SPPD & RAMPUNG — MODULE LOGO
 * File   : Logo.gs
 *
 * Menyimpan logo instansi ke Google Drive dan menyimpan file ID
 * di PropertiesService agar tidak memenuhi sel Google Sheets.
 */

/** Folder Drive untuk berkas logo. */
function logoDriveFolder_() {
  var name = DRIVE_FOLDER_NAME || 'SIAP SPPD - BERKAS';
  var folder = null;
  try {
    var folders = DriveApp.getFoldersByName(name);
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(name);
    }
  } catch (e) {
    folder = DriveApp.getFoldersByName(name).hasNext() ? DriveApp.getFoldersByName(name).next() : null;
  }
  return folder;
}

/** Batas ukuran logo: sama dengan keterangan pada UI (maks 2 MB). */
var LOGO_MAX_BYTES_ = 2 * 1024 * 1024;

/** MIME yang diterima untuk logo (sama dengan filter file pada UI). */
var LOGO_MIME_OK_ = ['image/png', 'image/jpeg'];

/** Nama berkas logo yang aman untuk Drive (tanpa karakter path). */
function logoSafeName_(name, mimeType) {
  var base = toStr_(name).trim() || 'logo';
  base = base.replace(/[\\/:*?"<>|]/g, '').trim() || 'logo';
  if (base.length > 80) base = base.slice(0, 80);
  var ext = (mimeType === 'image/png') ? '.png' : '.jpg';
  if (!/\.(png|jpe?g)$/i.test(base)) base += ext;
  return base;
}

/** Upload logo ke Drive dan simpan file ID. */
function uploadLogo(token, fileId, fileName) {
  return safeRun_('Logo berhasil diunggah.', function () {
    var sess = requireSession_(token);
    requireWrite_(sess, 'pengaturan');
    var sid = getSpreadsheetId_();
    var ss = SpreadsheetApp.openById(sid);
    var folder = logoDriveFolder_();
    var blob = null;
    var file = null;
    try {
      file = DriveApp.getFileById(fileId);
      blob = file.getBlob();
      blob.setName(fileName || file.getName());
    } catch (e) {
      throwError_('File tidak ditemukan. Pastikan fileId benar.');
    }
    /* Validasi MIME & ukuran disamakan dengan uploadLogoBytes — getLogoImage
       menghasilkan data URL yang hanya dapat dipakai UI bila MIME gambar
       dan ukuran wajar (maks 2 MB). */
    var mime = toStr_(blob.getContentType()).trim().toLowerCase();
    if (LOGO_MIME_OK_.indexOf(mime) === -1) {
      throwError_('Format logo tidak didukung. Gunakan PNG atau JPG.');
    }
    if (blob.getBytes().length > LOGO_MAX_BYTES_) {
      throwError_('Logo terlalu besar (maks 2 MB).');
    }
    try {
      var existing = folder.getFilesByName(fileName || '');
      while (existing.hasNext()) { existing.next().setTrashed(true); }
    } catch (e) {}
    var logoFile = folder.createFile(blob);
    var props = PropertiesService.getScriptProperties();
    props.setProperty(PROP_LOGO_FILE_ID, logoFile.getId());
    /* Update juga di INSTANSI sheet. */
    var inst = fetchInstansiRecord_();
    if (inst) {
      updateObject_(openSheet_(getDef_('INSTANSI')), getDef_('INSTANSI'), SINGLETON_KEY_INSTANSI, { logo: logoFile.getId(), updatedAt: nowIso_() });
    }
    /* BUG YANG DIPERBAIKI: cache instansi tidak pernah dibuang di sini
       (uploadLogoBytes melakukannya), jadi frontend tetap menampilkan logo
       lama hingga TTL habis. */
    refCacheBust_('instansi');
    return { fileId: logoFile.getId(), name: logoFile.getName(), url: logoFile.getUrl() };
  });
}

/**
 * Upload logo dari BROWSER ke Google Drive (dipakai tombol "Pilih Logo").
 * Browser tidak bisa membuat berkas Drive secara langsung, sehingga frontend
 * mengirim isi berkas sebagai base64:
 *   payload = { name, mimeType, data }   (data = base64 TANPA prefix data:)
 * Alur: validasi -> buat berkas fisik di folder Drive aplikasi -> simpan
 * File ID di PropertiesService + kolom `logo` sheet INSTANSI -> kembalikan
 * File ID agar frontend dapat memakainya kembali setelah refresh/reopen.
 */
function uploadLogoBytes(token, payload) {
  return safeRun_('Logo berhasil diunggah ke Google Drive.', function () {
    var sess = requireSession_(token);
    requireWrite_(sess, 'pengaturan');
    var p = payload || {};
    var mimeType = toStr_(p.mimeType).trim().toLowerCase();
    if (LOGO_MIME_OK_.indexOf(mimeType) === -1) {
      throwError_('Format logo tidak didukung. Gunakan PNG atau JPG.');
    }
    var b64 = toStr_(p.data).replace(/\s+/g, '');
    if (!isFilled_(b64)) throwError_('Data logo kosong.');
    /* Estimasi ukuran byte dari panjang base64 (tanpa decode penuh dulu). */
    var estBytes = Math.floor(b64.length * 3 / 4);
    if (estBytes > LOGO_MAX_BYTES_) {
      throwError_('Logo terlalu besar (maks 2 MB).');
    }
    var bytes;
    try {
      bytes = Utilities.base64Decode(b64);
    } catch (e) {
      throwError_('Data logo tidak valid (base64 rusak).');
    }
    if (!bytes || !bytes.length) throwError_('Data logo kosong.');
    if (bytes.length > LOGO_MAX_BYTES_) {
      throwError_('Logo terlalu besar (maks 2 MB).');
    }
    var folder = logoDriveFolder_();
    if (!folder) {
      throwError_('Folder Google Drive aplikasi tidak dapat dibuka. Pastikan akun memiliki akses Drive.');
    }
    var fileName = logoSafeName_(p.name, mimeType);
    var props = PropertiesService.getScriptProperties();
    var oldId = '';
    try { oldId = toStr_(props.getProperty(PROP_LOGO_FILE_ID)).trim(); } catch (e) { oldId = ''; }
    var logoFile;
    try {
      logoFile = folder.createFile(Utilities.newBlob(bytes, mimeType, fileName));
    } catch (e) {
      throwError_('Gagal membuat berkas logo di Google Drive. Pastikan akun memiliki akses tulis Drive.', e);
    }
    var newId = logoFile.getId();
    props.setProperty(PROP_LOGO_FILE_ID, newId);
    /* Berkas logo lama dibuang (trash) agar tidak menumpuk dan tidak
       mengganggu pemakaian logo baru. Kegagalan di sini tidak menggagalkan
       upload karena berkas baru sudah tersimpan dengan benar. */
    if (isFilled_(oldId) && oldId !== newId) {
      try { DriveApp.getFileById(oldId).setTrashed(true); } catch (e) {}
    }
    /* Referensi permanen di database: kolom `logo` sheet INSTANSI. */
    var inst = fetchInstansiRecord_();
    if (inst) {
      updateObject_(openSheet_(getDef_('INSTANSI')), getDef_('INSTANSI'), SINGLETON_KEY_INSTANSI, { logo: newId, updatedAt: nowIso_() });
    }
    refCacheBust_('instansi');
    bumpDataRev_('INSTANSI');
    return { fileId: newId, name: logoFile.getName(), url: logoFile.getUrl() };
  });
}

/**
 * Ambil ISI logo sebagai data URL (base64) agar frontend dapat
 * menampilkannya kembali setelah refresh/reopen tanpa menyimpan base64
 * raksasa di sel Sheets. Gagal membaca berkas TIDAK melempar error —
 * mengembalikan dataUrl kosong sehingga cetak tetap berjalan (tanpa logo).
 */
function getLogoImage(token) {
  return safeRun_('Logo instansi berhasil dibaca.', function () {
    requireSession_(token);
    var props = PropertiesService.getScriptProperties();
    var fileId = '';
    try { fileId = toStr_(props.getProperty(PROP_LOGO_FILE_ID)).trim(); } catch (e) { fileId = ''; }
    var inst = fetchInstansiRecord_();
    if (inst && isFilled_(inst.logo)) fileId = toStr_(inst.logo).trim();
    if (!isFilled_(fileId)) return { fileId: '', dataUrl: '' };
    try {
      var file = DriveApp.getFileById(fileId);
      var blob = file.getBlob();
      var mime = toStr_(blob.getContentType()).trim() || 'image/png';
      var dataUrl = 'data:' + mime + ';base64,' + Utilities.base64Encode(blob.getBytes());
      return { fileId: file.getId(), name: file.getName(), mimeType: mime, dataUrl: dataUrl, url: file.getUrl() };
    } catch (e) {
      logError_('getLogoImage', e);
      return { fileId: fileId, dataUrl: '' };
    }
  });
}

/** Dapatkan URL logo instansi. */
function getLogo(token) {
  return safeRun_('Logo instansi berhasil dibaca.', function () {
    requireSession_(token);
    var props = PropertiesService.getScriptProperties();
    var fileId = props.getProperty(PROP_LOGO_FILE_ID);
    var inst = fetchInstansiRecord_();
    if (inst && toStr_(inst.logo).trim()) fileId = toStr_(inst.logo);
    if (!fileId) return { logo: '' };
    try {
      var file = DriveApp.getFileById(fileId);
      return { fileId: file.getId(), name: file.getName(), url: file.getUrl() };
    } catch (e) {
      return { logo: '' };
    }
  });
}

/** Hapus logo dari Drive dan PropertiesService. */
function deleteLogo(token) {
  return safeRun_('Logo berhasil dihapus.', function () {
    var sess = requireSession_(token);
    requireWrite_(sess, 'pengaturan');
    var props = PropertiesService.getScriptProperties();
    var fileId = props.getProperty(PROP_LOGO_FILE_ID);
    if (fileId) {
      try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {}
      props.deleteProperty(PROP_LOGO_FILE_ID);
    }
    var inst = fetchInstansiRecord_();
    if (inst) {
      updateObject_(openSheet_(getDef_('INSTANSI')), getDef_('INSTANSI'), SINGLETON_KEY_INSTANSI, { logo: '', updatedAt: nowIso_() });
    }
    /* Cache referensi dibuang — frontend langsung melihat logo hilang. */
    refCacheBust_('instansi');
    return { success: true };
  });
}

/* ==========================================================================
 * END FILE ASLI: Logo.gs
 * ========================================================================== */

/* ==========================================================================
 * BEGIN FILE ASLI: WebApp.gs
 * ========================================================================== */
/**
 * ============================================================================
 * WEB APP / ENTRY POINT + HTTP API
 * ============================================================================
 * Modul ini melayani DUA mode sekaligus, tanpa menghilangkan yang lama:
 *
 * 1. MODE LAMA (tetap didukung penuh)
 *    doGet tanpa parameter menyajikan frontend (file HTML "Index") sehingga
 *    aplikasi masih bisa dipakai sebagai Web App seperti sebelumnya.
 *
 * 2. MODE BARU — API HTTP untuk frontend LOKAL
 *    Frontend lokal (dibuka dari localhost / file://) TIDAK memakai
 *    google.script.run, melainkan `fetch()` ke URL Web App ini.
 *
 *        Local Frontend  ->  fetch()  ->  Web App GAS  ->  Code.gs  ->  Sheets
 *
 *    doPost menerima JSON:
 *        { "action": "getSPPD", "args": ["<token>"] }
 *        { "action": "__batch", "calls": [ {action,args}, ... ] }
 *        { "action": "ping" }
 *    dan SELALU membalas dengan envelope yang konsisten:
 *        { "success": true,  "data": ..., "message": "..." }
 *        { "success": false, "data": null, "message": "...", "error": "..." }
 *
 * CATATAN CORS (penting untuk frontend lokal):
 *   Body dikirim sebagai `text/plain` (bukan `application/json`) agar termasuk
 *   "simple request" sehingga browser TIDAK mengirim preflight OPTIONS — GAS
 *   Web App tidak menangani OPTIONS. Google menambahkan header
 *   `Access-Control-Allow-Origin: *` pada respons Web App yang di-deploy
 *   dengan akses "Anyone", sehingga fetch dari localhost/file:// dapat membaca
 *   responsnya. Karena itu `doPost` juga membaca `e.parameter.payload` sebagai
 *   cadangan bila body tidak terkirim sebagai JSON.
 *
 * CATATAN KEAMANAN:
 *   - Endpoint ini TIDAK memuat ID/URL Spreadsheet pada respons publik.
 *   - Operasi administratif (setupDatabase, setSpreadsheetId, resetData,
 *     periksaDatabase, getDatabaseConfig) SENGAJA TIDAK diizinkan lewat HTTP —
 *     jalankan dari editor Apps Script.
 *   - Autentikasi tetap memakai token sesi aplikasi (fungsi requireSession_
 *     di tiap endpoint), jadi data tetap terlindungi.
 *
 * Catatan deployment:
 * - Nama file HTML di project GAS (bila ingin mode lama) harus "Index".
 * - Deploy -> New deployment -> Web app
 *     Execute as     : Me
 *     Who has access : Anyone        (diperlukan agar frontend lokal bisa fetch)
 * - Salin URL .../exec ke aplikasi: Pengaturan -> D. Konfigurasi Aplikasi.
 * ============================================================================
 */

/**
 * Daftar aksi yang BOLEH dipanggil lewat HTTP.
 * Sengaja berupa daftar putih (whitelist) agar fungsi internal tidak pernah
 * terekspos. Tambahkan aksi baru di sini bila ada endpoint baru.
 */
var API_ACTIONS_ = [
  /* Auth */
  'login', 'logout', 'getSession',
  /* Dashboard */
  'getDashboardData', 'getDataVersion',
  /* SPPD */
  'getSPPD', 'getSPPDById', 'createSPPD', 'updateSPPD', 'deleteSPPD', 'getSPPDRelations',
  /* Rampung */
  'getRampung', 'getRampungById', 'createRampung', 'updateRampung', 'deleteRampung',
  'verifikasiRampung', 'getRampungDraftFromSppd', 'createRampungFromSppd',
  /* Program */
  'getProgram', 'getProgramById', 'createProgram', 'updateProgram', 'deleteProgram',
  /* Standar Harga */
  'getStandarHarga', 'getStandarHargaById', 'createStandarHarga', 'updateStandarHarga',
  'deleteStandarHarga',
  /* Pejabat */
  'getPejabat', 'getPejabatById', 'updatePejabat', 'deletePejabat',
  /* Pengaturan */
  'getKop', 'getKopById', 'updateKop', 'deleteKop',
  'getInstansi', 'updateInstansi', 'getCetakConfig', 'updateCetakConfig',
  /* Dasar Hukum */
  'getDasarHukum', 'getDasarHukumById', 'updateDasarHukum', 'deleteDasarHukum',
  /* User */
  'getUsers', 'createUser', 'saveUser', 'updateUserPub', 'deleteUser',
  /* Logo */
  'uploadLogo', 'uploadLogoBytes', 'getLogo', 'getLogoImage', 'deleteLogo',
  /* Print (data dokumen dari sisi server) */
  'printSPPD', 'printRampung', 'printPreview'
];

/** Batas jumlah panggilan dalam satu batch (menjaga durasi eksekusi GAS). */
var API_BATCH_MAX_ = 20;

/* ==========================================================================
 * ENTRY POINT
 * ========================================================================== */

function doGet(e) {
  var action = apiParam_(e, 'action');

  /* Health check JSON — dipakai tombol "Tes Koneksi" dan pengecekan manual
     lewat browser: <URL>/exec?action=ping */
  if (action === 'ping' || action === 'health') {
    return jsonOut_(apiPing_());
  }

  /* MODE LAMA (UTAMA): sajikan frontend langsung dari project Apps Script.
     Nama berkas HTML yang lazim adalah "Index". Beberapa variasi penamaan
     tetap dicoba agar aplikasi langsung berjalan tanpa perlu menyesuaikan
     nama berkas terlebih dahulu. */
  var kandidatNama = ['Index', 'index', 'Index.html', 'index.html'];
  var galatTerakhir = null;
  for (var i = 0; i < kandidatNama.length; i++) {
    try {
      return HtmlService.createHtmlOutputFromFile(kandidatNama[i])
        .setTitle('SIAP SPPD & Rampung')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } catch (err) {
      galatTerakhir = err; /* coba nama berkas berikutnya */
    }
  }

  /* Tidak ada berkas HTML frontend pada project -> tampilkan halaman bantuan. */
  return HtmlService.createHtmlOutput(
    '<h3>Backend API aktif</h3>' +
    '<p>Web App ini berjalan sebagai <b>backend/API</b>. Frontend dijalankan secara lokal.</p>' +
    '<p>Cek koneksi: <a href="?action=ping" target="_blank">?action=ping</a></p>' +
    '<p>Bila ingin memakai mode lama (frontend di-host GAS), tambahkan file HTML bernama ' +
    '<b>Index</b> pada project Apps Script, lalu Deploy ulang.</p>' +
    '<pre style="white-space:pre-wrap">' +
    String(galatTerakhir && galatTerakhir.message ? galatTerakhir.message : galatTerakhir) + '</pre>'
  ).setTitle('SIAP SPPD & Rampung — Backend API');
}

function doPost(e) {
  try {
    var payload = parseApiPayload_(e);
    if (!payload) {
      return jsonOut_(fail_('Permintaan tidak berisi JSON yang valid.',
        'Body kosong atau bukan JSON. Kirim {"action":"...","args":[...]}.'));
    }

    var action = String(payload.action || '').trim();

    if (action === 'ping' || action === 'health') return jsonOut_(apiPing_());
    if (action === '__batch') return jsonOut_(apiBatch_(payload.calls));

    return jsonOut_(apiDispatch_(action, payload.args));
  } catch (err) {
    logError_('doPost', err);
    return jsonOut_(fail_('Terjadi kesalahan pada server API.', err));
  }
}

/* ==========================================================================
 * PARSER & KELUARAN
 * ========================================================================== */

/** Ambil parameter GET dengan aman. */
function apiParam_(e, name) {
  try {
    return (e && e.parameter && e.parameter[name]) ? String(e.parameter[name]).trim() : '';
  } catch (err) {
    return '';
  }
}

/**
 * Baca payload JSON dari body POST.
 * Urutan: (1) body mentah sebagai JSON — inilah yang dikirim frontend lokal,
 *         (2) `e.parameter.payload` sebagai cadangan (form-encoded).
 */
function parseApiPayload_(e) {
  var raw = '';
  try {
    raw = (e && e.postData && e.postData.contents) ? String(e.postData.contents) : '';
  } catch (err) { raw = ''; }

  if (raw) {
    try { return JSON.parse(raw); } catch (err1) { /* coba cara lain */ }
  }

  var alt = apiParam_(e, 'payload');
  if (alt) {
    try { return JSON.parse(alt); } catch (err2) { /* bukan JSON */ }
  }
  return null;
}

/** Balas JSON dengan envelope standar. */
function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ==========================================================================
 * DISPATCHER
 * ========================================================================== */

/** Benarkah aksi ini boleh dipanggil lewat HTTP? */
function isApiAction_(action) {
  if (!action) return false;
  for (var i = 0; i < API_ACTIONS_.length; i++) {
    if (API_ACTIONS_[i] === action) return true;
  }
  return false;
}

/** Ambil fungsi global berdasarkan nama (aman: hanya yang ada di whitelist). */
function resolveApiFunction_(action) {
  var fn = null;
  try {
    if (typeof globalThis !== 'undefined' && globalThis) fn = globalThis[action];
    if (typeof fn !== 'function' && typeof this !== 'undefined') fn = this[action];
  } catch (err) { fn = null; }
  return (typeof fn === 'function') ? fn : null;
}

/**
 * Jalankan satu aksi API dan SELALU kembalikan envelope.
 * Fungsi backend sudah mengembalikan {success,data,message} (lewat safeRun_),
 * jadi envelope itu diteruskan apa adanya. Bila suatu fungsi mengembalikan
 * nilai mentah, nilainya dibungkus menjadi envelope yang sah.
 */
function apiDispatch_(action, args) {
  var aksi = String(action || '').trim();
  if (!aksi) return fail_('Aksi API tidak boleh kosong.');

  if (!isApiAction_(aksi)) {
    return fail_('Aksi API tidak dikenal atau tidak diizinkan: "' + aksi + '".',
      'Aksi di luar whitelist ditolak demi keamanan.');
  }

  var fn = resolveApiFunction_(aksi);
  if (!fn) {
    return fail_('Fungsi backend "' + aksi + '" tidak ditemukan pada project Apps Script.',
      'Pastikan seluruh berkas .gs sudah terpasang, atau gunakan Code.gs gabungan.');
  }

  var argList = Array.isArray(args) ? args
    : (args === undefined || args === null ? [] : [args]);

  try {
    var hasil = fn.apply(null, argList);
    if (hasil && typeof hasil === 'object' && typeof hasil.success === 'boolean') {
      return hasil; /* envelope dari safeRun_ / fail_ — diteruskan apa adanya */
    }
    return ok_(hasil, 'Berhasil.');
  } catch (err) {
    logError_('apiDispatch_/' + aksi, err);
    return fail_((err && err.message) ? err.message : 'Terjadi kesalahan.', err);
  }
}

/**
 * Batch: banyak aksi dalam SATU permintaan HTTP.
 * Dipakai frontend saat boot agar tidak perlu 5+ round-trip berurutan.
 * Hasil: data berupa array envelope sesuai urutan permintaan.
 */
function apiBatch_(calls) {
  if (!Array.isArray(calls) || !calls.length) {
    return fail_('Daftar panggilan batch kosong.');
  }
  if (calls.length > API_BATCH_MAX_) {
    return fail_('Terlalu banyak panggilan dalam satu batch (maks ' + API_BATCH_MAX_ + ').');
  }

  var hasil = [];
  for (var i = 0; i < calls.length; i++) {
    var c = calls[i] || {};
    hasil.push(apiDispatch_(c.action, c.args));
  }
  return ok_(hasil, 'Batch selesai (' + hasil.length + ' panggilan).');
}

/* ==========================================================================
 * HEALTH CHECK — dipakai tombol "Tes Koneksi"
 * ========================================================================== */

/**
 * Memeriksa: API hidup, versi, waktu server, dan apakah Google Sheets
 * (database) benar-benar dapat dibuka serta memiliki Sheet yang diharapkan.
 * Tidak mengembalikan ID maupun URL Spreadsheet.
 */
function apiPing_() {
  var info = {
    api: 'SIAP SPPD & Rampung',
    versi: getAppVersion_(),
    waktuServer: new Date().toISOString(),
    zonaWaktu: Session.getScriptTimeZone(),
    sheetsTerhubung: false,
    sheetsPesan: '',
    database: '',
    jumlahSheet: 0,
    sheetHilang: []
  };

  try {
    var ss = getSpreadsheet_();
    info.sheetsTerhubung = true;
    info.database = ss.getName();
    var semua = ss.getSheets();
    info.jumlahSheet = semua.length;

    /* Pastikan Sheet inti benar-benar ada — bukti struktur database utuh. */
    var ada = {};
    for (var i = 0; i < semua.length; i++) {
      ada[String(semua[i].getName()).trim()] = true;
    }
    var wajib = sheetOrder_();
    for (var j = 0; j < wajib.length; j++) {
      var def = getDef_(wajib[j]);
      var nama = def ? def.name : wajib[j];
      if (!ada[String(nama).trim()]) info.sheetHilang.push(nama);
    }
    info.sheetsPesan = info.sheetHilang.length
      ? ('Database terbuka, tetapi ' + info.sheetHilang.length + ' Sheet belum ada. ' +
         'Jalankan setupDatabase() di editor Apps Script.')
      : 'Database Google Sheets terbuka dan seluruh Sheet tersedia.';
  } catch (err) {
    info.sheetsTerhubung = false;
    info.sheetsPesan = 'Database Google Sheets TIDAK dapat dibuka: ' +
      String(err && err.message ? err.message : err) +
      '. Jalankan setupDatabase() atau setSpreadsheetId() dari editor Apps Script.';
    logError_('apiPing_', err);
  }

  return ok_(info, info.sheetsTerhubung
    ? 'API aktif dan database terhubung.'
    : 'API aktif, tetapi database belum siap.');
}

/** Versi aplikasi dari Script Properties (bila belum ada, pakai default). */
function getAppVersion_() {
  try {
    var v = PropertiesService.getScriptProperties().getProperty(PROP_APP_VERSION);
    return v ? String(v) : '1.0.0';
  } catch (err) {
    return '1.0.0';
  }
}

/* ==========================================================================
 * END FILE ASLI: WebApp.gs
 * ========================================================================== */

/* ==========================================================================
 * BEGIN FILE ASLI: Api.gs
 * ========================================================================== */
/**
 * SIAP SPPD & RAMPUNG — API ENTRY POINT (INDEKS)
 * File   : Api.gs
 *
 * PENTING: File ini TIDAK lagi berisi deklarasi fungsi.
 * Seluruh fungsi publik yang dipanggil frontend lewat google.script.run
 * didefinisikan langsung sebagai fungsi GLOBAL pada file modulnya
 * (GAS menggabungkan semua file .gs ke satu namespace global):
 *
 *   Auth          : login, logout, getSession            (Auth.gs)
 *   Database      : setupDatabase, periksaDatabase,      (Database.gs)
 *                   getDatabaseConfig, setSpreadsheetId
 *   SPPD          : getSPPD, getSPPDById, createSPPD,    (SPPD.gs)
 *                   updateSPPD, deleteSPPD, getSPPDRelations
 *   Rampung       : getRampung, getRampungById,          (Rampung.gs)
 *                   createRampung, updateRampung, deleteRampung,
 *                   verifikasiRampung, getRampungDraftFromSppd,
 *                   createRampungFromSppd
 *   Program       : getProgram, getProgramById,          (Program.gs)
 *                   createProgram, updateProgram, deleteProgram
 *   StandarHarga  : getStandarHarga, getStandarHargaById,(StandarHarga.gs)
 *                   createStandarHarga, updateStandarHarga, deleteStandarHarga
 *   Pejabat       : getPejabat, getPejabatById,          (Pejabat.gs)
 *                   updatePejabat, deletePejabat
 *   Pengaturan    : getKop, getKopById, updateKop,       (Pengaturan.gs)
 *                   deleteKop, getInstansi, updateInstansi,
 *                   getCetakConfig, updateCetakConfig
 *   DasarHukum    : getDasarHukum, getDasarHukumById,    (DasarHukum.gs)
 *                   updateDasarHukum, deleteDasarHukum
 *   User          : getUsers, createUser, saveUser,      (User.gs)
 *                   updateUserPub, deleteUser
 *   Print         : printSPPD, printRampung, printPreview (Print.gs)
 *   Logo          : uploadLogo, getLogo, deleteLogo      (Logo.gs)
 *   Combined      : getDashboardData                      (Code.gs)
 *
 * Catatan: versi sebelumnya membungkus tiap fungsi dengan pemanggilan
 * `Modul.fungsi(...)` (mis. SPPD.getSPPD) — pola itu tidak cocok dengan
 * implementasi global dan membuat definisi ganda di namespace global.
 * Deklarasi ganda dihapus agar file ini aman pada urutan pemuatan apa pun.
 *
 * ============================================================================
 * AKSES API LEWAT HTTP (arsitektur Local Frontend + Google Sheets)
 * ============================================================================
 * Sejak frontend dapat dijalankan secara LOKAL, fungsi-fungsi di atas juga
 * dapat dipanggil lewat HTTP ke Web App (bukan hanya google.script.run).
 * Penghubungnya ada di WebApp.gs:
 *
 *   POST <URL_WEB_APP>/exec     Content-Type: text/plain
 *   { "action": "getSPPD", "args": ["<token>"] }
 *   -> { "success": true, "data": [...], "message": "..." }
 *
 *   POST { "action": "__batch", "calls": [ {"action":"getKop","args":["<token>"]}, ... ] }
 *   -> { "success": true, "data": [ <envelope>, <envelope> ], "message": "..." }
 *
 *   GET  <URL_WEB_APP>/exec?action=ping
 *   -> { "success": true, "data": { sheetsTerhubung, database, jumlahSheet, ... } }
 *
 * Hanya fungsi yang terdaftar pada `API_ACTIONS_` (WebApp.gs) yang boleh
 * dipanggil lewat HTTP. Fungsi administratif — setupDatabase, setSpreadsheetId,
 * resetData, periksaDatabase, getDatabaseConfig — SENGAJA dikecualikan dan
 * hanya dapat dijalankan dari editor Apps Script.
 */

/* ==========================================================================
 * COMBINED API: getDashboardData
 *
 * Ditempatkan di modul (bukan hanya di Code.gs) agar hasil generate ulang
 * Code.gs tetap memuatnya, dan agar opsi deploy multi-file tetap lengkap.
 * ========================================================================== */

function getDashboardData(token) {
  /* PENTING: helper bernama `safeRun_` (dengan garis bawah). Memanggil `safeRun`
     tanpa garis bawah akan melempar ReferenceError di runtime Google sehingga
     Dashboard SELALU gagal dimuat. */
  return safeRun_('Dashboard data berhasil dibaca.', function () {
    var sess = requireSession_(token);
    var result = {};

    /* Parallel fetch — semua data dibaca dalam satu execution context.
       SPPD & Rampung difilter kepemilikan: Admin = semua, Pengguna = miliknya. */
    if (canReadModule_(sess, 'sppd')) {
      result.sppd = filterOwnedList_(sess, fetchSPPDList_());
    } else {
      result.sppd = [];
    }

    if (canReadModule_(sess, 'rampung')) {
      result.rampung = filterOwnedList_(sess, fetchRampungList_());
    } else {
      result.rampung = [];
    }

    result.program = fetchProgramList_();
    result.standarHarga = fetchStandarHargaList_();
    result.pejabat = fetchPejabatList_();
    result.kop = fetchKopList_();
    result.dasarHukum = fetchDasarHukumList_();

    var instRec = fetchInstansiRecord_();
    result.instansi = instRec ? instansiFromRecord_(instRec) : instansiFromRecord_({});

    var cetakRec = fetchCetakConfigRecord_();
    result.cetakConfig = cetakRec ? cetakConfigFromRecord_(cetakRec) : cetakConfigFromRecord_({});

    return result;
  });
}

/**
 * VERSI DATA — dipakai frontend untuk memutuskan perlu-tidaknya memuat ulang.
 *
 * Sangat murah: hanya membaca Script Properties. Fungsi ini TIDAK membuka
 * Google Sheets dan tidak melakukan pembacaan range sama sekali, sehingga
 * frontend bisa sering memeriksanya tanpa membebani kuota Apps Script.
 *
 * Frontend membandingkan revisi ini dengan revisi data yang sedang di-cache:
 *   - sama    -> data cache dipakai, TIDAK ada permintaan data baru;
 *   - berbeda -> hanya data yang berubah itu yang dimuat ulang.
 */
function getDataVersion(token) {
  return safeRun_('Versi data berhasil dibaca.', function () {
    requireSession_(token);
    return { revs: readDataRevs_(), serverTime: nowIso_() };
  });
}

/* ==========================================================================
 * END FILE ASLI: Api.gs
 * ========================================================================== */
