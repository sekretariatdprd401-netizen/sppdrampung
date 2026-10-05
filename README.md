# SIAP SPPD & RAMPUNG

**Sistem Administrasi Perjalanan Dinas** — aplikasi web untuk mengelola SPPD
(Surat Perintah Perjalanan Dinas), Rampung (Laporan Pertanggungjawaban),
Program/Kegiatan, Standar Harga, Pejabat, Dasar Hukum, dan Pengaturan
instansi, dengan Google Sheets sebagai database utama.

Aplikasi ini didistribusikan dalam dua komponen yang saling terhubung
lewat HTTPS JSON:

```
Browser  →  GitHub Pages (frontend statis)
                │
                │ fetch POST /exec   Content-Type: text/plain
                │ { action, args:[…] }
                ▼
          Google Apps Script Web App (backend REST API)
                │
                ├── SpreadsheetApp  →  Google Sheets  (database 15 sheet)
                ├── DriveApp        →  Google Drive   (logo instansi)
                ├── CacheService / PropertiesService → sesi + config
                └── LockService     → serialisasi operasi tulis
```

## Struktur repository

```
.
├── index.html                     # SPA frontend (Bootstrap 5 + Chart.js dari CDN)
├── assets/
│   └── js/
│       ├── config.js              # URL backend & preferensi transport
│       └── api.js                 # Transport HTTP (fetch → Web App GAS)
├── docs/
│   ├── arsitektur.md
│   ├── panduan-deployment-github.md
│   ├── panduan-deployment-gas.md
│   ├── panduan-setup-sheets.md
│   ├── panduan-pengujian.md
│   └── keamanan.md
├── Code.gs                        # Backend gabungan 16 modul — sumber kebenaran
├── appsscript.json                # Manifest project GAS (scopes, web-app config)
├── .nojekyll                      # nonaktifkan Jekyll pada GitHub Pages
├── .gitignore
└── README.md
```

> `Code.gs` dan `appsscript.json` berada di root project supaya dapat langsung
> dipakai `clasp push` maupun paste manual ke editor Apps Script. Tidak ada
> duplikasi ke subfolder `backend/` — menghindari divergensi dua sumber
> kebenaran (lihat prompt §6: *jangan memaksakan pemisahan file bila justru
> membuat aplikasi tidak stabil*).

## Fitur aplikasi

| Modul            | Cakupan |
|------------------|---------|
| Dashboard        | Ringkasan SPPD + Rampung, versi data, grafik |
| SPPD             | CRUD + pelaksana (child) + status + relasi |
| Rampung          | CRUD + rincian per pegawai + komponen biaya + verifikasi + draft-dari-SPPD |
| Program          | CRUD + daftar kode rekening (child) |
| Standar Harga    | CRUD per daerah tujuan |
| Pejabat          | PA / PPTK / Bendahara / SPT / SPD |
| Dasar Hukum      | Daftar dasar hukum yang dikutip di dokumen |
| Pengaturan       | Identitas instansi, kop surat (multi), konfigurasi cetak, logo |
| User             | Manajemen akun Admin & Pengguna (khusus admin) |
| Cetak            | 11 template HTML siap `window.print()` (SPT, SPD, Kwitansi, Rill, Rekap) |

## Cara menjalankan

### 1. Frontend — GitHub Pages
Lihat [`docs/panduan-deployment-github.md`](docs/panduan-deployment-github.md).

### 2. Backend — Google Apps Script Web App
Lihat [`docs/panduan-deployment-gas.md`](docs/panduan-deployment-gas.md).

### 3. Database — Google Sheets
Lihat [`docs/panduan-setup-sheets.md`](docs/panduan-setup-sheets.md).
Dibuat otomatis oleh `setupDatabase()` dari editor Apps Script — **tidak ada
data dummy**.

## Konfigurasi frontend

Setelah backend di-deploy, tempel URL `/exec` ke salah satu:

1. **`assets/js/config.js`** → `backendUrl: 'https://script.google.com/macros/s/…/exec'`, commit.
2. **DevTools Console** pada halaman yang sudah terbuka (override per-browser, disimpan di `localStorage`):
   ```js
   SIAP.setBackendUrl('https://script.google.com/macros/s/…/exec');
   SIAP.info();     // konfirmasi transport aktif
   SIAP.ping();     // panggil health check ?action=ping
   ```

## Transport & mode koneksi

Aplikasi memilih transport secara otomatis saat boot:

| Kondisi saat boot                       | Transport aktif | Keterangan |
|-----------------------------------------|-----------------|------------|
| Disajikan oleh Apps Script (ada `google.script.run`) | `gas`  | Mode legacy, tetap didukung penuh |
| URL backend terisi, tidak ada `google.script.run`   | `web`  | Mode migrasi — GitHub Pages ke Web App GAS |
| Keduanya kosong                          | `demo` | Database tiruan di `localStorage` — untuk uji coba offline |

Override manual dari console:
```js
SIAP.setTransport('auto');   // perilaku default
SIAP.setTransport('web');    // paksa REST API (butuh backendUrl terisi)
SIAP.setTransport('demo');   // paksa mode demo
```

## Login default

Akun admin awal dibuat oleh `setupDatabase()` bila sheet `USER` masih kosong:

- Username : `admin`
- Password : `admin`

**Wajib diganti** setelah login pertama (lihat [`docs/keamanan.md`](docs/keamanan.md)).
Pengguna non‑admin: username = NIP, password = NIP (atau password khusus bila
diisi admin di menu User).

## Keamanan — ringkasan

- Repository **tidak** menyimpan Spreadsheet ID, API key, service account,
  atau kredensial apa pun. URL `/exec` memang publik — itu memang
  disengaja; endpoint menolak semua aksi selain `login` dan `ping` tanpa
  token sesi.
- Token sesi diterbitkan server, disimpan di `CacheService` + `ScriptProperties`,
  TTL 8 jam, diverifikasi setiap request lewat `requireSession_()`.
- Hak akses ditegakkan **di backend** (bukan hanya menyembunyikan tombol di
  frontend): `canReadModule_`, `canWriteModule_`, `requireAdmin_`,
  `filterOwnedList_` (baris `createdBy`).
- `setupDatabase`, `setSpreadsheetId`, `resetData`, `periksaDatabase`,
  `getDatabaseConfig` **tidak ada di whitelist API** dan hanya bisa
  dijalankan dari editor Apps Script.

Rincian penuh + daftar isu yang diketahui ada di [`docs/keamanan.md`](docs/keamanan.md).

## Pembaruan aplikasi

| Yang berubah                          | Langkah |
|---------------------------------------|---------|
| Frontend (HTML/CSS/JS)                | `git commit` → `git push` → GitHub Pages rebuild otomatis (± 1 menit) |
| Backend (`Code.gs`)                   | Upload ke project Apps Script (paste manual atau `clasp push`) → **`Deploy → Manage deployments → Edit → New version`**. JANGAN `New deployment` — URL `/exec` akan berubah. |
| Skema database (tambah sheet / kolom) | Edit `sheetDefs_()` di `Config.gs` → jalankan `setupDatabase()` dari editor Apps Script. Struktur lama tidak dihapus; kolom baru ditambahkan ke kanan. |

## Pengujian regresi

Daftar periksa lengkap sebelum/ sesudah deploy ada di
[`docs/panduan-pengujian.md`](docs/panduan-pengujian.md).

## Lisensi

Changelog & lisensi dapat ditambahkan pada iterasi berikutnya
(`LICENSE`, `docs/changelog.md`).
