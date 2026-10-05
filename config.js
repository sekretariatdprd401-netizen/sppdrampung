/**
 * ============================================================================
 * SIAP SPPD & RAMPUNG — KONFIGURASI TERPUSAT FRONTEND
 * File   : assets/js/config.js
 * Dimuat : sebelum script aplikasi lainnya (lihat index.html <head>).
 *
 * TUGAS FILE INI: menyediakan SATU tempat untuk semua nilai yang berubah
 * per-deployment. Frontend TIDAK boleh menyimpan kredensial, token, atau
 * Spreadsheet ID di sini — semua rahasia tetap di sisi server (GAS).
 * ============================================================================
 */
(function (global) {
  'use strict';

  /**
   * Nilai default. Ini di-commit ke repository GitHub, jadi:
   *   - backendUrl boleh kosong di awal; pengisiannya lewat salah satu dari
   *     dua jalur di bawah.
   *   - JANGAN taruh Spreadsheet ID, API key, service account JSON, atau
   *     password di file ini. Repository ini publik setelah deploy.
   *
   * CARA MENGISI backendUrl:
   *   1) (Paling mudah) Edit file ini, tempel URL Web App GAS, commit.
   *        backendUrl: 'https://script.google.com/macros/s/AKfycbXXXX/exec'
   *   2) (Tanpa rebuild) Buka DevTools Console di halaman aplikasi dan jalankan:
   *        SIAP.setBackendUrl('https://script.google.com/macros/s/.../exec');
   *     Nilai disimpan di localStorage browser itu dan MENIMPA config.js.
   *     Untuk menghapus override: SIAP.clearBackendUrl();
   */
  var DEFAULT_CONFIG = {
    /**
     * URL Web App GAS yang sudah di-deploy (wajib diakhiri '/exec').
     * Biarkan '' bila ingin diisi lewat localStorage per-browser.
     */
    backendUrl: '',

    /**
     * Transport mana yang dipakai saat boot.
     *   'auto'  : (default) — deteksi otomatis.
     *             google.script.run tersedia  -> 'gas'
     *             backendUrl terisi           -> 'web'
     *             keduanya tidak              -> 'demo' (localStorage tiruan)
     *   'gas'   : paksa google.script.run (hanya berlaku saat disajikan GAS).
     *   'web'   : paksa fetch() ke backendUrl.
     *   'demo'  : paksa mode demo offline (TANPA backend; data tidak ke Sheets).
     */
    transport: 'auto',

    /** Timeout fetch() per request (ms). Samai pola lama: 60.000 ms. */
    requestTimeoutMs: 60000,

    /** Kunci localStorage untuk override backendUrl + preferensi transport. */
    storageKeys: {
      backendUrl: 'siap_sppd_backend_url',
      transport: 'siap_sppd_transport'
    },

    /** Metadata ringkas untuk halaman "Tentang" / diagnosa. */
    app: {
      name: 'SIAP SPPD & RAMPUNG',
      version: '1.0.0',
      schemaVersion: '2026.09.16'
    }
  };

  /* Jangan timpa kalau sudah dimuat (mis. hot-reload). */
  if (!global.APP_CONFIG) {
    global.APP_CONFIG = DEFAULT_CONFIG;
  } else {
    /* Merge dangkal — biarkan properti yang sudah diset pengguna bertahan. */
    for (var k in DEFAULT_CONFIG) {
      if (Object.prototype.hasOwnProperty.call(DEFAULT_CONFIG, k) &&
          typeof global.APP_CONFIG[k] === 'undefined') {
        global.APP_CONFIG[k] = DEFAULT_CONFIG[k];
      }
    }
  }
})(window);
