/**
 * ============================================================================
 * SIAP SPPD & RAMPUNG — TRANSPORT API (HTTP JSON ke Web App GAS)
 * File   : assets/js/api.js
 * Dimuat : SETELAH assets/js/config.js, SEBELUM script aplikasi inline.
 *
 * Tugas: menyediakan satu fungsi `API.call(action, args)` yang meniru bentuk
 * `gasRequest(method, ...args)` secara persis (Promise, unwrap envelope,
 * pesan error yang sama). Dengan begitu tidak ada call site di index.html
 * yang perlu disunting selain dispatcher `gasRequest` itu sendiri.
 *
 * PERSYARATAN KONTRAK (sudah dipenuhi backend WebApp.gs):
 *   POST <backendUrl>          Content-Type: text/plain;charset=utf-8
 *        body: JSON { action, args:[...] }
 *   Response 200 JSON          { success, data, message, error? }
 *
 *   Body sengaja text/plain supaya browser TIDAK mengirim preflight OPTIONS,
 *   karena Google Apps Script tidak menangani OPTIONS. Jangan ganti ke
 *   application/json.
 *
 * Deploy GAS harus memilih akses "Anyone" agar Google menambahkan
 * header 'Access-Control-Allow-Origin: *' pada respons.
 * ============================================================================
 */
(function (global) {
  'use strict';

  var CFG = (global.APP_CONFIG || {});
  var KEYS = (CFG.storageKeys || {
    backendUrl: 'siap_sppd_backend_url',
    transport: 'siap_sppd_transport'
  });

  /* ---------------------------------------------------------------- helpers */
  function _lsGet(k) {
    try { return global.localStorage ? global.localStorage.getItem(k) : null; }
    catch (e) { return null; }
  }
  function _lsSet(k, v) {
    try { if (v === null || typeof v === 'undefined') { global.localStorage.removeItem(k); }
          else { global.localStorage.setItem(k, String(v)); } } catch (e) {}
  }
  function _trimUrl(u) { return String(u || '').trim(); }
  function _isExecUrl(u) {
    u = _trimUrl(u);
    return /^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/(exec|dev)$/.test(u);
  }

  /* ------------------------------------------------------------ konfigurasi */

  /** URL aktif: localStorage override -> config.js default -> kosong. */
  function getUrl() {
    var ov = _trimUrl(_lsGet(KEYS.backendUrl));
    if (ov) return ov;
    return _trimUrl(CFG.backendUrl);
  }
  function isConfigured() { return !!getUrl(); }

  /** Preferensi transport: localStorage override -> config.js default. */
  function getPreferredTransport() {
    var ov = _trimUrl(_lsGet(KEYS.transport));
    if (ov) return ov;
    return _trimUrl(CFG.transport) || 'auto';
  }

  /** Ketersediaan runtime google.script.run (mode GAS legacy). */
  function hasGoogleScriptRun() {
    return typeof global.google !== 'undefined'
        && !!(global.google.script && global.google.script.run);
  }

  /**
   * Tentukan mode koneksi aktif: 'gas' | 'web' | 'demo'.
   * 'auto' mengikuti: google.script.run > backendUrl terisi > demo.
   * Transport eksplisit ('gas'/'web'/'demo') selalu ditaati bila valid.
   */
  function resolveTransport() {
    var pref = getPreferredTransport();
    if (pref === 'demo') return 'demo';
    if (pref === 'gas') return hasGoogleScriptRun() ? 'gas' : (isConfigured() ? 'web' : 'demo');
    if (pref === 'web') return isConfigured() ? 'web' : (hasGoogleScriptRun() ? 'gas' : 'demo');
    /* 'auto' atau nilai lain */
    if (hasGoogleScriptRun()) return 'gas';
    if (isConfigured()) return 'web';
    return 'demo';
  }

  function transportLabel(t) {
    if (t === 'gas') return 'Google Apps Script (google.script.run)';
    if (t === 'web') return 'GitHub Pages + Web App GAS (REST API)';
    return 'Demo (localStorage, tanpa backend)';
  }

  /* ---------------------------------------------------------- fetch primitive */

  function makeTimeoutSignal(ms) {
    if (typeof global.AbortController === 'undefined') {
      return { signal: undefined, cancel: function () {} };
    }
    var ctrl = new global.AbortController();
    var timer = setTimeout(function () { try { ctrl.abort(); } catch (e) {} }, ms || 60000);
    return {
      signal: ctrl.signal,
      cancel: function () { clearTimeout(timer); }
    };
  }

  /**
   * Kirim satu payload ke backendUrl, kembalikan envelope penuh
   * { success, data, message, error? }.  Jangan unwrap di sini — biarkan
   * pemanggil memutuskan (batch butuh envelope, gasRequest butuh data).
   */
  function postEnvelope(payload) {
    var url = getUrl();
    if (!url) {
      return Promise.reject(new Error('URL backend belum dikonfigurasi. Isi assets/js/config.js atau jalankan SIAP.setBackendUrl("<URL>/exec").'));
    }
    var body;
    try { body = JSON.stringify(payload); }
    catch (e) { return Promise.reject(new Error('Payload tidak dapat diserialisasi: ' + (e && e.message ? e.message : e))); }

    var to = makeTimeoutSignal(CFG.requestTimeoutMs || 60000);
    return global.fetch(url, {
      method: 'POST',
      /* text/plain agar request termasuk "simple request" (tanpa preflight). */
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: body,
      /* GAS 302 -> script.googleusercontent.com; fetch harus follow. */
      redirect: 'follow',
      credentials: 'omit',
      signal: to.signal,
      cache: 'no-store'
    }).then(function (res) {
      if (!res.ok) {
        throw new Error('HTTP ' + res.status + ' dari backend. Periksa deployment Web App GAS.');
      }
      return res.json();
    }).catch(function (err) {
      if (err && err.name === 'AbortError') {
        throw new Error('Permintaan ke server terlalu lama (timeout). Silakan coba kembali.');
      }
      if (err instanceof TypeError) {
        /* fetch() melempar TypeError untuk masalah jaringan/CORS. */
        throw new Error('Tidak dapat menghubungi backend (jaringan / CORS). Pastikan URL diakhiri /exec dan deployment memakai akses "Anyone".');
      }
      throw err;
    }).then(function (env) {
      to.cancel();
      return env;
    }, function (err) {
      to.cancel();
      throw err;
    });
  }

  /* ------------------------------------------------------ API publik (SIAP) */

  /**
   * Satu panggilan aksi. Mengembalikan `data` (bukan envelope) — signature
   * identik dengan `gasRequest(method, ...args)` yang lama.
   */
  function call(action, argsArray) {
    return postEnvelope({ action: String(action || ''), args: Array.isArray(argsArray) ? argsArray : [] })
      .then(function (env) {
        if (!env || env.success !== true) {
          throw new Error((env && env.message) || 'Permintaan backend gagal.');
        }
        return env.data;
      });
  }

  /**
   * Batch server-side. Payload: {action:'__batch', calls:[{action,args}, ...]}
   * Mengembalikan array envelope sesuai urutan permintaan.
   * Backend membatasi 20 panggilan per batch — jika lebih, frontend harus
   * chunking (lihat chunkBatch).
   */
  function batchEnvelope(calls) {
    var list = (calls || []).filter(Boolean);
    if (!list.length) return Promise.resolve([]);
    return postEnvelope({ action: '__batch', calls: list }).then(function (env) {
      if (!env || env.success !== true) {
        throw new Error((env && env.message) || 'Batch gagal.');
      }
      return env.data || [];
    });
  }

  /** Potong daftar panggilan menjadi chunk <= max (default 20, sesuai server). */
  function chunkBatch(calls, max) {
    var limit = max || 20;
    var out = [];
    var cur = [];
    (calls || []).forEach(function (c) {
      if (!c) return;
      cur.push(c);
      if (cur.length >= limit) { out.push(cur); cur = []; }
    });
    if (cur.length) out.push(cur);
    return out;
  }

  /** Gabung hasil beberapa chunk batch menjadi satu array envelope. */
  function batchAll(calls) {
    var chunks = chunkBatch(calls, 20);
    var result = [];
    var chain = Promise.resolve();
    chunks.forEach(function (ch) {
      chain = chain.then(function () {
        return batchEnvelope(ch).then(function (envs) {
          Array.prototype.push.apply(result, envs || []);
        });
      });
    });
    return chain.then(function () { return result; });
  }

  /** GET ?action=ping — dipakai panel "Tes Koneksi". */
  function ping() {
    var url = getUrl();
    if (!url) {
      return Promise.reject(new Error('URL backend belum dikonfigurasi.'));
    }
    var sep = url.indexOf('?') === -1 ? '?' : '&';
    return global.fetch(url + sep + 'action=ping', {
      method: 'GET',
      credentials: 'omit',
      cache: 'no-store'
    }).then(function (r) {
      if (!r.ok) throw new Error('HTTP ' + r.status + ' dari backend.');
      return r.json();
    });
  }

  /* ------------------------------------------------------- kontrol runtime */

  function setBackendUrl(url) {
    var u = _trimUrl(url);
    if (!u) { throw new Error('URL kosong. Contoh: https://script.google.com/macros/s/AKfycb…/exec'); }
    if (!_isExecUrl(u)) {
      /* Perbolehkan tapi beri peringatan keras. */
      console.warn('[SIAP] URL tampak bukan Web App GAS standar. Pastikan URL berakhir /exec atau /dev.');
    }
    _lsSet(KEYS.backendUrl, u);
    return getUrl();
  }
  function clearBackendUrl() { _lsSet(KEYS.backendUrl, null); return getUrl(); }
  function setTransport(pref) {
    var p = _trimUrl(pref).toLowerCase();
    if (['auto', 'gas', 'web', 'demo'].indexOf(p) === -1) {
      throw new Error('Transport harus "auto" | "gas" | "web" | "demo".');
    }
    _lsSet(KEYS.transport, p);
    return getPreferredTransport();
  }
  function clearTransport() { _lsSet(KEYS.transport, null); return getPreferredTransport(); }
  function info() {
    return {
      transport: resolveTransport(),
      transportLabel: transportLabel(resolveTransport()),
      backendUrl: getUrl(),
      hasGoogleScriptRun: hasGoogleScriptRun(),
      configDefaultUrl: _trimUrl(CFG.backendUrl),
      configTransport: _trimUrl(CFG.transport) || 'auto',
      version: (CFG.app && CFG.app.version) || '',
      schema: (CFG.app && CFG.app.schemaVersion) || ''
    };
  }

  /* ------------------------------------------------------------- exports */

  var API = {
    call: call,
    batchEnvelope: batchEnvelope,
    batchAll: batchAll,
    chunkBatch: chunkBatch,
    ping: ping,
    getUrl: getUrl,
    isConfigured: isConfigured,
    hasGoogleScriptRun: hasGoogleScriptRun,
    resolveTransport: resolveTransport,
    transportLabel: transportLabel,
    getPreferredTransport: getPreferredTransport
  };
  global.API = API;

  /** Namespace kecil untuk dipanggil dari DevTools Console. */
  global.SIAP = {
    setBackendUrl: setBackendUrl,
    clearBackendUrl: clearBackendUrl,
    setTransport: setTransport,
    clearTransport: clearTransport,
    info: info,
    ping: ping,
    API: API
  };
})(window);
