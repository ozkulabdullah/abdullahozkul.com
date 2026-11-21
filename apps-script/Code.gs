/**
 * abdullahozkul.com — Skor tablosu API'si (v2: kullanıcı adı + PIN)
 * Google Apps Script web uygulaması. "abdullahozkul.com Skor Tablosu" dosyasını kullanır:
 *   - "Skorlar" sayfası: her oyun sonucu bir satır
 *   - "Oyuncular" sayfası: kullanıcı adı ve PIN özeti (PIN'in kendisi saklanmaz)
 *
 * GET  ?action=top&game=drone                         → en iyi 10
 * POST {action:'login',  name, pin}                    → adı ilk kez kullanıyorsa PIN'le kaydeder, yoksa PIN'i doğrular
 * POST {action:'submit', name, pin, game, score, device}
 */

const SHEET_ID = '1r-TSLvGwZJYlp8eX95w8gQIw_0xRp6BVOVbV7nJgC_w';
const GAMES = { drone: 5000, g2048: 1000000 };   // oyun: izin verilen en yüksek skor
const TOP_N = 10;
const MAX_FAILS = 5;          // hatalı PIN denemesi sınırı
const LOCK_SECONDS = 600;     // sınır aşılınca bekleme süresi

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    return json_({ ok: true, game: p.game, top: top_(p.game) });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  let p = {};
  try { p = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (_) {}
  try {
    if (p.action === 'login') return json_(login_(p));
    if (p.action === 'submit') return json_(submit_(p));
    throw new Error('Bilinmeyen işlem');
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

/* ---------- İşlemler ---------- */

function login_(p) {
  const a = auth_(p.name, p.pin);
  return { ok: true, name: a.name, created: a.created, best: bestFor_(a.key) };
}

function submit_(p) {
  const a = auth_(p.name, p.pin);
  const game = String(p.game || '');
  if (!(game in GAMES)) throw new Error('Bilinmeyen oyun');
  const score = Math.floor(Number(p.score));
  if (!isFinite(score) || score < 0 || score > GAMES[game]) throw new Error('Geçersiz skor');
  const device = String(p.device || '').slice(0, 40);

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    scores_().appendRow([new Date(), game, a.name, score, device]);
  } finally {
    lock.releaseLock();
  }
  CacheService.getScriptCache().remove('top_' + game);
  return { ok: true, game: game, top: top_(game), best: bestFor_(a.key) };
}

/* ---------- Kimlik (ad + PIN) ---------- */

function auth_(rawName, rawPin) {
  const name = cleanName_(rawName);
  if (name.length < 2) throw new Error('Kullanıcı adı en az 2 karakter olmalı');
  const pin = String(rawPin || '');
  if (!/^\d{4}$/.test(pin)) throw new Error('PIN 4 haneli bir sayı olmalı');

  const key = name.toLocaleLowerCase('tr');
  const cache = CacheService.getScriptCache();
  const failKey = 'fail_' + key;
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= MAX_FAILS) throw new Error('Çok fazla hatalı deneme. 10 dakika sonra tekrar dene.');

  const hash = hash_(key, pin);
  const sh = players_();
  let row = findPlayer_(sh, key);

  if (!row) {
    const lock = LockService.getScriptLock();
    lock.waitLock(5000);
    try {
      row = findPlayer_(sh, key);               // bu arada başkası almış olabilir
      if (!row) {
        sh.appendRow([key, name, hash, new Date()]);
        return { key: key, name: name, created: true };
      }
    } finally {
      lock.releaseLock();
    }
  }

  if (String(row[2]) !== hash) {
    cache.put(failKey, String(fails + 1), LOCK_SECONDS);
    throw new Error('Bu kullanıcı adı başka bir PIN ile alınmış. PIN hatalı.');
  }
  cache.remove(failKey);
  return { key: key, name: String(row[1]), created: false };
}

function findPlayer_(sh, key) {
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) if (String(rows[i][0]) === key) return rows[i];
  return null;
}

function hash_(key, pin) {
  const props = PropertiesService.getScriptProperties();
  let salt = props.getProperty('SALT');
  if (!salt) { salt = Utilities.getUuid(); props.setProperty('SALT', salt); }
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + key + '|' + pin, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

/* ---------- Skorlar ---------- */

function top_(game) {
  if (!(game in GAMES)) throw new Error('Bilinmeyen oyun');
  const cache = CacheService.getScriptCache();
  const hit = cache.get('top_' + game);
  if (hit) return JSON.parse(hit);

  const best = {};
  scores_().getDataRange().getValues().slice(1).forEach(function (r) {
    if (r[1] !== game) return;
    const name = String(r[2]);
    const key = name.toLocaleLowerCase('tr');
    const score = Number(r[3]) || 0;
    if (!best[key] || score > best[key].score) best[key] = { name: name, score: score };
  });
  const list = Object.keys(best).map(function (k) { return best[k]; })
    .sort(function (a, b) { return b.score - a.score; })
    .slice(0, TOP_N);
  cache.put('top_' + game, JSON.stringify(list), 60);
  return list;
}

function bestFor_(key) {
  const out = {};
  Object.keys(GAMES).forEach(function (g) { out[g] = 0; });
  scores_().getDataRange().getValues().slice(1).forEach(function (r) {
    if (String(r[2]).toLocaleLowerCase('tr') !== key || !(r[1] in out)) return;
    out[r[1]] = Math.max(out[r[1]], Number(r[3]) || 0);
  });
  return out;
}

/* ---------- Yardımcılar ---------- */

function scores_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheetByName('Skorlar') || ss.getSheets()[0];
}

function players_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName('Oyuncular');
  if (!sh) {
    sh = ss.insertSheet('Oyuncular');
    sh.getRange('A:C').setNumberFormat('@');      // adlar ve özetler metin olarak kalsın
    sh.appendRow(['anahtar', 'kullanici', 'pin_ozeti', 'olusturma']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function cleanName_(s) {
  return String(s || '')
    .replace(/[<>"'`\\=+@]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 16);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
