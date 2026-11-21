/**
 * abdullahozkul.com — Skor tablosu API'si
 * Google Apps Script web uygulaması. Skorları "abdullahozkul.com Skor Tablosu"
 * adlı Google Sheets dosyasına yazar ve en iyi skorları döndürür.
 *
 * Kullanım (GET):
 *   ?action=top&game=drone
 *   ?action=submit&game=g2048&name=Ali&score=1234&device=xyz
 */

const SHEET_ID = '1r-TSLvGwZJYlp8eX95w8gQIw_0xRp6BVOVbV7nJgC_w';
const GAMES = { drone: 5000, g2048: 1000000 };   // oyun: izin verilen en yüksek skor
const TOP_N = 10;

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === 'submit') return json_(submit_(p));
    return json_({ ok: true, game: p.game, top: top_(p.game) });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function submit_(p) {
  const game = String(p.game || '');
  if (!(game in GAMES)) throw new Error('Bilinmeyen oyun');
  const name = cleanName_(p.name);
  if (name.length < 2) throw new Error('Kullanıcı adı en az 2 karakter olmalı');
  const score = Math.floor(Number(p.score));
  if (!isFinite(score) || score < 0 || score > GAMES[game]) throw new Error('Geçersiz skor');
  const device = String(p.device || '').slice(0, 40);

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    sheet_().appendRow([new Date(), game, name, score, device]);
  } finally {
    lock.releaseLock();
  }
  CacheService.getScriptCache().remove('top_' + game);
  return { ok: true, game: game, top: top_(game) };
}

function top_(game) {
  if (!(game in GAMES)) throw new Error('Bilinmeyen oyun');
  const cache = CacheService.getScriptCache();
  const hit = cache.get('top_' + game);
  if (hit) return JSON.parse(hit);

  const rows = sheet_().getDataRange().getValues().slice(1);
  const best = {};
  rows.forEach(function (r) {
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

function sheet_() {
  return SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
}

function cleanName_(s) {
  return String(s || '')
    .replace(/[<>"'`\\]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 16);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
