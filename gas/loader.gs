/* ============================================================
 *  お金 GAS ローダー v1  —  これを1回貼れば、以後この画面を触らなくていい
 *  ------------------------------------------------------------
 *  本体コードは GitHub から自動で取ってきて実行する:
 *    https://github.com/gogojoe312-max/okane  →  gas/code.gs
 *  Claude が GitHub に push すれば、そのまま本番に反映される（最大6時間、
 *  すぐ反映したい時は reloadCode() を1回実行）。
 * ============================================================ */

var SRC_URL = 'https://raw.githubusercontent.com/gogojoe312-max/okane/main/gas/code.gs';
var CACHE_KEY = 'okane_src_v1';
var CACHE_SEC = 21600; // 6時間

/* ---------- 本体コードの取得 ---------- */
function fetchSrc_() {
  var r = UrlFetchApp.fetch(SRC_URL + '?cb=' + Date.now(), { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('GitHubから取得失敗 HTTP ' + r.getResponseCode());
  var s = r.getContentText();
  if (s.indexOf('function doPost') < 0 || s.indexOf('function scanMail') < 0) {
    throw new Error('取得したコードが不正（doPost/scanMailが無い）');
  }
  saveBackup_(s);
  return s;
}
function src_() {
  var c = CacheService.getScriptCache();
  var s = c.get(CACHE_KEY);
  if (s) return s;
  try { s = fetchSrc_(); }
  catch (e) { s = loadBackup_(); if (!s) throw e; }   // GitHubが落ちてても直前の内容で動く
  c.put(CACHE_KEY, s, CACHE_SEC);
  return s;
}
function reloadCode() {
  CacheService.getScriptCache().remove(CACHE_KEY);
  var s = fetchSrc_();
  CacheService.getScriptCache().put(CACHE_KEY, s, CACHE_SEC);
  Logger.log('reload ok: ' + s.length + '文字 / ' + s.split('\n')[0]);
  return s.length;
}

/* ---------- バックアップ（スクリプトプロパティに分割保存） ---------- */
function saveBackup_(s) {
  var p = PropertiesService.getScriptProperties();
  var size = 8000, n = Math.ceil(s.length / size), o = { SRC_N: String(n) };
  for (var i = 0; i < n; i++) o['SRC_' + i] = s.substr(i * size, size);
  p.setProperties(o);
}
function loadBackup_() {
  var p = PropertiesService.getScriptProperties(), n = Number(p.getProperty('SRC_N') || 0);
  if (!n) return '';
  var s = '';
  for (var i = 0; i < n; i++) s += (p.getProperty('SRC_' + i) || '');
  return s;
}

/* ---------- 本体の関数を呼ぶ ---------- */
function call_(name, arg) {
  var mine = { doPost: doPost, scanMail: scanMail, setup: setup };
  var f = eval(src_() + '\n;' + name + ';');
  if (typeof f !== 'function' || f === mine[name]) throw new Error('本体コードの読込に失敗: ' + name);
  return f(arg);
}

/* ---------- エントリポイント ---------- */
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents);
    var tok = PropertiesService.getScriptProperties().getProperty('TOKEN');
    if (req.action === 'reload') {
      if (req.token !== tok) return json_({ error: 'bad token' });
      return json_({ ok: true, len: reloadCode() });
    }
    if (req.action === 'srcinfo') {
      if (req.token !== tok) return json_({ error: 'bad token' });
      var s = src_();
      return json_({ ok: true, len: s.length, head: s.split('\n')[0].slice(0, 100) });
    }
  } catch (err) { /* 通常のリクエストは下へ */ }
  try { return call_('doPost', e); }
  catch (err2) { return json_({ error: String(err2 && err2.message || err2) }); }
}
function doGet() {
  return ContentService.createTextOutput('okane gas loader: ok').setMimeType(ContentService.MimeType.TEXT);
}
function scanMail() { return call_('scanMail', 3); }
function setup() { return call_('setup'); }

/* ---------- 権限（OAuthスコープ）を今までと同じにするためのダミー ---------- */
function _scopes_() {
  if (new Date().getTime() < 0) {
    GmailApp.search('x'); GmailApp.createLabel('x');
    SpreadsheetApp.create('x'); SpreadsheetApp.openById('x');
    UrlFetchApp.fetch('https://example.com');
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  }
}
