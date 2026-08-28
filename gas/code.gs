/* ================== お金 GAS v1 ==================
 * セットアップ:
 * 1. gogojoe312 の Google アカウントで script.google.com → 新規プロジェクト
 * 2. このコードを貼る
 * 3. プロジェクト設定 → スクリプト プロパティに2つ追加:
 *      ANTHROPIC_API_KEY = (AnthropicのAPIキー)
 *      TOKEN             = (合言葉。アプリの設定にも同じ文字列を入れる)
 * 4. 一度 setup() を手動実行（シート作成 + Gmail権限の承認）
 * 5. デプロイ → 新しいデプロイ → ウェブアプリ
 *      実行ユーザー: 自分 / アクセス: 全員
 *      → 出たURLをアプリの設定「GAS Web App URL」に貼る
 * 6. トリガー → 追加 → scanMail / 時間主導型 / 分タイマー / 15分おき
 * ================================================== */

var MODEL_CHAT = 'claude-sonnet-4-6';
var MODEL_LIGHT = 'claude-haiku-4-5-20251001';
var SHEET_NAME = 'tx';
var HEAD = ['id','date','time','amount','store','cat','keihi','src','card','memo','createdAt','updatedAt'];
var MAIL_QUERY = 'newer_than:2d -label:OKANE_DONE (from:rakuten-card.co.jp OR from:mail.rakuten-card.co.jp OR from:vpass.ne.jp OR from:smbc-card.com OR from:contact.vpass.ne.jp OR (from:amazon.co.jp subject:ご注文))';
var DEFAULT_CATS = ['食費','日用品','交通','趣味・娯楽','サブスク・固定費','その他'];

/* ---------- entry ---------- */
function doPost(e){
  var out;
  try{
    var req = JSON.parse(e.postData.contents);
    if(req.token !== prop_('TOKEN')) throw new Error('bad token');
    var p = req.payload || {};
    switch(req.action){
      case 'ping':   out = {ok:true, ver:'gas-v1'}; break;
      case 'pull':   out = pull_(p); break;
      case 'add':    out = add_(p); break;
      case 'del':    out = del_(p); break;
      case 'ai':     out = ai_(p); break;
      default: throw new Error('unknown action');
    }
  }catch(err){ out = {error: String(err && err.message || err)}; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
function prop_(k){ return PropertiesService.getScriptProperties().getProperty(k) || ''; }

/* ---------- sheet ---------- */
function setup(){
  sheet_(); // create
  GmailApp.createLabel('OKANE_DONE');
  Logger.log('setup ok: ' + SpreadsheetApp.openById(prop_('SHEET_ID')).getUrl());
}
function sheet_(){
  var id = prop_('SHEET_ID');
  var ss;
  if(id){ ss = SpreadsheetApp.openById(id); }
  else{
    ss = SpreadsheetApp.create('お金DB');
    PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  }
  var sh = ss.getSheetByName(SHEET_NAME);
  if(!sh){ sh = ss.insertSheet(SHEET_NAME); sh.appendRow(HEAD); }
  return sh;
}
function rows_(){
  var sh = sheet_();
  var v = sh.getDataRange().getValues();
  var out = [];
  for(var i=1;i<v.length;i++){
    var r = v[i]; if(!r[0]) continue;
    var o = {}; HEAD.forEach(function(h,j){ o[h]=r[j]; });
    o.amount = Number(o.amount)||0; o.keihi = (o.keihi===true||o.keihi==='TRUE'||o.keihi==='true'||o.keihi===1);
    o.createdAt = Number(o.createdAt)||0; o.updatedAt = Number(o.updatedAt)||0;
    if(o.date instanceof Date) o.date = Utilities.formatDate(o.date,'Asia/Tokyo','yyyy-MM-dd');
    out.push(o);
  }
  return out;
}
function pull_(p){
  var since = Number(p.since)||0;
  var all = rows_();
  var tx = since ? all.filter(function(t){ return (t.updatedAt||t.createdAt) > since; }) : all;
  return { tx: tx, now: Date.now() };
}
function add_(p){
  var sh = sheet_();
  var existing = {}; rows_().forEach(function(t){ existing[t.id]=true; });
  var n=0, up=0;
  (p.tx||[]).forEach(function(t){
    if(!t || !t.id) return;
    t.updatedAt = Date.now();
    if(existing[t.id]){ updateRow_(sh, t); up++; }
    else{ sh.appendRow(HEAD.map(function(h){ return t[h]!==undefined?t[h]:''; })); n++; }
  });
  return { added:n, updated:up };
}
function updateRow_(sh, t){
  var v = sh.getDataRange().getValues();
  for(var i=1;i<v.length;i++){
    if(v[i][0]===t.id){
      sh.getRange(i+1,1,1,HEAD.length).setValues([HEAD.map(function(h){ return t[h]!==undefined?t[h]:v[i][HEAD.indexOf(h)]; })]);
      return;
    }
  }
}
function del_(p){
  var sh = sheet_();
  var v = sh.getDataRange().getValues();
  for(var i=v.length-1;i>=1;i--){ if(v[i][0]===p.id) sh.deleteRow(i+1); }
  return { ok:true };
}

/* ---------- Claude ---------- */
function claude_(model, system, messages, maxTok){
  var key = prop_('ANTHROPIC_API_KEY');
  if(!key) throw new Error('ANTHROPIC_API_KEY未設定');
  var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages',{
    method:'post', contentType:'application/json',
    headers:{ 'x-api-key':key, 'anthropic-version':'2023-06-01' },
    payload: JSON.stringify({ model:model, max_tokens:maxTok||1200, system:system||'', messages:messages }),
    muteHttpExceptions:true
  });
  var j = JSON.parse(res.getContentText());
  if(j.error) throw new Error(j.error.message||'API error');
  var txt=''; (j.content||[]).forEach(function(b){ if(b.type==='text') txt+=b.text; });
  return txt;
}
function tryJson_(s){
  try{ return JSON.parse(s.replace(/```json|```/g,'').trim()); }catch(e){ return null; }
}

/* ---------- ai actions ---------- */
function ai_(p){
  if(p.mode==='chat')   return aiChat_(p);
  if(p.mode==='receipt')return aiReceipt_(p);
  if(p.mode==='review') return aiReview_(p);
  throw new Error('unknown ai mode');
}
function aiChat_(p){
  var sys = 'あなたは家計アプリ「お金」のAI。ユーザーの支出データと予算を根拠に、短く率直に日本語で答える。'
    +'金額の質問には具体的な数字で答える。「買っていい？」には 買え/待て/やめとけ のどれかを冒頭に置き根拠を1-3行。'
    +'ユーザーの発言が質問でなく支出の記録（例: ラーメン980円払った）なら、応答の最後に <TX>[{"date":"YYYY-MM-DD","amount":数値,"store":"店名","cat":"カテゴリ","keihi":false}]</TX> を1回だけ付ける。'
    +'カテゴリは次から選ぶ: '+DEFAULT_CATS.join('/')+'。データに無いことは推測せず「データに無い」と言う。\n【現在のデータ】\n'
    + JSON.stringify(p.context||{});
  var text = claude_(MODEL_CHAT, sys, p.messages||[], 1200);
  var tx = null;
  var m = text.match(/<TX>([\s\S]*?)<\/TX>/);
  if(m){ tx = tryJson_(m[1]); text = text.replace(m[0],'').trim(); }
  return { text:text, tx:tx||[] };
}
function aiReceipt_(p){
  var sys='レシート/明細画像を読み、JSONだけを返す（前置き・コードブロック禁止）。'
    +'形式: {"tx":[{"date":"YYYY-MM-DD","amount":合計金額の数値,"store":"店名","cat":"カテゴリ","items":"主要品目を、区切りで"}]}。'
    +'カテゴリ候補: '+(p.cats||DEFAULT_CATS).join('/')+'。読めない項目は省略。日付が無ければ省略。';
  var msg=[{role:'user',content:[
    {type:'image',source:{type:'base64',media_type:p.mediaType||'image/jpeg',data:p.image}},
    {type:'text',text:'このレシートを読み取って'}]}];
  var j = tryJson_(claude_(MODEL_CHAT, sys, msg, 800)) || {};
  return { tx: j.tx||[] };
}
function aiReview_(p){
  var sys='家計アプリのAI。渡されたデータで今期のレビューを日本語で書く。構成: 収支の結論(黒字/赤字と額)→良かった点1つ→悪かった点(具体的な店・カテゴリ・金額)→来期への提案1つ。350字以内。データに無いことは書かない。';
  var text = claude_(MODEL_CHAT, sys, [{role:'user',content:JSON.stringify(p.context||{})}], 800);
  return { text:text };
}

/* ---------- Gmail scan (15分トリガー) ---------- */
function scanMail(){
  var label = GmailApp.getUserLabelByName('OKANE_DONE') || GmailApp.createLabel('OKANE_DONE');
  var threads = GmailApp.search(MAIL_QUERY, 0, 20);
  if(!threads.length) return;
  var sh = sheet_();
  var existing = {}; rows_().forEach(function(t){ existing[t.id]=true; });
  threads.forEach(function(th){
    th.getMessages().forEach(function(msg){
      try{
        var res = parseMail_(msg);
        (res||[]).forEach(function(t){
          if(!t || !t.amount) return;
          t.id = 'm' + msg.getId() + (t.sub||'');
          if(existing[t.id]) return;
          t.src='mail'; t.createdAt=Date.now(); t.updatedAt=Date.now();
          t.cat = t.cat || autoCat_(t.store||'');
          t.keihi = !!t.keihi;
          sh.appendRow(HEAD.map(function(h){ return t[h]!==undefined?t[h]:''; }));
          existing[t.id]=true;
        });
      }catch(e){ /* skip broken mail */ }
    });
    th.addLabel(label);
  });
}
function parseMail_(msg){
  var from = msg.getFrom();
  var sub = msg.getSubject();
  var body = msg.getPlainBody().slice(0, 4000);
  var dateStr = Utilities.formatDate(msg.getDate(),'Asia/Tokyo','yyyy-MM-dd');
  var timeStr = Utilities.formatDate(msg.getDate(),'Asia/Tokyo','HH:mm');

  // 1) regex quick pass (よくある形式)
  var m = body.match(/利用金額[:：]?\s*([0-9,]+)\s*円/) || body.match(/([0-9,]+)\s*円.{0,10}(利用|決済)/);
  var storeM = body.match(/利用先[:：]?\s*(.+)/) || body.match(/加盟店名?[:：]?\s*(.+)/);
  var dM = body.match(/利用日(?:時)?[:：]?\s*(\d{4})[\/年](\d{1,2})[\/月](\d{1,2})/);
  if(m){
    return [{
      date: dM ? (dM[1]+'-'+('0'+dM[2]).slice(-2)+'-'+('0'+dM[3]).slice(-2)) : dateStr,
      time: timeStr,
      amount: parseInt(m[1].replace(/,/g,''),10),
      store: storeM ? storeM[1].trim().slice(0,40) : '',
      card: cardLabel_(from, body)
    }];
  }
  // 2) AIフォールバック（読めなかったメールだけ・軽量モデル）
  var sys='カード利用通知/Amazon注文確認メールから支出をJSONだけで抽出（前置き禁止）。'
    +'形式: {"tx":[{"date":"YYYY-MM-DD","amount":数値,"store":"店名/品名","sub":"同メール内で複数ある時の連番"}]}。'
    +'支出情報が無いメール(広告・案内)なら {"tx":[]}。';
  var j = tryJson_(claude_(MODEL_LIGHT, sys, [{role:'user',content:'From: '+from+'\nSubject: '+sub+'\nDate: '+dateStr+'\n\n'+body}], 600));
  if(j && j.tx){ j.tx.forEach(function(t){ t.card = t.card || cardLabel_(from, body); t.time = t.time || timeStr; }); return j.tx; }
  return [];
}
function cardLabel_(from, body){
  if(/vpass|smbc/i.test(from)) return '三井住友';
  if(/rakuten-card/i.test(from)) return '楽天';
  if(/amazon/i.test(from)) return 'Amazon';
  return '';
}
function autoCat_(store){
  if(/セブン|ファミ|ローソン|スーパー|マルエツ|オーケー|食|弁当|カフェ|レストラン/i.test(store)) return '食費';
  if(/薬|ドラッグ|ダイソー|ニトリ|無印/i.test(store)) return '日用品';
  if(/JR|メトロ|タクシー|ETC|駐車|モバイルSuica|PASMO/i.test(store)) return '交通';
  if(/AMAZON|STEAM|APPLE|PLUGIN|音楽|映画|書店/i.test(store)) return '趣味・娯楽';
  if(/NETFLIX|SPOTIFY|U-NEXT|ADOBE|GOOGLE|保険|電力|ガス|水道|NTT|ソフトバンク|ドコモ|AU/i.test(store)) return 'サブスク・固定費';
  return 'その他';
}
