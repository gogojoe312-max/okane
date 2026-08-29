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
var HEAD = ['id','date','time','amount','store','cat','keihi','src','card','memo','createdAt','updatedAt','skip'];
var MAIL_QUERY = 'newer_than:3d -label:OKANE_DONE ('
  + 'from:rakuten-card.co.jp OR from:rakuten.co.jp OR from:vpass.ne.jp OR from:smbc-card.com OR from:smbc.co.jp'
  + ' OR subject:(カード利用のお知らせ) OR subject:(ご利用のお知らせ) OR subject:(ご利用内容確認) OR subject:(利用速報)'
  + ' OR subject:(デビット) OR subject:(ご利用明細) OR (from:amazon.co.jp subject:ご注文)'
  + ')';
var DEFAULT_CATS = ['食費','日用品','交通','服・美容','趣味・娯楽','サブスク・固定費','その他'];

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
      case 'diag':   out = diag_(p); break;
      case 'scan':   out = {found: scanMail()}; break;
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
    o.skip = (o.skip===true||o.skip==='TRUE'||o.skip==='true'||o.skip===1);
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
  if(p.mode==='classify')return aiClassify_(p);
  if(p.mode==='debt')return aiDebt_(p);
  throw new Error('unknown ai mode');
}
function aiDebt_(p){
  var sys='家計アプリのAI。渡された借金データと収支から、返済プランを日本語で具体的に助言する。'
    +'構成: ①現状の一言（残債と完済予定）②どれから優先して返すべきか（金利の高い順が基本だが、少額を先に消す方が続く場合はそう言う）'
    +'③毎月あといくら回せそうか（予算の余りから根拠を示す）④繰上返済したらどれだけ利息が減り何ヶ月早まるか（具体的な数字で1例）。'
    +'450字以内。データに無いことは書かない。断定しすぎず、数字は渡されたものだけ使う。';
  return { text: claude_(MODEL_CHAT, sys, [{role:'user',content:JSON.stringify(p.context||{})}], 1200) };
}
function aiClassify_(p){
  var cats=(p.cats&&p.cats.length)?p.cats:DEFAULT_CATS;
  var rules=p.rules||{};
  var rtxt='';
  var ks=Object.keys(rules);
  if(ks.length){
    rtxt='\n【ユーザーが過去に手で直した分類（最優先で踏襲し、似た店も同じ扱いにする）】\n';
    ks.slice(0,120).forEach(function(k){ rtxt+= k+' → '+rules[k]+'\n'; });
  }
  var sys='銀行・カード明細の各行を分類し、JSONだけを返す（前置き・コードブロック禁止）。'
    +'形式: {"items":[{"i":番号,"cat":"カテゴリ","skip":true/false,"why":"skip理由(短く)"}]}。'
    +'skip=trueにするのは支出でない行だけ: カード会社への引落(カード利用は別で取込済みのため二重計上になる)、ATM引出、自分の口座間振替、投資・積立の移動。'
    +'手数料・年会費・公共料金・家賃・買い物は支出なのでskip=false。'
    +'店名は全角/半角カナや略記が多い(例: ｾﾌﾞﾝ-ｲﾚﾌﾞﾝ=コンビニ、ｼﾞﾔﾊﾟﾝﾋﾞﾊﾞﾚﾂｼﾞ=自販機、GO=タクシー配車、AF/サロン系=理美容)。'
    +'日本の店舗・サービス名の実態から判断し、迷ったら金額の大きさと業種で推定する。「その他」は本当に判断不能な時だけ。'
    +'catは次から選ぶ: '+cats.join('/')+rtxt;
  var text=claude_(MODEL_CHAT, sys, [{role:'user',content:JSON.stringify(p.items||[])}], 4000);
  var j=tryJson_(text)||{};
  return { items: j.items||[] };
}
function aiChat_(p){
  var sys = 'あなたは家計アプリ「お金」のAI。ユーザーの支出データと予算を根拠に、短く率直に日本語で答える。'
    +'金額の質問には具体的な数字で答える。「買っていい？」には 買え/待て/やめとけ のどれかを冒頭に置き根拠を1-3行。'
    +'ユーザーの発言が質問でなく支出の記録（例: ラーメン980円払った）なら、応答の最後に <TX>[{"date":"YYYY-MM-DD","amount":数値,"store":"店名","cat":"カテゴリ","keihi":false}]</TX> を1回だけ付ける。'
    +'カテゴリは次から選ぶ: '+DEFAULT_CATS.join('/')+'。借金がある場合は残債・完済予定も踏まえて答える。データに無いことは推測せず「データに無い」と言う。\n【現在のデータ】\n'
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

/* ---------- 診断: 通知メールが届いているか ---------- */
function diag_(p){
  var out={queries:[],samples:[]};
  var qs=[
    {q:'newer_than:30d from:rakuten-card.co.jp', label:'楽天カード(公式ドメイン)'},
    {q:'newer_than:30d from:vpass.ne.jp OR from:smbc-card.com', label:'三井住友(公式ドメイン)'},
    {q:'newer_than:30d subject:(カード利用のお知らせ)', label:'件名:カード利用のお知らせ'},
    {q:'newer_than:30d subject:(ご利用のお知らせ)', label:'件名:ご利用のお知らせ'},
    {q:'newer_than:30d subject:(速報)', label:'件名:速報'},
    {q:MAIL_QUERY, label:'現在の取込条件'}
  ];
  qs.forEach(function(x){
    var th=[]; try{ th=GmailApp.search(x.q,0,5); }catch(e){}
    out.queries.push({label:x.label, count:th.length, q:x.q});
    if(th.length&&out.samples.length<4){
      var msg=th[0].getMessages()[0];
      var parsed=null; try{ parsed=parseMail_(msg); }catch(e){ parsed='ERR:'+e; }
      out.samples.push({from:msg.getFrom(), subject:msg.getSubject(),
        body:msg.getPlainBody().slice(0,240), parsed:parsed});
    }
  });
  // 直近のカード会社らしきメール（条件に関係なく）
  var recent=[]; try{ recent=GmailApp.search('newer_than:14d (カード OR 利用 OR 決済)',0,8); }catch(e){}
  out.recentSubjects=recent.map(function(t){var m=t.getMessages()[0];return m.getFrom()+' | '+m.getSubject();});
  return out;
}

/* ---------- Gmail scan (15分トリガー) ---------- */
function scanMail(){
  var label = GmailApp.getUserLabelByName('OKANE_DONE') || GmailApp.createLabel('OKANE_DONE');
  var threads = GmailApp.search(MAIL_QUERY, 0, 20);
  var n = 0;
  if(!threads.length){ Logger.log('scanMail: 該当メール0件 / query=' + MAIL_QUERY); return 0; }
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
          existing[t.id]=true; n++;
        });
      }catch(e){ /* skip broken mail */ }
    });
    th.addLabel(label);
  });
  Logger.log('scanMail: ' + threads.length + 'スレッド処理 / ' + n + '件登録');
  return n;
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
