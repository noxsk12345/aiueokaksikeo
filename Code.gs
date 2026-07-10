// ================================================
// Choice Chats — Code.gs
// ================================================

var OPENROUTER_API_KEY = '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';
var SS_NAME = 'ChoiceChats_DB';

// ------------------------------------------------
// エントリーポイント
// ------------------------------------------------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Choice Chats')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ------------------------------------------------
// スプレッドシート管理
// ------------------------------------------------
function getSpreadsheet() {
  var files = DriveApp.getFilesByName(SS_NAME);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }
  var ss = SpreadsheetApp.create(SS_NAME);
  var u = ss.getActiveSheet().setName('Users');
  u.appendRow(['userId','username','email','passwordHash','registeredAt','lastLogin','settingsJson','profileImage','token']);
  var c = ss.insertSheet('Chats');
  c.appendRow(['chatId','userId','chatName','messagesJson','createdAt','updatedAt']);
  return ss;
}

function usersSheet()  { return getSpreadsheet().getSheetByName('Users'); }
function chatsSheet()  { return getSpreadsheet().getSheetByName('Chats'); }

// ------------------------------------------------
// ユーティリティ
// ------------------------------------------------
function uuid()   { return Utilities.getUuid(); }
function now()    { return new Date().toISOString(); }
function sha256(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(function(b){ return (b < 0 ? b + 256 : b).toString(16).padStart(2,'0'); }).join('');
}

// ------------------------------------------------
// 認証
// ------------------------------------------------
function register(username, email, password) {
  try {
    var sh = usersSheet();
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][2] === email) return { ok: false, msg: 'このメールアドレスはすでに登録されています' };
    }
    var id = uuid(), token = uuid(), ts = now();
    sh.appendRow([id, username, email, sha256(password), ts, ts, '{}', '', token]);
    return { ok: true, user: { userId:id, username:username, email:email, registeredAt:ts, lastLogin:ts, settings:{}, profileImage:'' }, token:token };
  } catch(e) { return { ok: false, msg: e.toString() }; }
}

function login(email, password) {
  try {
    var sh = usersSheet();
    var data = sh.getDataRange().getValues();
    var hash = sha256(password);
    for (var i = 1; i < data.length; i++) {
      if (data[i][2] === email && data[i][3] === hash) {
        var token = uuid(), ts = now();
        sh.getRange(i+1, 6).setValue(ts);
        sh.getRange(i+1, 9).setValue(token);
        var settings = {};
        try { settings = JSON.parse(data[i][6] || '{}'); } catch(e) {}
        return { ok: true, user: { userId:data[i][0], username:data[i][1], email:data[i][2], registeredAt:data[i][4], lastLogin:ts, settings:settings, profileImage:data[i][7]||'' }, token:token };
      }
    }
    return { ok: false, msg: 'メールアドレスまたはパスワードが正しくありません' };
  } catch(e) { return { ok: false, msg: e.toString() }; }
}

function verifyToken(userId, token) {
  var data = usersSheet().getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === userId && data[i][8] === token) return { ok: true, row: i + 1 };
  }
  return { ok: false };
}

// ------------------------------------------------
// プロフィール更新
// ------------------------------------------------
function updateProfile(userId, token, username, email, newPassword, profileImage) {
  try {
    var v = verifyToken(userId, token);
    if (!v.ok) return { ok: false, msg: '認証エラー' };
    var sh = usersSheet();
    if (username)     sh.getRange(v.row, 2).setValue(username);
    if (email)        sh.getRange(v.row, 3).setValue(email);
    if (newPassword)  sh.getRange(v.row, 4).setValue(sha256(newPassword));
    if (profileImage !== undefined) sh.getRange(v.row, 8).setValue(profileImage);
    return { ok: true };
  } catch(e) { return { ok: false, msg: e.toString() }; }
}

// ------------------------------------------------
// 設定の保存・読み込み
// ------------------------------------------------
function saveSettings(userId, token, settingsJson) {
  try {
    var v = verifyToken(userId, token);
    if (!v.ok) return { ok: false, msg: '認証エラー' };
    usersSheet().getRange(v.row, 7).setValue(settingsJson);
    return { ok: true };
  } catch(e) { return { ok: false, msg: e.toString() }; }
}

// ------------------------------------------------
// チャット CRUD
// ------------------------------------------------
function saveChat(userId, token, chatId, chatName, messagesJson) {
  try {
    var v = verifyToken(userId, token);
    if (!v.ok) return { ok: false, msg: '認証エラー' };
    var sh = chatsSheet();
    var data = sh.getDataRange().getValues();
    var ts = now();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === chatId && data[i][1] === userId) {
        sh.getRange(i+1, 3).setValue(chatName);
        sh.getRange(i+1, 4).setValue(messagesJson);
        sh.getRange(i+1, 6).setValue(ts);
        return { ok: true };
      }
    }
    sh.appendRow([chatId, userId, chatName, messagesJson, ts, ts]);
    return { ok: true };
  } catch(e) { return { ok: false, msg: e.toString() }; }
}

function loadChats(userId, token) {
  try {
    var v = verifyToken(userId, token);
    if (!v.ok) return { ok: false, msg: '認証エラー' };
    var data = chatsSheet().getDataRange().getValues();
    var chats = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === userId) {
        var msgs = [];
        try { msgs = JSON.parse(data[i][3] || '[]'); } catch(e) {}
        chats.push({ chatId:data[i][0], chatName:data[i][2], messages:msgs, createdAt:data[i][4], updatedAt:data[i][5] });
      }
    }
    chats.sort(function(a,b){ return new Date(b.updatedAt) - new Date(a.updatedAt); });
    return { ok: true, chats: chats };
  } catch(e) { return { ok: false, msg: e.toString() }; }
}

function deleteChat(userId, token, chatId) {
  try {
    var v = verifyToken(userId, token);
    if (!v.ok) return { ok: false, msg: '認証エラー' };
    var sh = chatsSheet();
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === chatId && data[i][1] === userId) { sh.deleteRow(i+1); return { ok: true }; }
    }
    return { ok: false, msg: '見つかりません' };
  } catch(e) { return { ok: false, msg: e.toString() }; }
}

// ------------------------------------------------
// AIモデル一覧（OpenRouter対応モデル）
// ------------------------------------------------
function getModels() {
  return [
    { id:'openai/gpt-4o',                           name:'GPT-4o',              provider:'OpenAI',      icon:'https://www.google.com/s2/favicons?domain=openai.com&sz=32' },
    { id:'openai/gpt-4o-mini',                      name:'GPT-4o Mini',         provider:'OpenAI',      icon:'https://www.google.com/s2/favicons?domain=openai.com&sz=32' },
    { id:'openai/gpt-4-turbo',                      name:'GPT-4 Turbo',         provider:'OpenAI',      icon:'https://www.google.com/s2/favicons?domain=openai.com&sz=32' },
    { id:'openai/o1-preview',                       name:'o1 Preview',          provider:'OpenAI',      icon:'https://www.google.com/s2/favicons?domain=openai.com&sz=32' },
    { id:'openai/o1-mini',                          name:'o1 Mini',             provider:'OpenAI',      icon:'https://www.google.com/s2/favicons?domain=openai.com&sz=32' },
    { id:'anthropic/claude-3.5-sonnet',             name:'Claude 3.5 Sonnet',   provider:'Anthropic',   icon:'https://www.google.com/s2/favicons?domain=anthropic.com&sz=32' },
    { id:'anthropic/claude-3.5-haiku',              name:'Claude 3.5 Haiku',    provider:'Anthropic',   icon:'https://www.google.com/s2/favicons?domain=anthropic.com&sz=32' },
    { id:'anthropic/claude-3-opus',                 name:'Claude 3 Opus',       provider:'Anthropic',   icon:'https://www.google.com/s2/favicons?domain=anthropic.com&sz=32' },
    { id:'anthropic/claude-3-haiku',                name:'Claude 3 Haiku',      provider:'Anthropic',   icon:'https://www.google.com/s2/favicons?domain=anthropic.com&sz=32' },
    { id:'google/gemini-2.0-flash-001',             name:'Gemini 2.0 Flash',    provider:'Google',      icon:'https://www.google.com/s2/favicons?domain=gemini.google.com&sz=32' },
    { id:'google/gemini-pro-1.5',                   name:'Gemini 1.5 Pro',      provider:'Google',      icon:'https://www.google.com/s2/favicons?domain=gemini.google.com&sz=32' },
    { id:'google/gemini-flash-1.5',                 name:'Gemini 1.5 Flash',    provider:'Google',      icon:'https://www.google.com/s2/favicons?domain=gemini.google.com&sz=32' },
    { id:'meta-llama/llama-3.1-405b-instruct',      name:'Llama 3.1 405B',      provider:'Meta',        icon:'https://www.google.com/s2/favicons?domain=meta.com&sz=32' },
    { id:'meta-llama/llama-3.1-70b-instruct',       name:'Llama 3.1 70B',       provider:'Meta',        icon:'https://www.google.com/s2/favicons?domain=meta.com&sz=32' },
    { id:'meta-llama/llama-3.3-70b-instruct',       name:'Llama 3.3 70B',       provider:'Meta',        icon:'https://www.google.com/s2/favicons?domain=meta.com&sz=32' },
    { id:'mistralai/mistral-large',                 name:'Mistral Large',       provider:'Mistral',     icon:'https://www.google.com/s2/favicons?domain=mistral.ai&sz=32' },
    { id:'mistralai/mistral-nemo',                  name:'Mistral Nemo',        provider:'Mistral',     icon:'https://www.google.com/s2/favicons?domain=mistral.ai&sz=32' },
    { id:'mistralai/codestral-mamba',               name:'Codestral',           provider:'Mistral',     icon:'https://www.google.com/s2/favicons?domain=mistral.ai&sz=32' },
    { id:'deepseek/deepseek-r1',                    name:'DeepSeek R1',         provider:'DeepSeek',    icon:'https://www.google.com/s2/favicons?domain=deepseek.com&sz=32' },
    { id:'deepseek/deepseek-chat',                  name:'DeepSeek V3',         provider:'DeepSeek',    icon:'https://www.google.com/s2/favicons?domain=deepseek.com&sz=32' },
    { id:'x-ai/grok-2-1212',                        name:'Grok 2',              provider:'xAI',         icon:'https://www.google.com/s2/favicons?domain=x.ai&sz=32' },
    { id:'x-ai/grok-beta',                          name:'Grok Beta',           provider:'xAI',         icon:'https://www.google.com/s2/favicons?domain=x.ai&sz=32' },
    { id:'cohere/command-r-plus-08-2024',           name:'Command R+',          provider:'Cohere',      icon:'https://www.google.com/s2/favicons?domain=cohere.com&sz=32' },
    { id:'perplexity/llama-3.1-sonar-large-128k-online', name:'Sonar Large',  provider:'Perplexity',  icon:'https://www.google.com/s2/favicons?domain=perplexity.ai&sz=32' }
  ];
}

// ------------------------------------------------
// OpenRouter チャット送信
// ------------------------------------------------
function sendMessage(model, messages) {
  try {
    var res = UrlFetchApp.fetch('https://openrouter.ai/api/v1/chat/completions', {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': 'Bearer ' + OPENROUTER_API_KEY,
        'HTTP-Referer': 'https://script.google.com',
        'X-Title': 'Choice Chats'
      },
      payload: JSON.stringify({ model: model, messages: messages }),
      muteHttpExceptions: true
    });
    var json = JSON.parse(res.getContentText());
    if (json.error) return { ok: false, msg: json.error.message || 'APIエラー' };
    return { ok: true, content: json.choices[0].message.content };
  } catch(e) { return { ok: false, msg: e.toString() }; }
}

// ------------------------------------------------
// Google Fonts 一覧取得
// ------------------------------------------------
function getGoogleFonts() {
  try {
    var res = UrlFetchApp.fetch('https://fonts.google.com/metadata/fonts', { muteHttpExceptions: true });
    var json = JSON.parse(res.getContentText().replace(")]}'\n", ''));
    return json.familyMetadataList.map(function(f){ return f.family; }).sort();
  } catch(e) {
    return ['BIZ UDPGothic','Dela Gothic One','DotGothic16','Exo 2','Hachi Maru Pop','Inter',
      'Kaisei Decol','Klee One','Kosugi Maru','M PLUS 1p','Montserrat','Noto Sans JP',
      'Noto Serif JP','Orbitron','Oswald','Pacifico','Playfair Display','Poppins','Raleway',
      'Rampart One','Reggae One','RocknRoll One','Roboto','Sawarabi Gothic','Sawarabi Mincho',
      'Shippori Mincho','Space Grotesk','Stick','Train One','Ubuntu','Work Sans','Yomogi',
      'Zen Kaku Gothic New','Zen Maru Gothic'];
  }
}
function test() {
  getSpreadsheet();
}
