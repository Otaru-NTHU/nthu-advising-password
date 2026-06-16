/**
 * 導生密碼查詢系統 — 後端（Google Apps Script）
 * 綁定在一份私有的 Google 試算表上。整份名單只存在這裡，
 * 學生只能查到「自己那一筆」，前端永遠拿不到全班資料。
 *
 * 安全機制：
 *   1. 只做單筆比對（姓名 + 學號都要對），不提供「列出全部」的功能
 *   2. 同一個學號連續錯誤 5 次 → 鎖定 1 小時（擋對單一學號暴力猜姓名）
 *   3. 全站每 10 分鐘最多 80 次查詢（擋大量掃描）
 *   4. 每次查詢都記錄在「查詢紀錄」工作表（時間 / 學號 / 姓名 / 結果）
 *   5. 查無資料時回傳統一訊息，不洩漏「學號存在但姓名錯」這種線索
 */

// ===================== 設定（依需要調整） =====================
var SHEET_NAME      = '名單';       // 放學生資料的工作表名稱
var LOG_SHEET_NAME  = '查詢紀錄';   // 自動建立、自動寫入的紀錄表
var COL_PROGRAM     = '系級';
var COL_ID          = '學號';
var COL_NAME        = '姓名';
var COL_STATUS      = '就學狀態';
// 密碼欄不用寫死欄名：程式會自動找「標題含『密碼』」的那一欄，
// 所以每學期叫「115上導師密碼 / 115下導師密碼」都能自動對應。

var MAX_FAILS_PER_ID = 5;     // 同一學號連續錯誤上限
var FAIL_WINDOW_SEC  = 3600;  // 鎖定時間（秒）= 1 小時
var GLOBAL_MAX       = 80;    // 全站每時間窗最大查詢次數
var GLOBAL_WINDOW_SEC= 600;   // 全站時間窗（秒）= 10 分鐘
// ============================================================

function doPost(e) {
  try {
    var body = {};
    try { body = JSON.parse(e.postData.contents); } catch (err) {}
    var name = normalize_(body.name);
    var id   = normalize_(body.id);

    if (!name || !id) return json_({ ok: false, code: 'EMPTY', msg: '請輸入姓名與學號' });

    var cache = CacheService.getScriptCache();

    // 全站節流
    var gCount = Number(cache.get('g') || 0);
    if (gCount >= GLOBAL_MAX) {
      return json_({ ok: false, code: 'BUSY', msg: '系統忙碌中，請稍後再試' });
    }
    cache.put('g', gCount + 1, GLOBAL_WINDOW_SEC);

    // 同一學號鎖定
    var fKey  = 'f_' + id;
    var fails = Number(cache.get(fKey) || 0);
    if (fails >= MAX_FAILS_PER_ID) {
      log_(id, name, '鎖定');
      return json_({ ok: false, code: 'LOCKED', msg: '嘗試次數過多，請 1 小時後再試' });
    }

    var rec = lookup_(id, name);
    if (!rec) {
      cache.put(fKey, fails + 1, FAIL_WINDOW_SEC);
      log_(id, name, '失敗');
      return json_({ ok: false, code: 'NOT_FOUND', msg: '查無資料，請確認姓名與學號是否正確' });
    }

    cache.remove(fKey);           // 成功後清掉該學號的失敗計數
    log_(id, name, '成功');
    return json_({
      ok: true,
      name:     rec.name,
      id:       rec.id,
      program:  rec.program,
      status:   rec.status,
      term:     rec.term,         // 例：115上導師密碼
      password: rec.password
    });

  } catch (err) {
    return json_({ ok: false, code: 'ERROR', msg: '系統錯誤，請稍後再試' });
  }
}

// 直接用瀏覽器開端點網址時，不會洩漏任何資料
function doGet(e) {
  return json_({ ok: false, code: 'METHOD', msg: '請從查詢頁面操作' });
}

// ---------- 核心比對 ----------
function lookup_(id, name) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) return null;

  var values = sh.getDataRange().getDisplayValues(); // 用顯示值，保留前導 0
  if (values.length < 2) return null;

  var header     = values[0];
  var idxId      = header.indexOf(COL_ID);
  var idxName    = header.indexOf(COL_NAME);
  var idxProgram = header.indexOf(COL_PROGRAM);
  var idxStatus  = header.indexOf(COL_STATUS);

  var idxPwd = -1;
  for (var c = 0; c < header.length; c++) {
    if (String(header[c]).indexOf('密碼') >= 0) { idxPwd = c; break; }
  }
  if (idxId < 0 || idxName < 0 || idxPwd < 0) return null;

  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (normalize_(row[idxId]) === id && normalize_(row[idxName]) === name) {
      return {
        id:       row[idxId],
        name:     row[idxName],
        program:  idxProgram >= 0 ? row[idxProgram] : '',
        status:   idxStatus  >= 0 ? row[idxStatus]  : '',
        term:     String(header[idxPwd]),
        password: row[idxPwd]
      };
    }
  }
  return null;
}

// ---------- 文字正規化：去空白、全形→半形數字 ----------
function normalize_(v) {
  if (v === null || v === undefined) return '';
  var s = String(v);
  s = s.replace(/\u3000/g, ' ').trim();                         // 全形空白
  s = s.replace(/[\uFF10-\uFF19]/g, function (ch) {             // 全形數字→半形
    return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0);
  });
  s = s.replace(/\s+/g, '');                                    // 移除中間所有空白
  return s;
}

// ---------- 寫入查詢紀錄 ----------
function log_(id, name, result) {
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(LOG_SHEET_NAME);
    if (!sh) {
      sh = ss.insertSheet(LOG_SHEET_NAME);
      sh.appendRow(['時間', '學號', '姓名', '結果']);
    }
    sh.appendRow([new Date(), id, name, result]);
  } catch (err) { /* 紀錄失敗不影響查詢本身 */ }
}

// ---------- 統一回傳 JSON ----------
function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
