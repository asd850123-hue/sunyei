// =====================================================
// 每日出餐品質控管 — Google Apps Script
//
// 設定步驟：
//   1. 開啟 script.google.com，建立新專案
//   2. 貼上此程式碼，修改下方 SHEET_ID
//   3. 部署 → 新增部署 → 類型：網頁應用程式
//      執行者：我自己 / 存取：所有人（匿名）
//   4. 複製部署網址，貼到 quality_form.html 的 SCRIPT_URL
// =====================================================

const SHEET_ID    = 'YOUR_GOOGLE_SHEET_ID_HERE'; // ← 改這裡
const TAB_RECORD  = '品質紀錄';
const TAB_STAFF   = '人員名單';

// ── GET：讀取人員名單 ─────────────────────────────
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  if (action === 'getStaff') {
    const sheet = getSheet(TAB_STAFF);
    const rows = sheet.getDataRange().getValues();
    // A 欄：姓名（跳過標題列）
    const staff = rows.slice(1).map(r => r[0]).filter(n => n && String(n).trim());
    return jsonOk({ staff });
  }

  return ContentService.createTextOutput('OK');
}

// ── POST：新增/移除人員、送出紀錄 ────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    if (action === 'addStaff')    return handleAddStaff(data.name);
    if (action === 'removeStaff') return handleRemoveStaff(data.name);
    if (action === 'submitRecord') return handleSubmit(data);

    return jsonOk({ message: 'unknown action' });
  } catch (err) {
    return jsonErr(err.message);
  }
}

// ── 人員：新增 ───────────────────────────────────
function handleAddStaff(name) {
  if (!name) return jsonErr('name required');
  const sheet = getSheet(TAB_STAFF);
  const names = sheet.getDataRange().getValues().slice(1).map(r => String(r[0]).trim());
  if (names.includes(name)) return jsonOk({ message: 'already exists' });
  sheet.appendRow([name, now()]);
  return jsonOk({ message: 'added', name });
}

// ── 人員：移除 ───────────────────────────────────
function handleRemoveStaff(name) {
  if (!name) return jsonErr('name required');
  const sheet = getSheet(TAB_STAFF);
  const data  = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim() === name) {
      sheet.deleteRow(i + 1); // 1-based
      break;
    }
  }
  return jsonOk({ message: 'removed', name });
}

// ── 紀錄：送出 ───────────────────────────────────
function handleSubmit(data) {
  const sheet = getSheet(TAB_RECORD);
  const submitTime = now();

  data.records.forEach(r => {
    sheet.appendRow([
      data.date,
      data.period,
      r.num,
      r.person,
      r.item,
      r.鹹度    ? '✓' : '',
      r.熟度    ? '✓' : '',
      r.美觀度  ? '✓' : '',
      r.燒焦    ? '✓' : '',
      r.異物    ? '✓' : '',
      r.異物說明,
      r.type,
      data.manager,
      submitTime,
    ]);
  });

  return jsonOk({ success: true, count: data.records.length });
}

// ── 內部工具 ──────────────────────────────────────
function getSheet(name) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === TAB_RECORD) {
      const headers = ['日期','餐期','編號','負責人','品項名稱',
                       '鹹度','熟度','美觀度','燒焦','異物','異物說明',
                       '紀錄類別','店長/組長','提交時間'];
      sheet.appendRow(headers);
      styleHeader(sheet, headers.length, '#b91c1c');
      [1,5,11,12,14].forEach((c,_) => sheet.setColumnWidth(c, 100));
      sheet.setColumnWidth(5, 150);
      sheet.setColumnWidth(11, 180);
    }
    if (name === TAB_STAFF) {
      sheet.appendRow(['姓名','新增時間']);
      styleHeader(sheet, 2, '#374151');
    }
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function styleHeader(sheet, cols, bg) {
  const r = sheet.getRange(1, 1, 1, cols);
  r.setFontWeight('bold').setBackground(bg).setFontColor('#ffffff').setHorizontalAlignment('center');
}

function now() {
  return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');
}

function jsonOk(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonErr(msg) {
  return ContentService.createTextOutput(JSON.stringify({ success: false, error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}
