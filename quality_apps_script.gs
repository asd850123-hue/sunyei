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

const SHEET_ID    = '1hdDj9-qyvOxodwwLI1UpZMVBGVUKQeWeTSXzneT_I7M';
const TAB_RECORD  = '品質紀錄';
const TAB_STAFF   = '人員名單';

// ── GET：讀取人員名單 / 紀錄 ──────────────────────
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  if (action === 'getStaff') {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('staff');
    if (cached) return jsonOk(JSON.parse(cached));
    const sheet = getSheet(TAB_STAFF);
    const rows = sheet.getDataRange().getValues();
    const staff = rows.slice(1).map(r => r[0]).filter(n => n && String(n).trim());
    cache.put('staff', JSON.stringify({ staff }), 60);
    return jsonOk({ staff });
  }

  if (action === 'getRecords') {
    const { date, period } = e.parameter;
    const cache = CacheService.getScriptCache();
    const key = 'rec_' + date + '_' + period;
    const cached = cache.get(key);
    if (cached) return jsonOk(JSON.parse(cached));
    return handleGetRecords(date, period);
  }

  if (action === 'getMonthRecords') {
    const { year, month } = e.parameter;
    const cache = CacheService.getScriptCache();
    const key = 'month_' + year + '_' + month;
    const cached = cache.get(key);
    if (cached) return jsonOk(JSON.parse(cached));
    return handleGetMonthRecords(year, month);
  }

  return ContentService.createTextOutput('OK');
}

// ── 日期欄位格式化（Sheets 會把日期字串轉成 Date 物件）
function fmtDate(val) {
  if (val instanceof Date) return Utilities.formatDate(val, 'Asia/Taipei', 'yyyy-MM-dd');
  return String(val).trim();
}

// ── 快取工具 ─────────────────────────────────────
function clearCache(date, period) {
  const cache = CacheService.getScriptCache();
  cache.remove('rec_' + date + '_' + period);
  if (date) cache.remove('month_' + date.substring(0, 4) + '_' + parseInt(date.substring(5, 7)));
}

// ── 紀錄：讀取單一餐期 ───────────────────────────
function handleGetRecords(date, period) {
  const sheet = getSheet(TAB_RECORD);
  const rows = sheet.getDataRange().getValues();
  const matching = rows.slice(1).filter(r =>
    fmtDate(r[0]) === date && String(r[1]).trim() === period
  );

  const result = { records: [], manager: '' };
  if (matching.length) {
    const latestTime = matching.reduce((max, r) => {
      const t = String(r[13]); return t > max ? t : max;
    }, '');
    const latest = matching.filter(r => String(r[13]) === latestTime);
    result.records = latest.map(r => ({
      num: r[2], person: String(r[3]), item: String(r[4]),
      鹹度: r[5]==='✓', 熟度: r[6]==='✓', 美觀度: r[7]==='✓',
      燒焦: r[8]==='✓', 異物: r[9]==='✓', 異物說明: String(r[10]),
      type: String(r[11]),
      perfect: !r[5] && !r[6] && !r[7] && !r[8] && !r[9] && String(r[11])==='主管抽查'
    }));
    result.manager = String(latest[0][12]);
  }

  CacheService.getScriptCache().put('rec_' + date + '_' + period, JSON.stringify(result), 60);
  return jsonOk(result);
}

// ── 紀錄：讀取整月 ───────────────────────────────
function handleGetMonthRecords(year, month) {
  const sheet = getSheet(TAB_RECORD);
  const rows = sheet.getDataRange().getValues();
  const prefix = year + '-' + String(month).padStart(2, '0');
  const matching = rows.slice(1).filter(r => fmtDate(r[0]).startsWith(prefix));

  // 每個 date+period 只取最新一次送出
  const latestTimes = {};
  matching.forEach(r => {
    const key = fmtDate(r[0]) + '_' + r[1];
    const t = String(r[13]);
    if (!latestTimes[key] || t > latestTimes[key]) latestTimes[key] = t;
  });

  const deduped = matching.filter(r => String(r[13]) === latestTimes[fmtDate(r[0]) + '_' + r[1]]);

  const result = { records: [] };
  result.records = deduped.map(r => ({
    date: fmtDate(r[0]), period: String(r[1]),
    num: r[2], person: String(r[3]), item: String(r[4]),
    鹹度: r[5]==='✓', 熟度: r[6]==='✓', 美觀度: r[7]==='✓',
    燒焦: r[8]==='✓', 異物: r[9]==='✓', 異物說明: String(r[10]),
    type: String(r[11]), manager: String(r[12]),
    perfect: !r[5] && !r[6] && !r[7] && !r[8] && !r[9] && String(r[11])==='主管抽查'
  }));

  const key = 'month_' + year + '_' + month;
  CacheService.getScriptCache().put(key, JSON.stringify(result), 60);
  return jsonOk(result);
}

// ── 紀錄：清空特定日期+餐期 ─────────────────────
function handleClearRecords(date, period) {
  clearCache(date, period);
  const sheet = getSheet(TAB_RECORD);
  const data = sheet.getDataRange().getValues();
  // 從最後一列往前刪，避免索引偏移
  for (let i = data.length - 1; i >= 1; i--) {
    if (fmtDate(data[i][0]) === date && String(data[i][1]).trim() === period) {
      sheet.deleteRow(i + 1);
    }
  }
  return jsonOk({ success: true });
}

// ── POST：新增/移除人員、送出紀錄 ────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    if (action === 'addStaff')     { CacheService.getScriptCache().remove('staff'); return handleAddStaff(data.name); }
    if (action === 'removeStaff')  { CacheService.getScriptCache().remove('staff'); return handleRemoveStaff(data.name); }
    if (action === 'submitRecord') return handleSubmit(data);
    if (action === 'clearRecords') return handleClearRecords(data.date, data.period);

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

// ── 紀錄：送出（先清舊資料再寫入，保持 Sheet 乾淨）
function handleSubmit(data) {
  clearCache(data.date, data.period);
  const sheet = getSheet(TAB_RECORD);
  const submitTime = now();

  // 刪除同日期+餐期的舊資料
  const allRows = sheet.getDataRange().getValues();
  for (let i = allRows.length - 1; i >= 1; i--) {
    if (fmtDate(allRows[i][0]) === data.date && String(allRows[i][1]).trim() === data.period) {
      sheet.deleteRow(i + 1);
    }
  }

  // 寫入最新資料
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
