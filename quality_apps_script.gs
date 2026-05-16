// =====================================================
// 每日出餐品質控管 — Google Apps Script（多門市版）
//
// 設定步驟：
//   1. 開啟 script.google.com，建立新專案
//   2. 貼上此程式碼，修改下方 SHEET_ID
//   3. 部署 → 新增部署 → 類型：網頁應用程式
//      執行者：我自己 / 存取：所有人（匿名）
//   4. 複製部署網址，貼到 quality_form.html 的 SCRIPT_URL
// =====================================================

const SHEET_ID   = '1hdDj9-qyvOxodwwLI1UpZMVBGVUKQeWeTSXzneT_I7M';
const TAB_RECORD = '品質紀錄';

// 人員名單 tab 依門市獨立：人員名單_六家店 / 人員名單_竹科店
function staffTab(store) {
  return store ? `人員名單_${store}` : '人員名單';
}

// ── GET：讀取人員名單 / 紀錄 ──────────────────────
function doGet(e) {
  const p      = e && e.parameter || {};
  const action = p.action;
  const store  = String(p.store || '').trim();

  if (action === 'getStaff') {
    const cache = CacheService.getScriptCache();
    const ckey  = 'staff_' + store;
    const cached = cache.get(ckey);
    if (cached) return jsonOk(JSON.parse(cached));
    const sheet = getSheet(staffTab(store));
    const rows  = sheet.getDataRange().getValues();
    const staff = rows.slice(1).map(r => r[0]).filter(n => n && String(n).trim());
    cache.put(ckey, JSON.stringify({ staff }), 60);
    return jsonOk({ staff });
  }

  if (action === 'getRecords') {
    const { date, period } = p;
    const cache = CacheService.getScriptCache();
    const ckey  = `rec_${store}_${date}_${period}`;
    const cached = cache.get(ckey);
    if (cached) return jsonOk(JSON.parse(cached));
    return handleGetRecords(date, period, store);
  }

  if (action === 'getMonthRecords') {
    const { year, month } = p;
    const cache = CacheService.getScriptCache();
    const ckey  = `month_${store}_${year}_${month}`;
    const cached = cache.get(ckey);
    if (cached) return jsonOk(JSON.parse(cached));
    return handleGetMonthRecords(year, month, store);
  }

  return ContentService.createTextOutput('OK');
}

// ── 日期欄位格式化 ────────────────────────────────
function fmtDate(val) {
  if (val instanceof Date) return Utilities.formatDate(val, 'Asia/Taipei', 'yyyy-MM-dd');
  return String(val).trim();
}

// ── 快取工具 ──────────────────────────────────────
function clearCache(date, period, store) {
  const cache = CacheService.getScriptCache();
  cache.remove(`rec_${store}_${date}_${period}`);
  if (date) cache.remove(`month_${store}_${date.substring(0,4)}_${parseInt(date.substring(5,7))}`);
}

// ── 紀錄：讀取單一餐期 ───────────────────────────
function handleGetRecords(date, period, store) {
  const sheet = getSheet(TAB_RECORD);
  const rows  = sheet.getDataRange().getValues();
  const matching = rows.slice(1).filter(r =>
    fmtDate(r[0]) === date &&
    String(r[1]).trim() === period &&
    String(r[14] || '') === store
  );

  const result = { records: [] };
  if (matching.length) {
    const latestTime = matching.reduce((max, r) => {
      const t = String(r[13]); return t > max ? t : max;
    }, '');
    const latest = matching.filter(r => String(r[13]) === latestTime);
    result.records = latest.map(r => ({
      num: r[2], person: String(r[3]), item: String(r[4]),
      鹹度: r[5]==='✓', 熟度: r[6]==='✓', 美觀度: r[7]==='✓',
      燒焦: r[8]==='✓', 異物: r[9]==='✓', 異物說明: String(r[10]),
      type: String(r[11]), supervisor: String(r[12] || ''), interceptor: String(r[15] || ''),
      perfect: !r[5] && !r[6] && !r[7] && !r[8] && !r[9] && String(r[11])==='主管抽查'
    }));
  }

  CacheService.getScriptCache().put(`rec_${store}_${date}_${period}`, JSON.stringify(result), 60);
  return jsonOk(result);
}

// ── 紀錄：讀取整月 ───────────────────────────────
function handleGetMonthRecords(year, month, store) {
  const sheet  = getSheet(TAB_RECORD);
  const rows   = sheet.getDataRange().getValues();
  const prefix = year + '-' + String(month).padStart(2, '0');
  const matching = rows.slice(1).filter(r =>
    fmtDate(r[0]).startsWith(prefix) &&
    String(r[14] || '') === store
  );

  // 每個 date+period 只取最新一次送出
  const latestTimes = {};
  matching.forEach(r => {
    const key = fmtDate(r[0]) + '_' + r[1];
    const t = String(r[13]);
    if (!latestTimes[key] || t > latestTimes[key]) latestTimes[key] = t;
  });

  const deduped = matching.filter(r => String(r[13]) === latestTimes[fmtDate(r[0]) + '_' + r[1]]);

  const result = { records: deduped.map(r => ({
    date: fmtDate(r[0]), period: String(r[1]),
    num: r[2], person: String(r[3]), item: String(r[4]),
    鹹度: r[5]==='✓', 熟度: r[6]==='✓', 美觀度: r[7]==='✓',
    燒焦: r[8]==='✓', 異物: r[9]==='✓', 異物說明: String(r[10]),
    type: String(r[11]), supervisor: String(r[12] || ''), interceptor: String(r[15] || ''),
    perfect: !r[5] && !r[6] && !r[7] && !r[8] && !r[9] && String(r[11])==='主管抽查'
  })) };

  CacheService.getScriptCache().put(`month_${store}_${year}_${month}`, JSON.stringify(result), 60);
  return jsonOk(result);
}

// ── 紀錄：清空特定日期+餐期+門市 ────────────────
function handleClearRecords(date, period, store) {
  clearCache(date, period, store);
  const sheet = getSheet(TAB_RECORD);
  const data  = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (fmtDate(data[i][0]) === date &&
        String(data[i][1]).trim() === period &&
        String(data[i][14] || '') === store) {
      sheet.deleteRow(i + 1);
    }
  }
  return jsonOk({ success: true });
}

// ── POST：新增/移除人員、送出紀錄 ────────────────
function doPost(e) {
  try {
    const data   = JSON.parse(e.postData.contents);
    const action = data.action;
    const store  = String(data.store || '').trim();
    const ckey   = 'staff_' + store;

    if (action === 'addStaff')     { CacheService.getScriptCache().remove(ckey); return handleAddStaff(data.name, store); }
    if (action === 'removeStaff')  { CacheService.getScriptCache().remove(ckey); return handleRemoveStaff(data.name, store); }
    if (action === 'submitRecord') return handleSubmit(data);
    if (action === 'clearRecords') return handleClearRecords(data.date, data.period, store);

    return jsonOk({ message: 'unknown action' });
  } catch (err) {
    return jsonErr(err.message);
  }
}

// ── 人員：新增 ───────────────────────────────────
function handleAddStaff(name, store) {
  if (!name) return jsonErr('name required');
  const sheet = getSheet(staffTab(store));
  const names = sheet.getDataRange().getValues().slice(1).map(r => String(r[0]).trim());
  if (names.includes(name)) return jsonOk({ message: 'already exists' });
  sheet.appendRow([name, now()]);
  return jsonOk({ message: 'added', name });
}

// ── 人員：移除 ───────────────────────────────────
function handleRemoveStaff(name, store) {
  if (!name) return jsonErr('name required');
  const sheet = getSheet(staffTab(store));
  const data  = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim() === name) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return jsonOk({ message: 'removed', name });
}

// ── 紀錄：送出（先清同門市+日期+餐期的舊資料再寫入）
function handleSubmit(data) {
  const store = String(data.store || '').trim();
  clearCache(data.date, data.period, store);
  const sheet      = getSheet(TAB_RECORD);
  const submitTime = now();

  const allRows = sheet.getDataRange().getValues();
  for (let i = allRows.length - 1; i >= 1; i--) {
    if (fmtDate(allRows[i][0]) === data.date &&
        String(allRows[i][1]).trim() === data.period &&
        String(allRows[i][14] || '') === store) {
      sheet.deleteRow(i + 1);
    }
  }

  data.records.forEach(r => {
    sheet.appendRow([
      data.date, data.period, r.num, r.person, r.item,
      r.鹹度   ? '✓' : '',
      r.熟度   ? '✓' : '',
      r.美觀度 ? '✓' : '',
      r.燒焦   ? '✓' : '',
      r.異物   ? '✓' : '',
      r.異物說明, r.type, r.supervisor || '', submitTime, store, r.interceptor || '',
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
      const headers = ['日期','餐期','編號','製餐人員','品項名稱',
                       '鹹度','熟度','美觀度','燒焦','異物','異物說明',
                       '紀錄類別','抽查主管','提交時間','門市','攔截人員'];
      sheet.appendRow(headers);
      styleHeader(sheet, headers.length, '#b91c1c');
      sheet.setColumnWidth(1, 110); sheet.setColumnWidth(5, 150);
      sheet.setColumnWidth(11, 180); sheet.setColumnWidth(15, 90);
    }
    if (name.startsWith('人員名單')) {
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
