/**
 * MWD 早餐店 — QR 集點卡後端 (Google Apps Script)
 * --------------------------------------------------
 * 資料存放在「綁定這個 Apps Script 的 Google 試算表」。
 * 兩個工作表會自動建立：
 *   Points : phone | points | lastStampDate | updatedAt
 *   Config : key   | value      （todayCode / staffCode）
 *
 * 客人流程：掃 QR -> 輸入手機 -> 輸入今日通關碼 -> 蓋章 +1
 * 防作弊  ：加點需正確的 todayCode，且每支手機每天限蓋一次。
 */

const CFG = {
  POINTS_SHEET: 'Points',
  CONFIG_SHEET: 'Config',
  LOG_SHEET: 'Log',                 // 每次蓋章的座標紀錄（供日後分析）
  TARGET: 10,                       // 集滿幾點可兌換
  REWARD: '炸物拚盤＋Snoopy 鑰匙圈（送完為止）', // 獎勵內容
  MIN_SPEND_WEEKDAY: 180,           // 平日（週二～週五）最低消費
  MIN_SPEND_WEEKEND: 300,           // 假日（週六、週日）最低消費
  SHOP: 'MWD 泰山明志 BRUNCH',
  TZ: 'Asia/Taipei',
  CODE_MODE: 'auto',                // 'auto' = 系統依日期自動產生今日通關碼；'manual' = 用 Config 的 todayCode
  // ---- 定位（目前只記錄、不擋人；之後分析完再決定是否開啟封鎖）----
  REQUIRE_LOCATION: false,          // false = 不因距離擋人，但仍記錄座標到 Log
  SHOP_LAT: 25.055472543375735,     // 店家緯度
  SHOP_LNG: 121.43159737244797,     // 店家經度
  RADIUS_M: 150,                    // 範圍（公尺）：開啟封鎖時用、也用來標記紀錄是否在範圍內
};

/** 部署成網頁應用程式時的進入點 */
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : '';
  const file = (page === 'admin') ? 'Admin' : 'Index';   // 加 ?page=admin 開店員管理頁
  return HtmlService.createHtmlOutputFromFile(file)
    .setTitle('MWD 集點卡')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- 工具 ---------- */

function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

/** 第一次執行時自動建立表頭與預設設定 */
function ensureSetup_() {
  const p = getSheet_(CFG.POINTS_SHEET);
  if (p.getLastRow() === 0) {
    p.appendRow(['phone', 'points', 'lastStampDate', 'updatedAt']);
  }
  const c = getSheet_(CFG.CONFIG_SHEET);
  if (c.getLastRow() === 0) c.appendRow(['key', 'value']);
  ensureConfigKey_('todayCode', '1234');  // 只有 CODE_MODE='manual' 時才會用到
  ensureConfigKey_('staffCode', '8888');  // 兌換用，店員私下保管
  ensureConfigKey_('adminPin', '2468');   // 管理頁密碼，店員私下保管
  ensureConfigKey_('codeSalt', 'mwd-' + Math.random().toString(36).slice(2, 10)); // 自動碼用的秘密種子，請勿外流
}

function ensureConfigKey_(key, val) {
  const c = getSheet_(CFG.CONFIG_SHEET);
  const data = c.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) if (String(data[i][0]).trim() === key) return;
  c.appendRow([key, val]);
}

/** 依「日期＋秘密種子」算出某天的 4 位數通關碼 */
function getCodeForDate_(dateStr) {
  const seed = dateStr + '|' + getConfig_('codeSalt');
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, seed);
  let n = 0;
  for (let i = 0; i < 4; i++) n = (n * 256 + (bytes[i] & 0xff)) % 10000;
  return ('000' + n).slice(-4);
}

/** 今日通關碼：auto = 自動算出；manual = 讀 Config 的 todayCode */
function getTodayCode_() {
  if (String(CFG.CODE_MODE) === 'manual') return getConfig_('todayCode');
  return getCodeForDate_(todayStr_());
}

/** 店員管理頁呼叫：輸入管理 PIN，回傳今天的通關碼 */
function adminTodayCode(pin) {
  ensureSetup_();
  if (String(pin || '').trim() !== getConfig_('adminPin')) return { ok: false, msg: '管理 PIN 不正確' };
  return { ok: true, code: getTodayCode_(), date: todayStr_(), mode: CFG.CODE_MODE };
}

/** 管理頁列印用：回傳未來 days 天的日期與通關碼（可印成碼卡，免手寫） */
function upcomingCodes(pin, days) {
  ensureSetup_();
  if (String(pin || '').trim() !== getConfig_('adminPin')) return { ok: false, msg: '管理 PIN 不正確' };
  days = Math.min(Math.max(parseInt(days, 10) || 30, 1), 60);
  const wk = { 1: '一', 2: '二', 3: '三', 4: '四', 5: '五', 6: '六', 7: '日' };
  const base = new Date();
  const out = [];
  for (let i = 0; i < days; i++) {
    const dt = new Date(base.getTime());
    dt.setDate(dt.getDate() + i);
    const ds = Utilities.formatDate(dt, CFG.TZ, 'yyyy-MM-dd');
    const dow = Number(Utilities.formatDate(dt, CFG.TZ, 'u')); // 1=Mon..7=Sun
    const closed = (dow === 1); // 週一公休
    out.push({
      date: ds,
      md: Utilities.formatDate(dt, CFG.TZ, 'M/d'),
      wk: '週' + wk[dow],
      closed: closed,
      code: closed ? '—' : (CFG.CODE_MODE === 'manual' ? getConfig_('todayCode') : getCodeForDate_(ds)),
    });
  }
  return { ok: true, shop: CFG.SHOP, list: out };
}

function getConfig_(key) {
  ensureSetup_();
  const data = getSheet_(CFG.CONFIG_SHEET).getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) return String(data[i][1]).trim();
  }
  return '';
}

function normPhone_(phone) {
  // 去掉非數字，並去掉開頭的 0（避免「試算表把 0912... 存成數字 912... 掉了前導 0」造成查無資料）
  return String(phone || '').replace(/[^0-9]/g, '').replace(/^0+/, '');
}

function todayStr_() {
  return Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd');
}

/**
 * 把 lastStampDate 儲存格的值統一轉成 'yyyy-MM-dd' 字串再比較。
 * 試算表會把字串 '2026-07-19' 自動吃成「日期物件」，直接 String() 會變成
 * 'Sun Jul 19 2026...'，導致「今天是否已蓋過」永遠比對不到 → 同一天可重複加點。
 * 這裡不論存進去是字串或日期，都轉回同一種格式，確保每天限一次真正生效。
 */
function normDateCell_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, CFG.TZ, 'yyyy-MM-dd');
  return String(v || '').trim();
}

/** 依星期回傳今天的最低消費（週六/週日為假日） */
function minSpend_() {
  const dow = Number(Utilities.formatDate(new Date(), CFG.TZ, 'u')); // 1=Mon .. 7=Sun
  return (dow === 6 || dow === 7) ? CFG.MIN_SPEND_WEEKEND : CFG.MIN_SPEND_WEEKDAY;
}

function findRow_(sh, phone) {
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (normPhone_(data[i][0]) === phone) return i + 1; // 回傳 1-based 列號
  }
  return -1;
}

/** 記錄一次蓋章的座標與距離（供日後分析；不阻擋流程） */
function logStamp_(phone, lat, lng, distM, within, pointsAfter) {
  try {
    const sh = getSheet_(CFG.LOG_SHEET);
    if (sh.getLastRow() === 0) {
      sh.appendRow(['time', 'phone', 'lat', 'lng', 'distM', 'within' + CFG.RADIUS_M + 'm', 'pointsAfter']);
    }
    sh.appendRow([new Date(), phone, lat, lng, distM, within, pointsAfter]);
  } catch (err) {
    // 記錄失敗不影響蓋章
  }
}

/** 兩點間距離（公尺）— Haversine */
function distanceMeters_(lat1, lng1, lat2, lng2) {
  const R = 6371000, toRad = function (d) { return d * Math.PI / 180; };
  const dLat = toRad(lat2 - lat1), dLng = toRad(lng2 - lng1);
  const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLng / 2) * Math.sin(dLng / 2);
  return 2 * R * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

/* ---------- 給前端呼叫的 API（google.script.run） ---------- */

/** 查詢點數，不加點 */
function getStatus(phone) {
  ensureSetup_();
  phone = normPhone_(phone);
  if (phone.length < 8) return { ok: false, msg: '請輸入正確的手機號碼' };
  const sh = getSheet_(CFG.POINTS_SHEET);
  const row = findRow_(sh, phone);
  let points = 0, last = '';
  if (row > 0) {
    points = Number(sh.getRange(row, 2).getValue()) || 0;
    last = normDateCell_(sh.getRange(row, 3).getValue());
  }
  return {
    ok: true,
    phone: phone,
    points: points,
    target: CFG.TARGET,
    reward: CFG.REWARD,
    minSpend: minSpend_(),
    minWeekday: CFG.MIN_SPEND_WEEKDAY,
    minWeekend: CFG.MIN_SPEND_WEEKEND,
    requireLocation: CFG.REQUIRE_LOCATION,
    shop: CFG.SHOP,
    canRedeem: points >= CFG.TARGET,
    stampedToday: last === todayStr_(),
  };
}

/** 蓋章 +1（需今日通關碼＋在店家範圍內；每支手機每天限一次） */
function addStamp(phone, code, lat, lng) {
  ensureSetup_();
  phone = normPhone_(phone);
  if (phone.length < 8) return { ok: false, msg: '請輸入正確的手機號碼' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (String(code || '').trim() !== getTodayCode_()) {
      return { ok: false, msg: '今日通關碼不正確' };
    }
    // 算距離（有座標才算）；REQUIRE_LOCATION 開啟時才據此擋人
    const la = parseFloat(lat), ln = parseFloat(lng);
    const hasLoc = !(isNaN(la) || isNaN(ln));
    const dist = hasLoc ? distanceMeters_(la, ln, CFG.SHOP_LAT, CFG.SHOP_LNG) : null;
    if (CFG.REQUIRE_LOCATION) {
      if (!hasLoc) return { ok: false, msg: '請開啟定位權限才能蓋章' };
      if (dist > CFG.RADIUS_M) {
        return { ok: false, msg: '請在店內掃碼蓋章（目前距離約 ' + Math.round(dist) + ' 公尺）' };
      }
    }
    const today = todayStr_();
    const sh = getSheet_(CFG.POINTS_SHEET);
    let row = findRow_(sh, phone);
    if (row < 0) {
      sh.appendRow([phone, 0, '', '']);
      row = sh.getLastRow();
    }
    const last = normDateCell_(sh.getRange(row, 3).getValue());
    let points = Number(sh.getRange(row, 2).getValue()) || 0;
    if (last === today) {
      return { ok: false, msg: '今天已經蓋過囉，明天再來！', points: points, target: CFG.TARGET };
    }
    points += 1;
    sh.getRange(row, 2).setValue(points);
    // 以純文字寫入日期，避免試算表把 '2026-07-19' 轉成日期物件而讓比對失效
    const lastCell = sh.getRange(row, 3);
    lastCell.setNumberFormat('@');
    lastCell.setValue(today);
    sh.getRange(row, 4).setValue(new Date());
    logStamp_(phone, hasLoc ? la : '', hasLoc ? ln : '', dist == null ? '' : Math.round(dist),
             dist == null ? '' : (dist <= CFG.RADIUS_M), points);
    return {
      ok: true,
      points: points,
      target: CFG.TARGET,
      reward: CFG.REWARD,
      canRedeem: points >= CFG.TARGET,
      msg: '蓋章成功，+1 點！',
    };
  } finally {
    lock.releaseLock();
  }
}

/** 兌換（需店員兌換碼；成功後扣除 target 點） */
function redeem(phone, staffCode) {
  ensureSetup_();
  phone = normPhone_(phone);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (String(staffCode || '').trim() !== getConfig_('staffCode')) {
      return { ok: false, msg: '店員兌換碼不正確' };
    }
    const sh = getSheet_(CFG.POINTS_SHEET);
    const row = findRow_(sh, phone);
    if (row < 0) return { ok: false, msg: '查無此號碼' };
    let points = Number(sh.getRange(row, 2).getValue()) || 0;
    if (points < CFG.TARGET) return { ok: false, msg: '點數不足，無法兌換' };
    points -= CFG.TARGET;
    sh.getRange(row, 2).setValue(points);
    sh.getRange(row, 4).setValue(new Date());
    return { ok: true, points: points, target: CFG.TARGET, msg: '兌換成功！' };
  } finally {
    lock.releaseLock();
  }
}

/* ==================================================================
 *  維護工具（人工清理用；不影響客人蓋章流程）
 *  - previewPhone(phone)  ：唯讀，預覽某支號碼在 Points / Log 的資料
 *  - flagTestPhone(phone) ：Points 刪除該號碼、Log 只在 note 欄標 'test'（不刪列）
 *  遵守專案 SOP：Log 為稽核紀錄「只標記、不刪」；只在 Points 修改狀態；動手前先備份。
 *  帶參數的函式無法在編輯器直接「執行」，請用下方「🛠 維護」選單操作。
 * ================================================================== */

/** 開啟試算表時建立「🛠 維護」選單（用對話框輸入號碼，避免手動點格子） */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛠 維護')
    .addItem('一鍵備份（建立副本）', 'promptBackupNow_')
    .addSeparator()
    .addItem('預覽號碼資料（唯讀）', 'promptPreviewPhone_')
    .addItem('掃描可疑成對（唯讀）', 'promptScanSuspiciousPairs_')
    .addItem('標記測試號碼（Points 刪除 / Log 標記）', 'promptFlagTestPhone_')
    .addItem('修正點數（Points 改值＋Log 留紀錄）', 'promptFixPoints_')
    .addToUi();
}

/** 找出 Log 的 note 欄；沒有就自動補在最後一欄，回傳 1-based 欄號 */
function ensureLogNoteCol_() {
  const sh = getSheet_(CFG.LOG_SHEET);
  if (sh.getLastRow() === 0) {   // Log 全空 → 先補標準表頭（含 note）
    sh.appendRow(['time', 'phone', 'lat', 'lng', 'distM', 'within' + CFG.RADIUS_M + 'm', 'pointsAfter', 'note']);
    return 8;
  }
  const lastCol = sh.getLastColumn();
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  for (let i = 0; i < header.length; i++) {
    if (String(header[i]).trim().toLowerCase() === 'note') return i + 1;
  }
  const col = lastCol + 1;       // 沒有 note 欄 → 加在最後一欄
  sh.getRange(1, col).setValue('note');
  return col;
}

/** 收集某支號碼在 Points / Log 的所有列（供預覽與清理共用；phone 需已正規化） */
function collectPhoneRows_(phone) {
  const pData = getSheet_(CFG.POINTS_SHEET).getDataRange().getValues();
  const pointsRows = [];
  for (let i = 1; i < pData.length; i++) {
    if (normPhone_(pData[i][0]) === phone) {
      pointsRows.push({ row: i + 1, points: pData[i][1], lastStampDate: normDateCell_(pData[i][2]) });
    }
  }
  const gData = getSheet_(CFG.LOG_SHEET).getDataRange().getValues();
  const logRows = [];
  for (let i = 1; i < gData.length; i++) {
    if (normPhone_(gData[i][1]) === phone) {
      logRows.push({ row: i + 1, time: gData[i][0], pointsAfter: gData[i][6] });
    }
  }
  return { pointsRows: pointsRows, logRows: logRows };
}

/**
 * 唯讀預覽：某支號碼在 Points / Log 有哪些資料，完全不改動。
 * 結果會寫到執行紀錄（Logger），也會回傳物件。用選單執行時會跳出彈窗顯示。
 */
function previewPhone(phone) {
  ensureSetup_();
  phone = normPhone_(phone);
  if (phone.length < 8) return { ok: false, msg: '請提供正確的手機號碼（純數字）' };
  const f = collectPhoneRows_(phone);
  const summary =
    '號碼 ' + phone + '\n' +
    'Points 分頁：' + f.pointsRows.length + ' 列' + (f.pointsRows.length ? '（清理時會刪除）' : '') + '\n' +
    f.pointsRows.map(function (r) {
      return '  · 第 ' + r.row + ' 列  points=' + r.points + '  lastStampDate=' + r.lastStampDate;
    }).join('\n') + (f.pointsRows.length ? '\n' : '') +
    'Log 分頁：' + f.logRows.length + ' 列（清理時「不刪」，只在 note 欄標 test）\n' +
    f.logRows.map(function (r) {
      return '  · 第 ' + r.row + ' 列  time=' + r.time + '  pointsAfter=' + r.pointsAfter;
    }).join('\n');
  Logger.log(summary);
  return { ok: true, phone: phone, pointsRows: f.pointsRows, logRows: f.logRows, summary: summary };
}

/**
 * 清理測試號碼：Points 刪除該號碼所有列、Log 只在 note 欄標 'test'（不刪列）。
 * 遵守 SOP：Log 為稽核紀錄只標記不刪。建議動手前先手動建立備份副本。
 */
function flagTestPhone(phone) {
  ensureSetup_();
  phone = normPhone_(phone);
  if (phone.length < 8) return { ok: false, msg: '請提供正確的手機號碼（純數字）' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const f = collectPhoneRows_(phone);

    // 1) Log：只標記，不刪。把每一列的 note 欄設成 'test'（已含 test 則略過）
    const noteCol = ensureLogNoteCol_();
    const gsh = getSheet_(CFG.LOG_SHEET);
    let flagged = 0;
    f.logRows.forEach(function (r) {
      const cell = gsh.getRange(r.row, noteCol);
      const cur = String(cell.getValue() || '').trim();
      if (cur.toLowerCase().indexOf('test') === -1) {
        cell.setValue(cur ? (cur + '; test') : 'test');
      }
      flagged++;
    });

    // 2) Points：刪除該號碼所有列（由下往上刪，列號才不會位移）
    const psh = getSheet_(CFG.POINTS_SHEET);
    const rowsDesc = f.pointsRows.map(function (r) { return r.row; }).sort(function (a, b) { return b - a; });
    rowsDesc.forEach(function (rowNum) { psh.deleteRow(rowNum); });

    const msg = '號碼 ' + phone + '：Points 刪除 ' + rowsDesc.length + ' 列、Log 標記 test ' + flagged + ' 列';
    Logger.log(msg);
    return { ok: true, phone: phone, pointsDeleted: rowsDesc.length, logFlagged: flagged, msg: msg };
  } finally {
    lock.releaseLock();
  }
}

/* ---- 「🛠 維護」選單用的對話框包裝（先預覽、再確認、才動手）---- */

function promptPreviewPhone_() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('預覽號碼資料（唯讀）', '請輸入要查詢的手機號碼：', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const out = previewPhone(res.getResponseText());
  ui.alert(out.ok ? out.summary : out.msg);
}

function promptFlagTestPhone_() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('標記測試號碼', '請輸入要清理的手機號碼：', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const phone = res.getResponseText();

  const pre = previewPhone(phone);                 // 先預覽
  if (!pre.ok) { ui.alert(pre.msg); return; }
  const confirm = ui.alert(
    '確認清理？（建議先建立備份副本）',
    pre.summary + '\n\n將執行：Points 刪除上列、Log 只標記 test（不刪）。要繼續嗎？',
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  const out = flagTestPhone(phone);
  ui.alert(out.ok ? ('完成：\n' + out.msg) : out.msg);
}

/** 一鍵備份：複製整份試算表到雲端硬碟，命名含日期時間；回傳新檔名稱與網址 */
function backupNow_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stamp = Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd_HHmm');
  const name = 'MWD集點卡_backup_' + stamp;
  const copy = ss.copy(name);
  Logger.log('已備份：' + name + '  ' + copy.getUrl());
  return { ok: true, name: name, url: copy.getUrl(), id: copy.getId() };
}

/**
 * 修正點數：把某支號碼的 Points 點數改成正確值，並在 Log 補一筆「人工修正」紀錄（不刪舊資料）。
 * 用於修掉像 bug 灌多的點數（例：965106328 從 8 改回 4）。correctPoints 需為 0 以上整數。
 */
function fixInflatedPoints(phone, correctPoints) {
  ensureSetup_();
  phone = normPhone_(phone);
  if (phone.length < 8) return { ok: false, msg: '請提供正確的手機號碼（純數字）' };
  const target = parseInt(correctPoints, 10);
  if (isNaN(target) || target < 0) return { ok: false, msg: '請輸入正確的點數（0 以上整數）' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const psh = getSheet_(CFG.POINTS_SHEET);
    const row = findRow_(psh, phone);
    if (row < 0) return { ok: false, msg: '查無此號碼（Points 沒有這一列）' };
    const before = Number(psh.getRange(row, 2).getValue()) || 0;
    if (before === target) return { ok: false, msg: '點數已經是 ' + target + '，不需修改' };

    psh.getRange(row, 2).setValue(target);     // 改點數
    psh.getRange(row, 4).setValue(new Date());  // 更新 updatedAt

    // Log 補一筆人工修正紀錄（append-only；note 說明前後值）
    const noteCol = ensureLogNoteCol_();
    const gsh = getSheet_(CFG.LOG_SHEET);
    const rowVals = [new Date(), phone, '', '', '', '', target];
    while (rowVals.length < noteCol - 1) rowVals.push('');   // 補齊到 note 欄前一格
    rowVals[noteCol - 1] = '人工修正 ' + before + '→' + target;
    gsh.appendRow(rowVals);

    const msg = '號碼 ' + phone + ' 點數：' + before + ' → ' + target + '（已在 Log 留紀錄）';
    Logger.log(msg);
    return { ok: true, phone: phone, before: before, after: target, msg: msg };
  } finally {
    lock.releaseLock();
  }
}

/** 兩個字串是否「長度相同、只差一個字元」（抓相近號碼／打錯字的分身） */
function nearIdentical_(a, b) {
  a = String(a); b = String(b);
  if (a.length !== b.length) return false;
  var diff = 0;
  for (var i = 0; i < a.length; i++) if (a[i] !== b[i]) { diff++; if (diff > 1) return false; }
  return diff === 1;
}

/**
 * 唯讀掃描：找出 Log 中「同一天、windowSec 秒內成對蓋章」的號碼組，用來抓
 * 跨號碼重複集點（例：一人用兩支相近號碼灌點）。已略過 note 標 test 的列。
 * 回傳：strong＝號碼只差一碼的組；repeat＝不同號碼但重複成對（≥2 天）的組。
 */
function scanSuspiciousPairs(windowSec) {
  ensureSetup_();
  var WIN = parseInt(windowSec, 10);
  if (isNaN(WIN) || WIN <= 0) WIN = 600;   // 預設 10 分鐘

  var data = getSheet_(CFG.LOG_SHEET).getDataRange().getValues();
  var header = data[0] || [];
  var noteCol = -1;
  for (var i = 0; i < header.length; i++) if (String(header[i]).trim().toLowerCase() === 'note') noteCol = i;

  // 依日期分組，收集 { 毫秒時間, 手機 }
  var byDay = {};
  for (var r = 1; r < data.length; r++) {
    var t = data[r][0];
    if (!(t instanceof Date)) continue;                       // 沒有有效時間就跳過
    var phone = normPhone_(data[r][1]);
    if (phone.length < 8) continue;
    if (noteCol >= 0 && String(data[r][noteCol] || '').toLowerCase().indexOf('test') !== -1) continue;
    var day = Utilities.formatDate(t, CFG.TZ, 'yyyy-MM-dd');
    (byDay[day] = byDay[day] || []).push({ ms: t.getTime(), phone: phone });
  }

  // 同一天、時間差 <= WIN 的不同號碼配對
  var pairs = {};
  for (var d in byDay) {
    var lst = byDay[d];
    lst.sort(function (x, y) { return x.ms - y.ms; });
    for (var a1 = 0; a1 < lst.length; a1++) {
      for (var b1 = a1 + 1; b1 < lst.length; b1++) {
        if (lst[a1].phone === lst[b1].phone) continue;
        var gap = Math.round(Math.abs(lst[b1].ms - lst[a1].ms) / 1000);
        if (gap > WIN) continue;
        var p1 = lst[a1].phone, p2 = lst[b1].phone;
        if (p1 > p2) { var tmp = p1; p1 = p2; p2 = tmp; }
        var key = p1 + '|' + p2;
        var rec = pairs[key] || (pairs[key] = { a: p1, b: p2, days: {}, minGap: 1e9 });
        if (rec.days[d] == null || gap < rec.days[d]) rec.days[d] = gap;
        if (gap < rec.minGap) rec.minGap = gap;
      }
    }
  }

  var list = [];
  for (var k in pairs) {
    var pr = pairs[k], c = 0;
    for (var dd in pr.days) c++;
    list.push({ a: pr.a, b: pr.b, count: c, minGap: pr.minGap, near: nearIdentical_(pr.a, pr.b), days: pr.days });
  }
  list.sort(function (x, y) {
    if (x.near !== y.near) return x.near ? -1 : 1;
    if (x.count !== y.count) return y.count - x.count;
    return x.minGap - y.minGap;
  });
  var strong = list.filter(function (x) { return x.near; });
  var repeat = list.filter(function (x) { return !x.near && x.count >= 2; });

  function fmtDays(days) { var o = []; for (var dk in days) o.push(dk + '(隔' + days[dk] + '秒)'); return o.join(', '); }
  var lines = ['可疑成對偵測（同一天 ' + WIN + ' 秒內成對蓋章；已略過 test 列）', ''];
  lines.push('【高度可疑】號碼只差一碼：' + (strong.length ? '' : '無'));
  strong.forEach(function (x) { lines.push('  ' + x.a + ' & ' + x.b + '  共 ' + x.count + ' 天  ' + fmtDays(x.days)); });
  lines.push('');
  lines.push('【值得留意】不同號碼但重複成對（≥2 天，多半是一起來的客人）：' + (repeat.length ? '' : '無'));
  repeat.forEach(function (x) { lines.push('  ' + x.a + ' & ' + x.b + '  共 ' + x.count + ' 天  ' + fmtDays(x.days)); });

  var summary = lines.join('\n');
  Logger.log(summary);
  return { ok: true, window: WIN, strong: strong, repeat: repeat, summary: summary };
}

/* ---- 「🛠 維護」選單用的對話框包裝 ---- */

function promptScanSuspiciousPairs_() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('掃描可疑成對（唯讀）', '同一天多少秒內算「成對」？（預設 600＝10 分鐘）', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var w = res.getResponseText().trim();
  var out = scanSuspiciousPairs(w === '' ? 600 : w);
  ui.alert(out.summary);
}

function promptBackupNow_() {
  const ui = SpreadsheetApp.getUi();
  const c = ui.alert('建立備份副本', '將複製整份試算表到你的雲端硬碟，繼續？', ui.ButtonSet.YES_NO);
  if (c !== ui.Button.YES) return;
  const out = backupNow_();
  ui.alert('備份完成：\n' + out.name + '\n' + out.url);
}

function promptFixPoints_() {
  const ui = SpreadsheetApp.getUi();
  const r1 = ui.prompt('修正點數', '請輸入手機號碼：', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const phone = r1.getResponseText();
  const pre = previewPhone(phone);
  if (!pre.ok) { ui.alert(pre.msg); return; }

  const r2 = ui.prompt('修正點數', pre.summary + '\n\n要把點數改成多少？（0 以上整數）', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;

  const confirm = ui.alert('確認修正？（建議先備份）',
    '號碼 ' + normPhone_(phone) + ' → 點數改成 ' + r2.getResponseText() + '，並在 Log 留紀錄。要繼續嗎？',
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  const out = fixInflatedPoints(phone, r2.getResponseText());
  ui.alert(out.ok ? ('完成：\n' + out.msg) : out.msg);
}
