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
  REWARD: '指定餐點或飲料一份',       // 獎勵內容
  MIN_SPEND_WEEKDAY: 180,           // 平日（週二～週五）最低消費
  MIN_SPEND_WEEKEND: 280,           // 假日（週六、週日）最低消費
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
  return String(phone || '').replace(/[^0-9]/g, '');
}

function todayStr_() {
  return Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd');
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
    last = String(sh.getRange(row, 3).getValue() || '');
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
    const last = String(sh.getRange(row, 3).getValue() || '');
    let points = Number(sh.getRange(row, 2).getValue()) || 0;
    if (last === today) {
      return { ok: false, msg: '今天已經蓋過囉，明天再來！', points: points, target: CFG.TARGET };
    }
    points += 1;
    sh.getRange(row, 2).setValue(points);
    sh.getRange(row, 3).setValue(today);
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
