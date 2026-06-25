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
  TARGET: 10,                       // 集滿幾點可兌換
  REWARD: '指定餐點或飲料一份',       // 獎勵內容
  MIN_SPEND_WEEKDAY: 180,           // 平日（週二～週五）最低消費
  MIN_SPEND_WEEKEND: 280,           // 假日（週六、週日）最低消費
  SHOP: 'MWD 泰山明志 BRUNCH',
  TZ: 'Asia/Taipei',
  // ---- 定位檢查（防止密碼外流後遠端亂蓋）----
  REQUIRE_LOCATION: true,           // 是否啟用定位檢查
  SHOP_LAT: 25.055472543375735,     // 店家緯度
  SHOP_LNG: 121.43159737244797,     // 店家經度
  RADIUS_M: 150,                    // 允許蓋章的範圍（公尺）
};

/** 部署成網頁應用程式時的進入點 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
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
  if (c.getLastRow() === 0) {
    c.appendRow(['key', 'value']);
    c.appendRow(['todayCode', '1234']); // 每天早上改這格（櫃台小白板上的數字）
    c.appendRow(['staffCode', '8888']); // 兌換用，店員私下保管，別公開
  }
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
    if (String(code || '').trim() !== getConfig_('todayCode')) {
      return { ok: false, msg: '今日通關碼不正確' };
    }
    if (CFG.REQUIRE_LOCATION) {
      const la = parseFloat(lat), ln = parseFloat(lng);
      if (isNaN(la) || isNaN(ln)) {
        return { ok: false, msg: '請開啟定位權限才能蓋章' };
      }
      const dist = distanceMeters_(la, ln, CFG.SHOP_LAT, CFG.SHOP_LNG);
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
