/**
 * 車險到期追蹤系統　後端（Google Apps Script）v2.0
 *
 * 這份程式不含任何個人設定。每個人的設定都放在「專案設定 > 指令碼屬性」：
 *   API_KEY      密鑰（執行 setup() 會自動產生）
 *   NOTIFY_EMAIL 通知信箱（沒設定時寄給部署者本人）
 *   REMIND_DAYS  提醒天數，例如 60,30,7（可在網頁的「設定」修改）
 *   COV_OPTIONS  任意險勾選項目（可在網頁的「設定」修改）
 *   NOTIFY_HOUR  每天幾點檢查並寄信，0 到 23，預設 22（可在網頁的「設定」修改）
 *   SHEET_ID     只有「程式不是從試算表建立」時才需要填
 *   SHEET_NAME   資料所在的工作表名稱，預設「客戶資料」，找不到時用第一個工作表
 *   APP_URL      （選填）網頁網址，會附在提醒信最後
 */
const VERSION = '2.1';
const TZ = 'Asia/Taipei';
const HEADERS = ['id', 'name', 'prep', 'vehicles', 'note', 'updatedAt'];
const PREP = ['待聯繫', '已聯繫', '已記錄'];
const RENEW = ['', '已通知', '已報價', '已約訪', '已見面／已續保', '未見面／未續保'];
const RENEW_CLOSED = ['已見面／已續保', '未見面／未續保'];
const DATE_TYPES = [['comp', '強制險'], ['vol', '任意險']];
const DEFAULTS = {
  REMIND_DAYS: '60,30,7',
  NOTIFY_HOUR: '22',
  SHEET_NAME: '客戶資料',
  COV_OPTIONS: JSON.stringify(['第三人', '超額500', '超額1000', '超額2000', '甲式', '乙式', '丙式', '竊盜', '駕駛人傷害', '乘客體傷', '刑事訴訟', '道路救援'])
};

/* ---------- 第一次使用：確認執行按鈕旁的下拉選單顯示 setup，再按「執行」 ---------- */

function setup() {
  const p = props_();
  if (!p.getProperty('API_KEY')) p.setProperty('API_KEY', Utilities.getUuid().replace(/-/g, ''));
  if (!p.getProperty('REMIND_DAYS')) p.setProperty('REMIND_DAYS', DEFAULTS.REMIND_DAYS);
  if (!p.getProperty('COV_OPTIONS')) p.setProperty('COV_OPTIONS', DEFAULTS.COV_OPTIONS);
  getSheet_(); // 確認找得到試算表，並補齊欄位標題

  if (!p.getProperty('NOTIFY_HOUR')) p.setProperty('NOTIFY_HOUR', DEFAULTS.NOTIFY_HOUR);
  const hour = installTrigger_();

  Logger.log('設定完成，每天 ' + hour + ':00 到 ' + (hour + 1) + ':00 之間檢查一次。');
  Logger.log('你的密鑰：' + p.getProperty('API_KEY'));
  Logger.log('通知信會寄到：' + notifyEmail_());
}

// 忘記密鑰時執行，會把密鑰印在執行記錄
function showKey() { Logger.log('你的密鑰：' + props_().getProperty('API_KEY')); }

// 懷疑密鑰外流時執行，舊密鑰立即失效，所有裝置都要重新輸入
function resetKey() {
  props_().setProperty('API_KEY', Utilities.getUuid().replace(/-/g, ''));
  showKey();
}

/* ---------- 對外介面 ---------- */

// GET 只回報服務狀態，不回傳任何客戶資料
function doGet() {
  return json_({ ok: true, service: 'car-expiry-tracker', version: VERSION });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  const key = props_().getProperty('API_KEY');
  if (!key) return json_({ ok: false, error: 'not_setup' });
  if (!req || typeof req.key !== 'string' || req.key !== key) return json_({ ok: false, error: 'unauthorized' });
  try {
    switch (req.action) {
      case 'list':         return json_({ ok: true, data: { clients: listClients_(), settings: getSettings_(), version: VERSION } });
      case 'save':         return json_({ ok: true, data: withLock_(() => saveClient_(req.client)) });
      case 'delete':       return json_({ ok: true, data: withLock_(() => deleteClient_(req.id)) });
      case 'saveSettings': return json_({ ok: true, data: saveSettings_(req.settings) });
      case 'testMail':     return json_({ ok: true, data: sendTestMail_() });
      default:             return json_({ ok: false, error: 'unknown_action' });
    }
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

// 依 NOTIFY_HOUR 重建每日觸發器；Google 會在該小時內的某個時間執行
function installTrigger_() {
  const hour = parseHour_(prop_('NOTIFY_HOUR'));
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['dailyCheck', 'checkExpiry'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyCheck').timeBased().everyDays(1).atHour(hour).inTimezone(TZ).create();
  return hour;
}

/* ---------- 每日檢查（由 setup 建立的觸發器呼叫） ---------- */

function dailyCheck() {
  const alerts = withLock_(() => {
    const sh = getSheet_();
    const { rows, idx } = readAll_(sh);
    const stages = getSettings_().remindDays;
    const today = todayNum_();
    const found = [];
    rows.forEach(x => {
      let dirty = false;
      x.obj.vehicles.forEach(v => {
        if (RENEW_CLOSED.indexOf(v.status) >= 0) return;
        DATE_TYPES.forEach(([k, label]) => {
          const n = dayNum_(v[k]);
          if (n == null) return;
          const d = n - today;
          const hit = stages.filter(s => d <= s);
          if (d < 0) hit.push(-1); // -1 代表「已過期」這一階段
          const fresh = hit.filter(s => v.notified[k].indexOf(s) < 0);
          if (!fresh.length) return;
          v.notified[k] = v.notified[k].concat(fresh);
          dirty = true;
          found.push({ name: x.obj.name, plate: v.plate, label: label, date: v[k], d: d, status: v.status });
        });
      });
      if (dirty) {
        const cell = sh.getRange(x.row, idx.vehicles + 1);
        cell.setNumberFormat('@');
        cell.setValue(JSON.stringify(x.obj.vehicles));
      }
    });
    return found;
  });
  if (!alerts.length) return;
  alerts.sort((a, b) => a.d - b.d);
  const subject = '【車險到期提醒】' + alerts.length + ' 筆進入提醒期';
  const body = '以下車輛今天進入新的提醒階段：\n\n' + alerts.map(alertLine_).join('\n') + mailFooter_();
  MailApp.sendEmail(notifyEmail_(), subject, body);
}

// 舊版觸發器叫的是 checkExpiry，保留這個名稱避免舊觸發器出錯
function checkExpiry() { dailyCheck(); }

/* ---------- 資料存取 ---------- */

function listClients_() {
  return readAll_(getSheet_()).rows.map(x => x.obj);
}

function saveClient_(c) {
  if (!c || typeof c !== 'object') throw new Error('缺少客戶資料');
  const nm = String(c.name || '').trim();
  if (!nm) throw new Error('請填寫客戶姓名');
  const sh = getSheet_();
  const { rows, head } = readAll_(sh);
  const clean = {
    id: c.id ? String(c.id) : newId_(),
    name: nm.slice(0, 50),
    prep: PREP.indexOf(c.prep) >= 0 ? c.prep : '待聯繫',
    note: String(c.note || '').slice(0, 500),
    vehicles: (Array.isArray(c.vehicles) ? c.vehicles : []).slice(0, 20).map((v, i) => normVehicle_(v, i)),
    updatedAt: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm')
  };
  const old = rows.filter(x => x.obj.id === clean.id)[0];
  clean.vehicles.forEach((v, i) => {
    const ov = old ? (old.obj.vehicles.filter(o => o.vid === v.vid)[0]) : null;
    if (!ov) { v.notified = { comp: [], vol: [] }; return; }
    // 已寄過的提醒以後端紀錄為準；到期日改了就重新計算提醒
    DATE_TYPES.forEach(([k]) => { v.notified[k] = ov[k] === v[k] ? ov.notified[k] : []; });
    // 換了新的到期日（新保期），而且這次沒有同時改狀態，就把狀態清空
    const datesChanged = ov.comp !== v.comp || ov.vol !== v.vol;
    if (datesChanged && v.status === ov.status) { v.status = ''; v.reason = ''; }
  });
  const rowVals = head.map((h, j) => {
    if (h === 'vehicles') return JSON.stringify(clean.vehicles);
    if (HEADERS.indexOf(h) >= 0) return String(clean[h] == null ? '' : clean[h]);
    return old ? old.raw[j] : ''; // 自己在試算表加的其他欄位，原樣保留
  });
  const rowNo = old ? old.row : sh.getLastRow() + 1;
  const rg = sh.getRange(rowNo, 1, 1, head.length);
  rg.setNumberFormat('@');
  rg.setValues([rowVals]);
  return clean;
}

function deleteClient_(id) {
  const sh = getSheet_();
  const hit = readAll_(sh).rows.filter(x => x.obj.id === String(id))[0];
  if (!hit) throw new Error('找不到這位客戶，可能已被刪除');
  sh.deleteRow(hit.row);
  return { id: String(id) };
}

function readAll_(sh) {
  const values = sh.getDataRange().getValues();
  const head = values[0].map(String);
  const idx = {};
  head.forEach((h, i) => { idx[h] = i; });
  const rows = values.slice(1)
    .map((r, i) => ({ row: i + 2, raw: r, obj: rowToClient_(r, idx) }))
    .filter(x => x.obj.id);
  return { rows, idx, head };
}

function rowToClient_(r, idx) {
  const g = k => (idx[k] == null || r[idx[k]] == null) ? '' : String(r[idx[k]]);
  let vehicles = [];
  try { vehicles = JSON.parse(g('vehicles') || '[]'); } catch (e) {}
  if (!Array.isArray(vehicles)) vehicles = [];
  return {
    id: g('id'),
    name: g('name'),
    prep: PREP.indexOf(g('prep')) >= 0 ? g('prep') : '已記錄', // 舊資料沒有這一欄，視為已記錄
    note: g('note'),
    vehicles: vehicles.map((v, i) => normVehicle_(v, i)),
    updatedAt: g('updatedAt')
  };
}

function normVehicle_(v, i) {
  v = v || {};
  const n = v.notified || {};
  const arr = a => Array.isArray(a) ? a.filter(x => typeof x === 'number') : [];
  const date = s => /^\d{4}-\d{2}-\d{2}$/.test(String(s || '')) ? String(s) : '';
  return {
    vid: v.vid ? String(v.vid) : 'v' + i,
    type: v.type === '機車' ? '機車' : '汽車',
    plate: String(v.plate || '').trim().toUpperCase().slice(0, 12),
    comp: date(v.comp),
    vol: date(v.vol),
    coverage: Array.isArray(v.coverage) ? v.coverage.map(String).slice(0, 40) : [],
    status: RENEW.indexOf(v.status) >= 0 ? v.status : '',
    reason: String(v.reason || '').slice(0, 200),
    notified: { comp: arr(n.comp), vol: arr(n.vol) }
  };
}

/* ---------- 設定 ---------- */

function getSettings_() {
  let cov = [];
  try { cov = JSON.parse(prop_('COV_OPTIONS')); } catch (e) {}
  if (!Array.isArray(cov)) cov = JSON.parse(DEFAULTS.COV_OPTIONS);
  return { remindDays: parseDays_(prop_('REMIND_DAYS')), notifyEmail: notifyEmail_(), covOptions: cov, notifyHour: parseHour_(prop_('NOTIFY_HOUR')) };
}

function saveSettings_(s) {
  s = s || {};
  const p = props_();
  if (s.remindDays != null) {
    const days = parseDays_(Array.isArray(s.remindDays) ? s.remindDays.join(',') : String(s.remindDays));
    p.setProperty('REMIND_DAYS', days.join(','));
  }
  if (s.notifyEmail != null) {
    const em = String(s.notifyEmail).trim();
    if (em && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(em)) throw new Error('通知信箱格式不正確');
    if (em) p.setProperty('NOTIFY_EMAIL', em); else p.deleteProperty('NOTIFY_EMAIL');
  }
  if (Array.isArray(s.covOptions)) {
    const seen = {};
    const cov = s.covOptions.map(x => String(x).trim().slice(0, 20)).filter(x => x && !seen[x] && (seen[x] = true)).slice(0, 40);
    p.setProperty('COV_OPTIONS', JSON.stringify(cov));
  }
  if (s.notifyHour != null && String(s.notifyHour) !== '') {
    const h = Number(s.notifyHour);
    if (!Number.isInteger(h) || h < 0 || h > 23) throw new Error('檢查時間要是 0 到 23 的整數');
    if (h !== parseHour_(prop_('NOTIFY_HOUR'))) {
      const before = prop_('NOTIFY_HOUR');
      p.setProperty('NOTIFY_HOUR', String(h));
      try { installTrigger_(); }
      catch (err) {
        p.setProperty('NOTIFY_HOUR', before);
        throw new Error('無法從網頁更改檢查時間，請到 Apps Script 編輯器執行一次 setup 後再試。');
      }
    }
  }
  return getSettings_();
}

function parseHour_(v) { const h = Number(v); return Number.isInteger(h) && h >= 0 && h <= 23 ? h : 22; }

function parseDays_(str) {
  const seen = {};
  const d = String(str || '').split(/[,，\s]+/).map(Number)
    .filter(n => Number.isInteger(n) && n >= 1 && n <= 365 && !seen[n] && (seen[n] = true))
    .sort((a, b) => b - a).slice(0, 6);
  return d.length ? d : [60, 30, 7];
}

/* ---------- 郵件 ---------- */

function sendTestMail_() {
  const clients = listClients_();
  const max = getSettings_().remindDays[0];
  const today = todayNum_();
  const items = [];
  clients.forEach(c => c.vehicles.forEach(v => {
    if (RENEW_CLOSED.indexOf(v.status) >= 0) return;
    DATE_TYPES.forEach(([k, label]) => {
      const n = dayNum_(v[k]);
      if (n == null) return;
      const d = n - today;
      if (d <= max) items.push({ name: c.name, plate: v.plate, label: label, date: v[k], d: d, status: v.status });
    });
  }));
  items.sort((a, b) => a.d - b.d);
  const subject = '【測試信】共 ' + clients.length + ' 位客戶，' + max + ' 天內未結案 ' + items.length + ' 筆';
  const body = (items.length ? max + ' 天內到期、尚未結案的項目：\n\n' + items.map(alertLine_).join('\n') : '目前沒有 ' + max + ' 天內到期、尚未結案的項目。')
    + '\n\n這是測試信，不會影響正式提醒的紀錄。' + mailFooter_();
  const to = notifyEmail_();
  MailApp.sendEmail(to, subject, body);
  return { to: to, count: items.length };
}

function alertLine_(a) {
  const when = a.d < 0 ? '已過期 ' + (-a.d) + ' 天' : (a.d === 0 ? '今天到期' : '剩 ' + a.d + ' 天');
  return '・' + a.name + '（' + a.plate + '）' + a.label + ' ' + a.date.replace(/-/g, '/') + '，' + when
    + (a.status ? '（目前：' + a.status + '）' : '（尚未處理）');
}

function mailFooter_() {
  const url = props_().getProperty('APP_URL');
  return '\n\n處理完請到系統更新狀態；標記為「已續保」或「未續保」後就不會再提醒。' + (url ? '\n' + url : '');
}

/* ---------- 工具函式 ---------- */

function props_() { return PropertiesService.getScriptProperties(); }
function prop_(k) { const v = props_().getProperty(k); return v == null ? (DEFAULTS[k] || '') : v; }
function notifyEmail_() { return props_().getProperty('NOTIFY_EMAIL') || Session.getEffectiveUser().getEmail(); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function withLock_(fn) { const l = LockService.getScriptLock(); l.waitLock(20000); try { return fn(); } finally { l.releaseLock(); } }
function newId_() { return 'c' + Date.now().toString(36) + Math.floor(Math.random() * 46656).toString(36); }
function dayNum_(s) { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s || ''); return m ? Date.UTC(+m[1], +m[2] - 1, +m[3]) / 864e5 : null; }
function todayNum_() { return dayNum_(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd')); }

function getSheet_() {
  const id = props_().getProperty('SHEET_ID');
  const ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('找不到試算表：請在指令碼屬性加上 SHEET_ID');
  const sh = ss.getSheetByName(prop_('SHEET_NAME')) || ss.getSheets()[0];
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
  } else {
    const head = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
    HEADERS.forEach(h => { if (head.indexOf(h) < 0) { head.push(h); sh.getRange(1, head.length).setValue(h); } });
  }
  return sh;
}
