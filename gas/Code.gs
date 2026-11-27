/**
 * 可愛反派狠餃色｜線上訂單 → Google Sheet
 * -------------------------------------------------------------
 * 使用方式（詳細步驟見 gas/README.md）：
 *   1. 建立一份空白 Google 試算表 →「擴充功能」→「Apps Script」
 *   2. 把這個檔案的內容整個貼上，存檔
 *   3. 執行一次 setup()（第一次會要求授權）
 *   4. 「部署」→「新增部署作業」→ 類型選「網頁應用程式」
 *      執行身分：我　／　誰可以存取：所有人
 *   5. 把產生的網址（…/exec）貼到網站 index.html 的 ORDER_API
 */

const SHEET = { ORDERS: '訂單', ITEMS: '訂單明細', PRODUCTS: '商品', SETTINGS: '設定' };
const STATUSES = ['新訂單', '製作中', '可取餐', '已出貨', '已完成', '已取消'];
const STATUS_COLORS = { '新訂單': '#FFE3EF', '製作中': '#FFF3D1', '可取餐': '#DFF7EA', '已出貨': '#E2F0FF', '已完成': '#EEEEEE', '已取消': '#F4CCCC' };
const ORDER_HEADERS = ['訂單編號', '建立時間', '狀態', '取餐方式', '門市', '取餐時間', '內用人數', '宅配地址', '姓名', '手機', 'Email', '付款方式', '備註', '品項', '小計', '運費', '應付總額', '來源'];
const ITEM_HEADERS = ['訂單編號', '建立時間', '品項代碼', '品名', '單價', '數量', '小計'];
const MODE_LABEL = { dinein: '內用', takeout: '外帶自取', delivery: '冷凍宅配' };

/* 與網站相同的預設商品（setup 時寫入「商品」工作表，之後以工作表為準） */
const DEFAULT_PRODUCTS = [
  ['p1', '招牌黃金鍋貼', 80], ['p2', '韭黃鮮蝦鍋貼', 120], ['p3', '麻辣牛肉鍋貼', 130], ['p4', '起司玉米鍋貼', 100],
  ['d1', '高麗菜豬肉水餃', 70], ['d2', '鮮蝦干貝水餃', 150], ['d3', '紅油抄手', 90], ['d4', '香菇素食水餃', 75],
  ['s1', '酸辣湯', 45], ['s2', '玉米濃湯', 45], ['s3', '古早味冬瓜茶', 35], ['s4', '濃醇無糖豆漿', 35],
  ['c1', '涼拌小黃瓜', 40], ['c2', '紅油皮蛋豆腐', 50], ['c3', '滷味拼盤', 70], ['c4', '辣味炒毛豆', 45],
  ['g1', '經典家庭組（30 入）', 520], ['g2', '招財鍋貼禮盒（50 入）', 980], ['g3', '職人饗宴禮盒（80 入）', 1480]
];
const DEFAULT_SETTINGS = [
  ['店家通知 Email', '', '新訂單會寄到這個信箱，留空則寄給試算表擁有者'],
  ['宅配運費', 160, '單位：元'],
  ['免運門檻', 1500, '宅配小計達此金額免運'],
  ['訂單編號前綴', 'CV', '例：CV251127-001'],
  ['寄送顧客確認信', '是', '填「是」或「否」'],
  ['接受訂單', '是', '填「否」可暫停線上訂單（例如公休）']
];

/* ============================================================
   試算表選單
============================================================ */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🥟 狠餃色訂單')
    .addItem('初始化 / 修復工作表', 'setup')
    .addItem('今日營收摘要', 'showTodaySummary')
    .addToUi();
}

/** 建立工作表、標題列、狀態下拉選單與顏色。重複執行不會清掉既有資料。 */
function setup() {
  const ss = SpreadsheetApp.getActive();
  const orders = ensureSheet_(ss, SHEET.ORDERS, ORDER_HEADERS);
  ensureSheet_(ss, SHEET.ITEMS, ITEM_HEADERS);
  const products = ensureSheet_(ss, SHEET.PRODUCTS, ['品項代碼', '品名', '單價', '上架']);
  const settings = ensureSheet_(ss, SHEET.SETTINGS, ['項目', '值', '說明']);

  if (products.getLastRow() < 2) {
    products.getRange(2, 1, DEFAULT_PRODUCTS.length, 4).setValues(DEFAULT_PRODUCTS.map(r => r.concat('是')));
  }
  if (settings.getLastRow() < 2) {
    settings.getRange(2, 1, DEFAULT_SETTINGS.length, 3).setValues(DEFAULT_SETTINGS);
  }

  // 狀態欄：下拉選單 + 依狀態上色
  const statusCol = ORDER_HEADERS.indexOf('狀態') + 1;
  const statusRange = orders.getRange(2, statusCol, orders.getMaxRows() - 1, 1);
  statusRange.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(false).build());
  const rules = STATUSES.map(s => SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$C2="${s}"`).setBackground(STATUS_COLORS[s])
    .setRanges([orders.getRange(2, 1, orders.getMaxRows() - 1, ORDER_HEADERS.length)]).build());
  orders.setConditionalFormatRules(rules);
  orders.getRange('B:B').setNumberFormat('yyyy/mm/dd hh:mm');
  orders.getRange('O:Q').setNumberFormat('"NT$"#,##0');
  orders.getRange('J:J').setNumberFormat('@');

  const first = ss.getSheets()[0];
  if (first.getName() !== SHEET.ORDERS && first.getLastRow() === 0) ss.deleteSheet(first);  // 刪掉預設空白「工作表1」
  SpreadsheetApp.getActive().toast('工作表已準備完成，接著請部署為網頁應用程式。', '狠餃色訂單', 8);
}

function ensureSheet_(ss, name, headers) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#FFD23F');
  sh.setFrozenRows(1);
  return sh;
}

/* ============================================================
   網站呼叫入口
============================================================ */
function doPost(e) {
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (data.website) return json_({ ok: true, orderNo: 'IGNORED' });   // 機器人陷阱欄位
    return json_(createOrder_(data));
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: err.message || String(err) });
  }
}

/** GET：?action=ping 檢查連線；?action=status&no=訂單編號&phone=手機末三碼 查訂單狀態 */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'status') {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.ORDERS);
    const rows = sh.getDataRange().getValues();
    const no = String(p.no || '').trim(), tail = String(p.phone || '').replace(/\D/g, '').slice(-3);
    for (let i = rows.length - 1; i > 0; i--) {
      if (rows[i][0] === no && tail.length === 3 && String(rows[i][9]).slice(-3) === tail) {
        return json_({ ok: true, orderNo: no, status: rows[i][2], total: rows[i][16] });
      }
    }
    return json_({ ok: false, error: '查無此訂單' });
  }
  return json_({ ok: true, service: '狠餃色訂單 API', time: new Date().toISOString() });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ============================================================
   建立訂單
============================================================ */
function createOrder_(d) {
  const settings = readSettings_();
  if (settings['接受訂單'] === '否') throw new Error('目前暫停線上訂單，請改電話或現場點餐。');

  // ---- 基本驗證（前端也有驗，這裡再驗一次）----
  const mode = String(d.mode || '');
  if (!MODE_LABEL[mode]) throw new Error('取餐方式不正確');
  const name = clean_(d.name, 40), phone = String(d.phone || '').replace(/[\s-]/g, '');
  const email = clean_(d.email, 120), address = clean_(d.address, 200);
  if (!name) throw new Error('請填寫姓名');
  if (!/^09\d{8}$/.test(phone)) throw new Error('手機號碼格式不正確');
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Email 格式不正確');
  if (mode === 'delivery' && address.length < 8) throw new Error('請填寫完整宅配地址');
  if (!Array.isArray(d.items) || !d.items.length) throw new Error('購物車是空的');

  // ---- 以「商品」工作表的價格重新計算，避免被竄改 ----
  const catalog = readProducts_();
  const lines = d.items.map(it => {
    const p = catalog[String(it.id)];
    const qty = Math.floor(Number(it.qty));
    if (!p) throw new Error(`商品 ${it.id} 不存在或已下架`);
    if (!(qty >= 1 && qty <= 99)) throw new Error(`「${p.name}」數量不正確`);
    return { id: String(it.id), name: p.name, price: p.price, qty, sub: p.price * qty };
  });
  const subtotal = lines.reduce((s, l) => s + l.sub, 0);
  const shipFee = Number(settings['宅配運費']) || 0, freeShip = Number(settings['免運門檻']) || Infinity;
  const shipping = mode === 'delivery' && subtotal < freeShip ? shipFee : 0;
  const total = subtotal + shipping;

  // ---- 產生訂單編號並寫入（加鎖避免同時下單編號重複）----
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let orderNo, now;
  try {
    now = new Date();
    orderNo = nextOrderNo_(settings['訂單編號前綴'] || 'CV', now);
    const ss = SpreadsheetApp.getActive();
    const summary = lines.map(l => `${l.name}×${l.qty}`).join('、');
    ss.getSheetByName(SHEET.ORDERS).appendRow([
      orderNo, now, '新訂單', MODE_LABEL[mode],
      mode === 'delivery' ? '' : clean_(d.storeName, 40),
      mode === 'delivery' ? '' : clean_(d.time, 40),
      mode === 'dinein' ? clean_(d.people, 10) : '',
      mode === 'delivery' ? address : '',
      name, "'" + phone, email, clean_(d.pay, 20), clean_(d.note, 300), summary,
      subtotal, shipping, total, clean_(d.source, 20) || 'web'
    ]);
    const itemSheet = ss.getSheetByName(SHEET.ITEMS);
    const itemRows = lines.map(l => [orderNo, now, l.id, l.name, l.price, l.qty, l.sub]);
    itemSheet.getRange(itemSheet.getLastRow() + 1, 1, itemRows.length, ITEM_HEADERS.length).setValues(itemRows);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  // ---- Email 通知（失敗不影響訂單成立）----
  const order = { orderNo, now, mode, name, phone, email, address, lines, subtotal, shipping, total,
    storeName: clean_(d.storeName, 40), time: clean_(d.time, 40), people: clean_(d.people, 10), pay: clean_(d.pay, 20), note: clean_(d.note, 300) };
  try { notifyStore_(order, settings); } catch (err) { console.warn('店家通知寄送失敗', err); }
  if (email && settings['寄送顧客確認信'] !== '否') {
    try { notifyCustomer_(order); } catch (err) { console.warn('顧客確認信寄送失敗', err); }
  }

  return { ok: true, orderNo, subtotal, shipping, total };
}

/** 每天從 001 開始編號：CV251127-001 */
function nextOrderNo_(prefix, now) {
  const tz = Session.getScriptTimeZone() || 'Asia/Taipei';
  const day = Utilities.formatDate(now, tz, 'yyMMdd');
  const props = PropertiesService.getScriptProperties();
  const key = 'seq_' + day;
  const n = Number(props.getProperty(key) || 0) + 1;
  props.setProperty(key, String(n));
  return `${prefix}${day}-${String(n).padStart(3, '0')}`;
}

function readProducts_() {
  const rows = SpreadsheetApp.getActive().getSheetByName(SHEET.PRODUCTS).getDataRange().getValues().slice(1);
  const map = {};
  rows.forEach(([id, name, price, on]) => {
    if (id && String(on).trim() !== '否' && Number(price) > 0) map[String(id).trim()] = { name: String(name), price: Number(price) };
  });
  return map;
}

function readSettings_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.SETTINGS);
  const out = {};
  if (sh) sh.getDataRange().getValues().slice(1).forEach(([k, v]) => { if (k) out[String(k).trim()] = typeof v === 'string' ? v.trim() : v; });
  return out;
}

function clean_(v, max) {
  return String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max)
    .replace(/^[=+\-@]/, "'$&");   // 防止被當成試算表公式
}

/* ============================================================
   Email
============================================================ */
function orderText_(o) {
  const where = o.mode === 'delivery' ? `宅配地址：${o.address}` : `門市：${o.storeName}\n時間：${o.time}${o.mode === 'dinein' ? `\n人數：${o.people}` : ''}`;
  return [
    `訂單編號：${o.orderNo}`,
    `取餐方式：${MODE_LABEL[o.mode]}`,
    where,
    `付款方式：${o.pay}`,
    '',
    ...o.lines.map(l => `・${l.name} × ${l.qty}　NT$${l.sub}`),
    '',
    `小計：NT$${o.subtotal}`,
    `運費：NT$${o.shipping}`,
    `應付總額：NT$${o.total}`,
    o.note ? `\n備註：${o.note}` : ''
  ].join('\n');
}

function notifyStore_(o, settings) {
  const to = settings['店家通知 Email'] || Session.getEffectiveUser().getEmail();
  if (!to) return;
  MailApp.sendEmail({
    to,
    subject: `🥟 新訂單 ${o.orderNo}｜${MODE_LABEL[o.mode]}｜NT$${o.total}`,
    body: `${orderText_(o)}\n\n顧客：${o.name}（${o.phone}${o.email ? '，' + o.email : ''}）\n\n試算表：${SpreadsheetApp.getActive().getUrl()}`
  });
}

function notifyCustomer_(o) {
  MailApp.sendEmail({
    to: o.email,
    name: '可愛反派狠餃色',
    subject: `【可愛反派狠餃色】訂單已收到 ${o.orderNo}`,
    body: `${o.name} 您好：\n\n感謝您的訂購，我們已收到您的訂單！\n\n${orderText_(o)}\n\n` +
      (o.mode === 'delivery' ? '我們將於 1–3 個工作天內以黑貓低溫宅配出貨。' : '取餐時請告知訂單編號或手機末三碼。') +
      '\n\n可愛反派狠餃色 敬上'
  });
}

/* ============================================================
   今日營收摘要（選單功能）
============================================================ */
function showTodaySummary() {
  const tz = Session.getScriptTimeZone() || 'Asia/Taipei';
  const today = Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd');
  const rows = SpreadsheetApp.getActive().getSheetByName(SHEET.ORDERS).getDataRange().getValues().slice(1)
    .filter(r => r[1] instanceof Date && Utilities.formatDate(r[1], tz, 'yyyy/MM/dd') === today && r[2] !== '已取消');
  const total = rows.reduce((s, r) => s + Number(r[16] || 0), 0);
  const byMode = {};
  rows.forEach(r => { byMode[r[3]] = (byMode[r[3]] || 0) + 1; });
  SpreadsheetApp.getUi().alert(`📅 ${today}\n\n訂單數：${rows.length} 筆\n營收：NT$${total.toLocaleString()}\n\n` +
    Object.entries(byMode).map(([k, v]) => `${k}：${v} 筆`).join('\n'));
}
