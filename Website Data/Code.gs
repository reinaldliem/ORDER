/**********************************************************************
 * BACKEND ORDER WEBSITE — Google Apps Script (v2: login + peran)
 * --------------------------------------------------------------------
 * Menyimpan data ke Google Sheets + login ID/PIN dengan peran:
 *   - admin : lihat semua order, kelola barang/harga/akun, export.
 *   - sales : hanya lihat & kelola order miliknya sendiri.
 *
 * Akun default pertama kali:  ID = admin   PIN = 1234   (peran admin)
 * GANTI PIN ini lewat menu Pengaturan setelah login!
 *
 * Setiap kali Code.gs diubah: Deploy > Manage deployments > edit (pensil)
 *   > Version: New version > Deploy. URL tetap sama.
 *********************************************************************/

var SHEET_ORDERS = 'Orders';
var SHEET_CONFIG = 'Config';

function getSS() { return SpreadsheetApp.getActiveSpreadsheet(); }

function ensureSheets() {
  var ss = getSS();
  var o = ss.getSheetByName(SHEET_ORDERS);
  if (!o) {
    o = ss.insertSheet(SHEET_ORDERS);
    o.appendRow(['no', 'date', 'dateKey', 'sales', 'customer', 'items', 'total', 'createdAt', 'editedAt']);
  }
  o.getRange(1, 1, o.getMaxRows(), 9).setNumberFormat('@');

  var c = ss.getSheetByName(SHEET_CONFIG);
  if (!c) {
    c = ss.insertSheet(SHEET_CONFIG);
    c.appendRow(['key', 'value']);
    c.appendRow(['barangList', JSON.stringify(['Produk A', 'Produk B', 'Produk C', 'Produk D'])]);
    c.appendRow(['priceMap', JSON.stringify({})]);
    c.appendRow(['users', JSON.stringify([{ id: 'admin', pin: '1234', role: 'admin' }])]);
  }
  return { orders: o, config: c };
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function parseJSON(v, fallback) {
  try { return JSON.parse(v); } catch (e) { return fallback; }
}

/* ---- Config helpers ---- */
function readConfig(sh) {
  var data = sh.config.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) map[data[i][0]] = data[i][1];
  return map;
}

function getUsers(sh) {
  var users = parseJSON(readConfig(sh).users, null);
  if (!users || !users.length) users = [{ id: 'admin', pin: '1234', role: 'admin' }];
  return users;
}

function setConfig(sh, key, value) {
  var data = sh.config.getDataRange().getValues();
  var val = JSON.stringify(value);
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) { sh.config.getRange(i + 1, 2).setValue(val); return; }
  }
  sh.config.appendRow([key, val]);
}

/* ---- Auth ---- */
function authenticate(sh, id, pin) {
  var users = getUsers(sh);
  for (var i = 0; i < users.length; i++) {
    if (String(users[i].id) === String(id) && String(users[i].pin) === String(pin)) return users[i];
  }
  return null;
}

/* ---- Orders ---- */
function readOrders(sh) {
  var data = sh.orders.getDataRange().getValues();
  var orders = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[0]) continue;
    orders.push({
      no: r[0], date: r[1], dateKey: r[2], sales: r[3], customer: r[4],
      items: parseJSON(r[5], []), total: Number(r[6]) || 0,
      createdAt: r[7], editedAt: r[8]
    });
  }
  orders.reverse();
  return orders;
}

function findRow(sheet, no) {
  var col = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
  for (var i = 1; i < col.length; i++) if (col[i][0] === no) return i + 1;
  return -1;
}

/* ============================ READ ============================ */
function doGet(e) {
  var sh = ensureSheets();
  var p = e && e.parameter ? e.parameter : {};
  var user = authenticate(sh, p.id, p.pin);
  if (!user) return json({ ok: false, error: 'auth' });

  var cfg = readConfig(sh);
  var users = getUsers(sh);
  var orders = readOrders(sh);
  if (user.role !== 'admin') orders = orders.filter(function (o) { return o.sales === user.id; });

  return json({
    ok: true,
    role: user.role,
    me: user.id,
    orders: orders,
    salesList: users.filter(function (u) { return u.role === 'sales'; }).map(function (u) { return u.id; }),
    barangList: parseJSON(cfg.barangList, []),
    priceMap: parseJSON(cfg.priceMap, {}),
    users: user.role === 'admin' ? users : []
  });
}

/* ============================ WRITE ============================ */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = ensureSheets();
    var body = JSON.parse(e.postData.contents);
    var auth = body.auth || {};
    var user = authenticate(sh, auth.id, auth.pin);
    if (!user) return json({ ok: false, error: 'auth' });

    var action = body.action;
    var admin = user.role === 'admin';

    if (action === 'saveOrder') {
      var o = body.order;
      if (!admin) o.sales = user.id;                       // sales hanya bisa atas namanya
      var prev = findRow(sh.orders, o.no);
      if (prev > 0 && !admin) {                            // sales hanya boleh edit miliknya
        var existingSales = sh.orders.getRange(prev, 4).getValue();
        if (existingSales !== user.id) return json({ ok: false, error: 'forbidden' });
      }
      saveOrder(sh, o);

    } else if (action === 'deleteOrder') {
      var row = findRow(sh.orders, body.no);
      if (row > 0) {
        if (!admin && sh.orders.getRange(row, 4).getValue() !== user.id)
          return json({ ok: false, error: 'forbidden' });
        sh.orders.deleteRow(row);
      }

    } else if (action === 'setConfig') {
      if (!admin) return json({ ok: false, error: 'forbidden' });
      setConfig(sh, body.key, body.value);

    } else if (action === 'saveUser') {
      if (!admin) return json({ ok: false, error: 'forbidden' });
      var users = getUsers(sh);
      var u = body.user;
      var idx = -1;
      for (var i = 0; i < users.length; i++) if (users[i].id === u.id) idx = i;
      if (idx >= 0) users[idx] = u; else users.push(u);
      setConfig(sh, 'users', users);

    } else if (action === 'deleteUser') {
      if (!admin) return json({ ok: false, error: 'forbidden' });
      var us = getUsers(sh).filter(function (x) { return x.id !== body.id; });
      if (!us.some(function (x) { return x.role === 'admin'; }))
        return json({ ok: false, error: 'need-admin' });   // jangan sampai admin habis
      setConfig(sh, 'users', us);

    } else {
      return json({ ok: false, error: 'unknown-action' });
    }

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function saveOrder(sh, o) {
  var row = [
    o.no, o.date || '', o.dateKey || '', o.sales || '', o.customer || '',
    JSON.stringify(o.items || []), o.total || 0,
    o.createdAt || new Date().toISOString(), o.editedAt || ''
  ];
  var existing = findRow(sh.orders, o.no);
  var target = existing > 0 ? existing : sh.orders.getLastRow() + 1;
  var range = sh.orders.getRange(target, 1, 1, row.length);
  range.setNumberFormat('@');
  range.setValues([row]);
}
