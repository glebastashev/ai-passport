var SHEET_ID   = '10JdWe9bwt7yVvwCSfcBmt_CLZGwxGiXznDG27ob8WN8';
var BOT_TOKEN  = '8705329586:AAGw7DX9o0kmZ0KjHYY8-ly5rDQP-ELGo14';
var SHEET_NAME = 'Заявки';
var AUTH_NAME  = 'Входы';

function api_(method, payload) {
  var res = UrlFetchApp.fetch('https://api.telegram.org/bot' + BOT_TOKEN + '/' + method, {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(payload || {}), muteHttpExceptions: true
  });
  return JSON.parse(res.getContentText());
}

function sheet_(name) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/** аватар держим в кэше, а не в таблице: иначе каждый опрос тащит мегабайты */
function avatar_(userId) {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('ava_' + userId);
  if (hit !== null) return hit;
  var photo = '';
  try {
    var ph = api_('getUserProfilePhotos', { user_id: userId, limit: 1 });
    if (ph.ok && ph.result.total_count) {
      var sizes = ph.result.photos[0];
      var f = api_('getFile', { file_id: sizes[Math.min(1, sizes.length - 1)].file_id });
      if (f.ok) {
        var blob = UrlFetchApp.fetch('https://api.telegram.org/file/bot' + BOT_TOKEN + '/' + f.result.file_path).getBlob();
        photo = 'data:image/jpeg;base64,' + Utilities.base64Encode(blob.getBytes());
      }
    }
  } catch (err) {
    photo = '';
  }
  try { cache.put('ava_' + userId, photo, 21600); } catch (err) {}
  return photo;
}

/** забираем новые нажатия Start; замок не даёт двум опросам обработать одно и то же */
function pull_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return;
  try {
    var props = PropertiesService.getScriptProperties();
    var offset = Number(props.getProperty('offset') || 0);
    var r = api_('getUpdates', { offset: offset, timeout: 0, allowed_updates: ['message'] });
    if (!r.ok || !r.result.length) return;

    var sh = sheet_(AUTH_NAME);
    var rows = [];
    r.result.forEach(function (upd) {
      offset = Math.max(offset, upd.update_id + 1);
      var msg = upd.message;
      if (!msg || !msg.text || msg.text.indexOf('/start') !== 0) return;
      var code = msg.text.split(' ')[1] || '';
      var u = msg.from;
      if (code) rows.push([new Date(), code, u.id, u.username || '',
        [u.first_name, u.last_name].filter(String).join(' ')]);
      api_('sendMessage', {
        chat_id: msg.chat.id,
        text: 'Готово, возвращайся на страницу паспорта: анкета уже открылась.'
      });
    });
    props.setProperty('offset', String(offset));
    if (rows.length) {
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, 5).setValues(rows);
      SpreadsheetApp.flush();
    }
  } finally {
    lock.releaseLock();
  }
}

var ORDER = ['sid', 'обновлено', 'этап', 'tg_name', 'tg_username', 'tg_id',
             'phone', 'job', 'income', 'why_me'];

/** одна анкета — одна строка: по sid находим её и обновляем, а не добавляем новую */
function doPost(e) {
  var data = (e && e.parameter) ? e.parameter : {};
  var lock = LockService.getScriptLock();
  lock.tryLock(20000);
  try {
    var sh = sheet_(SHEET_NAME);
    var last = sh.getLastRow();
    var headers = last > 0 ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].filter(String) : [];
    if (!headers.length) headers = ORDER.slice();
    Object.keys(data).forEach(function (k) { if (headers.indexOf(k) === -1) headers.push(k); });
    if (headers.indexOf('обновлено') === -1) headers.push('обновлено');
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');

    var sid = String(data.sid || '');
    var target = 0;
    if (sid && last > 1) {
      var col = headers.indexOf('sid') + 1;
      if (col > 0) {
        var vals = sh.getRange(2, col, last - 1, 1).getValues();
        for (var i = vals.length - 1; i >= 0; i--) {
          if (String(vals[i][0]) === sid) { target = i + 2; break; }
        }
      }
    }

    var row = headers.map(function (h) {
      if (h === 'обновлено') return new Date();
      return data[h] !== undefined ? data[h] : '';
    });

    if (target) {
      var old = sh.getRange(target, 1, 1, headers.length).getValues()[0];
      row = row.map(function (v, i) { return (v === '' && old[i] !== '') ? old[i] : v; });
      sh.getRange(target, 1, 1, headers.length).setValues([row]);
    } else {
      target = sh.getLastRow() + 1;
      sh.getRange(target, 1, 1, headers.length).setValues([row]);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true, row: target }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var cbName = p.callback || 'cb';

  if (p.check) {
    pull_();
    var sh = sheet_(AUTH_NAME);
    var last = sh.getLastRow();
    var out = { ok: false };
    if (last > 1) {
      var codes = sh.getRange(2, 2, last - 1, 1).getValues();   // читаем только колонку с кодом
      for (var i = codes.length - 1; i >= 0; i--) {
        if (String(codes[i][0]) === String(p.check)) {
          var row = sh.getRange(i + 2, 1, 1, 5).getValues()[0];
          out = { ok: true, id: row[2], username: row[3], name: row[4], photo: avatar_(row[2]) };
          break;
        }
      }
    }
    return ContentService.createTextOutput(cbName + '(' + JSON.stringify(out) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService.createTextOutput('ok');
}
