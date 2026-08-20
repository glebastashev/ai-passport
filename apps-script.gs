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

function avatar_(userId) {
  try {
    var ph = api_('getUserProfilePhotos', { user_id: userId, limit: 1 });
    if (!ph.ok || !ph.result.total_count) return '';
    var sizes = ph.result.photos[0];
    var fileId = sizes[Math.min(1, sizes.length - 1)].file_id;
    var f = api_('getFile', { file_id: fileId });
    if (!f.ok) return '';
    var blob = UrlFetchApp.fetch('https://api.telegram.org/file/bot' + BOT_TOKEN + '/' + f.result.file_path).getBlob();
    return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch (err) {
    return '';
  }
}

/** забираем новые нажатия Start и пишем их в лист «Входы» */
function pull_() {
  var props = PropertiesService.getScriptProperties();
  var offset = Number(props.getProperty('offset') || 0);
  var r = api_('getUpdates', { offset: offset, timeout: 0, allowed_updates: ['message'] });
  if (!r.ok || !r.result.length) return;

  var sh = sheet_(AUTH_NAME);
  r.result.forEach(function (upd) {
    offset = Math.max(offset, upd.update_id + 1);
    var msg = upd.message;
    if (!msg || !msg.text || msg.text.indexOf('/start') !== 0) return;
    var code = msg.text.split(' ')[1] || '';
    var u = msg.from;
    if (code) {
      sh.appendRow([new Date(), code, u.id, u.username || '',
        [u.first_name, u.last_name].filter(String).join(' '), '']);
    }
    api_('sendMessage', {
      chat_id: msg.chat.id,
      text: 'Готово, возвращайся на страницу паспорта: анкета уже открылась.'
    });
  });
  props.setProperty('offset', String(offset));
}

function doPost(e) {
  var data = (e && e.parameter) ? e.parameter : {};
  var sh = sheet_(SHEET_NAME);
  var headers = sh.getLastRow() > 0 ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
  Object.keys(data).forEach(function (k) { if (headers.indexOf(k) === -1) headers.push(k); });
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sh.appendRow(headers.map(function (h) { return data[h] || ''; }));
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};

  if (p.check) {
    pull_();                                   // сначала забираем свежие нажатия
    var sh = sheet_(AUTH_NAME);
    var rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues() : [];
    for (var i = rows.length - 1; i >= 0; i--) {
      if (String(rows[i][1]) === String(p.check)) {
        var photo = rows[i][5];
        if (!photo) {
          photo = avatar_(rows[i][2]);
          if (photo) sh.getRange(i + 2, 6).setValue(photo);
        }
        var out = { ok: true, id: rows[i][2], username: rows[i][3], name: rows[i][4], photo: photo };
        return ContentService.createTextOutput((p.callback || 'cb') + '(' + JSON.stringify(out) + ')')
          .setMimeType(ContentService.MimeType.JAVASCRIPT);
      }
    }
    return ContentService.createTextOutput((p.callback || 'cb') + '({"ok":false})')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService.createTextOutput('ok');
}
