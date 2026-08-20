/**
 * Бэкенд паспорта AI-Инженера на Google Apps Script.
 * Делает две вещи:
 *   1) принимает заполненные анкеты и пишет их в таблицу;
 *   2) работает вебхуком бота: человек жмёт Start, страница узнаёт его имя, ник и аватар.
 *
 * Установка:
 *  1. Создай таблицу на Google Диске.
 *  2. Расширения → Apps Script, вставь этот код.
 *  3. В строке BOT_TOKEN впиши токен бота (здесь он безопасен: код на сервере).
 *  4. Развернуть → Новое развёртывание → Веб-приложение.
 *     Запуск от имени: я. Доступ: все, включая анонимных.
 *  5. Скопируй ссылку .../exec и:
 *     — вставь её в index.html в SHEET_WEBHOOK;
 *     — один раз открой в браузере ссылку .../exec?setup=1, чтобы бот начал слать апдейты сюда.
 */

var BOT_TOKEN  = 'ВСТАВЬ_ТОКЕН_БОТА';
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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function doPost(e) {
  // ---- апдейт от бота: человек нажал Start с кодом ----
  var raw = e && e.postData ? e.postData.contents : '';
  if (raw && raw.charAt(0) === '{') {
    var upd = JSON.parse(raw);

    // Telegram повторяет доставку, пока не получит быстрый ответ.
    // Один и тот же апдейт обрабатываем строго один раз.
    var cache = CacheService.getScriptCache();
    var key = 'upd' + upd.update_id;
    if (cache.get(key)) return ContentService.createTextOutput('ok');
    cache.put(key, '1', 21600);

    var msg = upd.message;
    if (msg && msg.text && msg.text.indexOf('/start') === 0) {
      var code = msg.text.split(' ')[1] || '';
      var u = msg.from;
      if (code) {
        // аватар здесь не качаем: это долго и приводит к повторным доставкам
        sheet_(AUTH_NAME).appendRow([
          new Date(), code, u.id, u.username || '',
          [u.first_name, u.last_name].filter(String).join(' ')
        ]);
      }
      api_('sendMessage', {
        chat_id: msg.chat.id,
        text: 'Готово, возвращайся на страницу паспорта: данные уже подтянулись.'
      });
    }
    return ContentService.createTextOutput('ok');
  }

  // ---- обычная отправка анкеты ----
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

  // один раз: привязать вебхук бота к этому развёртыванию
  if (p.setup) {
    var url = ScriptApp.getService().getUrl();
    var r = api_('setWebhook', { url: url });
    return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
  }

  // страница спрашивает: этот код уже подтвердили?
  if (p.check) {
    var sh = sheet_(AUTH_NAME);
    var rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues() : [];
    for (var i = rows.length - 1; i >= 0; i--) {
      if (String(rows[i][1]) === String(p.check)) {
        var out = { ok: true, id: rows[i][2], username: rows[i][3], name: rows[i][4] };
        return ContentService.createTextOutput((p.callback || 'cb') + '(' + JSON.stringify(out) + ')')
          .setMimeType(ContentService.MimeType.JAVASCRIPT);
      }
    }
    return ContentService.createTextOutput((p.callback || 'cb') + '({"ok":false})')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService.createTextOutput('ok');
}
