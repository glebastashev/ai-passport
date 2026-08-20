/**
 * Приём ответов из паспорта AI-Инженера в Google-таблицу.
 *
 * Как поставить:
 * 1. Создай таблицу на Google Диске.
 * 2. Расширения → Apps Script, вставь этот код вместо содержимого файла.
 * 3. Развернуть → Новое развёртывание → тип «Веб-приложение».
 *    Запуск от имени: я. Доступ: все, включая анонимных.
 * 4. Скопируй ссылку вида https://script.google.com/macros/s/AK.../exec
 *    и вставь её в index.html в константу SHEET_WEBHOOK.
 */

var SHEET_NAME = 'Заявки';

function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

  var data = (e && e.parameter) ? e.parameter : {};
  var headers = sh.getLastRow() > 0
    ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
    : [];

  // новые поля дописываем в шапку, порядок колонок не ломается
  Object.keys(data).forEach(function (k) {
    if (headers.indexOf(k) === -1) headers.push(k);
  });
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  var row = headers.map(function (h) { return data[h] || ''; });
  sh.appendRow(row);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService.createTextOutput('ok');
}
