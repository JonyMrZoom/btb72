/**
 * BACK TO BACK 2026 — приём заявок с формы регистрации в Google Таблицу.
 *
 * Ключ SECRET должен совпадать с CONFIG.SECRET в файле registration.html
 */

const SECRET     = 'b2b2026-smeni-menya';
const SHEET_NAME = 'Заявки';

/** Заголовки столбцов. Порядок менять нельзя. */
const HEADERS = [
  'Дата заявки',      // A
  'Статус оплаты',    // B
  'Фамилия Имя',      // C
  'Никнейм',          // D
  'Телефон',          // E
  'Номинации',        // F
  'Мастер-классы',    // G
  'Сумма, ₽',         // H
  'Тариф',            // I
  'Город и команда',  // J
  'Согласие',         // K
  'Комментарий'       // L
];

const STATUS_COL = 2;
const PHONE_COL  = 5;
const STATUSES = ['Ждём скриншот', 'Скриншот отправлен', 'Оплачено', 'Отмена', 'Возврат'];

/* ========================================================================== */

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.secret !== SECRET) return out({ ok: false, error: 'bad secret' });

    const sh = getSheet();

    // Повторная отправка той же заявки — меняем статус, новую строку не плодим
    if (data.row) {
      const row = Number(data.row);
      if (row > 1 && row <= sh.getLastRow()) {
        const cur = sh.getRange(row, STATUS_COL).getValue();
        if (cur === 'Ждём скриншот' || cur === '') {
          sh.getRange(row, STATUS_COL).setValue(data.status || cur);
        }
        return out({ ok: true, row: row, updated: true });
      }
    }

    if (data.updateOnly) return out({ ok: true, skipped: true });

    sh.appendRow(buildRow(data));
    const row = sh.getLastRow();
    markDuplicatePhone(sh, row, data.phone);

    return out({ ok: true, row: row });

  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return out({ ok: true, service: 'BACK TO BACK registration', time: new Date() });
}

/* ========================================================================== */

function buildRow(d) {
  return [
    new Date(),
    d.status || 'Ждём скриншот',
    d.name || '',
    d.nick || '',
    "'" + (d.phone || ''),        // апостроф — чтобы телефон не стал числом
    (d.noms || []).join(', '),
    (d.mks  || []).join(', '),
    Number(d.sum) || 0,
    d.tariff || '',
    d.city || '',
    d.agree || '',
    ''
  ];
}

/** Помечает повторную заявку с того же телефона. */
function markDuplicatePhone(sh, row, phone) {
  if (!phone) return;
  const clean = String(phone).replace(/\D/g, '');
  if (clean.length < 10) return;
  const last = sh.getLastRow();
  if (last < 3) return;
  const all = sh.getRange(2, PHONE_COL, last - 1, 1).getValues().flat();
  const hits = all.filter(function (p) {
    return String(p).replace(/\D/g, '') === clean;
  }).length;
  if (hits > 1) {
    sh.getRange(row, HEADERS.length).setValue('Возможный дубль — телефон уже есть в списке');
  }
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  return sh;
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ==========================================================================
   Запусти один раз вручную: оформит таблицу, добавит выпадающий список
   статусов, закрепит шапку и раскрасит строки по статусу.
   ========================================================================== */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setFontSize(11)
    .setBackground('#2233DD').setFontColor('#FFFFFF')
    .setVerticalAlignment('middle');
  sh.setRowHeight(1, 38);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(4);

  const widths = [140, 160, 190, 170, 140, 250, 200, 100, 110, 190, 170, 230];
  widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUSES, true).setAllowInvalid(false).build();
  sh.getRange(2, STATUS_COL, 2000, 1).setDataValidation(rule);

  const range = sh.getRange(2, 1, 2000, HEADERS.length);
  const rules = [
    ['$B2="Оплачено"',           '#D8F0DC'],
    ['$B2="Скриншот отправлен"', '#FFF3C4'],
    ['$B2="Отмена"',             '#F2D4D4'],
    ['$B2="Возврат"',            '#E6E1F5']
  ].map(function (r) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=' + r[0]).setBackground(r[1]).setRanges([range]).build();
  });
  sh.setConditionalFormatRules(rules);

  sh.getRange(2, 8, 2000, 1).setNumberFormat('#,##0 ₽');
  sh.getRange(2, 1, 2000, 1).setNumberFormat('dd.MM.yyyy HH:mm');

  buildDashboard(ss);
  SpreadsheetApp.getUi().alert('Готово. Лист «Заявки» и сводка настроены.');
}

/** Сводка: заявки, оплаты, деньги, разбивка по номинациям. */
function buildDashboard(ss) {
  const name = 'Сводка';
  const old = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old);
  const d = ss.insertSheet(name, 0);

  const rows = [
    ['ПОКАЗАТЕЛЬ', 'ЗНАЧЕНИЕ'],
    ['Всего заявок',         '=COUNTA(Заявки!C2:C)'],
    ['Оплачено',             '=COUNTIF(Заявки!B2:B;"Оплачено")'],
    ['Скриншот на проверке', '=COUNTIF(Заявки!B2:B;"Скриншот отправлен")'],
    ['Ждём скриншот',        '=COUNTIF(Заявки!B2:B;"Ждём скриншот")'],
    ['', ''],
    ['Собрано, ₽',           '=SUMIF(Заявки!B2:B;"Оплачено";Заявки!H2:H)'],
    ['Ожидается, ₽',         '=SUMIF(Заявки!B2:B;"Скриншот отправлен";Заявки!H2:H)'],
    ['', ''],
    ['Kids',      '=COUNTIF(Заявки!F2:F;"*Kids*")'],
    ['Junior',    '=COUNTIF(Заявки!F2:F;"*Junior*")'],
    ['PowerMove', '=COUNTIF(Заявки!F2:F;"*PowerMove*")'],
    ['B-girl',    '=COUNTIF(Заявки!F2:F;"*B-girl*")'],
    ['Bboy PRO',  '=COUNTIF(Заявки!F2:F;"*Bboy PRO*")'],
    ['Footwork',  '=COUNTIF(Заявки!F2:F;"*Footwork*")'],
    ['TopRock',   '=COUNTIF(Заявки!F2:F;"*TopRock*")'],
    ['', ''],
    ['Заявок с мастер-классами', '=COUNTIF(Заявки!G2:G;"*Мастер*")']
  ];
  d.getRange(1, 1, rows.length, 2).setValues(rows);
  d.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#2233DD').setFontColor('#FFFFFF');
  d.setColumnWidth(1, 240);
  d.setColumnWidth(2, 140);
  d.getRange(7, 2, 2, 1).setNumberFormat('#,##0 ₽');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('BACK TO BACK')
    .addItem('Настроить таблицу', 'setupSheet')
    .addToUi();
}
