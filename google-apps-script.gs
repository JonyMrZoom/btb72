/**
 * BACK TO BACK 2026 — таблица заявок.
 *
 * Что делает:
 *   • принимает заявки с формы регистрации;
 *   • раскладывает номинации по отдельным столбцам, чтобы их можно было фильтровать;
 *   • строит дашборд с наполнением номинаций и долей городов;
 *   • переживает ручные правки: сортировку, вставку и удаление строк.
 *
 * Установка:
 *   1. Создай пустую Google Таблицу.
 *   2. Расширения → Apps Script, вставь этот код, поменяй SECRET.
 *   3. Сохрани, потом меню «BACK TO BACK» → «Собрать таблицу».
 *   4. Развернуть → Новое развёртывание → Веб-приложение → доступ «У всех».
 *   5. Ссылку /exec вставь в форму в поле ENDPOINT, туда же тот же SECRET.
 */

const SECRET = 'b2b2026-smeni-menya';

const SHEET = 'Заявки';
const DASH  = 'Дашборд';
const HELP  = 'Служебное';
const INFO  = 'Инструкция';

/** Номинации. Порядок задаёт порядок столбцов-флагов и строк на дашборде. */
const NOMS = ['Kids', 'Junior', 'PowerMove', 'B-girl', 'Bboy PRO', 'Footwork', 'TopRock'];

const BASE = ['ID', 'Дата заявки', 'Статус оплаты', 'Фамилия Имя', 'Никнейм', 'Телефон',
              'Город', 'Номинации', 'Мастер-классы', 'Сумма, ₽', 'Тариф', 'Согласие', 'Комментарий'];
const HEADERS = BASE.concat(NOMS);

const C = { ID:1, DATE:2, STATUS:3, NAME:4, NICK:5, PHONE:6, CITY:7,
            NOMS:8, MK:9, SUM:10, TARIFF:11, AGREE:12, NOTE:13, FLAG:14 };

const STATUSES = ['Ждём скриншот', 'Скриншот отправлен', 'Оплачено', 'Отмена', 'Возврат'];
const MARK = '✓';

const BLUE = '#2233DD';

/* ========================================================================== */
/*  ПРИЁМ ЗАЯВОК                                                              */
/* ========================================================================== */

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.secret !== SECRET) return out({ ok:false, error:'bad secret' });

    const sh = getSheet();

    // Заявка уже есть — меняем статус. Ищем по ID, а не по номеру строки,
    // поэтому сортировка и ручные вставки строк ничего не ломают.
    if (d.id) {
      const row = findRowById(sh, d.id);
      if (row) {
        const cur = String(sh.getRange(row, C.STATUS).getValue());
        if (cur === 'Ждём скриншот' || cur === '') {
          sh.getRange(row, C.STATUS).setValue(d.status || cur);
        }
        return out({ ok:true, id:d.id, updated:true });
      }
      if (d.updateOnly) return out({ ok:true, notFound:true });
    }

    if (d.updateOnly) return out({ ok:true, skipped:true });

    const id = d.id || newId();
    sh.appendRow(buildRow(d, id));
    const row = sh.getLastRow();
    sh.getRange(row, C.DATE).setNumberFormat('dd.MM.yyyy HH:mm');
    sh.getRange(row, C.SUM).setNumberFormat('#,##0 ₽');
    markDuplicate(sh, row, d.phone);

    return out({ ok:true, id:id, row:row });

  } catch (err) {
    return out({ ok:false, error:String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return out({ ok:true, service:'BACK TO BACK registration', time:new Date() });
}

function newId() {
  return 'B' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
}

/** Поиск строки по ID. Работает при любом порядке строк. */
function findRowById(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const ids = sh.getRange(2, C.ID, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return 0;
}

function buildRow(d, id) {
  const noms = (d.noms || []).join(', ');
  const row = [
    id,
    new Date(),
    d.status || 'Ждём скриншот',
    d.name || '',
    d.nick || '',
    "'" + (d.phone || ''),          // апостроф — иначе телефон станет числом
    d.city || '',
    noms,
    (d.mks || []).join(', '),
    Number(d.sum) || 0,
    d.tariff || '',
    d.agree || '',
    ''
  ];
  return row.concat(flagsFor(noms));
}

/** Раскладывает строку с номинациями по флагам ✓ для каждого столбца. */
function flagsFor(text) {
  const t = String(text || '');
  return NOMS.map(function (n) { return t.indexOf(n) >= 0 ? MARK : ''; });
}

function markDuplicate(sh, row, phone) {
  if (!phone) return;
  const clean = String(phone).replace(/\D/g, '');
  if (clean.length < 10) return;
  const last = sh.getLastRow();
  if (last < 3) return;
  const all = sh.getRange(2, C.PHONE, last - 1, 1).getValues();
  let hits = 0;
  all.forEach(function (p) { if (String(p[0]).replace(/\D/g, '') === clean) hits++; });
  if (hits > 1) sh.getRange(row, C.NOTE).setValue('Возможный дубль — телефон уже есть в списке');
}

/* ========================================================================== */
/*  РУЧНЫЕ ПРАВКИ                                                             */
/* ========================================================================== */

/**
 * Если менеджер руками поправил столбец «Номинации» — флаги пересчитываются сами.
 * Без этого фильтры по номинациям разошлись бы с текстом.
 */
function onEdit(e) {
  try {
    const sh = e.range.getSheet();
    if (sh.getName() !== SHEET) return;
    const row = e.range.getRow();
    const col = e.range.getColumn();
    if (row < 2) return;
    if (col > C.NOMS || col + e.range.getNumColumns() - 1 < C.NOMS) return;

    const n = e.range.getNumRows();
    const vals = sh.getRange(row, C.NOMS, n, 1).getValues();
    const flags = vals.map(function (v) { return flagsFor(v[0]); });
    sh.getRange(row, C.FLAG, n, NOMS.length).setValues(flags);
  } catch (err) {
    // тихо: onEdit не должен ломать работу с таблицей
  }
}

/** Пересобирает флаги для всех строк. Пригодится после массовой правки или вставки. */
function recalcFlags() {
  const sh = getSheet();
  const last = sh.getLastRow();
  if (last < 2) return;
  const vals = sh.getRange(2, C.NOMS, last - 1, 1).getValues();
  const flags = vals.map(function (v) { return flagsFor(v[0]); });
  sh.getRange(2, C.FLAG, last - 1, NOMS.length).setValues(flags);
  SpreadsheetApp.getActive().toast('Флаги номинаций пересчитаны: ' + (last - 1) + ' строк', 'Готово', 5);
}

/** Проставляет ID строкам, добавленным вручную. Без ID заявка не обновится с формы. */
function fillMissingIds() {
  const sh = getSheet();
  const last = sh.getLastRow();
  if (last < 2) return;
  const rng = sh.getRange(2, C.ID, last - 1, 1);
  const vals = rng.getValues();
  let n = 0;
  for (let i = 0; i < vals.length; i++) {
    if (!String(vals[i][0]).trim()) { vals[i][0] = newId(); n++; }
  }
  rng.setValues(vals);
  SpreadsheetApp.getActive().toast('Добавлено ID: ' + n, 'Готово', 5);
}

/* ========================================================================== */
/*  СБОРКА ТАБЛИЦЫ                                                            */
/* ========================================================================== */

function setupWorkbook() {
  const ss = SpreadsheetApp.getActive();
  buildSheet(ss);
  buildHelper(ss);
  buildDashboard(ss);
  buildInfo(ss);

  // порядок листов
  ss.setActiveSheet(ss.getSheetByName(DASH));
  ss.moveActiveSheet(1);
  ss.setActiveSheet(ss.getSheetByName(SHEET));
  ss.moveActiveSheet(2);

  const def = ss.getSheetByName('Лист1') || ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0) ss.deleteSheet(def);

  SpreadsheetApp.getUi().alert('Готово. Листы «Дашборд», «Заявки», «Служебное» и «Инструкция» собраны.');
}

function getSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  return sh;
}

function buildSheet(ss) {
  const sh = ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);

  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setFontSize(10).setFontColor('#FFFFFF')
    .setVerticalAlignment('middle').setWrap(true);
  sh.getRange(1, 1, 1, BASE.length).setBackground(BLUE);
  sh.getRange(1, C.FLAG, 1, NOMS.length).setBackground('#0E1A8C');
  sh.setRowHeight(1, 40);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(5);

  const widths = [110, 140, 155, 185, 155, 135, 130, 240, 150, 95, 105, 165, 230];
  widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  for (let i = 0; i < NOMS.length; i++) sh.setColumnWidth(C.FLAG + i, 92);

  sh.hideColumns(C.ID);   // технический столбец, участнику не показывается

  // статусы списком
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUSES, true).setAllowInvalid(false).build();
  sh.getRange(2, C.STATUS, 3000, 1).setDataValidation(rule);

  // подсветка строк по статусу
  const rng = sh.getRange(2, 1, 3000, HEADERS.length);
  const rules = [
    ['$C2="Оплачено"',           '#DCF2E1'],
    ['$C2="Скриншот отправлен"', '#FFF4C9'],
    ['$C2="Отмена"',             '#F7DCDC'],
    ['$C2="Возврат"',            '#EAE5F8']
  ].map(function (r) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=' + r[0]).setBackground(r[1]).setRanges([rng]).build();
  });
  // дубли телефонов
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($F2<>"",COUNTIF($F$2:$F$3000,$F2)>1)')
    .setFontColor('#C62828').setBold(true)
    .setRanges([sh.getRange(2, C.PHONE, 3000, 1)]).build());
  sh.setConditionalFormatRules(rules);

  sh.getRange(2, C.FLAG, 3000, NOMS.length)
    .setHorizontalAlignment('center').setFontColor(BLUE).setFontWeight('bold');
  sh.getRange(2, C.SUM, 3000, 1).setNumberFormat('#,##0 ₽');
  sh.getRange(2, C.DATE, 3000, 1).setNumberFormat('dd.MM.yyyy HH:mm');

  const f = sh.getFilter();
  if (f) f.remove();
  sh.getRange(1, 1, Math.max(sh.getLastRow(), 2), HEADERS.length).createFilter();
}

function buildHelper(ss) {
  let h = ss.getSheetByName(HELP);
  if (!h) h = ss.insertSheet(HELP);
  h.clear();
  h.getRange('A1').setValue('Города — считается автоматически, руками не трогать');
  h.getRange('A1').setFontWeight('bold');
  h.getRange('A2').setFormula(
    '=IFERROR(QUERY(Заявки!G2:G,"select G, count(G) where G is not null and G <> \'\' ' +
    'group by G order by count(G) desc label G \'Город\', count(G) \'Заявок\'",0),{"Город","Заявок"})');
  h.setColumnWidth(1, 220);
  h.setColumnWidth(2, 100);
  h.hideSheet();
}

function buildDashboard(ss) {
  let d = ss.getSheetByName(DASH);
  if (!d) d = ss.insertSheet(DASH);
  d.clear();
  d.getCharts().forEach(function (c) { d.removeChart(c); });

  d.setHiddenGridlines(true);
  [300, 110, 90, 170, 40, 300, 110, 90, 170].forEach(function (w, i) { d.setColumnWidth(i + 1, w); });

  d.getRange('A1').setValue('BACK TO BACK 2026 — ДАШБОРД')
    .setFontSize(18).setFontWeight('bold').setFontColor(BLUE);
  d.getRange('A2').setValue('Все цифры считаются по листу «Заявки». Руками здесь ничего не вводится.')
    .setFontSize(10).setFontColor('#8D8D85');

  head(d, 'A3:D3', 'ЗАЯВКИ');
  rows(d, 4, [
    ['Всего заявок',         '=COUNTA(Заявки!D2:D)'],
    ['Оплачено',             '=COUNTIF(Заявки!C2:C,"Оплачено")'],
    ['Скриншот на проверке', '=COUNTIF(Заявки!C2:C,"Скриншот отправлен")'],
    ['Ждём скриншот',        '=COUNTIF(Заявки!C2:C,"Ждём скриншот")'],
    ['Отмены и возвраты',    '=COUNTIF(Заявки!C2:C,"Отмена")+COUNTIF(Заявки!C2:C,"Возврат")']
  ]);

  head(d, 'A10:D10', 'ДЕНЬГИ');
  rows(d, 11, [
    ['Собрано',      '=SUMIF(Заявки!C2:C,"Оплачено",Заявки!J2:J)',            '#,##0 ₽'],
    ['На проверке',  '=SUMIF(Заявки!C2:C,"Скриншот отправлен",Заявки!J2:J)',  '#,##0 ₽'],
    ['Ждём оплату',  '=SUMIF(Заявки!C2:C,"Ждём скриншот",Заявки!J2:J)',       '#,##0 ₽'],
    ['Средний чек',  '=IFERROR(SUMIF(Заявки!C2:C,"Оплачено",Заявки!J2:J)/COUNTIF(Заявки!C2:C,"Оплачено"),0)', '#,##0 ₽']
  ]);

  // ---- наполнение номинаций ----
  head(d, 'A17:D17', 'НАПОЛНЕНИЕ НОМИНАЦИЙ');
  d.getRange('A18:D18').setValues([['Номинация', 'Заявок', 'Доля', '']])
    .setFontWeight('bold').setFontColor('#55554E').setFontSize(10);
  for (let i = 0; i < NOMS.length; i++) {
    const r = 19 + i;
    const col = colLetter(C.FLAG + i);
    d.getRange(r, 1).setValue(NOMS[i]);
    d.getRange(r, 2).setFormula('=COUNTIF(Заявки!' + col + '2:' + col + ',"' + MARK + '")');
    d.getRange(r, 3).setFormula('=IFERROR(B' + r + '/$B$4,0)').setNumberFormat('0.0%');
    d.getRange(r, 4).setFormula(
      '=SPARKLINE(B' + r + ',{"charttype","bar";"max",MAX($B$19:$B$25);"color1","' + BLUE + '"})');
  }
  d.getRange(19, 1, NOMS.length, 4).setBorder(true, true, true, true, true, true, '#E4E2DB',
    SpreadsheetApp.BorderStyle.SOLID);
  d.getRange(19, 2, NOMS.length, 1).setFontWeight('bold').setFontColor(BLUE);

  // ---- мастер-классы ----
  head(d, 'A27:D27', 'МАСТЕР-КЛАССЫ');
  rows(d, 28, [
    ['Берут 1 мастер-класс',  '=COUNTIF(Заявки!I2:I,"1 мастер-класс")'],
    ['Берут 2 мастер-класса', '=COUNTIF(Заявки!I2:I,"2 мастер-класса")'],
    ['Берут 3 мастер-класса', '=COUNTIF(Заявки!I2:I,"3 мастер-класса")'],
    ['Всего мест на занятиях', '=COUNTIF(Заявки!I2:I,"1 мастер-класс")+COUNTIF(Заявки!I2:I,"2 мастер-класса")*2+COUNTIF(Заявки!I2:I,"3 мастер-класса")*3']
  ]);

  // ---- города ----
  head(d, 'F3:I3', 'ГОРОДА УЧАСТНИКОВ');
  d.getRange('F4:I4').setValues([['Город', 'Заявок', 'Доля', '']])
    .setFontWeight('bold').setFontColor('#55554E').setFontSize(10);
  for (let i = 0; i < 15; i++) {
    const r = 5 + i;
    d.getRange(r, 6).setFormula('=IFERROR(Служебное!A' + (r - 2) + ',"")');
    d.getRange(r, 7).setFormula('=IFERROR(Служебное!B' + (r - 2) + ',"")');
    d.getRange(r, 8).setFormula('=IFERROR(G' + r + '/SUM($G$5:$G$19),"")').setNumberFormat('0.0%');
    d.getRange(r, 9).setFormula(
      '=IF(G' + r + '="","",SPARKLINE(G' + r + ',{"charttype","bar";"max",MAX($G$5:$G$19);"color1","#4A5BF2"}))');
  }
  d.getRange('F21').setValue('Показаны 15 самых частых городов. Полный список — на скрытом листе «Служебное».')
    .setFontSize(9).setFontColor('#8D8D85');

  buildCharts(ss, d);
}

function buildCharts(ss, d) {
  d.getCharts().forEach(function (c) { d.removeChart(c); });

  const pie = d.newChart().asPieChart()
    .addRange(ss.getSheetByName(HELP).getRange('A2:B17'))
    .setNumHeaders(1)
    .setOption('title', 'Города участников, доля от всех заявок')
    .setOption('pieSliceText', 'percentage')
    .setOption('legend', { position: 'right' })
    .setOption('width', 520).setOption('height', 320)
    .setPosition(23, 6, 0, 0)
    .build();
  d.insertChart(pie);

  const bar = d.newChart().asColumnChart()
    .addRange(d.getRange('A18:B25'))
    .setNumHeaders(1)
    .setOption('title', 'Сколько заявок в каждой номинации')
    .setOption('legend', { position: 'none' })
    .setOption('colors', [BLUE])
    .setOption('width', 520).setOption('height', 300)
    .setPosition(34, 1, 0, 0)
    .build();
  d.insertChart(bar);
}

function head(sh, a1, text) {
  sh.getRange(a1).merge().setValue(text)
    .setBackground(BLUE).setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(11).setVerticalAlignment('middle');
  sh.setRowHeight(sh.getRange(a1).getRow(), 26);
}

function rows(sh, start, list) {
  list.forEach(function (item, i) {
    const r = start + i;
    sh.getRange(r, 1).setValue(item[0]).setFontColor('#141410');
    const c = sh.getRange(r, 2).setFormula(item[1])
      .setFontWeight('bold').setFontColor(BLUE).setHorizontalAlignment('right');
    if (item[2]) c.setNumberFormat(item[2]);
  });
  sh.getRange(start, 1, list.length, 2)
    .setBorder(true, true, true, true, true, true, '#E4E2DB', SpreadsheetApp.BorderStyle.SOLID);
}

function colLetter(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
  return s;
}

function buildInfo(ss) {
  let i = ss.getSheetByName(INFO);
  if (!i) i = ss.insertSheet(INFO);
  i.clear();
  i.setColumnWidth(1, 4);
  i.setColumnWidth(2, 110);
  i.setHiddenGridlines(true);

  const lines = [
    ['h', 'РУЧНЫЕ ПРАВКИ — ЧТО МОЖНО'],
    ['p', ''],
    ['p', 'Сортировать строки как угодно, включая фильтры и сортировку по столбцам.'],
    ['p', 'Вставлять и удалять строки в любом месте.'],
    ['p', 'Править любые ячейки: фамилию, телефон, сумму, номинации, комментарий.'],
    ['p', 'Заявки с формы всё равно допишутся в конец и попадут в нужные строки —'],
    ['p', 'скрипт ищет заявку по скрытому столбцу A «ID», а не по номеру строки.'],
    ['p', ''],
    ['h', 'РУЧНЫЕ ПРАВКИ — ЧЕГО НЕЛЬЗЯ'],
    ['p', ''],
    ['p', 'Переименовывать листы «Заявки», «Дашборд», «Служебное» — скрипт их не найдёт.'],
    ['p', 'Менять порядок столбцов на листе «Заявки» — форма пишет по позициям.'],
    ['p', 'Стирать значения в скрытом столбце A «ID» — по нему находится заявка.'],
    ['p', 'Править что-либо на листе «Дашборд»: там формулы, они перезапишутся.'],
    ['p', ''],
    ['h', 'ФИЛЬТР ПО НОМИНАЦИЯМ'],
    ['p', ''],
    ['p', 'Справа от данных семь столбцов, по одному на номинацию. Галочка означает,'],
    ['p', 'что участник в ней заявлен. Нажми на воронку в шапке нужного столбца и'],
    ['p', 'оставь только галочку — получишь список участников этой номинации.'],
    ['p', 'Если поправишь столбец «Номинации» руками, галочки пересчитаются сами.'],
    ['p', ''],
    ['h', 'МЕНЮ «BACK TO BACK»'],
    ['p', ''],
    ['p', 'Собрать таблицу — пересобирает оформление, дашборд и графики.'],
    ['p', 'Пересчитать номинации — перестраивает галочки по всем строкам сразу.'],
    ['p', 'Проставить ID — выдаёт ID строкам, которые ты добавил вручную.'],
    ['p', ''],
    ['h', 'СТАТУСЫ'],
    ['p', ''],
    ['p', 'Ждём скриншот — заявка создана, оплату не подтверждали. По таким прозванивать.'],
    ['p', 'Скриншот отправлен — участник подтвердил отправку чека. Требует проверки.'],
    ['p', 'Оплачено — деньги на счету. Строка зеленеет, сумма идёт в «Собрано».'],
    ['p', 'Отмена — участник отказался. Возврат — деньги вернули.']
  ];
  let r = 2;
  lines.forEach(function (l) {
    const c = i.getRange(r, 2).setValue(l[1]);
    if (l[0] === 'h') {
      c.setFontWeight('bold').setFontSize(12).setFontColor('#FFFFFF').setBackground(BLUE);
      i.setRowHeight(r, 24);
    } else {
      c.setFontSize(11).setFontColor('#141410');
    }
    r++;
  });
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('BACK TO BACK')
    .addItem('Собрать таблицу', 'setupWorkbook')
    .addItem('Пересчитать номинации', 'recalcFlags')
    .addItem('Проставить ID вручную добавленным', 'fillMissingIds')
    .addToUi();
}
