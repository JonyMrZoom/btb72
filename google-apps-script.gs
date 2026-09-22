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
const LISTS = 'Списки';
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

/* --------------------------------------------------------------------------
   Разделитель аргументов в формулах зависит от языка таблицы:
   в английской это запятая, в русской — точка с запятой.
   Определяем его один раз пробной формулой и подставляем куда нужно.
   В шаблонах формул разделитель пишется знаком |
   -------------------------------------------------------------------------- */
let SEP = null;

function sep() {
  if (SEP) return SEP;
  const ss = SpreadsheetApp.getActive();
  const tmp = ss.insertSheet('__probe__');
  try {
    tmp.getRange('A1').setFormula('=SUM(1,2)');
    SpreadsheetApp.flush();
    SEP = (tmp.getRange('A1').getValue() === 3) ? ',' : ';';
  } catch (e) {
    SEP = ';';
  }
  ss.deleteSheet(tmp);
  return SEP;
}

/** Подставляет верный разделитель в шаблон формулы. */
function F(tpl) {
  return tpl.split('|').join(sep());
}

/* ========================================================================== */
/*  ПРИЁМ ЗАЯВОК                                                              */
/* ========================================================================== */

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.secret !== SECRET) return out({ ok:false, error:'bad secret' });

    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET) || getSheet();

    // Заявка уже есть — меняем статус. Ищем по ID, а не по номеру строки,
    // поэтому сортировка и ручные вставки строк ничего не ломают.
    if (d.id) {
      const row = findRowById(sh, d.id);
      if (row) {
        const cell = sh.getRange(row, C.STATUS);
        const cur = String(cell.getValue());
        if (cur === 'Ждём скриншот' || cur === '') cell.setValue(d.status || cur);
        return out({ ok:true, id:d.id, updated:true });
      }
      if (d.updateOnly) return out({ ok:true, notFound:true });
    }

    if (d.updateOnly) return out({ ok:true, skipped:true });

    // appendRow атомарен, блокировка не нужна: она только добавляла задержку
    const id = d.id || newId();
    sh.appendRow(buildRow(d, id));

    // форматы даты и суммы уже стоят на всём столбце — на строку их не ставим
    return out({ ok:true, id:id });

  } catch (err) {
    return out({ ok:false, error:String(err) });
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

    // поправили номинации руками → пересчитываем галочки в столбцах-фильтрах
    if (col > C.NOMS || col + e.range.getNumColumns() - 1 < C.NOMS) return;
    const n = e.range.getNumRows();
    const vals = sh.getRange(row, C.NOMS, n, 1).getValues();
    sh.getRange(row, C.FLAG, n, NOMS.length)
      .setValues(vals.map(function (v) { return flagsFor(v[0]); }));
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

/**
 * Apps Script пишет формулы с запятыми, а в русской локали разделитель — точка с запятой.
 * Из-за этого все формулы с двумя аргументами падали в #ERROR!.
 * Переводим таблицу в англоязычную локаль: на отображение денег и дат это не влияет,
 * форматы у нас заданы явно.
 */
function ensureLocale(ss) {
  if (ss.getSpreadsheetLocale() !== 'en_US') {
    ss.setSpreadsheetLocale('en_US');
    SpreadsheetApp.flush();
  }
}

function setupWorkbook() {
  const ss = SpreadsheetApp.getActive();
  ensureLocale(ss);
  buildHelper(ss);
  buildSheet(ss);
  buildDashboard(ss);
  buildNomSheets(ss);
  buildInfo(ss);

  // порядок листов
  ss.setActiveSheet(ss.getSheetByName(DASH));
  ss.moveActiveSheet(1);
  ss.setActiveSheet(ss.getSheetByName(SHEET));
  ss.moveActiveSheet(2);
  NOMS.forEach(function (n, i) {
    const sh = ss.getSheetByName(n);
    if (sh) { ss.setActiveSheet(sh); ss.moveActiveSheet(3 + i); }
  });

  const def = ss.getSheetByName('Лист1') || ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0) ss.deleteSheet(def);

  SpreadsheetApp.getUi().alert('Готово. Собран дашборд, лист заявок и по вкладке на каждую номинацию.');
}


/**
 * Лист, пришедший из файла, может иметь ровно столько столбцов, сколько было
 * в исходнике. Новые столбцы вроде «Явки» тогда оказываются за краем сетки
 * и любая запись в них падает с ошибкой. Здесь лист при необходимости расширяется.
 */
/** Убирает столбец «Явка», если он остался от прошлой версии. */
function dropAttendanceColumn(sh) {
  if (!sh) return;
  const width = sh.getMaxColumns();
  for (let c = width; c > HEADERS.length; c--) {
    const head = String(sh.getRange(1, c).getValue()).trim();
    if (head === 'Явка' || head === '') {
      if (head === 'Явка') sh.deleteColumn(c);
    }
  }
  sh.getRange(1, HEADERS.length + 1, sh.getMaxRows(),
              Math.max(sh.getMaxColumns() - HEADERS.length, 1))
    .clearDataValidations();
}

function ensureColumns(sh, need) {
  const have = sh.getMaxColumns();
  if (have < need) sh.insertColumnsAfter(have, need - have);
  return sh;
}

function getSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  ensureColumns(sh, HEADERS.length);
  return sh;
}

function buildSheet(ss) {
  const sh = ensureColumns(ss.getSheetByName(SHEET) || ss.insertSheet(SHEET), HEADERS.length);

  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setFontSize(10).setFontColor('#FFFFFF')
    .setVerticalAlignment('middle').setWrap(true);
  sh.getRange(1, 1, 1, BASE.length).setBackground(BLUE);
  sh.getRange(1, C.FLAG, 1, NOMS.length).setBackground('#0E1A8C');
  sh.setRowHeight(1, 40);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(5);

  const widths = [110, 125, 150, 175, 140, 130, 115, 200, 140, 90, 105, 165, 200];
  widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  for (let i = 0; i < NOMS.length; i++) sh.setColumnWidth(C.FLAG + i, 84);

  // прячем то, что при работе почти не нужно, — иначе столбцы с номинациями
  // уезжают за край экрана и их никто не находит
  sh.showColumns(1, HEADERS.length);
  sh.hideColumns(C.ID);
  sh.hideColumns(C.NOMS);    // слепленный текст, для фильтра не годится
  sh.hideColumns(C.TARIFF);
  sh.hideColumns(C.AGREE);

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

  // фильтр на весь лист: иначе новые заявки окажутся вне его диапазона
  const f = sh.getFilter();
  if (f) f.remove();
  sh.getRange(1, 1, sh.getMaxRows(), HEADERS.length).createFilter();
}

function buildHelper(ss) {
  let h = ss.getSheetByName(HELP);
  if (!h) h = ss.insertSheet(HELP);
  h.clear();
  h.getRange('A1').setValue('Считается автоматически, руками не трогать').setFontWeight('bold');

  // сводка по городам
  h.getRange('A2').setFormula(F(
    '=IFERROR(QUERY(Заявки!G2:G|"select G, count(G) where G is not null and G <> \'\' ' +
    'group by G order by count(G) desc label G \'Город\', count(G) \'Заявок\'"|0)|"")'));

  // справочник: номинация → буква столбца с галочками
  const map = NOMS.map(function (n, i) { return [n, colLetter(C.FLAG + i)]; });
  h.getRange('D1:E1').setValues([['Номинация', 'Столбец']]).setFontWeight('bold');
  h.getRange(2, 4, map.length, 2).setValues(map);

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
  [280, 120, 90, 190, 40, 260, 120, 90, 190].forEach(function (w, i) { d.setColumnWidth(i + 1, w); });

  d.getRange('A1').setValue('BACK TO BACK 2026')
    .setFontSize(18).setFontWeight('bold').setFontColor(BLUE);
  d.setRowHeight(1, 28);

  const strip = [
    ['Всего заявок', '=COUNTA(Заявки!D2:D)'],
    ['Оплачено',     F('=COUNTIF(Заявки!C2:C|"Оплачено")')],
    ['Собрано, ₽',   F('=SUMIF(Заявки!C2:C|"Оплачено"|Заявки!J2:J)')]
  ];
  strip.forEach(function (it, i) {
    const col = 1 + i * 3;
    d.getRange(2, col).setValue(it[0]).setFontSize(10).setFontColor('#8D8D85');
    const c = d.getRange(3, col).setFormula(it[1])
      .setFontSize(22).setFontWeight('bold').setFontColor(BLUE);
    if (i === 2) c.setNumberFormat('#,##0 ₽');
  });
  d.setRowHeight(3, 32);

  // ---- участники по номинациям ----
  head(d, 'A5:D5', 'УЧАСТНИКОВ В НОМИНАЦИИ');
  d.getRange('A6:D6').setValues([['Номинация', 'Участников', 'Доля', '']])
    .setFontWeight('bold').setFontColor('#55554E').setFontSize(10);
  for (let i = 0; i < NOMS.length; i++) {
    const r = 7 + i;
    const col = colLetter(C.FLAG + i);
    d.getRange(r, 1).setValue(NOMS[i]).setFontSize(12);
    d.getRange(r, 2).setFormula(F('=COUNTIF(Заявки!' + col + '2:' + col + '|"' + MARK + '")'))
      .setFontWeight('bold').setFontColor(BLUE).setFontSize(13).setHorizontalAlignment('right');
    d.getRange(r, 3).setFormula(F('=IFERROR(B' + r + '/$B$3|0)'))
      .setNumberFormat('0%').setHorizontalAlignment('right');
    d.getRange(r, 4).setFormula(F(
      '=IF(B' + r + '=0|""|REPT("▮"|MAX(1|ROUND(B' + r + '/MAX($B$7:$B$13)*16|0))))'))
      .setFontColor('#4A5BF2');
    d.setRowHeight(r, 24);
  }
  d.getRange(7, 1, NOMS.length, 4)
    .setBorder(true, true, true, true, true, true, '#E4E2DB', SpreadsheetApp.BorderStyle.SOLID);

  // ---- города ----
  head(d, 'F5:I5', 'ГОРОДА УЧАСТНИКОВ');
  d.getRange('F6:I6').setValues([['Город', 'Участников', 'Доля', '']])
    .setFontWeight('bold').setFontColor('#55554E').setFontSize(10);
  for (let i = 0; i < 15; i++) {
    const r = 7 + i;
    d.getRange(r, 6).setFormula(F('=IFERROR(Служебное!A' + (r - 4) + '|"")')).setFontSize(12);
    d.getRange(r, 7).setFormula(F('=IFERROR(Служебное!B' + (r - 4) + '|"")'))
      .setFontWeight('bold').setFontColor(BLUE).setFontSize(13).setHorizontalAlignment('right');
    d.getRange(r, 8).setFormula(F('=IFERROR(G' + r + '/SUM($G$7:$G$21)|"")'))
      .setNumberFormat('0%').setHorizontalAlignment('right');
    d.getRange(r, 9).setFormula(F(
      '=IF(N(G' + r + ')=0|""|REPT("▮"|MAX(1|ROUND(G' + r + '/MAX($G$7:$G$21)*16|0))))'))
      .setFontColor('#4A5BF2');
  }
  d.getRange('F23').setValue('Показаны 15 самых частых городов.')
    .setFontSize(9).setFontColor('#8D8D85');

  buildCharts(ss, d);
}

/** Два графика под таблицами: столбики по номинациям и круг по городам. */
function buildCharts(ss, d) {
  d.getCharts().forEach(function (c) { d.removeChart(c); });

  const bar = d.newChart().asColumnChart()
    .addRange(d.getRange('A6:B13'))
    .setNumHeaders(1)
    .setOption('title', 'Участников в каждой номинации')
    .setOption('legend', { position: 'none' })
    .setOption('colors', [BLUE])
    .setOption('width', 500).setOption('height', 300)
    .setPosition(24, 1, 0, 0)
    .build();
  d.insertChart(bar);

  const pie = d.newChart().asPieChart()
    .addRange(ss.getSheetByName(HELP).getRange('A2:B17'))
    .setNumHeaders(1)
    .setOption('title', 'Города участников')
    .setOption('pieSliceText', 'percentage')
    .setOption('legend', { position: 'right' })
    .setOption('width', 500).setOption('height', 300)
    .setPosition(24, 6, 0, 0)
    .build();
  d.insertChart(pie);
}

/**
 * По вкладке на каждую номинацию. Список собирается сам и обновляется,
 * как только приходит новая заявка или ты правишь данные руками.
 */
function buildNomSheets(ss) {
  ss = ss || SpreadsheetApp.getActive();
  const main = ss.getSheetByName(SHEET);
  if (!main) { SpreadsheetApp.getUi().alert('Нет листа «Заявки».'); return; }
  ensureColumns(main, HEADERS.length);

  const old = ss.getSheetByName(LISTS);
  if (old) ss.deleteSheet(old);
  dropAttendanceColumn(main);

  NOMS.forEach(function (name, i) {
    const col = colLetter(C.FLAG + i);
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    sh.clear();
    // clear() не снимает выпадающие списки и примечания — убираем отдельно,
    // иначе на месте фамилий остаются стрелки и красные пометки об ошибке
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns())
      .clearDataValidations().clearNote();
    sh.showColumns(1, sh.getMaxColumns());
    sh.setFrozenColumns(0);
    sh.setConditionalFormatRules([]);
    sh.setHiddenGridlines(true);
    sh.setTabColor(BLUE);
    [240, 190, 150, 150, 170, 110].forEach(function (w, k) { sh.setColumnWidth(k + 1, w); });

    sh.getRange('A1').setValue(name.toUpperCase())
      .setFontSize(20).setFontWeight('bold').setFontColor(BLUE);
    sh.setRowHeight(1, 30);

    sh.getRange('A2').setFormula(F(
      '="Участников: "&COUNTIF(Заявки!' + col + '2:' + col + '|"' + MARK + '")'))
      .setFontSize(13).setFontWeight('bold');
    sh.getRange('C2').setFormula(F(
      '="Оплатили: "&COUNTIFS(Заявки!' + col + '2:' + col + '|"' + MARK + '"' +
      '|Заявки!C2:C|"Оплачено")'))
      .setFontSize(13).setFontColor('#159A45').setFontWeight('bold');
    sh.getRange('A3').setValue('Список собирается сам. Руками здесь ничего не вводится.')
      .setFontSize(9).setFontColor('#8D8D85');

    sh.getRange('A5').setFormula(F(
      '=IFERROR(QUERY(Заявки!$A$2:$T|' +
      '"select D, E, F, G, C, J where ' + col + ' = \'' + MARK + '\' order by D ' +
      'label D \'Фамилия Имя\', E \'Никнейм\', F \'Телефон\', G \'Город\', ' +
      'C \'Статус оплаты\', J \'Сумма\'"|0)|"В этой номинации пока никого нет")'));

    sh.getRange('A5:F5').setBackground(BLUE).setFontColor('#FFFFFF')
      .setFontWeight('bold').setFontSize(10);
    sh.setRowHeight(5, 26);
    sh.getRange('F6:F500').setNumberFormat('#,##0 ₽');
    sh.setFrozenRows(5);

    const rng = sh.getRange('A6:F500');
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$E6="Оплачено"').setBackground('#DCF2E1').setRanges([rng]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$E6="Скриншот отправлен"').setBackground('#FFF4C9').setRanges([rng]).build()
    ]);
  });
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

/** Пересобирает только витрины, данные не трогает. */
function rebuildViews() {
  const ss = SpreadsheetApp.getActive();
  ensureLocale(ss);
  buildHelper(ss);
  buildDashboard(ss);
  buildNomSheets(ss);
  ss.toast('Дашборд и вкладки номинаций пересобраны', 'Готово', 5);
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('BACK TO BACK')
    .addItem('Собрать таблицу', 'setupWorkbook')
    .addItem('Пересобрать дашборд и списки номинаций', 'rebuildViews')
    .addItem('Найти дубли по телефону', 'checkDuplicates')
    .addItem('Пересчитать номинации', 'recalcFlags')
    .addItem('Проставить ID вручную добавленным', 'fillMissingIds')
    .addToUi();
}
