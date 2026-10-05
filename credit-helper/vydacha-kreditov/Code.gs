/**
 * ================================================================
 *   КРЕДИТНЫЙ ПОМОЩНИК — Согдийский филиал (ЗАО ДСБ)
 *   Таблица «Выдача кредитов». v13 — тот же код, что и у Истаравшана,
 *   настроенный под колонки этого файла (см. CFG ниже).
 *
 *   Отличия от файла Истаравшана:
 *   • ФИО в E, Остаток в O, Эксперт в S, Телефон в R; колонки «Статус»
 *     нет — кредит закрыт, когда Остаток ≤ 0.
 *   • Есть «Поручитель» (U) и «Залог» (V) — блок «Залоговое покрытие»
 *     считает настоящую разбивку: золото, недвижимость, авто, только
 *     поручитель, без обеспечения, и сумму оценки залога.
 *   • Дата оплаты пишется в колонку W.
 *   • Продукт «Дастраст» в расчёты не входит.
 *
 *   Остальное — как в v13: галочка «Оплачено» с датой, «Кому звонить»,
 *   полное меню, единая логика статусов (см. описание в Code.gs Истаравшана).
 *
 *   ⚠ Этот файл собирается из credit-helper/Code.gs скриптом
 *   build_vydacha.py — правьте Code.gs, а не этот файл.
 *
 *  УСТАНОВКА: Расширения → Apps Script, СКОПИРОВАТЬ старый код себе
 *  (на всякий случай), стереть, вставить этот целиком, сохранить,
 *  обновить вкладку таблицы, затем
 *  Кредитный помощник → 📅 Снять устаревшие галочки «Оплачено»,
 *  Кредитный помощник → 🟢 Пересчитать подсветку,
 *  Кредитный помощник → 🔄 Обновить дашборд.
 * ================================================================
 */

const CFG = {
  SHEET_DATA_CANDIDATES: ['Кредитный лист', 'кредитный лист', 'Лист1'],
  SHEET_DASHBOARD:       'Дашборд',
  SHEET_HISTORY:         'История портфеля',
  SHEET_DASTRAS:         'Дастраст',
  BRANCH_NAME:           'Согдийский филиал',
  BANK_NAME:             'ЗАО «Душанбе Сити Банк»',

  EXPERTS: ['Дадабаев Н', 'Курбонов Х', 'Ахмедов Б', 'Ахроров М', 'Карими О'],
  PRODUCT_GOLD: ['Ломбард', 'Ломбарди'],
  EXPERT_BONUS_RATE: 0.003,

  // За сколько дней до платежа подсвечивать оранжевым («скоро»)
  DUE_SOON_DAYS: 7,

  // Насколько раньше срока оплата ещё засчитывается за этот платёж
  EARLY_PAY_DAYS: 10,

  // Сколько строк показывать в блоке «Кому звонить»
  CALL_LIST_LIMIT: 30,

  // Строки с продуктом «Дастраст» в основном листе пропускать?
  SKIP_DASTRAS_IN_MAIN: true,
  DASTRAS_PRODUCTS: ['Дастраст', 'Дастрас'],

  // Известные продукты (для «Чистки данных»)
  PRODUCTS: ['Истеъмоли', 'Ипотека', 'Ломбард', 'Ломбарди', 'Тиҷорати', 'Автокредит', 'Дастраст', 'Дастрас'],

  // Служебные листы — никогда не принимаются за лист с кредитами
  SERVICE_SHEETS: ['Закрытые', 'Остатки активных'],

  // Колонки «Статус» в этом файле нет (STATUS: 0) — закрыт = остаток ≤ 0
  STATUS_CLOSED_VALUE: 'закрыл',

  // A №, B Оплачено, C Выдача кредита, D осталось платежей, E ФИО,
  // F Сумма TJS, G сумма USD, H Продукт, I Ставка, J Срок, K по графику,
  // L Оплата, M Процент, N Тело, O Остаток, P ID клиента, Q Номер договора,
  // R Номер тел, S Эксперт, T Цель, U Поручитель, V Залог, W Дата оплаты (новая)
  COL: {
    N: 1, PAID: 2, ISSUE_DATE: 3, FIO: 5, AMOUNT: 6, PRODUCT: 8, RATE: 9, TERM: 10,
    SCHEDULE: 11, PAYMENT: 12, INTEREST: 13, PRINCIPAL: 14, BALANCE: 15,
    CLIENT_ID: 16, EXPERT: 19, PHONE: 18,
    STATUS: 0, CLOSE_DATE: 0, UNDERWRITER: 0, PAID_DATE: 23,
    GUARANTOR: 21, COLLATERAL: 22
  },
  LAST_COL: 23,
  PAID_DATE_HEADER: 'Дата оплаты',

  DASTRAS_COL: { N: 1, FIO: 2, AMOUNT: 3, PRODUCT: 4, RATE: 5, TERM: 6, EXPERT: 13, PHONE: 14 },
  DASTRAS_LAST_COL: 14,

  COLOR_PAID:     '#E6F4EA',
  COLOR_OVERDUE:  '#FCE8E6',
  COLOR_DUE_SOON: '#FEF7E0',
  COLOR_NEW:      '#F1F3F4'
};


// ================================================================
// МЕНЮ
// ================================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Кредитный помощник')
    .addItem('🔄 Обновить дашборд',                    'buildDashboard')
    .addItem('🟢 Пересчитать подсветку',               'recolorAllRows')
    .addItem('📅 Снять устаревшие галочки «Оплачено»', 'clearStaleTicks')
    .addSeparator()
    .addItem('💵 Платежи на неделю',                   'collectPaymentsWeek')
    .addItem('💵 Платежи на месяц',                    'collectPaymentsMonth')
    .addSeparator()
    .addItem('📸 Сохранить снимок портфеля',           'savePortfolioSnapshot')
    .addItem('📄 PDF дашборда на Диск',                'exportReportPdf')
    .addSeparator()
    .addSubMenu(ui.createMenu('🧮 Расчёты')
      .addItem('Средневзвешенная ставка', 'calcWeightedRate')
      .addItem('Средний чек',             'calcAvgTicket')
      .addItem('Прогноз годового дохода', 'calcYearForecast')
      .addItem('Залоговое покрытие',      'calcCollateralCoverage'))
    .addSubMenu(ui.createMenu('⚙ Настройка')
      .addItem('🔍 Диагностика подсветки',                         'diagHighlight')
      .addItem('🧽 Снять чередование и усл. форматирование',        'clearBandingAndRules')
      .addSeparator()
      .addItem('⏰ Включить авто-режим',                            'enableAutoMode')
      .addItem('⛔ Выключить авто-режим',                           'disableAutoMode'))
    .addToUi();

  ui.createMenu('Чистка данных')
    .addItem('🧹 Проверить и почистить',              'cleanupData')
    .addItem('📋 Только показать отчёт (без правок)', 'cleanupReport')
    .addSeparator()
    .addItem('📦 Перенести Дастрас на отдельный лист', 'moveDastrasToSeparateSheet')
    .addToUi();
}


// ================================================================
// Поиск листа с данными
// ================================================================
function getDataSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (const name of CFG.SHEET_DATA_CANDIDATES) {
    const sh = ss.getSheets().find(s => s.getName().trim().toLowerCase() === name.toLowerCase());
    if (sh) return sh;
  }
  const service = new Set([CFG.SHEET_DASHBOARD, CFG.SHEET_HISTORY, CFG.SHEET_DASTRAS]
    .concat(CFG.SERVICE_SHEETS).map(n => n.trim().toLowerCase()));
  const first = ss.getSheets().find(s => !service.has(s.getName().trim().toLowerCase()));
  if (!first) throw new Error('Не найден лист с данными кредитов');
  return first;
}

/** Колонка T «Дата оплаты» — подписываем заголовок, если пусто */
function ensurePaidDateColumn_(sh) {
  const cell = sh.getRange(1, CFG.COL.PAID_DATE);
  if (cell.getValue() === '') {
    cell.setValue(CFG.PAID_DATE_HEADER).setFontWeight('bold');
  }
}

function moveDastrasToSeparateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getDataSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert('Нет данных'); return; }

  let dastSh = ss.getSheetByName(CFG.SHEET_DASTRAS);
  if (!dastSh) dastSh = ss.insertSheet(CFG.SHEET_DASTRAS);

  const values = sh.getRange(2, 1, lastRow - 1, CFG.LAST_COL).getValues();
  const C = CFG.COL;
  const D = CFG.DASTRAS_COL;

  // Кто уже есть на листе «Дастрас» — чтобы не задвоить при повторном запуске
  const existing = new Set();
  const dLast = dastSh.getLastRow();
  if (dLast >= 1) {
    dastSh.getRange(1, 1, dLast, CFG.DASTRAS_LAST_COL).getValues().forEach(r => {
      const f = (r[D.FIO - 1] || '').toString().trim().toLowerCase();
      if (f) existing.add(f);
    });
  }

  const dastRows = [];
  const rowsToDelete = [];
  let skipped = 0;

  for (let i = 0; i < values.length; i++) {
    const r = values[i];
    if (!isDastras_(r[C.PRODUCT - 1])) continue;

    const fio = (r[C.FIO - 1] || '').toString().trim();
    rowsToDelete.push(i + 2);

    if (existing.has(fio.toLowerCase())) { skipped++; continue; }

    const target = new Array(CFG.DASTRAS_LAST_COL).fill('');
    target[D.N - 1]       = r[C.N - 1];
    target[D.FIO - 1]     = fio;
    target[D.AMOUNT - 1]  = dastrasLimit_(r);
    target[D.PRODUCT - 1] = r[C.PRODUCT - 1];
    target[D.RATE - 1]    = r[C.RATE - 1];
    target[D.TERM - 1]    = r[C.TERM - 1];
    target[D.EXPERT - 1]  = r[C.EXPERT - 1];
    target[D.PHONE - 1]   = r[C.PHONE - 1];
    dastRows.push(target);
  }

  if (!rowsToDelete.length) {
    SpreadsheetApp.getUi().alert('Продукт «Дастрас» на основном листе не найден — переносить нечего.');
    return;
  }

  if (dastRows.length) {
    dastSh.getRange(dastSh.getLastRow() + 1, 1, dastRows.length, CFG.DASTRAS_LAST_COL).setValues(dastRows);
  }
  rowsToDelete.sort((a, b) => b - a).forEach(r => sh.deleteRow(r));

  SpreadsheetApp.getUi().alert('Готово',
    `Удалено из основного листа: ${rowsToDelete.length} строк «Дастрас».\n` +
    `Добавлено на лист «${CFG.SHEET_DASTRAS}»: ${dastRows.length}.\n` +
    (skipped ? `Пропущено как уже существующие: ${skipped}.\n` : '') +
    'Обнови дашборд — Дастрас в общих суммах больше не мешается.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/** Строки «Дастрас» в основном листе (они пропускаются в расчётах) */
function countDastrasInMain_() {
  if (!CFG.SKIP_DASTRAS_IN_MAIN) return { count: 0, limitSum: 0 };
  const sh = getDataSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { count: 0, limitSum: 0 };
  const C = CFG.COL;
  let count = 0, limitSum = 0;
  sh.getRange(2, 1, lastRow - 1, CFG.LAST_COL).getValues().forEach(r => {
    if (isEmptyRow_(r) || !isDastras_(r[C.PRODUCT - 1])) return;
    count++;
    limitSum += dastrasLimit_(r);
  });
  return { count, limitSum };
}

function readDastrasTotals_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEET_DASTRAS);
  if (!sh) return { count: 0, limitSum: 0 };

  const lastRow = sh.getLastRow();
  if (lastRow < 1) return { count: 0, limitSum: 0 };

  const D = CFG.DASTRAS_COL;
  const a1 = sh.getRange(1, 1).getValue();
  const hasHeader = a1 !== '' && isNaN(parseFloat(a1));
  const startRow = hasHeader ? 2 : 1;
  if (startRow > lastRow) return { count: 0, limitSum: 0 };

  const values = sh.getRange(startRow, 1, lastRow - startRow + 1, CFG.DASTRAS_LAST_COL).getValues();
  let count = 0, limitSum = 0;
  values.forEach(r => {
    const fio = (r[D.FIO - 1] || '').toString().trim();
    if (!fio) return;
    count++;
    limitSum += toNum_(r[D.AMOUNT - 1]);
  });
  return { count, limitSum };
}


// ================================================================
// Закрыт ли кредит — единая проверка (Остаток ≤ 0 ИЛИ Статус = «Закрыл»)
// ================================================================
/**
 * Лимит «Дастрас» из строки основного листа. В «Сумме кредита» у таких
 * строк стоит заглушка (0 или 0,1), а лимит записан в колонку «По графику» —
 * берём большее из двух.
 */
function dastrasLimit_(r) {
  const C = CFG.COL;
  const amount = toNum_(r[C.AMOUNT - 1]);
  const sched = r[C.SCHEDULE - 1];
  return Math.max(amount, (sched instanceof Date) ? 0 : toNum_(sched));
}

function isDastras_(product) {
  return CFG.DASTRAS_PRODUCTS.includes((product || '').toString().trim());
}

function isClosed_(balance, statusText) {
  if (balance <= 0) return true;
  const s = (statusText || '').toString().trim().toLowerCase();
  return s === CFG.STATUS_CLOSED_VALUE;
}


// ================================================================
// Чтение данных
// ================================================================
function isEmptyRow_(r) {
  const C = CFG.COL;
  return !r[C.FIO - 1] && !r[C.AMOUNT - 1];
}

/** Строка листа (массив значений) → объект кредита */
function rowToLoan_(r, rowNum) {
  // Колонка с номером 0 = её нет в файле: r[-1] даёт undefined → пусто
  const C = CFG.COL;
  const balance  = toNum_(r[C.BALANCE - 1]);
  const status   = (r[C.STATUS - 1] || '').toString().trim();
  const closed   = isClosed_(balance, status);
  const interest = toNum_(r[C.INTEREST - 1]);
  const principal = toNum_(r[C.PRINCIPAL - 1]);

  return {
    row:          rowNum,
    n:            r[C.N - 1],
    paid:         r[C.PAID - 1] === true,
    paidDate:     toDate_(r[C.PAID_DATE - 1]),
    issueDate:    toDate_(r[C.ISSUE_DATE - 1]),
    fio:          (r[C.FIO - 1] || '').toString().trim(),
    amount:       toNum_(r[C.AMOUNT - 1]),
    product:      (r[C.PRODUCT - 1] || '').toString().trim(),
    rate:         parseRate_(r[C.RATE - 1]),
    term:         parseTerm_(r[C.TERM - 1]),
    maturity:     toDate_(r[C.SCHEDULE - 1]),
    payment:      toNum_(r[C.PAYMENT - 1]),
    interest:     interest,
    principal:    principal,
    balance:      balance,
    clientId:     (r[C.CLIENT_ID - 1] || '').toString(),
    expert:       (r[C.EXPERT - 1] || '').toString().trim(),
    phone:        (r[C.PHONE - 1] || '').toString(),
    status:       status,
    closeDate:    toDate_(r[C.CLOSE_DATE - 1]),
    underwriter:  (r[C.UNDERWRITER - 1] || '').toString().trim(),
    guarantor:    (r[C.GUARANTOR - 1] || '').toString().trim(),
    collateral:   (r[C.COLLATERAL - 1] || '').toString().trim(),
    closed:       closed,
    // Для сумм портфеля: закрытые кредиты в ОД и доход не входят
    od:           closed ? 0 : balance,
    activeInterest:  closed ? 0 : interest,
    activePrincipal: closed ? 0 : principal
  };
}

function readLoans_() {
  const sh = getDataSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const values = sh.getRange(2, 1, lastRow - 1, CFG.LAST_COL).getValues();
  const loans = [];
  for (let i = 0; i < values.length; i++) {
    if (isEmptyRow_(values[i])) continue;
    const loan = rowToLoan_(values[i], i + 2);
    if (CFG.SKIP_DASTRAS_IN_MAIN && isDastras_(loan.product)) continue;
    loans.push(loan);
  }
  return loans;
}


// ================================================================
// СТАТУС ПЛАТЕЖА — единая логика для подсветки, дашборда, сбора, диагностики
// ================================================================
/**
 * Возвращает { state, due, tick, expired }:
 *   state: 'paid' | 'overdue' | 'soon' | 'new' | 'later' | 'closed' | 'nodate' | 'skip'
 *   due:   дата платежа, о котором речь (или null)
 *   tick:  ''        — галочки нет или она с датой и в силе
 *          'legacy'  — галочка стоит без даты оплаты (поставлена до v13)
 *          'stale'   — галочка стоит, но дата оплаты относится к прошлому платежу
 *   expired: true — срок кредита («По графику») вышел, а остаток есть
 *
 * День платежа = день выдачи (выдан 24-го → платит 24-го каждый месяц).
 * Новая выдача: выдан в текущем месяце, первый платёж — в следующем.
 */
function paymentState_(x, today) {
  const res = { state: 'later', due: null, tick: '', expired: false };
  if (isDastras_(x.product)) { res.state = 'skip'; return res; }
  if (x.closed) { res.state = 'closed'; return res; }
  if (x.paid && !(x.paidDate instanceof Date)) res.tick = 'legacy';
  if (!(x.issueDate instanceof Date)) { res.state = x.paid ? 'paid' : 'nodate'; return res; }

  const isNew = isNewIssue_(x.issueDate, today);
  const due = isNew ? firstDueDate_(x.issueDate) : nextDueDate_(x.issueDate, today);
  const maturity = x.maturity instanceof Date ? startOfDay_(x.maturity) : null;
  res.due = due;

  if (coversDue_(x, due, false)) {
    res.state = 'paid';
    // Этот месяц оплачен, но следующий платёж (например 1-го числа) уже близко
    if (!isNew) {
      const next = dueInMonth_(x.issueDate, today.getFullYear(), today.getMonth() + 1);
      if ((!maturity || next <= maturity) &&
          daysBetween_(today, next) <= CFG.DUE_SOON_DAYS &&
          !coversDue_(x, next, true)) {
        res.state = 'soon';
        res.due = next;
      }
    }
    return res;
  }

  if (x.paid) res.tick = 'stale';
  if (isNew) { res.state = 'new'; return res; }

  // Срок кредита уже вышел, а остаток есть — это просрочка независимо от дня
  if (maturity && maturity < today) {
    res.state = 'overdue';
    res.expired = true;
    return res;
  }

  const diff = daysBetween_(today, due);
  res.state = diff < 0 ? 'overdue' : diff <= CFG.DUE_SOON_DAYS ? 'soon' : 'later';
  return res;
}

/**
 * Покрывает ли галочка «Оплачено» платёж на дату d.
 * Оплата с датой засчитывается, если сделана не раньше чем за EARLY_PAY_DAYS
 * дней до платежа. Галочка без даты (старая) — только за текущий платёж.
 */
function coversDue_(x, d, isNext) {
  if (!x.paid) return false;
  if (!(x.paidDate instanceof Date)) return !isNext;
  return startOfDay_(x.paidDate) >= addDays_(d, -CFG.EARLY_PAY_DAYS);
}

function stateColor_(state) {
  switch (state) {
    case 'paid':    return CFG.COLOR_PAID;
    case 'overdue': return CFG.COLOR_OVERDUE;
    case 'soon':    return CFG.COLOR_DUE_SOON;
    case 'new':     return CFG.COLOR_NEW;
    default:        return null;
  }
}

function stateMark_(state) {
  switch (state) {
    case 'paid':    return '🟢';
    case 'overdue': return '🔴';
    case 'soon':    return '🟠';
    case 'new':     return '⬜';
    default:        return '⚪';
  }
}

/** Цвет для строки листа (массив значений) или null */
function rowColor_(r, today) {
  if (isEmptyRow_(r)) return null;
  return stateColor_(paymentState_(rowToLoan_(r, 0), today).state);
}

/**
 * Новая выдача — кредит выдан в ТЕКУЩЕМ календарном месяце и первый платёж
 * ещё не наступил (первый платёж — день выдачи в следующем месяце после
 * месяца выдачи; до этой даты по кредиту рано считать просрочку).
 */
function isNewIssue_(issueDate, today) {
  if (!(issueDate instanceof Date)) return false;
  const sameMonth = issueDate.getFullYear() === today.getFullYear() &&
                    issueDate.getMonth() === today.getMonth();
  if (!sameMonth) return false;
  return today < firstDueDate_(issueDate);
}

/**
 * День платежа (= день выдачи) в месяце (y, m). m может выходить за 0..11.
 * Если в месяце меньше дней (выдан 31-го, а это февраль) — последний день месяца.
 */
function dueInMonth_(issueDate, y, m) {
  const first = new Date(y, m, 1);
  const yy = first.getFullYear(), mm = first.getMonth();
  const daysInMonth = new Date(yy, mm + 1, 0).getDate();
  return new Date(yy, mm, Math.min(issueDate.getDate(), daysInMonth));
}

/** Дата первого платежа = день выдачи в СЛЕДУЮЩЕМ месяце после выдачи. */
function firstDueDate_(issueDate) {
  return dueInMonth_(issueDate, issueDate.getFullYear(), issueDate.getMonth() + 1);
}

/** Дата платежа в ТЕКУЩЕМ месяце. */
function nextDueDate_(issueDate, today) {
  if (!(issueDate instanceof Date)) return null;
  return dueInMonth_(issueDate, today.getFullYear(), today.getMonth());
}

function startOfDay_(d) {
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
}

function addDays_(d, n) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
}

function daysBetween_(from, to) {
  return Math.round((startOfDay_(to) - startOfDay_(from)) / 86400000);
}


// ================================================================
// ПОДСВЕТКА СТРОК
// ================================================================
/**
 *   🟢 зелёный   — оплачен ближайший платёж (галочка + дата оплаты)
 *   🔴 красный   — день платежа прошёл, не оплачено; или срок кредита вышел
 *   🟠 оранжевый — платёж в ближайшие DUE_SOON_DAYS дней
 *   ⬜ серый     — новая выдача этого месяца, первый платёж ещё впереди
 *   ⚪ без цвета — платёж не скоро, кредит закрыт, Дастрас, нет даты выдачи
 */
function recolorAllRows() {
  const sh = getDataSheet_();
  const res = recolorSheetFast_(sh);
  const ui = SpreadsheetApp.getUi();

  let tail;
  if (res.legacy + res.stale > 0) {
    tail = '📅 Галочек «Оплачено» без даты (старые): ' + res.legacy + '\n' +
           '📅 Галочек за прошлый платёж (уже не в силе): ' + res.stale + '\n' +
           'Убрать их — «📅 Снять устаревшие галочки «Оплачено»».';
  } else if (res.overdue + res.soon === 0) {
    tail = 'Ни одной красной или оранжевой строки. Если такого быть не должно — ' +
           'запусти ⚙ Настройка → «Диагностика подсветки».';
  } else {
    tail = 'Дальше строки красятся сами при клике на галочку, дата оплаты пишется в колонку T.';
  }

  ui.alert('Подсветка пересчитана',
    'Обработано строк: ' + res.total + '\n\n' +
    '🟢 оплачено: ' + res.paid + '\n' +
    '🔴 просрочено: ' + res.overdue + '\n' +
    '🟠 платёж в ближайшие ' + CFG.DUE_SOON_DAYS + ' дней: ' + res.soon + '\n' +
    '⬜ новая выдача (платёж ещё не наступил): ' + res.newIssue + '\n' +
    '⚪ без подсветки: ' + res.none + '\n\n' +
    tail + '\n\n' +
    'Если цвета не видно — чередование цветов и условное форматирование ' +
    'рисуются поверх. Сними их: ⚙ Настройка → «Снять чередование и усл. форматирование».',
    ui.ButtonSet.OK);
}

/** Красит весь лист за один проход — читает и пишет пакетно, не по ячейке */
function recolorSheetFast_(sh) {
  const res = { total: 0, paid: 0, overdue: 0, soon: 0, newIssue: 0, none: 0, legacy: 0, stale: 0 };
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return res;
  ensurePaidDateColumn_(sh);

  const values = sh.getRange(2, 1, lastRow - 1, CFG.LAST_COL).getValues();
  const today = startOfDay_(new Date());
  const backgrounds = values.map(r => {
    let color = null;
    if (!isEmptyRow_(r)) {
      const st = paymentState_(rowToLoan_(r, 0), today);
      color = stateColor_(st.state);
      if (st.tick === 'legacy') res.legacy++;
      if (st.tick === 'stale')  res.stale++;
    }
    if (color === CFG.COLOR_PAID)          res.paid++;
    else if (color === CFG.COLOR_OVERDUE)  res.overdue++;
    else if (color === CFG.COLOR_DUE_SOON) res.soon++;
    else if (color === CFG.COLOR_NEW)      res.newIssue++;
    else res.none++;
    res.total++;
    return new Array(CFG.LAST_COL).fill(color);
  });

  sh.getRange(2, 1, backgrounds.length, CFG.LAST_COL).setBackgrounds(backgrounds);
  return res;
}

/** Красит несколько строк подряд (для onEdit) */
function recolorRows_(sh, startRow, numRows) {
  const range = sh.getRange(startRow, 1, numRows, CFG.LAST_COL);
  const today = startOfDay_(new Date());
  const backgrounds = range.getValues().map(r => new Array(CFG.LAST_COL).fill(rowColor_(r, today)));
  range.setBackgrounds(backgrounds);
}

/**
 * Пишет дату оплаты в колонку T при клике на галочку.
 * Клик по одной галочке: поставили — сегодня, сняли — дата стирается.
 * Вставка/правка диапазона, захватившего колонку B: только заполняем
 * пустые даты у поставленных галочек и стираем у снятых — чтобы
 * случайная вставка не «оплатила» старые галочки сегодняшним днём.
 */
function stampPaidDates_(sh, startRow, numRows, singleColumn) {
  const C = CFG.COL;
  const paid = sh.getRange(startRow, C.PAID, numRows, 1).getValues();
  const dateRange = sh.getRange(startRow, C.PAID_DATE, numRows, 1);
  const dates = dateRange.getValues();
  const today = startOfDay_(new Date());

  const out = paid.map((p, i) => {
    const old = dates[i][0];
    if (p[0] !== true) return [''];
    if (singleColumn || !toDate_(old)) return [today];
    return [old];
  });

  ensurePaidDateColumn_(sh);
  dateRange.setValues(out).setNumberFormat('dd.mm.yyyy');
}

function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getSheetId() !== getDataSheet_().getSheetId()) return;

    const startRow = Math.max(2, e.range.getRow());
    const endRow = e.range.getLastRow();
    if (endRow < startRow) return;
    const numRows = endRow - startRow + 1;

    const C = CFG.COL;
    const firstCol = e.range.getColumn(), lastCol = e.range.getLastColumn();
    if (firstCol <= C.PAID && lastCol >= C.PAID) {
      stampPaidDates_(sh, startRow, numRows, firstCol === lastCol);
    }
    recolorRows_(sh, startRow, numRows);
  } catch (err) {
    // Тихо игнорируем — не мешаем обычному редактированию
  }
}

/** Снимает галочки «Оплачено», которые уже не относятся к ближайшему платежу */
function clearStaleTicks() {
  const ui = SpreadsheetApp.getUi();
  const sh = getDataSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) { ui.alert('Нет данных'); return; }

  const C = CFG.COL;
  const values = sh.getRange(2, 1, lastRow - 1, CFG.LAST_COL).getValues();
  const today = startOfDay_(new Date());
  const stale = [], legacy = [];

  values.forEach((r, i) => {
    if (isEmptyRow_(r)) return;
    const st = paymentState_(rowToLoan_(r, i + 2), today);
    if (st.tick === 'stale')  stale.push(i);
    if (st.tick === 'legacy') legacy.push(i);
  });

  let clearLegacy = false;
  if (legacy.length) {
    const resp = ui.alert('Галочки без даты оплаты',
      `У ${legacy.length} кредитов галочка «Оплачено» стоит без даты оплаты ` +
      '(поставлена до этой версии) — непонятно, за какой месяц.\n\n' +
      'Снять их тоже?\n' +
      '«Да» — снять. Кто уже заплатил в этом месяце — отметьте заново, дата запишется сама.\n' +
      '«Нет» — оставить как есть (будут считаться оплатой за текущий месяц).',
      ui.ButtonSet.YES_NO);
    clearLegacy = resp === ui.Button.YES;
  }

  const toClear = stale.concat(clearLegacy ? legacy : []);
  if (!toClear.length) {
    ui.alert('Снимать нечего', 'Все галочки «Оплачено» относятся к ближайшему платежу.', ui.ButtonSet.OK);
    return;
  }

  // Пишем всю колонку B одним вызовом; дата в T остаётся как история
  const paidCol = values.map(r => [r[C.PAID - 1]]);
  toClear.forEach(i => { paidCol[i][0] = false; });
  sh.getRange(2, C.PAID, paidCol.length, 1).setValues(paidCol);

  const res = recolorSheetFast_(sh);
  ui.alert('Готово',
    `Снято галочек: ${toClear.length}` +
    ` (устаревших: ${stale.length}` + (clearLegacy ? `, без даты: ${legacy.length}` : '') + ').\n\n' +
    `Сейчас: 🔴 ${res.overdue} • 🟠 ${res.soon} • 🟢 ${res.paid} • ⬜ ${res.newIssue}`,
    ui.ButtonSet.OK);
}

/** Снимает чередование цветов и условное форматирование — они бьют подсветку */
function clearBandingAndRules() {
  const sh = getDataSheet_();
  const bandings = sh.getBandings();
  const rulesCount = sh.getConditionalFormatRules().length;

  bandings.forEach(b => b.remove());
  sh.clearConditionalFormatRules();

  SpreadsheetApp.getUi().alert('Готово',
    'Снято чередований цветов: ' + bandings.length + '\n' +
    'Снято правил условного форматирования: ' + rulesCount + '\n\n' +
    (bandings.length + rulesCount === 0
      ? 'Ничего не было — значит подсветку перекрывало не это.'
      : 'Теперь запусти «Пересчитать подсветку».'),
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/** Показывает по каждому кредиту, что посчиталось — где искать причину */
function diagHighlight() {
  const sh = getDataSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert('Нет данных'); return; }

  const values = sh.getRange(2, 1, lastRow - 1, CFG.LAST_COL).getValues();
  const today = startOfDay_(new Date());
  const f = fmtDate_;
  const tickText = { legacy: '  ✔ без даты', stale: '  ✔ устарела' };

  const lines = [];
  lines.push('Сегодня: ' + f(today));
  lines.push('Лист: «' + sh.getName() + '»');
  lines.push('Чередований: ' + sh.getBandings().length +
             ' • Правил усл. форматирования: ' + sh.getConditionalFormatRules().length);
  lines.push('');

  let shown = 0;
  values.forEach((r, i) => {
    if (isEmptyRow_(r) || shown >= 30) return;
    const l = rowToLoan_(r, i + 2);
    if (isDastras_(l.product)) return;

    const st = paymentState_(l, today);
    const diff = st.due ? daysBetween_(today, st.due) : '—';
    lines.push(stateMark_(st.state) + ' стр.' + l.row + '  ' + l.fio.substring(0, 22) +
      '  выдача ' + f(l.issueDate) +
      '  платёж ' + f(st.due) + ' (' + diff + ' дн.)' +
      (st.expired ? '  срок вышел ' + f(l.maturity) : '') +
      '  оплата ' + f(l.paidDate) + (tickText[st.tick] || '') +
      '  ост. ' + Math.round(l.balance) +
      '  статус: ' + (l.status || '—'));
    shown++;
  });

  SpreadsheetApp.getUi().alert('Диагностика подсветки', lines.join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK);
  Logger.log(lines.join('\n'));
}


// ================================================================
// ДАШБОРД
// ================================================================
function buildDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CFG.SHEET_DASHBOARD);
  if (!sh) sh = ss.insertSheet(CFG.SHEET_DASHBOARD);

  // Объединения переживают clear() — снимаем, иначе setValue падает
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.clear();
  sh.clearFormats();

  const loans = readLoans_();
  if (!loans.length) {
    sh.getRange('B2').setValue('Нет данных на листе с кредитами');
    return;
  }

  sh.setColumnWidth(1, 20);
  sh.setColumnWidth(2, 210);
  [3, 4, 5, 6, 7].forEach(c => sh.setColumnWidth(c, 115));
  sh.setColumnWidth(8, 150);
  sh.setColumnWidth(9, 110);

  let row = 1;
  row = drawHeader_(sh, row);
  row = drawSummaryCards_(sh, row, loans);
  row = drawCallList_(sh, row, loans);
  row = drawExperts_(sh, row, loans);
  row = drawProducts_(sh, row, loans);
  row = drawCollateral_(sh, row, loans);
  row = drawTop5_(sh, row, loans);
  row = drawFinancials_(sh, row, loans);
  row = drawKeyObservations_(sh, row, loans);

  sh.setHiddenGridlines(true);
  ss.setActiveSheet(sh);
  ss.toast('Дашборд обновлён', CFG.BRANCH_NAME, 3);
}

function drawHeader_(sh, row) {
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm');
  sh.getRange(row, 2).setValue('📋 ДАШБОРД КРЕДИТНОГО ПОРТФЕЛЯ')
    .setFontSize(16).setFontWeight('bold').setFontColor('#1F3864');
  row++;
  sh.getRange(row, 2).setValue(`${CFG.BRANCH_NAME} • ${CFG.BANK_NAME} • Обновлено: ${now}`)
    .setFontSize(9).setFontColor('#666666');
  return row + 2;
}

function drawSummaryCards_(sh, row, loans) {
  const totalIssued  = loans.reduce((s, l) => s + l.amount, 0);
  const totalBalance = loans.reduce((s, l) => s + l.od, 0);
  const active = loans.filter(l => !l.closed).length;
  const count = loans.length;
  const avg = count ? totalIssued / count : 0;
  const wRate = weightedRateByAmount_(loans);

  const cards = [
    { col: 2, label: 'ПОРТФЕЛЬ (ОД)',  value: fmtNum_(totalBalance),    caption: `сомони • ${active} активных` },
    { col: 4, label: 'ВСЕГО ВЫДАНО',   value: fmtNum_(totalIssued),     caption: `${count} кредитов` },
    { col: 6, label: 'СРЕДНИЙ ЧЕК',    value: fmtNum_(Math.round(avg)), caption: 'сомони на кредит' },
    { col: 8, label: 'СР.ВЗВ. СТАВКА', value: (wRate * 100).toFixed(2).replace('.', ',') + '%', caption: 'годовых по объёму выдач' }
  ];

  const cardTop = row, cardBottom = row + 3;
  cards.forEach(c => {
    sh.getRange(cardTop, c.col, cardBottom - cardTop + 1, 2)
      .setBackground('#F8F9FA')
      .setBorder(true, true, true, true, false, false, '#E3E6EA', SpreadsheetApp.BorderStyle.SOLID);
  });
  cards.forEach(c => sh.getRange(cardTop, c.col).setValue(c.label).setFontSize(9).setFontWeight('bold').setFontColor('#666666'));
  cards.forEach(c => sh.getRange(cardTop + 1, c.col).setValue(c.value).setFontSize(18).setFontWeight('bold').setFontColor('#1F3864'));
  cards.forEach(c => sh.getRange(cardTop + 2, c.col).setValue(c.caption).setFontSize(8).setFontStyle('italic').setFontColor('#999999'));

  row = cardBottom + 2;

  const dast = readDastrasTotals_();
  const dastMain = countDastrasInMain_();
  if (dast.count > 0 || dastMain.count > 0) {
    const parts = [];
    if (dast.count) parts.push(`лист «${CFG.SHEET_DASTRAS}»: ${dast.count} лимитов на ${fmtNum_(dast.limitSum)} сом.`);
    if (dastMain.count) parts.push(`в основном листе: ${dastMain.count} строк на ${fmtNum_(dastMain.limitSum)} сом.`);
    sh.getRange(row, 2).setValue(`📌 ${CFG.DASTRAS_PRODUCTS[0]} в цифры выше не входит — ` + parts.join(' • '))
      .setFontSize(9).setFontStyle('italic').setFontColor('#666666');
    row += 2;
  }
  return row;
}

/** «Кому звонить»: просрочка и платежи в ближайшие DUE_SOON_DAYS дней */
function drawCallList_(sh, row, loans) {
  sh.getRange(row, 2).setValue(`📞 КОМУ ЗВОНИТЬ — просрочка и платежи в ближайшие ${CFG.DUE_SOON_DAYS} дней`)
    .setFontSize(12).setFontWeight('bold').setFontColor('#1F3864');
  row += 2;

  const today = startOfDay_(new Date());
  const list = loans
    .map(l => ({ l: l, st: paymentState_(l, today) }))
    .filter(x => x.st.state === 'overdue' || x.st.state === 'soon')
    .sort((a, b) => {
      if (a.st.state !== b.st.state) return a.st.state === 'overdue' ? -1 : 1;
      return (a.st.due || 0) - (b.st.due || 0);
    });

  if (!list.length) {
    sh.getRange(row, 2).setValue(`✅ Просрочки нет, платежей в ближайшие ${CFG.DUE_SOON_DAYS} дней нет.`)
      .setFontSize(10).setFontColor('#2E7D32');
    return row + 3;
  }

  ['Статус', 'ФИО', 'Телефон', 'Эксперт', 'Дата платежа', 'Дней до платежа', 'Сумма платежа'].forEach((h, i) =>
    sh.getRange(row, 2 + i).setValue(h).setFontWeight('bold').setFontSize(9)
      .setFontColor('#444444').setBorder(false, false, true, false, false, false, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID));
  row++;

  const shown = list.slice(0, CFG.CALL_LIST_LIMIT);
  const data = shown.map(x => [
    x.st.state === 'overdue' ? (x.st.expired ? '🔴 срок вышел' : '🔴 просрочка') : '🟠 скоро',
    x.l.fio,
    fmtPhone_(x.l.phone),
    x.l.expert || '(не указан)',
    x.st.expired ? x.l.maturity : x.st.due,
    x.st.expired ? daysBetween_(today, x.l.maturity) : daysBetween_(today, x.st.due),
    x.st.expired ? x.l.balance : (x.l.payment || '—')
  ]);
  const range = sh.getRange(row, 2, data.length, 7);
  range.setValues(data).setFontSize(10)
    .setBackgrounds(shown.map(x => new Array(7).fill(stateColor_(x.st.state))));
  sh.getRange(row, 6, data.length, 1).setNumberFormat('dd.mm.yyyy');
  sh.getRange(row, 7, data.length, 1).setNumberFormat('0');
  sh.getRange(row, 8, data.length, 1).setNumberFormat('#,##0');
  row += data.length;

  const overdue = list.filter(x => x.st.state === 'overdue');
  const soon = list.filter(x => x.st.state === 'soon');
  sh.getRange(row, 2).setValue(
    `Итого: 🔴 ${overdue.length} на ${fmtNum_(overdue.reduce((s, x) => s + x.l.payment, 0))} сом. • ` +
    `🟠 ${soon.length} на ${fmtNum_(soon.reduce((s, x) => s + x.l.payment, 0))} сом.` +
    (list.length > shown.length ? `  (показано ${shown.length} из ${list.length})` : '') +
    '  • «срок вышел» — сумма = остаток ОД')
    .setFontSize(9).setFontStyle('italic').setFontColor('#666666');
  return row + 3;
}

function drawExperts_(sh, row, loans) {
  sh.getRange(row, 2).setValue('👤 РАСПРЕДЕЛЕНИЕ ПО ЭКСПЕРТАМ').setFontSize(12).setFontWeight('bold').setFontColor('#1F3864');
  row += 2;
  const list = groupLoans_(loans, l => l.expert || '(не указан)');
  return drawGroupTable_(sh, row, 'Эксперт', list, sumAmount_(loans)) + 2;
}

function drawProducts_(sh, row, loans) {
  sh.getRange(row, 2).setValue('🛍 РАСПРЕДЕЛЕНИЕ ПО ПРОДУКТАМ').setFontSize(12).setFontWeight('bold').setFontColor('#1F3864');
  row += 2;
  const list = groupLoans_(loans, l => l.product || '(не указан)');
  return drawGroupTable_(sh, row, 'Продукт', list, sumAmount_(loans)) + 2;
}

/** Группировка для таблиц экспертов/продуктов, по убыванию суммы выдач */
function groupLoans_(loans, keyFn) {
  const groups = {};
  loans.forEach(l => {
    const k = keyFn(l);
    if (!groups[k]) groups[k] = { count: 0, amount: 0, balance: 0, income: 0 };
    groups[k].count++; groups[k].amount += l.amount;
    groups[k].balance += l.od; groups[k].income += l.activeInterest;
  });
  return Object.entries(groups).sort((a, b) => b[1].amount - a[1].amount);
}

function sumAmount_(loans) {
  return loans.reduce((s, l) => s + l.amount, 0);
}

function drawGroupTable_(sh, row, firstColLabel, list, totalIssued) {
  const headers = [firstColLabel, 'Кол-во', 'Сумма выдач', 'Доля', 'Остаток ОД', 'Доход/мес', 'Визуализация'];
  headers.forEach((h, i) => {
    sh.getRange(row, 2 + i).setValue(h).setFontWeight('bold').setFontSize(9)
      .setFontColor('#444444').setBorder(false, false, true, false, false, false, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
  });
  row++;

  let totCount = 0, totAmount = 0, totBalance = 0, totIncome = 0;

  list.forEach(([name, g], idx) => {
    const share = totalIssued ? g.amount / totalIssued : 0;
    const zebra = idx % 2 === 1 ? '#FAFBFC' : null;

    const cells = [
      sh.getRange(row, 2).setValue(name).setFontSize(10),
      sh.getRange(row, 3).setValue(g.count).setNumberFormat('0').setFontSize(10),
      sh.getRange(row, 4).setValue(g.amount).setNumberFormat('#,##0').setFontSize(10),
      sh.getRange(row, 5).setValue(share).setNumberFormat('0.00%').setFontSize(10),
      sh.getRange(row, 6).setValue(g.balance).setNumberFormat('#,##0').setFontSize(10),
      sh.getRange(row, 7).setValue(g.income).setNumberFormat('#,##0').setFontSize(10)
    ];
    if (zebra) cells.forEach(c => c.setBackground(zebra));

    const barCell = sh.getRange(row, 8);
    if (zebra) barCell.setBackground(zebra);
    setBarCell_(barCell, share * 100);

    totCount += g.count; totAmount += g.amount; totBalance += g.balance; totIncome += g.income;
    row++;
  });

  const totals = [
    ['ИТОГО', null], [totCount, '0'], [totAmount, '#,##0'],
    [totalIssued ? totAmount / totalIssued : 0, '0.00%'],
    [totBalance, '#,##0'], [totIncome, '#,##0']
  ];
  totals.forEach((t, i) => {
    const cell = sh.getRange(row, 2 + i).setValue(t[0]).setFontWeight('bold').setFontSize(10)
      .setBorder(true, false, false, false, false, false, '#999999', SpreadsheetApp.BorderStyle.SOLID);
    if (t[1]) cell.setNumberFormat(t[1]);
  });
  sh.getRange(row, 8).setBorder(true, false, false, false, false, false, '#999999', SpreadsheetApp.BorderStyle.SOLID);
  return row + 1;
}

/** Двухцветный текстовый бар (без формул — не зависит от локали) */
function setBarCell_(cell, pct) {
  const totalBlocks = 10;
  const filled = Math.max(0, Math.min(totalBlocks, Math.round((pct / 100) * totalBlocks)));
  const text = '█'.repeat(totalBlocks);
  cell.setFontSize(10);

  const filledStyle = SpreadsheetApp.newTextStyle().setForegroundColor('#4472C4').build();
  const emptyStyle  = SpreadsheetApp.newTextStyle().setForegroundColor('#E3E6EA').build();
  const builder = SpreadsheetApp.newRichTextValue().setText(text);
  if (filled > 0) builder.setTextStyle(0, filled, filledStyle);
  if (filled < totalBlocks) builder.setTextStyle(filled, totalBlocks, emptyStyle);
  cell.setRichTextValue(builder.build());
}

/**
 * Тип обеспечения по колонкам «Залог» и «Поручитель».
 * Если колонок нет в файле — упрощённо по продукту («Ломбард» = золото).
 */
const COLLATERAL_TYPES = ['Золото', 'Недвижимость', 'Автотранспорт', 'Прочий залог',
                          'Только поручительство', 'Без обеспечения'];

function hasCollateralColumns_() {
  return CFG.COL.COLLATERAL > 0 || CFG.COL.GUARANTOR > 0;
}

function collateralType_(l) {
  if (!hasCollateralColumns_()) {
    return CFG.PRODUCT_GOLD.includes(l.product) ? 'Золото' : 'Без данных о залоге';
  }
  const z = l.collateral.toLowerCase();
  if (z) {
    if (/тилло|золот|тило/.test(z))                                   return 'Золото';
    if (/мошин|машин|авто|транспорт|наклиёт|toyota|lexus|hyundai|kia|mercedes|bmw|chevrolet|nissan|honda|opel/.test(z))
      return 'Автотранспорт';
    if (/манкул|амвол|хона|ҳавли|хавли|недвиж|квартир|бино|замин/.test(z)) return 'Недвижимость';
    return 'Прочий залог';
  }
  if (l.guarantor) return 'Только поручительство';
  if (CFG.PRODUCT_GOLD.includes(l.product)) return 'Золото';
  return 'Без обеспечения';
}

/**
 * Оценка залога из текста — число перед словом «сомони»:
 * «Тилло 95гр оценка 55 100 сомони» → 55100, «залог хона 402 235 сомони» → 402235.
 * Нет суммы — 0.
 */
function collateralValue_(text) {
  const t = (text || '').toString();
  const m = t.match(/(\d{1,3}(?:[ \u00A0]\d{3})+|\d+)\s*сом/i) ||
            t.match(/оценк\S*\s*(\d{1,3}(?:[ \u00A0]\d{3})+|\d+)/i);
  return m ? toNum_(m[1].replace(/[ \u00A0]/g, '')) : 0;
}

function drawCollateral_(sh, row, loans) {
  sh.getRange(row, 2).setValue('🏛 ЗАЛОГОВОЕ ПОКРЫТИЕ ПОРТФЕЛЯ').setFontSize(12).setFontWeight('bold').setFontColor('#1F3864');
  row++;
  if (!hasCollateralColumns_()) {
    sh.getRange(row, 2).setValue(
      '⚠ В файле нет колонок «Поручитель» и «Залог» — показана упрощённая разбивка по продукту. ' +
      'Полная классификация станет доступна, если добавить эти колонки.')
      .setFontSize(8).setFontStyle('italic').setFontColor('#C00000');
  } else {
    sh.getRange(row, 2).setValue(
      'По колонкам «Залог» и «Поручитель». Оценка — сумма в сомони из колонки «Залог»; покрытие = оценка / ОД.')
      .setFontSize(8).setFontStyle('italic').setFontColor('#999999');
  }
  row += 2;

  const active = loans.filter(l => !l.closed);
  const groups = {};
  active.forEach(l => {
    const t = collateralType_(l);
    if (!groups[t]) groups[t] = { count: 0, od: 0, value: 0 };
    groups[t].count++;
    groups[t].od += l.od;
    groups[t].value += collateralValue_(l.collateral);
  });
  const order = hasCollateralColumns_() ? COLLATERAL_TYPES : ['Золото', 'Без данных о залоге'];
  const totalOd = active.reduce((s, l) => s + l.od, 0);
  const showValue = hasCollateralColumns_();

  const headers = ['Тип обеспечения', 'Кол-во', 'Остаток ОД', 'Доля в ОД'].concat(showValue ? ['Оценка залога', 'Покрытие'] : []);
  headers.forEach((h, i) =>
    sh.getRange(row, 2 + i).setValue(h).setFontWeight('bold').setFontSize(9)
      .setFontColor('#444444').setBorder(false, false, true, false, false, false, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID));
  row++;

  order.filter(t => groups[t]).forEach(t => {
    const g = groups[t];
    sh.getRange(row, 2).setValue(t).setFontSize(10);
    sh.getRange(row, 3).setValue(g.count).setNumberFormat('0').setFontSize(10);
    sh.getRange(row, 4).setValue(g.od).setNumberFormat('#,##0').setFontSize(10);
    sh.getRange(row, 5).setValue(totalOd ? g.od / totalOd : 0).setNumberFormat('0.00%').setFontSize(10);
    if (showValue) {
      sh.getRange(row, 6).setValue(g.value || '—').setNumberFormat('#,##0').setFontSize(10);
      sh.getRange(row, 7).setValue(g.value && g.od ? g.value / g.od : '—').setNumberFormat('0%').setFontSize(10);
    }
    row++;
  });
  return row + 2;
}

function drawTop5_(sh, row, loans) {
  sh.getRange(row, 2).setValue('🏆 ТОП-5 КРУПНЕЙШИХ КРЕДИТОВ (по сумме выдачи)').setFontSize(12).setFontWeight('bold').setFontColor('#1F3864');
  row += 2;

  ['№', 'ФИО', 'Сумма', 'Продукт', 'Эксперт', 'Ставка', 'Остаток ОД', 'Доход/мес'].forEach((h, i) =>
    sh.getRange(row, 2 + i).setValue(h).setFontWeight('bold').setFontSize(9)
      .setFontColor('#444444').setBorder(false, false, true, false, false, false, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID));
  row++;

  loans.slice().sort((a, b) => b.amount - a.amount).slice(0, 5).forEach((l, idx) => {
    const medal = idx === 0 ? '🥇' : idx === 1 ? '🥈' : idx === 2 ? '🥉' : (idx + 1).toString();
    const zebra = idx % 2 === 1 ? '#FAFBFC' : null;
    const cells = [
      sh.getRange(row, 2).setValue(medal).setFontSize(10),
      sh.getRange(row, 3).setValue(l.fio).setFontSize(10),
      sh.getRange(row, 4).setValue(l.amount).setNumberFormat('#,##0').setFontSize(10),
      sh.getRange(row, 5).setValue(l.product).setFontSize(10),
      sh.getRange(row, 6).setValue(l.expert).setFontSize(10),
      sh.getRange(row, 7).setValue(l.rate).setNumberFormat('0.00%').setFontSize(10),
      sh.getRange(row, 8).setValue(l.od).setNumberFormat('#,##0').setFontSize(10),
      sh.getRange(row, 9).setValue(l.activeInterest).setNumberFormat('#,##0').setFontSize(10)
    ];
    if (zebra) cells.forEach(c => c.setBackground(zebra));
    row++;
  });
  return row + 2;
}


// ================================================================
// ФИНАНСОВЫЕ ПОКАЗАТЕЛИ
// ================================================================
function drawFinancials_(sh, row, loans) {
  sh.getRange(row, 2).setValue('💰 ФИНАНСОВЫЕ ПОКАЗАТЕЛИ (по данным на текущий месяц)')
    .setFontSize(12).setFontWeight('bold').setFontColor('#1F3864');
  row += 2;

  const active = loans.filter(l => !l.closed);
  const monthlyIncome = active.reduce((s, l) => s + l.interest, 0);
  const monthlyPrincipal = active.reduce((s, l) => s + l.principal, 0);
  const yearForecast = monthlyIncome * 12;
  const totalBalance = active.reduce((s, l) => s + l.balance, 0);
  const effYield = totalBalance ? yearForecast / totalBalance : 0;
  const rateByBalance = weightedRateByBalance_(loans);
  const withIncome = active.filter(l => l.interest > 0).length;

  const cards = [
    { col: 2, label: 'МЕС. % ДОХОД',        value: fmtNum_(monthlyIncome),    caption: `по ${withIncome} из ${active.length} активных` },
    { col: 4, label: 'МЕС. ПОГАШЕНИЕ ТЕЛА', value: fmtNum_(monthlyPrincipal), caption: 'сомони возвращается за месяц' },
    { col: 6, label: 'ГОД. ПРОГНОЗ %',      value: fmtNum_(yearForecast),     caption: 'при текущем портфеле' },
    { col: 8, label: 'ЭФФ. ДОХОДНОСТЬ',     value: (effYield * 100).toFixed(2).replace('.', ',') + '%', caption: 'фактическая по активному ОД' }
  ];

  cards.forEach(c => sh.getRange(row, c.col).setValue(c.label).setFontSize(9).setFontWeight('bold').setFontColor('#666666'));
  cards.forEach(c => sh.getRange(row + 1, c.col).setValue(c.value).setFontSize(16).setFontWeight('bold').setFontColor('#2E7D32'));
  cards.forEach(c => sh.getRange(row + 2, c.col).setValue(c.caption).setFontSize(8).setFontStyle('italic').setFontColor('#999999'));
  row += 4;

  sh.getRange(row, 2).setValue(
    `Ср.взв. ставка по активному ОД: ${(rateByBalance * 100).toFixed(2).replace('.', ',')}%  •  ` +
    `Мес. аннуитет всего: ${fmtNum_(monthlyIncome + monthlyPrincipal)} сом ` +
    `(${fmtNum_(monthlyIncome)} % + ${fmtNum_(monthlyPrincipal)} тело)`)
    .setFontSize(9).setFontStyle('italic').setFontColor('#666666');
  row += 2;

  if (withIncome < active.length * 0.5) {
    sh.getRange(row, 2).setValue(
      '⚠ Колонка «Процент» заполнена лишь у ' + withIncome + ' из ' + active.length +
      ' активных кредитов — доход и прогноз занижены во столько же раз.')
      .setFontSize(9).setFontColor('#C00000');
    row += 2;
  }
  return row + 1;
}


// ================================================================
// КЛЮЧЕВЫЕ НАБЛЮДЕНИЯ
// ================================================================
function drawKeyObservations_(sh, row, loans) {
  sh.getRange(row, 2).setValue('⚡ КЛЮЧЕВЫЕ НАБЛЮДЕНИЯ').setFontSize(12).setFontWeight('bold').setFontColor('#1F3864');
  row += 2;

  const totalBal = loans.reduce((s, l) => s + l.od, 0);
  if (hasCollateralColumns_()) {
    const unsecured = loans.filter(l => !l.closed && collateralType_(l) === 'Без обеспечения');
    const unsecuredBal = unsecured.reduce((s, l) => s + l.od, 0);
    const share = totalBal ? unsecuredBal / totalBal : 0;
    row = writeObservation_(sh, row, share >= 0.2 ? '⚠' : '🔒',
      `Без залога и без поручителя: ${unsecured.length} кредитов, ${(share * 100).toFixed(1)}% ОД (${fmtNum_(unsecuredBal)} сом.)`,
      unsecured.length
        ? 'Стоит проверить: ' + unsecured.sort((a, b) => b.od - a.od).slice(0, 6).map(l => l.fio).join(', ') +
          (unsecured.length > 6 ? '…' : '') + '.'
        : 'Все активные кредиты имеют залог или поручителя.');
  } else {
    const goldBal = loans.filter(l => collateralType_(l) === 'Золото').reduce((s, l) => s + l.od, 0);
    const goldShare = totalBal ? goldBal / totalBal : 0;
    row = writeObservation_(sh, row, '🔒',
      `Обеспечение (упрощённо): «Ломбард» — ${(goldShare * 100).toFixed(1)}% ОД, остальное — без данных о залоге`,
      'Полная классификация недоступна — в файле нет колонок «Поручитель» и «Залог».');
  }

  const totalIssuedBase = sumAmount_(loans);
  const topExpertEntry = groupLoans_(loans, l => l.expert || '(не указан)')[0];
  if (topExpertEntry) {
    const [expName, g] = topExpertEntry;
    const share = totalIssuedBase ? g.amount / totalIssuedBase : 0;
    row = writeObservation_(sh, row, share >= 0.5 ? '⚠' : '👥',
      `Концентрация: эксперт ${expName} — ${(share * 100).toFixed(1)}% портфеля`,
      share >= 0.5 ? 'Высокая концентрация — стоит диверсифицировать.'
        : share >= 0.35 ? 'Умеренная концентрация — есть куда диверсифицировать.'
        : 'Портфель хорошо диверсифицирован между экспертами.');
  }

  const byIncome = groupLoans_(loans, l => l.expert || '(не указан)').sort((a, b) => b[1].income - a[1].income);
  const totalIncomeBase = loans.reduce((s, l) => s + l.activeInterest, 0);
  if (byIncome.length && totalIncomeBase > 0) {
    const [expName, g] = byIncome[0];
    row = writeObservation_(sh, row, '📊',
      `По доходу лидирует: ${expName} — ${(g.income / totalIncomeBase * 100).toFixed(1)}% месячного дохода (${fmtNum_(g.income)} сом)`,
      'Учитываются только активные кредиты с заполненной колонкой «Процент».');
  }

  // Статусы платежей — та же логика, что и подсветка
  const today = startOfDay_(new Date());
  const states = loans.map(l => ({ l: l, st: paymentState_(l, today) }));

  const overdue = states.filter(x => x.st.state === 'overdue');
  const expired = overdue.filter(x => x.st.expired);
  const overdueSum = overdue.reduce((s, x) => s + x.l.payment, 0);
  row = writeObservation_(sh, row, overdue.length ? '🔴' : '✅',
    `Просрочено (день платежа прошёл, оплата не отмечена): ${overdue.length} кредитов на ${fmtNum_(overdueSum)} сом.`,
    overdue.length
      ? 'Либо реальная просрочка, либо не проставлены галочки «Оплачено».' +
        (expired.length ? ` Из них у ${expired.length} срок кредита уже вышел, а остаток есть.` : '')
      : 'Все наступившие платежи отмечены.');

  const fresh = states.filter(x => x.st.state === 'new');
  if (fresh.length) {
    row = writeObservation_(sh, row, '⬜',
      `Новые выдачи этого месяца: ${fresh.length} на ${fmtNum_(fresh.reduce((s, x) => s + x.l.amount, 0))} сом.`,
      'Первый платёж — в следующем месяце, просрочку по ним пока не считаем.');
  }

  const legacy = states.filter(x => x.st.tick === 'legacy').length;
  const stale = states.filter(x => x.st.tick === 'stale').length;
  if (legacy + stale) {
    row = writeObservation_(sh, row, '📅',
      `Галочки «Оплачено» требуют внимания: без даты — ${legacy}, за прошлый платёж — ${stale}`,
      'Кредитный помощник → «Снять устаревшие галочки «Оплачено»».');
  }

  // Расхождение между колонкой «Статус» и остатком (для ручной сверки)
  const mismatches = loans.filter(l => statusMismatch_(l.balance, l.status));
  if (mismatches.length) {
    row = writeObservation_(sh, row, '⚠',
      `Статус не совпадает с остатком у ${mismatches.length} кредитов`,
      'Стоит сверить вручную: ' + mismatches.slice(0, 8).map(l => l.fio).join(', ') +
      (mismatches.length > 8 ? '…' : '') + '.');
  }

  return row + 1;
}

/** Статус «Закрыл» при остатке > 0, или остаток 0 при статусе не «Закрыл» */
function statusMismatch_(balance, status) {
  if (!CFG.COL.STATUS) return false;
  const saysClosed = (status || '').toString().trim().toLowerCase() === CFG.STATUS_CLOSED_VALUE;
  return saysClosed !== (balance <= 0);
}

function writeObservation_(sh, row, icon, title, caption) {
  sh.getRange(row, 2).setValue(icon).setFontSize(11);
  sh.getRange(row, 3).setValue(title).setFontWeight('bold').setFontSize(10).setFontColor('#333333');
  row++;
  sh.getRange(row, 3).setValue(caption).setFontSize(9).setFontColor('#999999').setFontStyle('italic');
  return row + 2;
}


// ================================================================
// ЧИСТКА ДАННЫХ
// ================================================================
function cleanupData()   { runCleanup_(true); }
function cleanupReport() { runCleanup_(false); }

function runCleanup_(apply) {
  const sh = getDataSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert('Нет данных'); return; }

  const range = sh.getRange(2, 1, lastRow - 1, CFG.LAST_COL);
  const values = range.getValues();
  const headers = sh.getRange(1, 1, 1, CFG.LAST_COL).getValues()[0];
  const issues = [];
  let fixed = 0;
  const C = CFG.COL;

  for (let i = 0; i < values.length; i++) {
    const r = values[i];
    const rowNum = i + 2;

    [C.AMOUNT, C.PAYMENT, C.INTEREST, C.PRINCIPAL, C.BALANCE].forEach(idx => {
      const v = r[idx - 1];
      if (typeof v === 'string' && v.trim() !== '') {
        const num = parseRuNumber_(v);
        if (num !== null) {
          if (apply) values[i][idx - 1] = num;
          issues.push(`Строка ${rowNum}, «${headerName_(headers, idx)}»: "${v}" → ${num}`);
          fixed++;
        }
      }
    });

    [C.FIO, C.EXPERT, C.PRODUCT, C.STATUS, C.UNDERWRITER].filter(idx => idx > 0).forEach(idx => {
      const v = r[idx - 1];
      if (typeof v === 'string' && v !== v.trim()) {
        if (apply) values[i][idx - 1] = v.trim();
        issues.push(`Строка ${rowNum}, «${headerName_(headers, idx)}»: убраны лишние пробелы`);
        fixed++;
      }
    });

    const product = (r[C.PRODUCT - 1] || '').toString().trim();
    if (product && !CFG.PRODUCTS.includes(product)) {
      issues.push(`Строка ${rowNum}: неизвестный продукт "${product}"`);
    }
    const expert = (r[C.EXPERT - 1] || '').toString().trim();
    if (expert && !CFG.EXPERTS.includes(expert)) {
      issues.push(`Строка ${rowNum}: неизвестный эксперт "${expert}"`);
    }

    const status = (r[C.STATUS - 1] || '').toString().trim();
    const knownStatuses = ['Активный', 'Закрыл', ''];
    if (C.STATUS && status && !knownStatuses.includes(status)) {
      issues.push(`Строка ${rowNum}: неизвестный статус "${status}" (ожидалось «Активный» или «Закрыл»)`);
    }
    const balance = toNum_(r[C.BALANCE - 1]);
    if (r[C.FIO - 1] && !isDastras_(product) && statusMismatch_(balance, status)) {
      issues.push(`Строка ${rowNum}: остаток ${balance > 0 ? 'больше 0' : '= 0'}, а статус «${status || '(пусто)'}» — стоит сверить`);
    }

    // Дата выдачи нужна для подсветки
    if (!isDastras_(product) && r[C.FIO - 1] && !toDate_(r[C.ISSUE_DATE - 1])) {
      issues.push(`Строка ${rowNum}: нет даты выдачи — строка не будет подсвечиваться`);
    }
  }

  if (apply && fixed > 0) range.setValues(values);

  const msg = apply
    ? `Исправлено значений: ${fixed}. Всего замечаний: ${issues.length}`
    : `Найдено замечаний: ${issues.length}. Изменения НЕ применены`;
  const details = issues.slice(0, 40).join('\n') + (issues.length > 40 ? `\n… и ещё ${issues.length - 40}` : '');
  SpreadsheetApp.getUi().alert(msg, details || 'Замечаний нет', SpreadsheetApp.getUi().ButtonSet.OK);
}

/** Название колонки — из строки заголовков листа */
function headerName_(headers, colIdx) {
  const h = (headers[colIdx - 1] || '').toString().trim();
  return h || ('кол.' + colIdx);
}


// ================================================================
// РАСЧЁТЫ
// ================================================================
function calcWeightedRate() {
  const rate = weightedRateByAmount_(readLoans_());
  SpreadsheetApp.getUi().alert('Средневзвешенная ставка',
    (rate * 100).toFixed(2) + ' % (взвешено по объёму выдачи)',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function calcAvgTicket() {
  const loans = readLoans_();
  const avg = sumAmount_(loans) / (loans.length || 1);
  SpreadsheetApp.getUi().alert('Средний чек', fmtNum_(avg) + ' сомони',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function calcYearForecast() {
  const active = readLoans_().filter(l => !l.closed);
  const monthly = active.reduce((s, l) => s + l.interest, 0);
  const withIncome = active.filter(l => l.interest > 0).length;
  SpreadsheetApp.getUi().alert('Прогноз годового дохода',
    fmtNum_(monthly * 12) + ' сомони (сумма «Процент» по активным × 12)\n\n' +
    `Колонка «Процент» заполнена у ${withIncome} из ${active.length} активных кредитов — ` +
    'если не у всех, число занижено.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function calcCollateralCoverage() {
  const active = readLoans_().filter(l => !l.closed);
  const totalOd = active.reduce((s, l) => s + l.od, 0);
  const groups = {};
  active.forEach(l => {
    const t = collateralType_(l);
    if (!groups[t]) groups[t] = { count: 0, od: 0 };
    groups[t].count++; groups[t].od += l.od;
  });
  const lines = Object.entries(groups).sort((a, b) => b[1].od - a[1].od).map(([t, g]) =>
    `${t}: ${g.count} кред., ${fmtNum_(g.od)} сом. (${totalOd ? (g.od / totalOd * 100).toFixed(1) : 0}% ОД)`);
  SpreadsheetApp.getUi().alert('Залоговое покрытие',
    lines.join('\n') + (hasCollateralColumns_() ? '' :
      '\n\nВ файле нет колонок «Поручитель»/«Залог» — разбивка упрощённая, по продукту.'),
    SpreadsheetApp.getUi().ButtonSet.OK);
}


// ================================================================
// СБОР ПЛАТЕЖЕЙ
// ================================================================
function collectPaymentsWeek()  { collectPayments_(7);  }
function collectPaymentsMonth() { collectPayments_(30); }

/**
 * Неоплаченные платежи с датой в [сегодня, сегодня + days] — с учётом
 * следующего месяца, новых выдач и даты оплаты. Просрочка — отдельным списком сверху.
 */
function collectPayments_(days) {
  const loans = readLoans_();
  const today = startOfDay_(new Date());
  const upto = addDays_(today, days);

  const overdue = [], due = [];
  loans.forEach(l => {
    const st = paymentState_(l, today);
    if (st.state === 'closed' || st.state === 'nodate' || st.state === 'skip') return;
    if (st.state === 'overdue') { overdue.push({ l: l, d: st.expired ? l.maturity : st.due, expired: st.expired }); return; }

    const maturity = l.maturity instanceof Date ? startOfDay_(l.maturity) : null;
    const candidates = st.state === 'new'
      ? [{ d: st.due, next: false }]
      : [{ d: nextDueDate_(l.issueDate, today), next: false },
         { d: dueInMonth_(l.issueDate, today.getFullYear(), today.getMonth() + 1), next: true }];
    candidates.forEach(c => {
      if (c.d < today || c.d > upto) return;
      if (maturity && c.d > maturity) return;
      if (coversDue_(l, c.d, c.next)) return;
      due.push({ l: l, d: c.d });
    });
  });

  overdue.sort((a, b) => a.d - b.d);
  due.sort((a, b) => a.d - b.d);

  const fmt = x => `${Utilities.formatDate(x.d, Session.getScriptTimeZone(), 'dd.MM')} — ${x.l.fio} — ` +
    `${fmtNum_(x.expired ? x.l.balance : x.l.payment)} сом.` +
    (x.l.phone ? ` — ${fmtPhone_(x.l.phone)}` : '') + (x.expired ? ' (срок вышел, остаток)' : '');

  const parts = [];
  if (overdue.length) {
    parts.push(`🔴 ПРОСРОЧЕНО: ${overdue.length} на ${fmtNum_(overdue.reduce((s, x) => s + x.l.payment, 0))} сом.`);
    parts.push(overdue.slice(0, 15).map(fmt).join('\n') + (overdue.length > 15 ? `\n… и ещё ${overdue.length - 15}` : ''));
    parts.push('');
  }
  if (due.length) {
    parts.push(`🟠 К ОПЛАТЕ за ${days} дн.: ${due.length} на ${fmtNum_(due.reduce((s, x) => s + x.l.payment, 0))} сом.`);
    parts.push(due.slice(0, 25).map(fmt).join('\n') + (due.length > 25 ? `\n… и ещё ${due.length - 25}` : ''));
  } else {
    parts.push('В этот период неоплаченных платежей не найдено.');
  }

  const title = days === 7 ? 'Платежи на неделю' : 'Платежи на месяц';
  SpreadsheetApp.getUi().alert(title, parts.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}


// ================================================================
// PDF-ОТЧЁТ
// ================================================================
function exportReportPdf() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEET_DASHBOARD);
  if (!sh) { SpreadsheetApp.getUi().alert('Сначала обнови дашборд'); return; }

  const url = ss.getUrl().replace(/\/edit.*$/, '') +
    '/export?exportFormat=pdf&format=pdf&gid=' + sh.getSheetId() +
    '&portrait=false&fitw=true&sheetnames=false&printtitle=false&gridlines=false';
  const token = ScriptApp.getOAuthToken();
  const blob = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } })
                          .getBlob().setName(`${CFG.BRANCH_NAME}_дашборд_${dateStr_()}.pdf`);
  const file = DriveApp.createFile(blob);
  SpreadsheetApp.getUi().alert('PDF готов', 'Сохранён на Диск:\n' + file.getUrl(), SpreadsheetApp.getUi().ButtonSet.OK);
}


// ================================================================
// СНИМОК ПОРТФЕЛЯ (история)
// ================================================================
function savePortfolioSnapshot() {
  const lastRow = writePortfolioSnapshot_();
  SpreadsheetApp.getUi().alert('Снимок сохранён',
    `Записан на лист «${CFG.SHEET_HISTORY}», строка ${lastRow}`,
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/** Дописывает строку снимка на лист истории, возвращает номер строки */
function writePortfolioSnapshot_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CFG.SHEET_HISTORY);
  const headers = ['Дата снимка', 'Кол-во кредитов', 'Портфель ОД, сомони', 'Всего выдано, сомони']
    .concat(CFG.EXPERTS.map(e => e + ' (ОД)'), ['Дастрас (лимит)']);
  if (!sh) {
    sh = ss.insertSheet(CFG.SHEET_HISTORY);
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4472C4').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  }

  const loans = readLoans_();
  const balByExpert = {};
  CFG.EXPERTS.forEach(e => balByExpert[e] = 0);
  loans.forEach(l => { if (balByExpert.hasOwnProperty(l.expert)) balByExpert[l.expert] += l.od; });
  const dastTotals = readDastrasTotals_();

  sh.appendRow([
    new Date(), loans.length + dastTotals.count,
    loans.reduce((s, l) => s + l.od, 0),
    sumAmount_(loans)
  ].concat(CFG.EXPERTS.map(e => balByExpert[e]), [dastTotals.limitSum]));

  const lastRow = sh.getLastRow();
  sh.getRange(lastRow, 1).setNumberFormat('dd.MM.yyyy HH:mm');
  sh.getRange(lastRow, 2, 1, headers.length - 1).setNumberFormat('#,##0');
  return lastRow;
}


// ================================================================
// АВТО-РЕЖИМ
// ================================================================
function enableAutoMode() {
  removeAutoTriggers_();
  ScriptApp.newTrigger('recolorAllRowsSilent').timeBased().atHour(7).everyDays(1).create();
  ScriptApp.newTrigger('buildDashboard').timeBased().atHour(8).everyDays(1).create();
  ScriptApp.newTrigger('savePortfolioSnapshotSilent').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  SpreadsheetApp.getUi().alert('Авто-режим включён',
    'Подсветка — каждый день в 07:00.\nДашборд — каждый день в 08:00.\n' +
    'Снимок портфеля — каждый понедельник в 08:00.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/** Для триггера — без окон, иначе триггер падает */
function recolorAllRowsSilent() {
  recolorSheetFast_(getDataSheet_());
}

/** Для триггера — снимок без окна «Снимок сохранён» */
/** Для триггера — снимок без окна (getUi() в триггере недоступен) */
function savePortfolioSnapshotSilent() {
  writePortfolioSnapshot_();
}

function disableAutoMode() {
  removeAutoTriggers_();
  SpreadsheetApp.getActiveSpreadsheet().toast('Авто-режим выключен', CFG.BRANCH_NAME, 3);
}

function removeAutoTriggers_() {
  const names = new Set(['buildDashboard', 'savePortfolioSnapshot', 'savePortfolioSnapshotSilent', 'recolorAllRowsSilent']);
  ScriptApp.getProjectTriggers().forEach(t => {
    if (names.has(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });
}


// ================================================================
// УТИЛИТЫ
// ================================================================
function toNum_(v) {
  if (v === '' || v === null || v === undefined) return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  const n = parseRuNumber_(v);
  return n === null ? 0 : n;
}

function parseRuNumber_(v) {
  if (typeof v === 'number') return v;
  if (!v) return null;
  let s = v.toString().trim().replace(/\s| /g, '');
  if (s === '') return null;
  if (s.includes(',') && s.includes('.')) {
    s = s.replace(/\./g, '').replace(',', '.');
  } else if (s.includes(',')) {
    s = s.replace(',', '.');
  }
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

function parseRate_(v) {
  if (typeof v === 'number') return v > 1 ? v / 100 : v;
  if (!v) return 0;
  const m = v.toString().match(/([\d.,]+)/);
  return m ? parseFloat(m[1].replace(',', '.')) / 100 : 0;
}

function parseTerm_(v) {
  if (typeof v === 'number') return v;
  if (!v) return 0;
  const m = v.toString().match(/(\d+)/);
  return m ? parseInt(m[1], 10) : 0;
}

function weightedRateByAmount_(loans) {
  let num = 0, den = 0;
  loans.forEach(l => { num += l.rate * l.amount; den += l.amount; });
  return den ? num / den : 0;
}

function weightedRateByBalance_(loans) {
  let num = 0, den = 0;
  loans.forEach(l => { num += l.rate * l.od; den += l.od; });
  return den ? num / den : 0;
}

function fmtNum_(n) {
  return Math.round(Number(n) || 0).toLocaleString('ru-RU');
}

function fmtDate_(d) {
  return (d instanceof Date) ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd.MM.yyyy') : '—';
}

/**
 * 992987967675 → +992 98 796 7675. Несколько номеров в ячейке
 * («992927369626 \\ 992887791981») — каждый отдельно, через запятую.
 */
function fmtPhone_(v) {
  const text = (v || '').toString();
  const nums = text.match(/\d{9,12}/g);
  if (!nums) return text;
  return nums.map(d => {
    if (/^992\d{9}$/.test(d)) return '+992 ' + d.slice(3, 5) + ' ' + d.slice(5, 8) + ' ' + d.slice(8);
    if (/^\d{9}$/.test(d))     return '+992 ' + d.slice(0, 2) + ' ' + d.slice(2, 5) + ' ' + d.slice(5);
    return d;
  }).join(', ');
}

/** Дата из ячейки: настоящая дата или текст «дд.мм.гггг». Иначе null. */
function toDate_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  const m = (v || '').toString().trim().match(/^(\d{1,2})[.\/](\d{1,2})[.\/](\d{4})$/);
  if (!m) return null;
  const d = new Date(+m[3], +m[2] - 1, +m[1]);
  return d.getMonth() === +m[2] - 1 ? d : null;
}

function dateStr_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
