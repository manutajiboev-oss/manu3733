/**
 * ДАШБОРД ДОХОДА ПО ЭКСПЕРТАМ  —  веб-приложение (Google Apps Script, один файл)
 * Согдийский филиал ЗАО «Душанбе Сити Банк», г. Худжанд
 *
 * ЧТО ДЕЛАЕТ
 *   doGet() читает три ИСТОЧНИКА (каждый — свой файл Google Таблиц): выдачи
 *   («Кредитный лист»), архив и «Доходы», сопоставляет каждого клиента с экспертом
 *   (ключ: ФИО + сумма + дата выдачи), считает доход (тело / процент) по месяцам и
 *   открывает интерактивный дашборд с выбором эксперта и режимом «Вся команда».
 *
 * КАК РАЗВЕРНУТЬ
 *   1. Создай отдельный проект: script.google.com → «Создать проект».
 *   2. Вставь этот файл целиком в Код.gs.
 *   3. Впиши в CONFIG.SOURCES URL (или ID) КАЖДОГО файла-источника:
 *      ACTIVE_SS = «Выдача кредитов», ARCHIVE_SS = «Архив…», INCOME_SS = «Доходы…».
 *      (SS_ID — таблица по умолчанию, если какой-то источник оставить пустым.)
 *   4. Развернуть → Новое развёртывание → «Веб-приложение» →
 *      «Выполнять от имени: Я», «Доступ: только я» → Развернуть → открой URL.
 *   5. Если лист внутри файла не нашёлся — впиши имя вкладки в CONFIG.SHEETS.
 *
 * ДИАГНОСТИКА: запусти debugListSheets() → «Журнал выполнения» покажет вкладки
 *   каждого файла и что распозналось.
 */

// ======================= НАСТРОЙКИ =======================
var CONFIG = {
  // Таблица по умолчанию (если источник ниже не указан отдельно). Можно '' если скрипт привязан к таблице.
  SS_ID: '1yz0i7250HdkAEV6HI_2IkfQVk_1HMwekdCrDZ_EEyOw',

  // Источники — КАЖДЫЙ в своём файле Google Таблиц. Впиши URL или ID файла.
  // Пусто → берётся SS_ID выше. Годится и полный URL таблицы, и просто её ID.
  SOURCES: {
    ACTIVE_SS:  '',   // файл с «Кредитный лист» / выдачами (пусто = SS_ID)
    ARCHIVE_SS: '1162-LzkmA48fL8Uj9WfdoodTVmqxyMUa1G95R5emYQs',  // «Архив от 2022 до 2026»
    INCOME_SS:  '1Fuv5IWBn0wnu1xw-vqW_RusHj7yIL9MdSEN2zetSaLs'    // «Доходы_Янв_Авг_2026»
  },

  // Явные имена вкладок ВНУТРИ файла-источника. Оставь '' — авто-определение по заголовкам.
  SHEETS: {
    ACTIVE:  '',   // активные кредиты (есть колонки «Эксперт», «Продукт», «Остаток»)
    ARCHIVE: '',   // архив (есть колонки «Эксперт», «Статус», «Дата закрытия»)
    INCOME:  'Лист1'  // доходы: вкладка с данными (не «Дашборд доходов»). Пусто = авто-поиск.
  },

  EXCLUDE_DASTRAST: true,   // исключать продукт «Дастраст» из расчётов
  BONUS_RATE: 0.003,        // бонус эксперта = 0,3% от суммы выдач

  // Унификация имён экспертов (в разных листах пишут по-разному).
  // Ключ — подстрока имени (в нижнем регистре), значение — как показывать.
  EXPERT_ALIASES: {
    'неъматчон': 'Дадабаев Н.',
    'дадабаев':  'Дадабаев Н.',
    'ахроров':   'Ахроров М.',
    'ахмедов':   'Ахмедов Б.',
    'курбонов':  'Курбонов Х.',
    'карими':    'Карими О.'
  }
};

var MONTHS = ['Январь','Февраль','Март','Апрель','Май','Июнь','Июль','Август',
              'Сентябрь','Октябрь','Ноябрь','Декабрь'];
var MONTHS_SHORT = ['Янв','Фев','Мар','Апр','Май','Июн','Июл','Авг','Сен','Окт','Ноя','Дек'];

// ======================= ТОЧКА ВХОДА =======================
function doGet() {
  var html;
  try {
    html = renderHtml(buildPayload());
  } catch (err) {
    // показываем понятную ошибку вместо стандартной страницы Apps Script
    var msg = String(err && err.message || err).replace(/[&<>]/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c];
    });
    html = '<!doctype html><meta charset="utf-8"><body style="font-family:system-ui,sans-serif;padding:24px">' +
           '<h2>Не удалось построить дашборд</h2><p>' + msg + '</p>' +
           '<p style="color:#666">Запусти debugListSheets() в редакторе скриптов и проверь CONFIG.SOURCES / CONFIG.SHEETS.</p></body>';
  }
  return HtmlService.createHtmlOutput(html)
    .setTitle('Дашборд дохода по экспертам')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ======================= ДОСТУП К ТАБЛИЦАМ =======================
// извлекает ID из URL таблицы или принимает голый ID
function idFrom_(s) {
  if (!s) return '';
  var m = String(s).match(/[-\w]{25,}/);
  return m ? m[0] : String(s).trim();
}
// открывает таблицу по URL/ID; пусто → SS_ID; если и он пуст → активная таблица
function openSS_(idOrUrl) {
  var id = idFrom_(idOrUrl) || idFrom_(CONFIG.SS_ID);
  if (id) return SpreadsheetApp.openById(id);
  var a = SpreadsheetApp.getActiveSpreadsheet();
  if (!a) throw new Error('Не задан SS_ID/SOURCES и скрипт не привязан к таблице.');
  return a;
}

function norm_(v) {
  if (v === null || v === undefined) return '';
  return String(v).replace(/\s+/g, ' ').trim().toLowerCase();
}
function num_(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  var s = String(v).replace(/\s/g, '').replace(/,/g, '.').replace(/[^0-9.\-]/g, '');
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}
function ymd_(v) {
  if (v instanceof Date) {
    return v.getFullYear() + '-' + ('0'+(v.getMonth()+1)).slice(-2) + '-' + ('0'+v.getDate()).slice(-2);
  }
  if (v === null || v === undefined) return '';
  var s = String(v).trim();
  if (!s) return '';
  var m = s.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})/);   // дд.мм.гггг (текст)
  if (m) { var y = m[3].length === 2 ? '20'+m[3] : m[3];
    return y + '-' + ('0'+m[2]).slice(-2) + '-' + ('0'+m[1]).slice(-2); }
  var m2 = s.match(/^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})/);    // гггг-мм-дд
  if (m2) return m2[1] + '-' + ('0'+m2[2]).slice(-2) + '-' + ('0'+m2[3]).slice(-2);
  return s;
}
// нормализация номера договора: убираем пробелы, слэши/бэкслэши, регистр
function normDog_(v) {
  if (v === null || v === undefined) return '';
  return String(v).replace(/[\s\\\/]+/g, '').trim().toLowerCase();
}
// канонизация имени эксперта: убираем точки/лишние пробелы и приводим к единому виду по алиасам
function canonExpert_(raw) {
  if (raw === null || raw === undefined || String(raw).trim() === '') return '';
  var n = norm_(raw);
  var al = CONFIG.EXPERT_ALIASES || {};
  for (var key in al) { if (n.indexOf(key) !== -1) return al[key]; }
  return String(raw).replace(/\./g, ' ').replace(/\s+/g, ' ').trim();
}

// индекс колонки по списку вариантов заголовка (подстрока, регистронезависимо)
function colIdx_(headers, variants) {
  for (var i = 0; i < headers.length; i++) {
    var h = norm_(headers[i]);
    for (var j = 0; j < variants.length; j++) {
      if (h && h.indexOf(variants[j]) !== -1) return i;
    }
  }
  return -1;
}

// ======================= АВТО-ОПРЕДЕЛЕНИЕ ЛИСТОВ =======================
// подбирает нужный лист внутри файла: по явному имени или по сигнатуре заголовков
function pickSheet_(ss, type, explicitName) {
  if (explicitName) { var s = ss.getSheetByName(explicitName); if (s) return s; }
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    if (sh.getLastRow() < 2 || sh.getLastColumn() < 2) continue;
    var top = sh.getRange(1, 1, Math.min(3, sh.getLastRow()), sh.getLastColumn()).getValues();
    var r1 = (top[0] || []).map(norm_);
    var r2 = (top[1] || []).map(norm_);
    var has = function (row, kw) { for (var j=0;j<row.length;j++) if (row[j].indexOf(kw)!==-1) return true; return false; };
    if (type === 'income') {
      var mc = 0;
      for (var m = 0; m < MONTHS.length; m++) if (has(r1, MONTHS[m].toLowerCase())) mc++;
      if (mc >= 2 && (has(r2, 'тело') || has(r1, 'тело'))) return sh;
    } else if (type === 'archive') {
      if (has(r1,'эксперт') && has(r1,'статус') && (has(r1,'дата закры') || has(r1,'закрыт'))) return sh;
    } else if (type === 'active') {
      if (has(r1,'эксперт') && has(r1,'продукт') && has(r1,'остаток')) return sh;
    }
  }
  // запасной вариант — первый лист файла (частый случай, когда файл = один лист)
  return sheets.length ? sheets[0] : null;
}

// открывает каждый источник (возможно, в разных файлах) и находит в нём нужный лист
function findSheets_() {
  var actSS = openSS_(CONFIG.SOURCES && CONFIG.SOURCES.ACTIVE_SS);
  var arcSS = openSS_(CONFIG.SOURCES && CONFIG.SOURCES.ARCHIVE_SS);
  var incSS = openSS_(CONFIG.SOURCES && CONFIG.SOURCES.INCOME_SS);
  return {
    active:  pickSheet_(actSS, 'active',  CONFIG.SHEETS.ACTIVE),
    archive: pickSheet_(arcSS, 'archive', CONFIG.SHEETS.ARCHIVE),
    income:  pickSheet_(incSS, 'income',  CONFIG.SHEETS.INCOME),
    _ss: { active: actSS, archive: arcSS, income: incSS }
  };
}

// ======================= ПАРСИНГ =======================
// Возвращает карту кредитов: key(ФИО|сумма|дата) -> {expert, product, rate, ost, status, closeDate}
function readLoans_(activeSh, archiveSh) {
  var map = {};
  // ключ кредита: по номеру договора (если есть), иначе по ФИО+сумма+дата
  function put(fio, sum, date, dog, obj) {
    var d = ymd_(date);
    var nd = normDog_(dog);
    var key = nd ? ('D:' + nd) : ('N:' + norm_(fio) + '|' + Math.round(num_(sum)) + '|' + d);
    if (!map[key]) map[key] = { expert:'', product:'', rate:0, ost:0, status:'', closeDate:'', fio:fio, sum:num_(sum),
                                vyd:d, year:(d ? parseInt(d.slice(0,4),10) : 0), dog:nd };
    var t = map[key];
    for (var k in obj) if (obj[k] !== '' && obj[k] !== undefined && obj[k] !== null) t[k] = obj[k];
    return key;
  }

  // Активные
  if (activeSh) {
    var av = activeSh.getDataRange().getValues();
    var h = av[0];
    var cF = colIdx_(h, ['фио']),
        cS = colIdx_(h, ['сумма кредита','сумма']),
        cP = colIdx_(h, ['продукт']),
        cR = colIdx_(h, ['процентная ставка','ставка']),
        cO = colIdx_(h, ['остаток']),
        cE = colIdx_(h, ['эксперт']),
        cDg= colIdx_(h, ['номер договора','договор']),
        cD = colIdx_(h, ['дата выдач','выдач']);
    for (var i = 1; i < av.length; i++) {
      var row = av[i]; if (cF<0 || !row[cF]) continue;
      put(row[cF], cS>=0?row[cS]:0, cD>=0?row[cD]:'', cDg>=0?row[cDg]:'', {
        expert: cE>=0?canonExpert_(row[cE]):'',
        product: cP>=0?String(row[cP]).trim():'',
        rate: cR>=0?num_(row[cR]):0,
        ost: cO>=0?num_(row[cO]):0,
        status: 'акт',
        fromActive: true
      });
    }
  }
  // Архив
  if (archiveSh) {
    var rv = archiveSh.getDataRange().getValues();
    var hh = rv[0];
    var aF = colIdx_(hh, ['фио']),
        aS = colIdx_(hh, ['сумма']),
        aP = colIdx_(hh, ['продукт']),
        aE = colIdx_(hh, ['эксперт']),
        aSt= colIdx_(hh, ['статус']),
        aC = colIdx_(hh, ['дата закры','закрыт']),
        aDg= colIdx_(hh, ['номер договора','договор']),
        aD = colIdx_(hh, ['дата выдач','выдач']);
    for (var r = 1; r < rv.length; r++) {
      var rr = rv[r]; if (aF<0 || !rr[aF]) continue;
      var st = aSt>=0 ? norm_(rr[aSt]) : '';
      var closed = st.indexOf('закр') !== -1;
      put(rr[aF], aS>=0?rr[aS]:0, aD>=0?rr[aD]:'', aDg>=0?rr[aDg]:'', {
        expert:  aE>=0?canonExpert_(rr[aE]):'',
        product: aP>=0?String(rr[aP]).trim():'',
        status:  closed ? 'закр' : 'акт',
        closeDate: (closed && aC>=0) ? ymd_(rr[aC]) : '',
        fromArchive: true
      });
    }
  }
  return map;
}

// Читает лист доходов: возвращает {months:[короткие], rows:[{fio,sum,date,status,mp:[...],mt:[...]}]}
function readIncome_(sh) {
  if (!sh) throw new Error('Лист «Доходы» не найден. Впиши имя вкладки в CONFIG.SHEETS.INCOME.');
  var vals = sh.getDataRange().getValues();
  // найти строку заголовка месяцев (есть >=2 названия месяцев)
  var hdrRow = -1;
  for (var i = 0; i < Math.min(5, vals.length); i++) {
    var cnt = 0, row = vals[i].map(norm_);
    for (var m = 0; m < MONTHS.length; m++) for (var c = 0; c < row.length; c++) if (row[c].indexOf(MONTHS[m].toLowerCase())!==-1){cnt++;break;}
    if (cnt >= 2) { hdrRow = i; break; }
  }
  if (hdrRow < 0) throw new Error('В листе доходов не найдена строка с месяцами.');
  var subRow = hdrRow + 1;          // строка «Тело»/«%»
  var dataStart = subRow + 1;
  var h1 = vals[hdrRow], h2 = vals[subRow] || [];

  // колонки-идентификаторы (в строке заголовка месяцев или выше)
  var idHeaders = vals[hdrRow].map(norm_);
  var cFio = colIdx_(idHeaders, ['фио']),
      cSum = colIdx_(idHeaders, ['сумма']),
      cDate= colIdx_(idHeaders, ['дата выдач','выдач']),
      cSt  = colIdx_(idHeaders, ['статус']),
      cDg  = colIdx_(idHeaders, ['номер договора','договор']);

  // разметка месяцев: для каждой колонки с названием месяца берём её («Тело») и следующую («%»)
  var monthCols = [];  // {short, teloCol, procCol}
  var seenMi = {};
  for (var c = 0; c < h1.length; c++) {
    var name = norm_(h1[c]);
    for (var mi = 0; mi < MONTHS.length; mi++) {
      if (name.indexOf(MONTHS[mi].toLowerCase()) !== -1) {
        if (seenMi[mi]) break;          // месяц уже размечен (напр., второй столбец объединённой шапки)
        seenMi[mi] = true;
        var teloC = c, procC = c + 1;
        // если во 2-й строке порядок «%»/«Тело» — переставим
        if (norm_(h2[c]).indexOf('%') !== -1 || norm_(h2[c]).indexOf('процент') !== -1) { teloC = c+1; procC = c; }
        monthCols.push({ mi: mi, short: MONTHS_SHORT[mi], teloCol: teloC, procCol: procC });
        break;
      }
    }
  }
  monthCols.sort(function(a,b){ return a.mi - b.mi; });

  var rows = [];
  for (var r = dataStart; r < vals.length; r++) {
    var v = vals[r];
    var fio = cFio>=0 ? v[cFio] : '';
    if (!fio || !String(fio).trim()) continue;
    var mt = [], mp = [];
    for (var k = 0; k < monthCols.length; k++) {
      mt.push(Math.round(num_(v[monthCols[k].teloCol])*100)/100);
      mp.push(Math.round(num_(v[monthCols[k].procCol])*100)/100);
    }
    rows.push({
      fio: String(fio).trim(),
      sum: cSum>=0 ? num_(v[cSum]) : 0,
      date: cDate>=0 ? ymd_(v[cDate]) : '',
      status: cSt>=0 ? String(v[cSt]).trim() : '',
      dog: cDg>=0 ? v[cDg] : '',
      mt: mt, mp: mp
    });
  }
  // отбрасываем пустые месяцы в конце (колонка уже есть, но поступлений ещё нет — иначе «динамика −100%»)
  var keep = monthCols.length;
  while (keep > 1) {
    var any = false;
    for (var q = 0; q < rows.length && !any; q++) if (rows[q].mt[keep-1] || rows[q].mp[keep-1]) any = true;
    if (any) break;
    keep--;
  }
  if (keep < monthCols.length) {
    monthCols = monthCols.slice(0, keep);
    rows.forEach(function (x) { x.mt = x.mt.slice(0, keep); x.mp = x.mp.slice(0, keep); });
  }
  return { months: monthCols.map(function(x){return x.short;}), rows: rows };
}

// ======================= СБОРКА ДАННЫХ =======================
function buildPayload() {
  var f = findSheets_();
  if (!f.income) throw new Error('Не найден лист доходов в файле INCOME_SS. Впиши URL/ID файла «Доходы…» в CONFIG.SOURCES.INCOME_SS (и при нужде имя вкладки в CONFIG.SHEETS.INCOME).');

  var loans = readLoans_(f.active, f.archive);
  var inc = readIncome_(f.income);
  var months = inc.months, MN = months.length;

  // эксперт по каждому кредиту (для «Выдачи»): считаем портфель эксперта из loans-карты
  // 1) агрегируем клиентов дохода к экспертам
  var experts = {};   // name -> expert object
  function ensure(name) {
    if (!experts[name]) experts[name] = {
      name: name, clients: [], products: {}, portfolio: [],
      monthTelo: zeros(MN), monthProc: zeros(MN),
      kpiProc: 0, kpiTelo: 0, vydacha: 0, ost: 0,
      active: 0, issuedPeriod: 0, wsum: 0, wrate_num: 0
    };
    return experts[name];
  }
  function zeros(n){ var a=[]; for(var i=0;i<n;i++)a.push(0); return a; }
  var UNMATCHED = 'Не распознан';

  // период (год) — по датам выдачи строк дохода: самый частый год
  var yc = {};
  for (var ii = 0; ii < inc.rows.length; ii++) {
    var yy = inc.rows[ii].date ? parseInt(inc.rows[ii].date.slice(0,4),10) : 0;
    if (yy) yc[yy] = (yc[yy]||0)+1;
  }
  var periodYear = new Date().getFullYear();
  var best = -1;
  for (var y in yc) { if (yc[y] > best || (yc[y]===best && +y>periodYear)) { best = yc[y]; periodYear = +y; } }

  // 2) портфель по кредитам (loans): считаем ТОЛЬКО действующие кредиты (из «Кредитного листа», статус ≠ закр).
  //    Выдачи (кол-во и сумма), остаток и ставка — по одному и тому же набору действующих, чтобы число выдач = «активных».
  for (var key in loans) {
    var L = loans[key];
    if (!L.expert) continue;
    if (CONFIG.EXCLUDE_DASTRAST && norm_(L.product).indexOf('дастрас') !== -1) continue;
    var e = ensure(L.expert);
    if (L.fromActive && L.status !== 'закр') {
      e.active++; e.ost += L.ost;
      e.issuedPeriod++; e.vydacha += L.sum;               // выдачи = действующие кредиты (кол-во и сумма совпадают с «активных»)
      if (L.rate) { e.wsum += L.sum; e.wrate_num += L.sum * L.rate; }
      e.portfolio.push({ prod: L.product || '—', ost: L.ost, sum: L.sum, rate: L.rate || 0 });  // для фильтра по продукту
    }
  }

  // индекс для запасного матчинга по ФИО+сумма (если дата в файлах записана по-разному)
  var byFioSum = {};
  for (var lk in loans) {
    var lo = loans[lk];
    var fk = norm_(lo.fio) + '|' + Math.round(lo.sum);
    (byFioSum[fk] || (byFioSum[fk] = [])).push(lo);
  }

  // 3) доход: сопоставляем строки дохода с кредитами по ключу ФИО|сумма|дата, с запасным вариантом
  var unmatchedCount = 0;
  for (var i = 0; i < inc.rows.length; i++) {
    var row = inc.rows[i];
    // 1) основной ключ — номер договора (если он есть в строке дохода и в кредитах)
    var nd = normDog_(row.dog);
    var L = nd ? loans['D:' + nd] : null;
    // 2) запасной ключ — ФИО+сумма+дата
    if (!L) L = loans['N:' + norm_(row.fio) + '|' + Math.round(row.sum) + '|' + row.date];
    if (!L || !L.expert) {
      // запасной матчинг: по ФИО+сумма; при нескольких кредитах берём с совпадающей датой, иначе первый с экспертом
      var cand = byFioSum[norm_(row.fio) + '|' + Math.round(row.sum)];
      if (cand && cand.length) {
        var pick = null;
        for (var ci = 0; ci < cand.length; ci++) if (cand[ci].vyd === row.date && cand[ci].expert) { pick = cand[ci]; break; }
        if (!pick) for (var cj = 0; cj < cand.length; cj++) if (cand[cj].expert) { pick = cand[cj]; break; }
        if (pick) L = pick;
      }
    }
    var expName, prod, sum, status;
    if (L && L.expert) {
      if (CONFIG.EXCLUDE_DASTRAST && norm_(L.product).indexOf('дастрас') !== -1) continue;
      expName = L.expert; prod = L.product || '—'; sum = L.sum || row.sum; status = L.status;
    } else {
      expName = UNMATCHED; prod = '—'; sum = row.sum; status = row.status || ''; unmatchedCount++;
    }
    var e = ensure(expName);
    var telo = 0, proc = 0;
    for (var k = 0; k < MN; k++) { telo += row.mt[k]; proc += row.mp[k]; e.monthTelo[k]+=row.mt[k]; e.monthProc[k]+=row.mp[k]; }
    telo = r2(telo); proc = r2(proc);
    e.kpiTelo = r2(e.kpiTelo + telo); e.kpiProc = r2(e.kpiProc + proc);
    e.clients.push({ name: row.fio, prod: prod, sum: sum, status: status, telo: telo, proc: proc, mp: row.mp, mt: row.mt });
    var p = e.products[prod] || (e.products[prod] = { proc:0, telo:0, cnt:0, sum:0 });
    p.proc = r2(p.proc + proc); p.telo = r2(p.telo + telo); p.cnt++; p.sum += sum;
  }

  // финализация экспертов
  var list = [];
  for (var name in experts) {
    var e = experts[name];
    e.wrate = e.wsum ? r2(e.wrate_num / e.wsum * 100) : 0;
    e.bonus = r2(e.vydacha * CONFIG.BONUS_RATE);
    e.monthTelo = e.monthTelo.map(r2); e.monthProc = e.monthProc.map(r2);
    e.kpiTotal = r2(e.kpiProc + e.kpiTelo);
    delete e.wsum; delete e.wrate_num;
    list.push(e);
  }
  // сортировка: по доходу убыв., «Не распознан» в конец
  list.sort(function(a,b){
    if (a.name===UNMATCHED) return 1; if (b.name===UNMATCHED) return -1;
    return b.kpiProc - a.kpiProc;
  });

  // сводка по команде
  var team = { monthTelo: zeros(MN), monthProc: zeros(MN), byExpert: [], byProduct: {}, totalProc:0, totalTelo:0, vydacha:0, active:0 };
  list.forEach(function(e){
    if (e.name===UNMATCHED) return;
    team.totalProc = r2(team.totalProc + e.kpiProc);
    team.totalTelo = r2(team.totalTelo + e.kpiTelo);
    team.vydacha += e.vydacha; team.active += e.active;
    for (var k=0;k<MN;k++){ team.monthTelo[k]+=e.monthTelo[k]; team.monthProc[k]+=e.monthProc[k]; }
    team.byExpert.push({ name:e.name, proc:e.kpiProc, telo:e.kpiTelo, vydacha:e.vydacha, active:e.active });
    for (var pr in e.products){ var t=team.byProduct[pr]||(team.byProduct[pr]={proc:0,telo:0,cnt:0,sum:0});
      t.proc=r2(t.proc+e.products[pr].proc); t.telo=r2(t.telo+e.products[pr].telo); t.cnt+=e.products[pr].cnt; t.sum+=e.products[pr].sum; }
  });
  team.monthTelo=team.monthTelo.map(r2); team.monthProc=team.monthProc.map(r2);
  team.byExpert.sort(function(a,b){return b.proc-a.proc;});

  function r2(x){ return Math.round(x*100)/100; }

  // СВЕРКА: итог по ВСЕМ клиентам файла (включая «Не распознан») — должен совпасть с листом «Дашборд доходов»
  var grandProc = 0, grandTelo = 0;
  list.forEach(function(e){ grandProc += e.kpiProc; grandTelo += e.kpiTelo; });

  return {
    generatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone()||'Asia/Dushanbe', 'dd.MM.yyyy HH:mm'),
    months: months,
    periodYear: periodYear,
    periodLabel: months.length ? (months[0] + '–' + months[months.length-1] + ' ' + periodYear) : '',
    experts: list,
    team: team,
    unmatched: unmatchedCount,
    grandProc: r2(grandProc),
    grandTelo: r2(grandTelo),
    bonusRate: CONFIG.BONUS_RATE
  };
}

// диагностика — запусти вручную, смотри логи (Ctrl+Enter)
function debugListSheets() {
  var f = findSheets_();
  function tabs(ss){ return ss ? ss.getSheets().map(function(s){return '«'+s.getName()+'»';}).join(', ') : '—'; }
  Logger.log('Файл ACTIVE («'+(f._ss.active?f._ss.active.getName():'?')+'») вкладки: ' + tabs(f._ss.active));
  Logger.log('Файл ARCHIVE («'+(f._ss.archive?f._ss.archive.getName():'?')+'») вкладки: ' + tabs(f._ss.archive));
  Logger.log('Файл INCOME («'+(f._ss.income?f._ss.income.getName():'?')+'») вкладки: ' + tabs(f._ss.income));
  Logger.log('Определено → Активные: ' + (f.active?f.active.getName():'НЕ НАЙДЕН') +
             ' | Архив: ' + (f.archive?f.archive.getName():'НЕ НАЙДЕН') +
             ' | Доходы: ' + (f.income?f.income.getName():'НЕ НАЙДЕН'));
  var p = buildPayload();
  Logger.log('Эксперты: ' + p.experts.map(function(e){return e.name+' ('+e.kpiProc+')';}).join('; '));
  Logger.log('Не распознано строк дохода: ' + p.unmatched);
}

// ======================= HTML (вшит в этот же файл) =======================
function renderHtml(payload) {
  var json = JSON.stringify(payload)
    .replace(/<\//g, '<\\/')
    .replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
  // функция-замена: иначе «$&», «$'» и т.п. в данных (ФИО, продукты) испортят JSON
  return HTML_TEMPLATE.replace('/*__PAYLOAD__*/{}', function () { return json; });
}

var HTML_TEMPLATE = `<!doctype html><html><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@500;600&display=swap">
<title>Доход по экспертам</title>
<style>
:root{
  color-scheme:light;
  --bg:#eef1f8; --panel:#ffffff; --panel2:#f6f8fc; --ink:#0f1222; --ink2:#5b6076; --muted:#98a0b4;
  --line:#e6eaf2; --ring:rgba(15,18,34,.06);
  --accent:#4f46e5; --accent2:#6366f1; --telo:#b6bdd0;
  --up:#0ca678; --down:#e5484d;
  --c1:#4f46e5; --c2:#f97316; --c3:#10b981; --c4:#f5a524; --c5:#8b5cf6; --c6:#ef4444;
  --t-ind:#eef0ff; --t-teal:#e7f8f1; --t-amber:#fef4e2; --t-slate:#eef1f7; --t-rose:#fdeef0;
  --shadow:0 1px 2px rgba(15,18,34,.04), 0 6px 20px rgba(15,18,34,.05);
}
@media(prefers-color-scheme:dark){:root:not([data-theme="light"]){
  --bg:#0b0d16; --panel:#151826; --panel2:#1b1f30; --ink:#f4f5fb; --ink2:#b9bfd2; --muted:#7f869c;
  --line:#262a3d; --ring:rgba(255,255,255,.08);
  --accent:#7c83ff; --accent2:#8b90ff; --telo:#3b4160;
  --up:#2bd4a0; --down:#ff6b6b;
  --c1:#7c83ff; --c2:#fb923c; --c3:#2dd4a7; --c4:#fbbf3c; --c5:#a78bfa; --c6:#f87171;
  --t-ind:#1c1f3a; --t-teal:#0f2b26; --t-amber:#2b2411; --t-slate:#1a1e2e; --t-rose:#2b1620;
  --shadow:none;
}}
:root[data-theme="dark"]{
  --bg:#0b0d16; --panel:#151826; --panel2:#1b1f30; --ink:#f4f5fb; --ink2:#b9bfd2; --muted:#7f869c;
  --line:#262a3d; --ring:rgba(255,255,255,.08);
  --accent:#7c83ff; --accent2:#8b90ff; --telo:#3b4160;
  --up:#2bd4a0; --down:#ff6b6b;
  --c1:#7c83ff; --c2:#fb923c; --c3:#2dd4a7; --c4:#fbbf3c; --c5:#a78bfa; --c6:#f87171;
  --t-ind:#1c1f3a; --t-teal:#0f2b26; --t-amber:#2b2411; --t-slate:#1a1e2e; --t-rose:#2b1620;
  --shadow:none;
}
*{box-sizing:border-box}html,body{margin:0}
body{background:var(--bg);color:var(--ink);font-family:"Manrope",system-ui,-apple-system,"Segoe UI",sans-serif;line-height:1.45;-webkit-font-smoothing:antialiased;}
.wrap{max-width:1120px;margin:0 auto;padding:24px 16px 40px;}
.num{font-variant-numeric:tabular-nums;}
.mono{font-family:"IBM Plex Mono",ui-monospace,monospace;}

header.top{display:flex;justify-content:space-between;align-items:flex-start;gap:16px;flex-wrap:wrap;margin-bottom:16px;}
.eyebrow{font-size:11px;letter-spacing:.16em;text-transform:uppercase;color:var(--accent);font-weight:700;}
h1{font-size:27px;font-weight:800;margin:.12em 0 .35em;letter-spacing:-.02em;}
.meta{display:flex;gap:8px;flex-wrap:wrap}
.chip{font-size:12px;color:var(--ink2);background:var(--panel);border:1px solid var(--line);border-radius:999px;padding:4px 11px;font-weight:600;}
.chip b{color:var(--ink)}

.controls{display:flex;gap:10px;align-items:center;flex-wrap:wrap;background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:12px 14px;margin-bottom:18px;box-shadow:var(--shadow);}
.controls .grp{display:flex;gap:8px;align-items:center;}
.controls label{font-size:12px;color:var(--muted);font-weight:700;text-transform:uppercase;letter-spacing:.04em;}
select,.btn{appearance:none;cursor:pointer;background:var(--panel2);color:var(--ink);border:1px solid var(--line);border-radius:10px;padding:8px 12px;font:inherit;font-size:13px;font-weight:600;}
select:hover,.btn:hover{border-color:var(--accent);}
.sep{flex:1}
.presets{display:flex;gap:6px;}
.presets .btn{padding:7px 11px;font-size:12px;}
.presets .btn.on{background:var(--accent);color:#fff;border-color:var(--accent);}

.tiles{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:16px;}
.tile{background:var(--panel);border:1px solid var(--line);border-radius:16px;padding:15px 16px;box-shadow:var(--shadow);position:relative;overflow:hidden;}
.tile .k{font-size:11px;color:var(--muted);font-weight:700;letter-spacing:.03em;text-transform:uppercase;}
.tile .v{font-size:24px;font-weight:800;margin-top:7px;letter-spacing:-.02em;}
.tile .v small{font-size:12.5px;font-weight:700;color:var(--ink2);margin-left:3px;}
.tile .sub{font-size:11.5px;color:var(--ink2);margin-top:5px;display:flex;align-items:center;gap:5px;}
.tile.hero{background:linear-gradient(135deg,var(--accent),var(--accent2));border-color:transparent;color:#fff;}
.tile.hero .k,.tile.hero .v small{color:rgba(255,255,255,.85)}.tile.hero .v{color:#fff}.tile.hero .sub{color:rgba(255,255,255,.9)}
.tile.tint-teal{background:var(--t-teal)}.tile.tint-amber{background:var(--t-amber)}.tile.tint-slate{background:var(--t-slate)}.tile.tint-ind{background:var(--t-ind)}.tile.tint-rose{background:var(--t-rose)}
.delta{font-weight:700;font-size:12px;padding:1px 7px;border-radius:999px;display:inline-flex;align-items:center;gap:3px;}
.delta.up{color:var(--up);background:color-mix(in srgb,var(--up) 14%,transparent);}
.delta.down{color:var(--down);background:color-mix(in srgb,var(--down) 14%,transparent);}
.delta.flat{color:var(--muted);background:var(--panel2);}

.grid2{display:grid;grid-template-columns:1fr 1fr;gap:16px;}
.panel{background:var(--panel);border:1px solid var(--line);border-radius:18px;padding:18px 20px;box-shadow:var(--shadow);margin-bottom:16px;}
.panel h2{font-size:15px;font-weight:700;margin:0 0 2px;}
.panel .cap{font-size:12px;color:var(--muted);margin:0 0 14px;}
.legend{display:flex;gap:16px;flex-wrap:wrap;font-size:12px;color:var(--ink2);margin-bottom:8px;}
.legend span{display:inline-flex;align-items:center;gap:6px;}
.sw{width:11px;height:11px;border-radius:3px;display:inline-block;}
svg{display:block;width:100%;height:auto;overflow:visible;}
.axlabel{fill:var(--muted);font-size:11px;}
.gridline{stroke:var(--line);stroke-width:1;}
.baseline{stroke:var(--line);stroke-width:1.4;}
.val{fill:var(--ink);font-size:11px;font-weight:700;font-family:"IBM Plex Mono",monospace;}
.barlabel{fill:var(--ink2);font-size:12px;}

table{width:100%;border-collapse:collapse;font-size:12.5px;}
th,td{padding:9px 10px;text-align:right;border-bottom:1px solid var(--line);}
th:first-child,td:first-child{text-align:left;}
th{font-size:11px;text-transform:uppercase;letter-spacing:.04em;color:var(--muted);font-weight:700;}
tbody tr:hover{background:var(--panel2);}
td.name{color:var(--ink);font-weight:600;}
.pill{font-size:11px;padding:2px 9px;border-radius:999px;border:1px solid var(--ring);color:var(--ink2);white-space:nowrap;}
tr.tot td{font-weight:800;border-top:2px solid var(--line);border-bottom:none;background:var(--panel2);}
.dim{color:var(--muted);}
.tblwrap{overflow-x:auto;}
.arrow-up{color:var(--up);font-weight:700;}.arrow-down{color:var(--down);font-weight:700;}

.note{font-size:12px;color:var(--ink2);background:var(--panel2);border:1px solid var(--line);border-left:3px solid var(--accent);border-radius:10px;padding:11px 14px;margin-top:4px;}
footer{margin-top:22px;font-size:11.5px;color:var(--muted);text-align:center;}
#tt{position:fixed;pointer-events:none;opacity:0;transform:translate(-50%,-115%);transition:opacity .1s;background:var(--ink);color:var(--bg);font-size:11.5px;padding:7px 10px;border-radius:8px;z-index:9;white-space:nowrap;box-shadow:0 6px 18px rgba(0,0,0,.28);font-weight:600;}
#tt .tk{opacity:.7;font-size:10.5px;}
@media(max-width:820px){.tiles{grid-template-columns:repeat(2,1fr);}.grid2{grid-template-columns:1fr;}h1{font-size:22px;}}
@media(prefers-reduced-motion:reduce){*{transition:none!important}}
</style></head><body>
<div class="wrap">
  <header class="top">
    <div>
      <div class="eyebrow">Кредитный дашборд · доход</div>
      <h1 id="title">Доход по экспертам</h1>
      <div class="meta">
        <span class="chip">Период: <b id="period"></b></span>
        <span class="chip">Согдийский филиал ЗАО «Душанбе Сити Банк»</span>
        <span class="chip" id="gen"></span>
      </div>
    </div>
    <button class="btn" id="themeBtn">◐ Тема</button>
  </header>

  <div class="controls">
    <div class="grp"><label>Эксперт</label><select id="selExpert"></select></div>
    <div class="grp"><label>Продукт</label><select id="selProd"></select></div>
    <div class="grp"><label>Месяцы</label>
      <select id="selFrom"></select><span style="color:var(--muted)">—</span><select id="selTo"></select>
    </div>
    <span class="sep"></span>
    <div class="presets" id="presets"></div>
    <button class="btn" id="reloadBtn" title="Перечитать свежие данные из таблиц">⟳ Обновить</button>
    <button class="btn" id="dlBtn" title="Скачать текущий вид в CSV">⬇ Скачать CSV</button>
  </div>

  <section class="tiles" id="tiles"></section>

  <div class="panel">
    <h2>Поступления по месяцам</h2>
    <p class="cap">Возврат тела и процентный доход, сомони. Над столбцами — изменение %-дохода к предыдущему месяцу.</p>
    <div class="legend">
      <span><i class="sw" style="background:var(--telo)"></i>Тело (возврат ОД)</span>
      <span><i class="sw" style="background:var(--accent)"></i>Процент (доход)</span>
    </div>
    <div id="cMonth"></div>
  </div>

  <div class="grid2">
    <div class="panel"><h2 id="hA"></h2><p class="cap" id="capA"></p><div id="cA"></div></div>
    <div class="panel"><h2 id="hB"></h2><p class="cap" id="capB"></p><div id="cB"></div></div>
  </div>

  <div class="panel">
    <h2>Динамика по месяцам</h2>
    <p class="cap">Помесячно за выбранный диапазон, с изменением %-дохода к предыдущему месяцу.</p>
    <div class="tblwrap"><table id="tblDyn"></table></div>
  </div>

  <div class="panel">
    <h2 id="hTbl"></h2><p class="cap" id="capTbl"></p>
    <div class="tblwrap"><table id="tbl"></table></div>
  </div>

  <p class="note" id="note"></p>
  <footer>Доход = проценты. Тело — возврат основного долга. Доходность = %-доход к остатку ОД за выбранный период. Все цифры пересчитываются по выбранным месяцам.</footer>
</div>
<div id="tt"></div>
<script>
var DATA=/*__PAYLOAD__*/{};
var SVGNS="http://www.w3.org/2000/svg";
function el(t,a){var n=document.createElementNS(SVGNS,t);for(var k in a)n.setAttribute(k,a[k]);return n;}
// экранирование текста из таблиц перед вставкой в innerHTML (ФИО, продукты, имена экспертов)
function esc(s){return String(s==null?"":s).replace(/[&<>"]/g,function(ch){return {"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[ch];});}
function fmt(n,d){d=(d==null?2:d);if(n==null||isNaN(n))return"—";if(n===0)return"0";var neg=n<0;var s=Math.abs(n).toFixed(d).replace(".",",");var p=s.split(",");p[0]=p[0].replace(/\\B(?=(\\d{3})+(?!\\d))/g," ");return (neg?"−":"")+p.join(",");}
function f0(n){return fmt(n,0);}
function pct(n){return (n>=0?"":"−")+Math.abs(n).toFixed(1).replace(".",",")+"%";}
var tt=document.getElementById("tt");
function sTT(e,h){tt.innerHTML=h;tt.style.left=e.clientX+"px";tt.style.top=e.clientY+"px";tt.style.opacity="1";}
function hTT(){tt.style.opacity="0";}
var PC={"Автокредит":"var(--c1)","Ипотека":"var(--c2)","Истеъмоли":"var(--c3)","Ломбарди":"var(--c4)","Ломбард":"var(--c4)","Бизнес":"var(--c5)","Тиҷорати":"var(--c5)","Тиҷорати (бизнес)":"var(--c5)"};
function pcol(p){return PC[p]||PC[prodKey(p)]||"var(--c6)";}
function clr(id){document.getElementById(id).innerHTML="";}
function zeros(n){var a=[];for(var i=0;i<n;i++)a.push(0);return a;}

var M=DATA.months, N=M.length;
var TEAM=DATA.experts.filter(function(e){return e.name!=="Не распознан";});
var state={who:"__team__",from:0,to:N-1,prod:"__all__"};
var LAST={};
// приведение названий продуктов к единому ключу (Ломбард/Ломбарди, Тиҷорати/Бизнес)
function prodKey(p){p=(p||"").toString().trim();var l=p.toLowerCase();
  if(l.indexOf("ломбар")===0)return "Ломбарди";
  if(l.indexOf("тиҷ")===0||l.indexOf("тич")===0||l.indexOf("бизнес")===0)return "Тиҷорати (бизнес)";
  if(l.indexOf("исте")===0||l.indexOf("истеъ")===0)return "Истеъмоли";
  if(l.indexOf("ипотек")===0)return "Ипотека";
  if(l.indexOf("автокр")===0)return "Автокредит";
  return p||"—";}
function prodOK(p){return state.prod==="__all__"||prodKey(p)===state.prod;}

function calc(who,from,to){
  var experts = who==="__team__" ? TEAM : [DATA.experts[who]];
  // клиенты дохода (фильтр по продукту)
  var clients=[];
  experts.forEach(function(e){e.clients.forEach(function(c){if(prodOK(c.prod))clients.push(c);});});
  var monthP=zeros(N),monthT=zeros(N),prod={};
  clients.forEach(function(c){
    for(var i=0;i<N;i++){monthP[i]+=(c.mp[i]||0);monthT[i]+=(c.mt[i]||0);}
    var cP=0,cT=0;for(var j=from;j<=to;j++){cP+=(c.mp[j]||0);cT+=(c.mt[j]||0);}
    c._rp=cP;c._rt=cT;
    var pk=prodKey(c.prod);var p=prod[pk]||(prod[pk]={proc:0,cnt:0,sum:0});p.proc+=cP;p.cnt++;p.sum+=c.sum;
  });
  var kpiProc=0,kpiTelo=0;for(var i=from;i<=to;i++){kpiProc+=monthP[i];kpiTelo+=monthT[i];}
  // портфель (действующие кредиты) — фильтр по продукту
  var ost=0,active=0,vyd=0,wn=0,wd=0;
  experts.forEach(function(e){(e.portfolio||[]).forEach(function(L){if(prodOK(L.prod)){active++;ost+=L.ost;vyd+=L.sum;if(L.rate){wn+=L.sum*L.rate;wd+=L.sum;}}});});
  var wrate= wd?wn/wd*100:0;
  var bonus= vyd*(DATA.bonusRate||0.003);
  return {clients:clients,monthP:monthP,monthT:monthT,prod:prod,kpiProc:kpiProc,kpiTelo:kpiTelo,ost:ost,active:active,vyd:vyd,bonus:bonus,wrate:wrate,issued:active};
}
function deltaBadge(cur,prev){
  if(prev==null)return '<span class="delta flat">—</span>';
  if(prev===0)return '<span class="delta flat">—</span>';
  var d=(cur-prev)/prev*100;
  if(Math.abs(d)<0.05)return '<span class="delta flat">0%</span>';
  var cls=d>0?"up":"down";var ar=d>0?"▲":"▼";
  return '<span class="delta '+cls+'">'+ar+' '+pct(d)+'</span>';
}

/* ---------- controls ---------- */
(function(){
  var se=document.getElementById("selExpert");
  var o=document.createElement("option");o.value="__team__";o.textContent="◆ Вся команда";se.appendChild(o);
  DATA.experts.forEach(function(e,i){var op=document.createElement("option");op.value=String(i);op.textContent=e.name+"  ("+f0(e.kpiProc||sumArr(e.monthProc))+" сом.)";se.appendChild(op);});
  var sf=document.getElementById("selFrom"),st=document.getElementById("selTo");
  M.forEach(function(m,i){var a=document.createElement("option");a.value=i;a.textContent=m;sf.appendChild(a);var b=document.createElement("option");b.value=i;b.textContent=m;st.appendChild(b);});
  sf.value=0;st.value=N-1;
  // продукты — собираем из портфеля и клиентов
  var sp=document.getElementById("selProd");
  var pset={};
  DATA.experts.forEach(function(e){(e.portfolio||[]).forEach(function(L){pset[prodKey(L.prod)]=1;});e.clients.forEach(function(c){pset[prodKey(c.prod)]=1;});});
  var oAll=document.createElement("option");oAll.value="__all__";oAll.textContent="Все продукты";sp.appendChild(oAll);
  Object.keys(pset).filter(function(p){return p&&p!=="—";}).sort().forEach(function(p){var op=document.createElement("option");op.value=p;op.textContent=p;sp.appendChild(op);});
  sp.addEventListener("change",function(){state.prod=sp.value;render();});
  se.addEventListener("change",function(){state.who=se.value;render();});
  sf.addEventListener("change",function(){state.from=+sf.value;if(state.from>state.to){state.to=state.from;st.value=state.to;}render();});
  st.addEventListener("change",function(){state.to=+st.value;if(state.to<state.from){state.from=state.to;sf.value=state.from;}render();});
  // presets
  var pr=document.getElementById("presets");
  var presets=[["Весь период",0,N-1],["Последний мес",N-1,N-1]];
  if(N>=3)presets.splice(1,0,["Последние 3",N-3,N-1]);
  presets.forEach(function(p){var b=document.createElement("button");b.className="btn";b.textContent=p[0];
    b.addEventListener("click",function(){state.from=p[1];state.to=p[2];sf.value=state.from;st.value=state.to;render();});pr.appendChild(b);});
  document.getElementById("period").textContent=DATA.periodLabel;
  document.getElementById("gen").textContent="Обновлено "+DATA.generatedAt;
  render();
})();
function sumArr(a){var s=0;for(var i=0;i<a.length;i++)s+=a[i];return s;}

/* ---------- скачивание CSV ---------- */
function csvVal(s){s=(s==null?"":String(s));if(/[";\\n]/.test(s))return '"'+s.replace(/"/g,'""')+'"';return s;}
function csvNum(n){return (n==null||isNaN(n))?"":(Math.round(n*100)/100).toString().replace(".",",");}
function downloadCSV(){
  var A=LAST.A; if(!A)return; var who=state.who, rl=LAST.range;
  var L=[];
  L.push(["Дашборд дохода", who==="__team__"?"Вся команда":DATA.experts[who].name].map(csvVal).join(";"));
  L.push(["Период", rl+" "+DATA.periodYear].map(csvVal).join(";"));
  L.push(["Доход (%)", csvNum(A.kpiProc), "Тело", csvNum(A.kpiTelo), "Всего", csvNum(A.kpiProc+A.kpiTelo)].join(";"));
  L.push("");
  L.push("Динамика по месяцам");
  L.push(["Месяц","Тело","Процент","Всего"].join(";"));
  for(var i=state.from;i<=state.to;i++) L.push([M[i],csvNum(A.monthT[i]),csvNum(A.monthP[i]),csvNum(A.monthT[i]+A.monthP[i])].join(";"));
  L.push(["Итого",csvNum(A.kpiTelo),csvNum(A.kpiProc),csvNum(A.kpiProc+A.kpiTelo)].join(";"));
  L.push("");
  if(who==="__team__"){
    L.push("Сводка по экспертам");
    L.push(["Эксперт","Активных","Выдачи","Остаток ОД","Тело","Процент","Доходность %"].join(";"));
    TEAM.map(function(e){var idx=DATA.experts.indexOf(e);var c=calc(String(idx),state.from,state.to);
      return {n:e.name,act:c.active,vyd:c.vyd,ost:c.ost,telo:c.kpiTelo,proc:c.kpiProc};})
    .sort(function(a,b){return b.proc-a.proc;})
    .forEach(function(r){L.push([csvVal(r.n),r.act,csvNum(r.vyd),csvNum(r.ost),csvNum(r.telo),csvNum(r.proc),(r.ost>0?csvNum(r.proc/r.ost*100):"")].join(";"));});
  } else {
    L.push("Клиенты");
    L.push(["Клиент","Продукт","Сумма","Статус","Тело","Процент"].join(";"));
    A.clients.slice().sort(function(a,b){return b._rp-a._rp;}).forEach(function(c){
      L.push([csvVal(c.name),csvVal(c.prod),csvNum(c.sum),csvVal(c.status),csvNum(c._rt),csvNum(c._rp)].join(";"));});
  }
  var csv="﻿"+L.join("\\r\\n");
  var name="dohod_"+(who==="__team__"?"komanda":DATA.experts[who].name.replace(/[^\\wа-яА-ЯёЁ]+/g,"_"))+"_"+M[state.from]+"-"+M[state.to]+".csv";
  try{
    var blob=new Blob([csv],{type:"text/csv;charset=utf-8;"});
    var url=URL.createObjectURL(blob);
    var a=document.createElement("a");a.href=url;a.download=name;document.body.appendChild(a);a.click();
    setTimeout(function(){URL.revokeObjectURL(url);a.remove();},1500);
  }catch(e){
    // запасной вариант, если песочница блокирует Blob-скачивание
    window.open("data:text/csv;charset=utf-8,"+encodeURIComponent(csv),"_blank");
  }
}
document.getElementById("dlBtn").addEventListener("click",downloadCSV);
document.getElementById("reloadBtn").addEventListener("click",function(){location.reload();});

/* ---------- charts ---------- */
function monthChart(from,to,monthP,monthT){
  clr("cMonth");
  var idx=[];for(var i=from;i<=to;i++)idx.push(i);
  var W=980,H=300,mL=58,mR=14,mT=26,mB=34,iw=W-mL-mR,ih=H-mT-mB,n=idx.length,step=iw/n,bw=Math.min(58,step*0.58);
  var max=0;idx.forEach(function(i){max=Math.max(max,monthT[i]+monthP[i]);});if(max<=0)max=1;
  var sg=niceStep(max),nice=Math.ceil(max/sg)*sg,y=function(v){return mT+ih-(v/nice)*ih;};
  var svg=el("svg",{viewBox:"0 0 "+W+" "+H});
  for(var g=0;g<=nice+1;g+=sg){var yy=y(g);svg.appendChild(el("line",{class:"gridline",x1:mL,x2:W-mR,y1:yy,y2:yy}));var tx=el("text",{class:"axlabel",x:mL-8,y:yy+4,"text-anchor":"end"});tx.textContent=f0(g);svg.appendChild(tx);}
  svg.appendChild(el("line",{class:"baseline",x1:mL,x2:W-mR,y1:y(0),y2:y(0)}));
  idx.forEach(function(i,k){
    var cx=mL+step*k+step/2,x=cx-bw/2,t=monthT[i],p=monthP[i];
    if(t>0)svg.appendChild(el("rect",{x:x,y:y(t),width:bw,height:Math.max(y(0)-y(t),2),rx:4,fill:"var(--telo)"}));
    if(p>0){var yTop=y(t+p),h=y(t)-yTop-(t>0?2:0);svg.appendChild(el("rect",{x:x,y:yTop,width:bw,height:Math.max(h,2),rx:4,fill:"var(--accent)"}));}
    // value + delta vs previous month (global index i-1)
    if(p>0||t>0){
      var lab=el("text",{class:"val",x:cx,y:y(t+p)-16,"text-anchor":"middle"});lab.textContent=f0(p);svg.appendChild(lab);
      if(i>0){var prev=monthP[i-1];if(prev>0){var d=(p-prev)/prev*100;var col=d>=0?"var(--up)":"var(--down)";var ar=d>=0?"▲":"▼";
        var dl=el("text",{x:cx,y:y(t+p)-4,"text-anchor":"middle"});dl.setAttribute("font-size","10");dl.setAttribute("font-weight","700");dl.setAttribute("fill",col);dl.textContent=ar+" "+Math.abs(d).toFixed(0)+"%";svg.appendChild(dl);}}
    }
    var mx=el("text",{class:"axlabel",x:cx,y:H-12,"text-anchor":"middle"});mx.textContent=M[i];svg.appendChild(mx);
    var hit=el("rect",{x:cx-step/2,y:mT,width:step,height:ih,fill:"transparent"});
    hit.addEventListener("mousemove",function(e){sTT(e,"<b>"+M[i]+"</b><br><span class=tk>процент</span> "+fmt(p)+"<br><span class=tk>тело</span> "+fmt(t)+"<br><span class=tk>всего</span> "+fmt(t+p));});
    hit.addEventListener("mouseleave",hTT);svg.appendChild(hit);
  });
  document.getElementById("cMonth").appendChild(svg);
}
function niceStep(max){var raw=max/5,mag=Math.pow(10,Math.floor(Math.log(raw)/Math.LN10)),nn=raw/mag;var s=nn>=5?5:nn>=2?2:1;return s*mag;}
function hbars(id,items,colorFn,labelFn,tipFn){
  clr(id);if(!items.length){document.getElementById(id).innerHTML='<p class="cap">Нет данных за выбранный период.</p>';return;}
  var W=980,rowH=34,mT=6,mB=8,mL=210,mR=120,H=mT+mB+items.length*rowH,iw=W-mL-mR;
  var max=0;items.forEach(function(d){max=Math.max(max,d.v);});if(max<=0)max=1;var x=function(v){return v/max*iw;};
  var svg=el("svg",{viewBox:"0 0 "+W+" "+H});
  items.forEach(function(d,i){var cy=mT+i*rowH+rowH/2,bh=18,yy=cy-bh/2,w=Math.max(x(d.v),2);
    var nm=el("text",{class:"barlabel",x:mL-12,y:cy+4,"text-anchor":"end"});nm.textContent=d.label;svg.appendChild(nm);
    svg.appendChild(el("rect",{x:mL,y:yy,width:w,height:bh,rx:5,fill:colorFn(d)}));
    var vl=el("text",{class:"val",x:mL+w+8,y:cy+4});vl.textContent=labelFn(d);svg.appendChild(vl);
    var hit=el("rect",{x:0,y:cy-rowH/2,width:W,height:rowH,fill:"transparent"});hit.addEventListener("mousemove",function(e){sTT(e,tipFn(d));});hit.addEventListener("mouseleave",hTT);svg.appendChild(hit);});
  document.getElementById(id).appendChild(svg);
}

/* ---------- render ---------- */
function render(){
  var A=calc(state.who,state.from,state.to);
  var rangeLabel=(state.from===state.to)?M[state.from]:(M[state.from]+"–"+M[state.to]);
  LAST={A:A,range:rangeLabel};
  var pl=(state.prod==="__all__"?"":" · "+state.prod);
  document.getElementById("title").textContent = (state.who==="__team__" ? "Доход по экспертам — вся команда" : ("Доход эксперта: "+DATA.experts[state.who].name)) + pl;
  // KPI
  var total=A.kpiProc+A.kpiTelo;
  var yield_=A.ost>0?A.kpiProc/A.ost*100:0;
  // динамика: последний месяц диапазона vs предыдущий
  var last=A.monthP[state.to], prev=(state.to>0)?A.monthP[state.to-1]:null;
  var dPct=(prev&&prev>0)?((last-prev)/prev*100):null;
  var dCls=dPct==null?"flat":(dPct>=0?"up":"down");var dAr=dPct==null?"":(dPct>=0?"▲":"▼");
  var tiles=[
    {cls:"tile hero",k:"Процентный доход",v:fmt(A.kpiProc),u:"сом.",sub:rangeLabel},
    {cls:"tile tint-slate",k:"Возврат тела",v:fmt(A.kpiTelo),u:"сом.",sub:"основной долг"},
    {cls:"tile tint-teal",k:"Всего поступлений",v:fmt(total),u:"сом.",sub:"тело + процент"},
    {cls:"tile tint-ind",k:"Доходность к остатку",v:pct(yield_).replace("%",""),u:"%",sub:"%-доход / остаток ОД"},
    {cls:"tile",k:"Динамика ("+M[state.to]+")",v:(dPct==null?"—":pct(dPct)),u:"",sub:'<span class="delta '+dCls+'">'+dAr+' к пред. месяцу</span>'},
    {cls:"tile tint-amber",k:"Активных кредитов",v:A.active,u:"шт.",sub:"текущий портфель"},
    {cls:"tile",k:"Остаток ОД",v:f0(A.ost),u:"сом.",sub:"по активным"},
    {cls:"tile",k:(state.who==="__team__"?"Выдачи (действ.)":"Средневзв. ставка"),
     v:(state.who==="__team__"?f0(A.vyd):fmt(A.wrate,1)),u:(state.who==="__team__"?"сом.":"%"),
     sub:(state.who==="__team__"?A.issued+" действующих кред.":"бонус "+f0(A.bonus)+" сом.")}
  ];
  var box=document.getElementById("tiles");box.innerHTML="";
  tiles.forEach(function(t){var d=document.createElement("div");d.className=t.cls;d.innerHTML='<div class="k">'+t.k+'</div><div class="v num">'+t.v+' <small>'+t.u+'</small></div><div class="sub">'+(t.sub||"")+'</div>';box.appendChild(d);});

  // month chart
  monthChart(state.from,state.to,A.monthP,A.monthT);

  // panel A / B
  if(state.who==="__team__"){
    document.getElementById("hA").textContent="Доход по экспертам";
    document.getElementById("capA").textContent="Процентный доход за "+rangeLabel+", сомони.";
    var tot=A.kpiProc||1;
    var perExp=TEAM.map(function(e,i){var idx=DATA.experts.indexOf(e);var c=calc(String(idx),state.from,state.to);return {name:e.name,proc:c.kpiProc,ost:c.ost,vyd:c.vyd,active:c.active};}).sort(function(a,b){return b.proc-a.proc;});
    hbars("cA",perExp.map(function(e){return{label:e.name,v:e.proc,proc:e.proc,ost:e.ost};}),function(){return"var(--accent)";},
      function(d){return f0(d.proc)+" · "+(d.proc/tot*100).toFixed(1).replace(".",",")+"%";},
      function(d){return "<b>"+esc(d.label)+"</b><br>доход "+fmt(d.proc)+" сом.<br><span class=tk>доходность "+(d.ost>0?(d.proc/d.ost*100).toFixed(1).replace(".",","):"—")+"%</span>";});
    document.getElementById("hB").textContent="Доходность по экспертам";
    document.getElementById("capB").textContent="%-доход к остатку ОД за "+rangeLabel+".";
    hbars("cB",perExp.map(function(e){return{label:e.name,v:(e.ost>0?e.proc/e.ost*100:0),proc:e.proc,ost:e.ost};}).sort(function(a,b){return b.v-a.v;}),
      function(){return"var(--c3)";},function(d){return d.v.toFixed(1).replace(".",",")+"%";},
      function(d){return "<b>"+esc(d.label)+"</b><br>доходность "+d.v.toFixed(2).replace(".",",")+"%<br><span class=tk>доход "+f0(d.proc)+" / остаток "+f0(d.ost)+"</span>";});
  } else {
    var e=DATA.experts[state.who];var tot=A.kpiProc||1;
    document.getElementById("hA").textContent="Доход по клиентам";
    document.getElementById("capA").textContent="Процентный доход за "+rangeLabel+", сомони.";
    var cl=A.clients.filter(function(c){return c._rp>0;}).sort(function(a,b){return b._rp-a._rp;}).slice(0,12);
    hbars("cA",cl.map(function(c){return{label:c.name.replace(/ *\\(.*\\)/,""),v:c._rp,prod:c.prod,proc:c._rp,sum:c.sum};}),
      function(d){return pcol(d.prod);},function(d){return f0(d.proc)+" · "+(d.proc/tot*100).toFixed(1).replace(".",",")+"%";},
      function(d){return "<b>"+esc(d.label)+"</b><br><span class=tk>"+esc(d.prod)+" · "+f0(d.sum)+" сом.</span><br>доход "+fmt(d.proc)+" сом.";});
    document.getElementById("hB").textContent="Структура по продуктам";
    document.getElementById("capB").textContent="Доля продукта в %-доходе за "+rangeLabel+".";
    var keys=Object.keys(A.prod).sort(function(a,b){return A.prod[b].proc-A.prod[a].proc;});
    hbars("cB",keys.map(function(k){return{label:k,v:A.prod[k].proc,prod:k,cnt:A.prod[k].cnt};}),
      function(d){return pcol(d.prod);},function(d){return (d.v/tot*100).toFixed(1).replace(".",",")+"%";},
      function(d){return "<b>"+esc(d.label)+"</b><br><span class=tk>"+d.cnt+" кред.</span><br>доход "+fmt(d.v)+" сом.";});
  }

  // dynamics table
  var dh="<thead><tr><th>Месяц</th><th>Тело</th><th>Процент</th><th>Всего</th><th>Δ к пред.</th></tr></thead><tbody>";
  for(var i=state.from;i<=state.to;i++){
    var pv=(i>0)?A.monthP[i-1]:null;
    dh+="<tr><td class=name>"+M[i]+"</td><td class=num dim>"+fmt(A.monthT[i])+"</td><td class=num>"+fmt(A.monthP[i])+"</td><td class=num>"+fmt(A.monthT[i]+A.monthP[i])+"</td><td class=num>"+deltaBadge(A.monthP[i],pv)+"</td></tr>";
  }
  dh+="<tr class=tot><td>Итого</td><td class=num>"+fmt(A.kpiTelo)+"</td><td class=num>"+fmt(A.kpiProc)+"</td><td class=num>"+fmt(A.kpiProc+A.kpiTelo)+"</td><td></td></tr></tbody>";
  document.getElementById("tblDyn").innerHTML=dh;

  // detail table
  if(state.who==="__team__"){
    document.getElementById("hTbl").textContent="Сводка по экспертам";
    document.getElementById("capTbl").textContent="За "+rangeLabel+".";
    var rows=TEAM.map(function(e){var idx=DATA.experts.indexOf(e);var c=calc(String(idx),state.from,state.to);return {name:e.name,active:c.active,vyd:c.vyd,telo:c.kpiTelo,proc:c.kpiProc,ost:c.ost};}).sort(function(a,b){return b.proc-a.proc;});
    var tot=A.kpiProc||1;
    var h="<thead><tr><th>Эксперт</th><th>Активных</th><th>Выдачи</th><th>Остаток ОД</th><th>Тело</th><th>Процент</th><th>Доходность</th></tr></thead><tbody>";
    rows.forEach(function(r){h+="<tr><td class=name>"+esc(r.name)+"</td><td class=num>"+r.active+"</td><td class=num>"+f0(r.vyd)+"</td><td class=num dim>"+f0(r.ost)+"</td><td class=num dim>"+fmt(r.telo)+"</td><td class=num>"+fmt(r.proc)+"</td><td class=num>"+(r.ost>0?(r.proc/r.ost*100).toFixed(1).replace(".",",")+"%":"—")+"</td></tr>";});
    h+="<tr class=tot><td>ИТОГО</td><td class=num>"+A.active+"</td><td class=num>"+f0(A.vyd)+"</td><td class=num>"+f0(A.ost)+"</td><td class=num>"+fmt(A.kpiTelo)+"</td><td class=num>"+fmt(A.kpiProc)+"</td><td class=num>"+(A.ost>0?(A.kpiProc/A.ost*100).toFixed(1).replace(".",",")+"%":"—")+"</td></tr></tbody>";
    document.getElementById("tbl").innerHTML=h;
    var recon="<b>Сверка.</b> Всего дохода по файлу за весь период: %-доход <b>"+fmt(DATA.grandProc)+"</b> сом., тело <b>"+fmt(DATA.grandTelo)+"</b> сом. — должно совпасть с листом «Дашборд доходов». ";
    var um=DATA.unmatched?("Из них "+DATA.unmatched+" строк(и) без совпадения по ФИО+сумма+дата — в «Не распознан»."):"Все строки сопоставлены с экспертами.";
    document.getElementById("note").innerHTML=recon+um;
  } else {
    var e=DATA.experts[state.who];
    document.getElementById("hTbl").textContent="Детализация по клиентам";
    document.getElementById("capTbl").textContent="Клиенты эксперта с поступлениями за "+rangeLabel+".";
    var tot=A.kpiProc||1;
    var rows=A.clients.slice().sort(function(a,b){return b._rp-a._rp;});
    var sumTot=0;
    var h="<thead><tr><th>Клиент</th><th>Продукт</th><th>Сумма</th><th>Тело</th><th>Процент</th><th>Доля</th></tr></thead><tbody>";
    rows.forEach(function(c){sumTot+=(c.sum||0);var sh=c._rp>0?(c._rp/tot*100).toFixed(1).replace(".",",")+"%":"—";h+="<tr><td class=name>"+esc(c.name)+(c.status==="закр"?" <span class=dim>закрыт</span>":"")+"</td><td><span class=pill>"+esc(c.prod)+"</span></td><td class=num>"+f0(c.sum)+"</td><td class=num dim>"+(c._rt?fmt(c._rt):"—")+"</td><td class=num>"+(c._rp?fmt(c._rp):"—")+"</td><td class=num>"+sh+"</td></tr>";});
    // итог «Сумма» = сумма строк таблицы (а не выдачи по действующему портфелю — это другая выборка)
    h+="<tr class=tot><td>ИТОГО</td><td></td><td class=num>"+f0(sumTot)+"</td><td class=num>"+fmt(A.kpiTelo)+"</td><td class=num>"+fmt(A.kpiProc)+"</td><td class=num>100%</td></tr></tbody>";
    document.getElementById("tbl").innerHTML=h;
    document.getElementById("note").innerHTML="<b>Как считается.</b> Доход = проценты за выбранные месяцы. Крупное «тело» в отдельные месяцы — досрочные и полные погашения ОД (закрытие кредитов), не доход.";
  }
}

/* ---------- theme ---------- */
(function(){var btn=document.getElementById("themeBtn"),root=document.documentElement;var s=null;try{s=localStorage.getItem("dash-theme");}catch(e){}if(s)root.setAttribute("data-theme",s);
btn.addEventListener("click",function(){var cur=root.getAttribute("data-theme");var dark=cur==="dark"||(!cur&&matchMedia("(prefers-color-scheme:dark)").matches);var nx=dark?"light":"dark";root.setAttribute("data-theme",nx);try{localStorage.setItem("dash-theme",nx);}catch(e){}});})();
</script>
</body></html>
`;
