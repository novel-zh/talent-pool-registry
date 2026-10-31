/** Реестр специалистов. Персональные данные хранятся только в Google Таблицах. */
const ID_HEADER_ = 'Сотрудник_ID';
const SCHEMA_ = {
  main: {
    property: 'MAIN_SPREADSHEET_URL', tabProperty: 'MAIN_TAB_NAME', nameKey: 'name',
    fields: {
      number: ['Nr.', '№'], level: ['Stufe'], name: ['Name'], region: ['Region'],
      trackDE: ['Entwicklungstrack'], activity: ['Fachrichtung und Tätigkeit'],
      dec2025: ['Eingruppierung Stufe 30.12.2025'],
      jun2026: ['Eingruppierung Stufe 30.06.2026'],
      dec2026: ['Eingruppierung Stufe 30.12.2026'],
      birth: ['Geburtsdatum'], phone: ['Kontaktnummer'], email: ['E-Mail'],
      mainNote: ['Empfehlung/ Kommentare'], future: ['Zukunftige Entwiclung', 'Zukünftige Entwicklung'],
      track: ['Трек']
    }
  },
  bkr: {
    property: 'BKR_SPREADSHEET_URL', tabProperty: 'BKR_TAB_NAME', nameKey: 'analysisName',
    fields: {
      analysisNumber: ['№'], analysisName: ['ФИО участника'], analysisRegion: ['Регион'],
      goal: ['Цель участника (через 5 лет)'], recommendation: ['Рекомендация (РО/РП/ИБ)'],
      training: ['Обучения/проекты за 2026 г.'], plan2027: ['План на 2027 г. (рекомендации)']
    }
  }
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Реестр специалистов')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Запустить при настройке и для строк, добавленных напрямую в таблицы. */
function initializeEmployeeIds() {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const tables = [readTable_('main'), readTable_('bkr')];
    tables.forEach(t => {
      assertUniqueIds_(t, false);
      if (!t.sheet.getRange(1, 1, Math.max(t.sheet.getLastRow(), 1), Math.max(t.sheet.getLastColumn(), 1)).canEdit()) {
        throw new Error('Для назначения ID нужен доступ «Редактор» к обеим таблицам.');
      }
    });
    const counts = tables.map(t => countNames_(t));
    const known = new Map();
    tables.forEach(t => t.rows.forEach(r => {
      if (!r.id) return;
      const key = normalize_(r.values[SCHEMA_[t.kind].nameKey]);
      if (!known.has(key)) known.set(key, new Set());
      known.get(key).add(r.id);
    }));
    const newIds = new Map();
    const assignments = [];
    tables.forEach((t, ti) => t.rows.forEach(r => {
      if (r.id) return;
      const key = normalize_(r.values[SCHEMA_[t.kind].nameKey]);
      if (counts[0].get(key) > 1 || counts[1].get(key) > 1 || (known.has(key) && known.get(key).size > 1)) {
        throw new Error('Неоднозначное ФИО в строке ' + r.row + ' листа «' + t.sheet.getName() + '». Назначьте различным людям разные Сотрудник_ID вручную; одному человеку в двух таблицах — одинаковый ID.');
      }
      const existing = known.has(key) ? Array.from(known.get(key))[0] : null;
      if (!existing && !newIds.has(key)) newIds.set(key, 'EMP-' + Utilities.getUuid());
      assignments.push({table: ti, row: r.row, id: existing || newIds.get(key)});
    }));
    // Все проверки закончены до первой записи. Существующие ID не меняются.
    tables.forEach((t, ti) => {
      if (t.idCol < 0) {
        t.idCol = t.sheet.getLastColumn();
        if (t.idCol >= t.sheet.getMaxColumns()) t.sheet.insertColumnAfter(t.sheet.getMaxColumns());
        t.sheet.getRange(t.headerRow, t.idCol + 1).setValue(ID_HEADER_);
      }
      assignments.filter(a => a.table === ti).forEach(a => t.sheet.getRange(a.row, t.idCol + 1).setValue(a.id));
    });
    SpreadsheetApp.flush();
    const result = loadPersonnel();
    console.log('Настройка завершена. Людей: ' + result.people.length + '. Новых ID в строках: ' + assignments.length + '.');
    return {people: result.people.length, assigned: assignments.length};
  } finally { lock.releaseLock(); }
}

/** Проверка настроек из редактора; в журнал выводится только число записей. */
function checkConnection() {
  const result = loadPersonnel();
  console.log('Подключение работает. Людей: ' + result.people.length + '.');
  return {people: result.people.length};
}

/** Вызывается интерфейсом от имени вошедшего пользователя. Без общего кэша. */
function loadPersonnel() {
  const main = readTable_('main');
  const bkr = readTable_('bkr');
  assertUniqueIds_(main, true);
  assertUniqueIds_(bkr, true);
  const mainEditable = tableEditable_(main);
  const bkrEditable = tableEditable_(bkr);
  const merged = new Map();
  function initial(id) {
    const p = {id: id, hasBase: false, hasAnalysis: false, canEditMain: false, canEditBkr: false};
    Object.keys(SCHEMA_.main.fields).concat(Object.keys(SCHEMA_.bkr.fields)).forEach(k => p[k] = '');
    return p;
  }
  main.rows.forEach(r => {
    const p = initial(r.id);
    Object.assign(p, r.values, {hasBase: true, canEditMain: mainEditable, revisionMain: revision_(r)});
    merged.set(r.id, p);
  });
  bkr.rows.forEach(r => {
    const p = merged.get(r.id);
    if (!p) return; // Состав списка и ФИО определяются только основной таблицей.
    const values = Object.assign({}, r.values);
    delete values.analysisName;
    Object.assign(p, values, {hasAnalysis: true, canEditBkr: bkrEditable, revisionBkr: revision_(r)});
  });
  const people = Array.from(merged.values());
  people.forEach(p => delete p.analysisName);
  return {people: people, canCreate: mainEditable && bkrEditable,
    unlinkedAnalysisCount: bkr.rows.filter(r => !merged.has(r.id)).length,
    loadedAt: new Date().toISOString()};
}

function tableEditable_(table) {
  return table.sheet.getRange(table.headerRow, 1, Math.max(1, table.sheet.getLastRow() - table.headerRow + 1), table.width).canEdit();
}

/** Сохраняется один раздел/источник за раз, по постоянному ID, с проверкой версии. */
function savePersonnel(request) {
  if (!request || !['main', 'bkr'].includes(request.source) || typeof request.id !== 'string' || !request.changes || typeof request.changes !== 'object') {
    throw new Error('Некорректный запрос сохранения. Обновите страницу.');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    // Основная таблица — точка проверки допуска ко всему приложению, в том числе при записи БКР.
    const main = readTable_('main');
    const table = request.source === 'main' ? main : readTable_('bkr');
    assertUniqueIds_(table, true);
    const record = table.rows.find(r => r.id === request.id);
    if (!record) throw new Error('Запись удалена или её ID изменился. Нажмите «Обновить данные».');
    if (request.revision !== revision_(record)) throw new Error('Запись уже изменена в таблице или другим пользователем. Обновите данные перед редактированием.');
    const changes = prepareChanges_(table, record, request.changes);
    writeChanges_(changes);
    return {saved: changes.length};
  } finally { lock.releaseLock(); }
}

/** Сохранить единую форму; оба раздела проверяются до первой записи. */
function updatePersonnel(request) {
  if (!request || typeof request.id !== 'string' || !request.sections || typeof request.sections !== 'object') {
    throw new Error('Некорректный запрос сохранения. Обновите страницу.');
  }
  const keys = Object.keys(request.sections);
  if (!keys.length || keys.some(k => !['main', 'bkr'].includes(k))) throw new Error('Некорректные разделы записи.');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const tables = {main: readTable_('main'), bkr: readTable_('bkr')};
    Object.values(tables).forEach(t => assertUniqueIds_(t, true));
    if (!tables.main.rows.some(r => r.id === request.id)) throw new Error('Участник отсутствует в основном списке. Обновите данные.');
    let changes = [];
    keys.forEach(kind => {
      const section = request.sections[kind], table = tables[kind];
      if (!section || !section.changes || typeof section.changes !== 'object') throw new Error('Некорректные поля записи.');
      const record = table.rows.find(r => r.id === request.id);
      if (!record) throw new Error('Запись раздела удалена. Обновите данные.');
      if (section.revision !== revision_(record)) throw new Error('Запись уже изменена в таблице или другим пользователем. Обновите данные перед редактированием.');
      changes = changes.concat(prepareChanges_(table, record, section.changes));
    });
    writeChanges_(changes);
    return {saved: changes.length};
  } finally { lock.releaseLock(); }
}

function typedValue_(key, text, table) {
  if (text.length > 30000) throw new Error('Слишком длинное значение поля.');
  if (key === 'birth' && text) {
    if (!/^\d{2}\.\d{2}\.\d{4}$/.test(text)) throw new Error('Дата рождения: используйте формат ДД.ММ.ГГГГ.');
    let date;
    try { date = Utilities.parseDate(text, table.timeZone, 'dd.MM.yyyy'); } catch (e) { throw new Error('Некорректная дата рождения.'); }
    if (Utilities.formatDate(date, table.timeZone, 'dd.MM.yyyy') !== text) throw new Error('Некорректная дата рождения.');
    return date;
  }
  if (['number', 'analysisNumber', 'level'].includes(key) && /^\d+(?:\.\d+)?$/.test(text)) return Number(text);
  return text.startsWith('=') ? "'" + text : text;
}

function prepareChanges_(table, record, input) {
    const changes = [];
    Object.keys(input).forEach(key => {
      if (key === 'analysisName' || !Object.prototype.hasOwnProperty.call(SCHEMA_[table.kind].fields, key) || table.cols[key] === undefined) {
        throw new Error('Поле недоступно для изменения: ' + key);
      }
      const text = String(input[key] == null ? '' : input[key]).trim();
      if (text.length > 30000) throw new Error('Слишком длинное значение поля.');
      if (text === record.values[key]) return;
      if (key === 'name' && !text) throw new Error('ФИО не может быть пустым.');
      const cell = table.sheet.getRange(record.row, table.cols[key] + 1);
      if (!cell.canEdit()) throw new Error('Нет права редактировать этот раздел или поле защищено.');
      if (cell.getFormula()) throw new Error('В этом поле находится формула. Измените её непосредственно в таблице.');
      changes.push({cell: cell, value: typedValue_(key, text, table), oldValue: cell.getValue(), oldFormula: cell.getFormula()});
    });
    return changes;
}

function writeChanges_(changes) {
    const applied = [];
    try {
      changes.forEach(c => { applied.push(c); c.cell.setValue(c.value); });
      SpreadsheetApp.flush();
    } catch (e) {
      let restored = true;
      applied.reverse().forEach(c => {
        try {
          const oldLiteral = typeof c.oldValue === 'string' && c.oldValue.startsWith('=') ? "'" + c.oldValue : c.oldValue;
          c.cell.setValue(c.oldFormula || oldLiteral);
        } catch (restoreError) { restored = false; }
      });
      try { SpreadsheetApp.flush(); } catch (restoreError) { restored = false; }
      throw new Error(restored ? 'Не удалось сохранить. Предыдущие значения восстановлены. Обновите страницу и проверьте права.' : 'Сохранение прервалось. Часть полей могла измениться. Проверьте запись в исходной таблице.');
    }
}

/** Создать участника в обоих источниках с одним автоматически назначенным ID. */
function createPersonnel(request) {
  if (!request || !/^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(request.requestId || '') || !request.values || typeof request.values !== 'object') {
    throw new Error('Некорректный запрос добавления. Откройте форму заново.');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const tables = [readTable_('main'), readTable_('bkr')];
    tables.forEach(t => {
      assertUniqueIds_(t, true);
      if (!tableEditable_(t)) throw new Error('Для добавления участника нужен доступ «Редактор» к обеим таблицам.');
    });
    // Один запрос формы всегда получает один ID: повтор после потери ответа не создаёт дубль.
    const id = 'EMP-' + request.requestId.toLowerCase();
    const existing = tables.map(t => t.rows.find(r => r.id === id));
    if (existing.every(Boolean)) return {id: id, created: false};
    if (existing.some(Boolean)) throw new Error('Обнаружена незавершённая запись в одной таблице. Проверьте обе таблицы для ID ' + id + ' перед повторным добавлением.');
    const values = {};
    const allowed = Object.keys(SCHEMA_.main.fields).concat(Object.keys(SCHEMA_.bkr.fields)).filter(k => !['analysisName', 'number', 'analysisNumber'].includes(k));
    Object.keys(request.values).forEach(key => {
      if (!allowed.includes(key)) throw new Error('Поле недоступно для добавления: ' + key);
      values[key] = String(request.values[key] == null ? '' : request.values[key]).trim();
      if (values[key].length > 30000) throw new Error('Слишком длинное значение поля.');
    });
    if (!values.name) throw new Error('Укажите ФИО участника.');
    values.analysisName = values.name;
    if (!values.analysisRegion) values.analysisRegion = values.region || '';
    const plans = tables.map(t => {
      const row = new Array(t.width).fill('');
      Object.keys(SCHEMA_[t.kind].fields).forEach(key => {
        if (t.cols[key] === undefined) {
          if (values[key]) throw new Error('Не найден столбец для поля ' + key + '. Проверьте заголовки таблицы.');
          return;
        }
        const value = ['number', 'analysisNumber'].includes(key)
          ? String(t.rows.reduce((max, r) => Math.max(max, /^\d+$/.test(r.values[key] || '') ? Number(r.values[key]) : 0), 0) + 1)
          : values[key] || '';
        row[t.cols[key]] = typedValue_(key, value, t);
      });
      row[t.idCol] = id;
      return {table: t, rowNumber: t.sheet.getLastRow() + 1, values: row};
    });
    // Права и значения проверены для обоих источников до записи участника.
    plans.forEach(p => {
      const sheet = p.table.sheet;
      if (p.rowNumber > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), p.rowNumber - sheet.getMaxRows());
      p.range = sheet.getRange(p.rowNumber, 1, 1, p.table.width);
      if (!p.range.canEdit()) throw new Error('Новая строка защищена. Проверьте права на обе таблицы.');
      if (p.range.getValues()[0].some(v => v !== '') || p.range.getFormulas()[0].some(Boolean)) throw new Error('Строка для добавления уже занята. Повторите попытку.');
    });
    const attempted = [];
    try {
      plans.forEach(p => { attempted.push(p); p.range.setValues([p.values]); });
      SpreadsheetApp.flush();
    } catch (error) {
      let restored = true;
      attempted.reverse().forEach(p => {
        try {
          const current = p.range.getValues()[0];
          if (current[p.table.idCol] === id) p.range.clearContent();
          else if (current.some(v => v !== '')) restored = false;
        } catch (restoreError) { restored = false; }
      });
      try { SpreadsheetApp.flush(); } catch (restoreError) { restored = false; }
      throw new Error(restored ? 'Не удалось добавить участника. Записанные данные отменены. Проверьте подключение и повторите попытку.' : 'Добавление прервалось. Проверьте обе таблицы для ID ' + id + ': часть данных могла сохраниться.');
    }
    return {id: id, created: true};
  } finally { lock.releaseLock(); }
}

function normalize_(value) { return String(value == null ? '' : value).normalize('NFKC').replace(/\s+/g, ' ').trim().toLowerCase(); }
function headerKey_(value) { return normalize_(value).replace(/\s/g, ''); }
function countNames_(table) {
  const map = new Map();
  table.rows.forEach(r => { const key = normalize_(r.values[SCHEMA_[table.kind].nameKey]); map.set(key, (map.get(key) || 0) + 1); });
  return map;
}
function readTable_(kind) {
  const schema = SCHEMA_[kind];
  const props = PropertiesService.getScriptProperties();
  const configured = (props.getProperty(schema.property) || '').trim();
  const id = (configured.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/) || [])[1] || (/^[a-zA-Z0-9_-]{15,}$/.test(configured) ? configured : '');
  if (!id) throw new Error('В настройках проекта добавьте свойство ' + schema.property + ' со ссылкой на Google Таблицу.');
  let book;
  try { book = SpreadsheetApp.openById(id); } catch (e) {
    throw new Error('Не удалось открыть ' + schema.property + '. Проверьте Google-аккаунт, доступ и что файл сохранён в формате Google Таблиц, а не Excel.');
  }
  const tabName = (props.getProperty(schema.tabProperty) || '').trim();
  const sheets = tabName ? [book.getSheetByName(tabName)].filter(Boolean) : book.getSheets();
  if (!sheets.length) throw new Error('Лист «' + tabName + '» не найден. Проверьте ' + schema.tabProperty + '.');
  const candidates = [];
  sheets.forEach(sheet => {
    const width = sheet.getLastColumn(), height = sheet.getLastRow();
    if (!width || !height) return;
    const headers = sheet.getRange(1, 1, Math.min(height, 10), width).getDisplayValues();
    headers.forEach((row, ri) => {
      const cols = {};
      Object.keys(schema.fields).forEach(key => {
        const aliases = schema.fields[key].map(headerKey_);
        const positions = row.map((v, ci) => aliases.includes(headerKey_(v)) ? ci : -1).filter(ci => ci >= 0);
        if (positions.length > 1) throw new Error('Повторяется заголовок поля ' + key + ' на листе «' + sheet.getName() + '».');
        if (positions.length) cols[key] = positions[0];
      });
      const regionKey = kind === 'main' ? 'region' : 'analysisRegion';
      if (cols[schema.nameKey] === undefined || cols[regionKey] === undefined) return;
      const idCol = row.findIndex(v => headerKey_(v) === headerKey_(ID_HEADER_));
      candidates.push({sheet: sheet, headerRow: ri + 1, cols: cols, idCol: idCol, width: width});
    });
  });
  if (candidates.length !== 1) throw new Error('Не удалось однозначно найти заголовки ' + kind + '. Проверьте названия столбцов или задайте ' + schema.tabProperty + '.');
  const t = candidates[0];
  Object.assign(t, {kind: kind, timeZone: book.getSpreadsheetTimeZone(), rows: []});
  const n = t.sheet.getLastRow() - t.headerRow;
  if (n <= 0) return t;
  const range = t.sheet.getRange(t.headerRow + 1, 1, n, t.width);
  const display = range.getDisplayValues(), raw = range.getValues(), formulas = range.getFormulas();
  display.forEach((row, index) => {
    if (!String(row[t.cols[schema.nameKey]] || '').trim()) return;
    const values = {}, rowFormulas = {};
    Object.keys(schema.fields).forEach(key => {
      const col = t.cols[key];
      values[key] = col === undefined ? '' : String(row[col] == null ? '' : row[col]).trim();
      if (key === 'birth' && col !== undefined && Object.prototype.toString.call(raw[index][col]) === '[object Date]') {
        values[key] = Utilities.formatDate(raw[index][col], t.timeZone, 'dd.MM.yyyy');
      }
      rowFormulas[key] = col === undefined ? '' : formulas[index][col];
    });
    t.rows.push({row: t.headerRow + 1 + index, id: t.idCol < 0 ? '' : String(row[t.idCol] || '').trim(), values: values, formulas: rowFormulas});
  });
  return t;
}
function assertUniqueIds_(table, required) {
  if (required && table.idCol < 0) throw new Error('Сначала запустите initializeEmployeeIds из редактора Apps Script.');
  const seen = new Set();
  table.rows.forEach(r => {
    if (!r.id) { if (required) throw new Error('Нет Сотрудник_ID в строке ' + r.row + ' листа «' + table.sheet.getName() + '». Администратору нужно запустить initializeEmployeeIds.'); return; }
    if (seen.has(r.id)) throw new Error('Повторяется Сотрудник_ID на листе «' + table.sheet.getName() + '». У разных людей должны быть разные ID.');
    seen.add(r.id);
  });
}
function revision_(record) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify({id: record.id, values: record.values, formulas: record.formulas}), Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(bytes);
}
