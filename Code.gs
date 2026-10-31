/** Adaptive registry; each accessing user owns their connection settings. */
const CONNECTION_KEY_ = 'ADAPTIVE_REGISTRY_CONNECTION_V1';
const LIMITS_ = {rows: 5000, columns: 100, cells: 100000, characters: 2000000};

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Реестр кадрового резерва').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function include_(name) {
  if (!['Model', 'Client'].includes(name)) throw new Error('Неизвестный компонент.');
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}
function getInitialState() { const c = connection_(); return c ? readDataset_(c) : {connected: false}; }
function refreshDataset() { return readDataset_(requireConnection_()); }
function disconnectSpreadsheet() {
  assertUserContext_();
  PropertiesService.getUserProperties().deleteProperty(CONNECTION_KEY_);
  return {connected: false};
}
function connectSpreadsheet(request) {
  assertUserContext_();
  const source = parseSource_(request && request.url), book = openBook_(source.id), sheets = book.getSheets();
  const sheet = source.gid === null ? sheets.find(s => s.getLastRow() && s.getLastColumn()) || sheets[0]
    : sheets.find(s => s.getSheetId() === source.gid);
  if (!sheet) throw new Error('Лист из ссылки не найден. Скопируйте ссылку на нужный лист ещё раз.');
  const c = {spreadsheetId: source.id, sheetId: sheet.getSheetId(), headerRow: 0};
  const result = readDataset_(c, book);
  PropertiesService.getUserProperties().setProperty(CONNECTION_KEY_, JSON.stringify(c));
  return result;
}
function selectSheet(request) {
  const current = requireConnection_(), id = Number(request && request.sheetId);
  const header = request && request.headerRow !== '' && request.headerRow != null ? Number(request.headerRow) : 0;
  if (!Number.isInteger(id) || id < 0 || !Number.isInteger(header) || header < 0 || header > 30) {
    throw new Error('Проверьте лист и строку заголовков: 1–30 или «Авто».');
  }
  const next = {spreadsheetId: current.spreadsheetId, sheetId: id, headerRow: header};
  const result = readDataset_(next);
  PropertiesService.getUserProperties().setProperty(CONNECTION_KEY_, JSON.stringify(next));
  return result;
}
function connection_() {
  assertUserContext_();
  const saved = PropertiesService.getUserProperties().getProperty(CONNECTION_KEY_);
  if (!saved) return null;
  try {
    const c = JSON.parse(saved);
    if (!/^[a-zA-Z0-9_-]{15,}$/.test(c.spreadsheetId) || !Number.isInteger(c.sheetId) ||
        !Number.isInteger(c.headerRow) || c.headerRow < 0 || c.headerRow > 30) throw new Error();
    return c;
  } catch (e) { throw new Error('Подключение повреждено. Подключите таблицу заново.'); }
}
function requireConnection_() { const c = connection_(); if (!c) throw new Error('Сначала подключите Google Таблицу.'); return c; }
function assertUserContext_() {
  const active = Session.getActiveUser().getEmail(), effective = Session.getEffectiveUser().getEmail();
  if (!active || !effective || active.toLowerCase() !== effective.toLowerCase()) {
    throw new Error('Приложение должно работать от имени открывающего его пользователя. Проверьте параметр «Выполнять от имени» в развёртывании Apps Script и войдите в Google-аккаунт.');
  }
}
function parseSource_(value) {
  const text = String(value || '').trim();
  if (/^[a-zA-Z0-9_-]{15,}$/.test(text)) return {id: text, gid: null};
  const match = text.match(/^https:\/\/docs\.google\.com\/spreadsheets\/(?:u\/\d+\/)?d\/([a-zA-Z0-9_-]{15,})(?:[/?#].*)?$/i);
  if (!match) throw new Error('Вставьте ссылку https://docs.google.com/spreadsheets/d/… или ID Google Таблицы.');
  const gid = (text.match(/[?#&]gid=(\d+)(?:[&#]|$)/) || [])[1];
  return {id: match[1], gid: gid === undefined ? null : Number(gid)};
}
function openBook_(id) {
  try { return SpreadsheetApp.openById(id); }
  catch (e) { throw new Error('Не удалось открыть таблицу. Проверьте ссылку, Google-аккаунт и доступ. Нужен формат Google Таблиц, а не Excel.'); }
}
function normalize_(v) { return String(v == null ? '' : v).normalize('NFKC').replace(/\s+/g, ' ').trim().toLocaleLowerCase(); }
function isEmpty_(v) { return v === '' || v === null || v === undefined; }
function isDate_(v) { return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime()); }
function findHeader_(display, raw) {
  const first = display.findIndex(row => row.some(v => String(v).trim()));
  if (first < 0) throw new Error('Лист пуст. Добавьте названия столбцов.');
  // Prefer the first plausible multi-column heading over a merged document title.
  for (let i = first; i < display.length; i++) {
    if (display[i].filter(v => String(v).trim()).length >= 2) return i + 1;
  }
  return first + 1;
}
function inferColumn_(label, values, display) {
  const nonempty = values.filter(v => !isEmpty_(v));
  const identifier = /(^|[\s_])(id|ид|иин|iin|код|номер|№|телефон|phone|контакт|email|e-mail|nr\.?)([\s_]|$)/i.test(label);
  let type = 'text';
  if (!identifier && nonempty.length) {
    if (nonempty.every(isDate_)) type = 'date';
    else if (nonempty.every(v => typeof v === 'boolean')) type = 'boolean';
    else if (nonempty.every(v => typeof v === 'number' && Number.isFinite(v))) type = 'number';
  }
  const options = Array.from(new Set(display.filter(v => String(v).trim()).map(String)));
  const name = /(^|\s)(фио|фамилия|имя|name|surname|full name)(\s|$)/i.test(label);
  const filter = type !== 'text' ? type : options.length && options.length <= 30 && options.every(v => v.length <= 100) && !name ? 'category' : 'text';
  return {type: type, filter: filter, options: filter === 'category' ? options.sort((a,b) => a.localeCompare(b, 'ru')) : []};
}
function digest_(value) {
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(value), Utilities.Charset.UTF_8));
}
function rowRevision_(raw, formulas, schema) { return digest_([schema, raw.map(v => isDate_(v) ? {date: v.toISOString()} : v), formulas]); }
function columnLetter_(n) { let text = ''; while (n) { n--; text = String.fromCharCode(65 + n % 26) + text; n = Math.floor(n / 26); } return text; }

function readDataset_(c, suppliedBook) {
  const book = suppliedBook || openBook_(c.spreadsheetId), sheet = book.getSheets().find(s => s.getSheetId() === c.sheetId);
  if (!sheet) throw new Error('Подключённый лист удалён. Подключите таблицу заново.');
  const width = sheet.getLastColumn(), height = sheet.getLastRow();
  if (!width || !height) throw new Error('Лист пуст. Добавьте строку с названиями столбцов.');
  if (width > LIMITS_.columns) throw new Error('Больше ' + LIMITS_.columns + ' столбцов. Подготовьте отдельный лист с нужными данными.');
  const top = sheet.getRange(1, 1, Math.min(20, height), width);
  const headerRow = c.headerRow || findHeader_(top.getDisplayValues(), top.getValues());
  if (headerRow > height) throw new Error('Строка заголовков находится за пределами данных.');
  const headers = sheet.getRange(headerRow, 1, 1, width).getDisplayValues()[0];
  const count = Math.min(Math.max(0, height - headerRow), LIMITS_.rows, Math.floor(LIMITS_.cells / width));
  const range = count ? sheet.getRange(headerRow + 1, 1, count, width) : null;
  const sourceRaw = range ? range.getValues() : [], sourceDisplay = range ? range.getDisplayValues() : [], sourceFormulas = range ? range.getFormulas() : [];
  const raw = [], display = [], formulas = [], rowNumbers = [];
  let characters = 0, payloadLimited = false;
  for (let i = 0; i < sourceDisplay.length; i++) {
    const size = sourceDisplay[i].reduce((n,v) => n + String(v).length, 0);
    if (characters + size > LIMITS_.characters) { payloadLimited = true; break; }
    characters += size;
    if (sourceRaw[i].every(isEmpty_) && sourceDisplay[i].every(isEmpty_)) continue;
    raw.push(sourceRaw[i]); display.push(sourceDisplay[i].map(String)); formulas.push(sourceFormulas[i]); rowNumbers.push(headerRow + 1 + i);
  }
  const timezone = book.getSpreadsheetTimeZone(), names = new Map(), columns = [];
  for (let i = 0; i < width; i++) {
    const header = String(headers[i] || '').trim();
    if (!header && !raw.some(row => !isEmpty_(row[i])) && !display.some(row => row[i])) continue;
    const base = header || 'Столбец ' + columnLetter_(i + 1), number = (names.get(normalize_(base)) || 0) + 1;
    names.set(normalize_(base), number);
    // Suffixes are based on physical column position; keys never use user-controlled headers.
    let label = base + (number > 1 ? ' (' + number + ')' : '');
    while (columns.some(col => col.label === label)) label += ' · ' + columnLetter_(i + 1);
    columns.push(Object.assign({key: 'c' + i, index: i, label: label, sourceHeader: header}, inferColumn_(label, raw.map(r => r[i]), display.map(r => r[i]))));
  }
  if (!columns.length) throw new Error('Столбцы не найдены. Выберите строку заголовков вручную.');
  const schemaVersion = digest_([sheet.getSheetId(), headerRow, headers, columns.map(col => [col.key, col.type])]);
  const canEdit = sheet.getRange(headerRow, 1, Math.max(1, count + 1), width).canEdit();
  const rows = raw.map((row, ri) => {
    const values = {}, typed = {}, formulaFields = [];
    columns.forEach(col => {
      const value = row[col.index]; values[col.key] = display[ri][col.index];
      if (col.type === 'date') typed[col.key] = isDate_(value) ? Utilities.formatDate(value, timezone, 'yyyy-MM-dd') : '';
      else if (col.type === 'number' || col.type === 'boolean') typed[col.key] = isEmpty_(value) ? '' : value;
      if (formulas[ri][col.index]) formulaFields.push(col.key);
    });
    return {rowNumber: rowNumbers[ri], revision: rowRevision_(row, formulas[ri], schemaVersion), values: values, typed: typed, formulaFields: formulaFields};
  });
  const title = columns.find(col => /^(фио|фамилия|имя|name|full name|участник|сотрудник)$/i.test(col.sourceHeader)) || columns[0];
  return {connected: true, spreadsheetName: book.getName(), sourceUrl: book.getUrl() + '#gid=' + sheet.getSheetId(),
    sheets: book.getSheets().map(s => ({id: s.getSheetId(), name: s.getName(), empty: !s.getLastRow() || !s.getLastColumn()})),
    sheetId: sheet.getSheetId(), sheetName: sheet.getName(), headerRow: headerRow, headerMode: c.headerRow,
    columns: columns, rows: rows, titleKey: title.key, schemaVersion: schemaVersion, canEdit: canEdit, canAdd: canEdit,
    physicalDataRows: Math.max(0, height - headerRow), truncated: count < height - headerRow || payloadLimited, loadedAt: new Date().toISOString()};
}
function typedInput_(col, value, timezone) {
  const text = String(value == null ? '' : value);
  if (text.length > 30000) throw new Error('Поле «' + col.label + '» длиннее 30 000 символов.');
  if (text === '') return '';
  if (col.type === 'number') {
    if (!text.trim() || !Number.isFinite(Number(text))) throw new Error('Поле «' + col.label + '»: укажите число.');
    return Number(text);
  }
  if (col.type === 'boolean') {
    if (value === true || text === 'true') return true;
    if (value === false || text === 'false') return false;
    throw new Error('Поле «' + col.label + '»: выберите Да или Нет.');
  }
  if (col.type === 'date') {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) throw new Error('Поле «' + col.label + '»: выберите дату.');
    let date;
    try { date = Utilities.parseDate(text, timezone, 'yyyy-MM-dd'); } catch (e) { throw new Error('Некорректная дата.'); }
    if (Utilities.formatDate(date, timezone, 'yyyy-MM-dd') !== text) throw new Error('Некорректная дата.');
    return date;
  }
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}
function saveRow(request) {
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const c = requireConnection_(), book = openBook_(c.spreadsheetId), data = readDataset_(c, book);
    if (!request || request.schemaVersion !== data.schemaVersion || !Number.isInteger(request.rowNumber)) throw new Error('Структура таблицы изменилась. Обновите данные.');
    const record = data.rows.find(row => row.rowNumber === request.rowNumber);
    if (!record || record.revision !== request.revision) throw new Error('Строка изменена, перемещена или удалена. Обновите данные перед сохранением.');
    if (!request.changes || typeof request.changes !== 'object' || Array.isArray(request.changes)) throw new Error('Некорректные поля записи.');
    const sheet = book.getSheets().find(s => s.getSheetId() === c.sheetId);
    const changes = Object.keys(request.changes).map(key => {
      const col = data.columns.find(column => column.key === key);
      if (!col) throw new Error('Неизвестный столбец.');
      const cell = sheet.getRange(request.rowNumber, col.index + 1);
      if (!cell.canEdit()) throw new Error('Нет доступа к полю «' + col.label + '».');
      if (cell.getFormula()) throw new Error('Поле «' + col.label + '» вычисляется формулой.');
      return {cell: cell, original: cell.getValue(), value: typedInput_(col, request.changes[key], book.getSpreadsheetTimeZone())};
    });
    const applied = [];
    try { changes.forEach(change => { applied.push(change); change.cell.setValue(change.value); }); SpreadsheetApp.flush(); }
    catch (error) {
      let incomplete = false;
      applied.reverse().forEach(change => { try { const v = change.original; change.cell.setValue(typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : v); } catch (e) { incomplete = true; } });
      throw new Error(incomplete ? 'Сбой записи. Откат не завершён: проверьте строку в Google Таблице.' : 'Не удалось сохранить. Исходные значения восстановлены.');
    }
    return {saved: changes.length};
  } finally { lock.releaseLock(); }
}
function appendRow(request) {
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const c = requireConnection_(), book = openBook_(c.spreadsheetId), data = readDataset_(c, book);
    if (!request || request.schemaVersion !== data.schemaVersion) throw new Error('Структура таблицы изменилась. Обновите данные.');
    if (!data.canAdd) throw new Error('Для добавления нужен доступ к редактированию листа.');
    if (!/^[a-zA-Z0-9-]{16,80}$/.test(String(request.requestId || ''))) throw new Error('Некорректный номер запроса.');
    const props = PropertiesService.getUserProperties(), key = 'APPEND_V1_' + digest_([c.spreadsheetId, c.sheetId, request.requestId]);
    const saved = props.getProperty(key);
    if (saved) { const old = JSON.parse(saved); if (old.state !== 'done') throw new Error('Предыдущая попытка не завершена. Проверьте строку ' + old.row + ' в таблице.'); return {created: false, rowNumber: old.row}; }
    if (!request.values || typeof request.values !== 'object' || Array.isArray(request.values)) throw new Error('Некорректные поля записи.');
    const keys = Object.keys(request.values);
    if (keys.some(key => !data.columns.some(col => col.key === key))) throw new Error('Неизвестный столбец.');
    const sheet = book.getSheets().find(s => s.getSheetId() === c.sheetId), rowNumber = sheet.getLastRow() + 1;
    const values = Array(sheet.getLastColumn()).fill('');
    keys.forEach(key => { const col = data.columns.find(column => column.key === key); values[col.index] = typedInput_(col, request.values[key], book.getSpreadsheetTimeZone()); });
    if (values.every(isEmpty_)) throw new Error('Заполните хотя бы одно поле.');
    if (rowNumber > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 1);
    const range = sheet.getRange(rowNumber, 1, 1, values.length);
    if (!range.canEdit()) throw new Error('Нет доступа к новой строке.');
    props.setProperty(key, JSON.stringify({row: rowNumber, state: 'pending', time: Date.now()}));
    try { range.setValues([values]); SpreadsheetApp.flush(); }
    catch (error) {
      try { range.clearContent(); props.deleteProperty(key); }
      catch (rollback) { throw new Error('Сбой добавления. Проверьте строку ' + rowNumber + ' перед повторной попыткой.'); }
      throw new Error('Добавление отменено. Повторите сохранение.');
    }
    props.setProperty(key, JSON.stringify({row: rowNumber, state: 'done', time: Date.now()}));
    const history = Object.entries(props.getProperties()).filter(([k]) => k.startsWith('APPEND_V1_')).map(([k,v]) => [k, JSON.parse(v)]).filter(([,v]) => v.state === 'done').sort((a,b) => b[1].time - a[1].time);
    history.slice(100).forEach(([k]) => props.deleteProperty(k));
    return {created: true, rowNumber: rowNumber};
  } finally { lock.releaseLock(); }
}
