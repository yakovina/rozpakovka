// Приймач відповідей для сторінок «Анкета» (аркуш «Відповіді») і «Кейси» (аркуш «Кейси»).
// Кожна сторінка каже, в який аркуш писати (поле sheet). Без нього — «Відповіді», як і раніше.
// Будова аркуша: id · службові поля (імʼя — або хто + кейс) · оновлено · дані · далі по колонці на питання.

const DEFAULT_SHEET = 'Відповіді';

function sheet_(name, create) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh && create) sh = ss.insertSheet(name);
  return sh;
}
function sheetName_(p) {
  const s = p && p.sheet != null ? String(p.sheet).trim() : '';
  return s || DEFAULT_SHEET;
}
function fixed_(sh) {   // скільки службових колонок перед питаннями: шукаємо «дані» в заголовку
  const lastCol = sh.getLastColumn();
  if (!lastCol) return 0;
  const head = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const i = head.indexOf('дані');
  return i >= 0 ? i + 1 : 4;
}
function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function iso_(v) { try { return v ? new Date(v).toISOString() : ''; } catch (e) { return ''; } }
const EMPTY_MARKS = { '(не знаю)': 1, '(пропущено)': 1 };

// GET ?id=…            — чи є вже відповіді під цим ключем (інший пристрій, інший браузер)
// GET ?list=1&owner=…  — список кейсів цієї людини (тільки для аркушів, що не «Відповіді»)
function doGet(e) {
  const p = (e && e.parameter) || {};
  const name = sheetName_(p);
  if (p.list) return list_(name, p.owner);
  const id = p.id;
  if (!id) return ContentService.createTextOutput('ok');
  const sh = sheet_(name, false);
  if (!sh) return json_({ found: false });
  const row = findRow_(sh, id);
  if (!row) return json_({ found: false });
  const FIXED = fixed_(sh), lastCol = sh.getLastColumn();
  const vals = sh.getRange(row, 1, 1, lastCol).getValues()[0];
  const answers = {};
  for (let c = FIXED; c < lastCol; c++) {
    const v = vals[c];
    if (v !== '' && v != null) answers[c - FIXED + 1] = String(v);
  }
  const fields = vals.slice(1, FIXED - 2).map(v => v == null ? '' : String(v));
  return json_({ found: true, name: fields[0] || '', fields, updated: iso_(vals[FIXED - 2]), state: String(vals[FIXED - 1] || ''), answers });
}

function list_(name, owner) {
  owner = owner == null ? '' : String(owner).trim();
  if (name === DEFAULT_SHEET || !owner) return json_({ ok: false, error: 'no list' });
  const sh = sheet_(name, false);
  if (!sh) return json_({ ok: true, rows: [] });
  const last = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (last < 2) return json_({ ok: true, rows: [] });
  const FIXED = fixed_(sh);
  const vals = sh.getRange(2, 1, last - 1, lastCol).getValues();
  const pre = owner + '|', rows = [];
  vals.forEach(v => {
    const id = String(v[0] == null ? '' : v[0]);
    if (!id || id.indexOf(pre) !== 0) return;
    let done = 0;
    for (let c = FIXED; c < lastCol; c++) {
      const a = v[c];
      if (a !== '' && a != null && !EMPTY_MARKS[String(a)]) done++;
    }
    rows.push({ id, fields: v.slice(1, FIXED - 2).map(x => x == null ? '' : String(x)), updated: iso_(v[FIXED - 2]), done, total: lastCol - FIXED });
  });
  return json_({ ok: true, rows });
}

// POST {sheet?, id, name? | fields?, labels?, total, questions?, answers:{n:text}, clear:[n], state?}
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (!d || !d.id) return json_({ ok: false, error: 'no id' });
    const name = sheetName_(d);
    const sh = sheet_(name, true);
    const fields = Array.isArray(d.fields) ? d.fields.map(x => x == null ? '' : String(x)) : [d.name == null ? '' : String(d.name)];
    const labels = Array.isArray(d.labels) && d.labels.length === fields.length ? d.labels.map(String)
      : (fields.length === 1 ? ['імʼя'] : fields.map((x, i) => 'поле ' + (i + 1)));
    const n = (d.questions && d.questions.length) || Number(d.total) || 0;
    if (sh.getLastRow() === 0) {
      if (!n) return json_({ ok: false, error: 'no questions' });
      const head = ['id'].concat(labels, ['оновлено', 'дані']);
      for (let i = 1; i <= n; i++) {
        const q = d.questions && d.questions[i - 1];
        head.push(q ? (q.n + '. ' + q.q) : String(i));
      }
      sh.appendRow(head);
      sh.setFrozenRows(1);
      sh.setFrozenColumns(1 + labels.length);
      const F = head.length - n;
      sh.hideColumns(F);           // службова колонка зі станом чек-лістів
      sh.setColumnWidths(F + 1, n, 220);
    }
    const FIXED = fixed_(sh);
    const nQ = Math.max(n, sh.getLastColumn() - FIXED);
    const clr = {}; (d.clear || []).forEach(x => clr[String(x)] = 1);
    let row = findRow_(sh, d.id);
    let existing = null;
    if (row) existing = sh.getRange(row, 1, 1, FIXED + nQ).getValues()[0];
    else row = sh.getLastRow() + 1;
    const vals = [String(d.id)];
    for (let f = 0; f < FIXED - 3; f++) {
      const v = fields[f] != null ? fields[f] : '';
      vals.push(v || (existing && existing[1 + f] != null ? existing[1 + f] : ''));
    }
    vals.push(new Date());
    vals.push(d.state != null ? String(d.state) : (existing && existing[FIXED - 1] != null ? existing[FIXED - 1] : ''));
    for (let i = 1; i <= nQ; i++) {
      let a = d.answers ? d.answers[i] : '';
      a = a == null ? '' : String(a);
      const was = existing ? existing[FIXED + i - 1] : '';
      if (!a && clr[String(i)]) a = '';                               // людина навмисно стерла відповідь
      else if (!a && was !== '' && was != null) a = String(was);      // інакше порожнє не стирає вже збережене
      vals.push(a);
    }
    sh.getRange(row, 1, 1, vals.length).setValues([vals]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
