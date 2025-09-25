/**
 * AI Club Assignment Tracker — Consolidated Script
 * - ICS import (with course extraction)
 * - Canvas API import (points, links, descriptions)
 * - Cleaners (fill Course, default times, points-from-notes, Days Left)
 * - Syllabus parsing via OpenAI (optional)
 * - Calendar sync
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AI + Canvas')
    .addItem('Import from Canvas (ICS)', 'uiImportIcs')
    .addItem('Import from Canvas API (points)', 'uiImportFromCanvasApi')
    .addSeparator()
    .addItem('Set Canvas Domain', 'uiSetCanvasDomain')
    .addItem('Set Canvas API Token', 'uiSetCanvasToken')
    .addSeparator()
    .addItem('Parse Syllabus/Text', 'uiParseSyllabus')
    .addItem('Classify & Apply Weights', 'uiClassifyAndWeight')
    .addSeparator()
    .addItem('Clean Titles & Fill Course', 'cleanTitlesAndFillCourse')
    .addItem('Autofill & Clean', 'autofillAndClean')
    .addSeparator()
    .addItem('Add to Google Calendar', 'uiAddToCalendar')
    .addToUi();
}

// ==== CONFIG ====
const SHEET_NAME  = 'Assignments';
const PASTE_SHEET = 'Paste';

// Optional AI (only for syllabus parsing / policy mapping)
const OPENAI_API_KEY = 'YOUR_OPENAI_KEY_HERE';
const MODEL = 'gpt-4o';

// ==== HELPERS ====
function getSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  const headers = ['Course','Assignment Title','Type','Due Date','Due Time','Points','Category','Weight %','Source','Link','Status','Notes','Days Left'];
  if (sh.getLastRow() === 0) sh.appendRow(headers);
  return sh;
}
function getPasteSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(PASTE_SHEET);
  if (!sh) sh = ss.insertSheet(PASTE_SHEET);
  return sh;
}
function getBodyRange_(sh) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow <= 1) return null;
  return sh.getRange(2, 1, lastRow - 1, lastCol);
}
function writeRowsInBatches_(sh, startRow, values, batchSize) {
  const lastCol = sh.getLastColumn();
  const bs = batchSize || 300;
  for (let i = 0; i < values.length; i += bs) {
    const slice = values.slice(i, i + bs);
    sh.getRange(startRow + i, 1, slice.length, lastCol).setValues(slice);
    SpreadsheetApp.flush();
    Utilities.sleep(30);
  }
}
function stripHtml_(html) {
  if (!html) return '';
  return html.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
}
function guessTypeFromSummary(s) {
  s = (s || '').toLowerCase();
  if (s.includes('quiz')) return 'Quiz';
  if (s.includes('exam') || s.includes('midterm') || s.includes('final')) return 'Exam';
  if (s.includes('lab')) return 'Lab';
  if (s.includes('project')) return 'Project';
  if (s.includes('hw') || s.includes('homework') || s.includes('assignment')) return 'Homework';
  if (s.includes('reading')) return 'Reading';
  return 'Other';
}

// ==== COURSE + TITLE EXTRACTION (from bracketed tail in SUMMARY) ====
const COURSE_REGEX = /[A-Z]{2,6}-\d{2,4}(?:-[A-Z0-9]{2,4})?/; // e.g., CS-483-001, GEOL-101-P01
function extractCourseAndCleanTitle_(rawTitle) {
  if (!rawTitle) return { course: '', title: '' };
  let title = String(rawTitle).replace(/\\,/g, ',').trim();

  // collect all [ ... ] groups; pick first that contains a course code
  const groups = [...title.matchAll(/\[([^\]]+)\]/g)].map(m => m[1]);
  let course = '';
  for (const g of groups) {
    const inside = g.replace(/[()]/g, ' ');
    const parts = inside.split(/[,;|]/).map(x => x.trim()).filter(Boolean);
    for (const p of parts) {
      const m = p.match(COURSE_REGEX);
      if (m) { course = m[0]; break; }
    }
    if (course) break;
  }
  // strip final trailing [...] block for a cleaner title
  title = title.replace(/\s*\[[^\]]+\]\s*$/, '').trim();
  return { course, title };
}

// ==== ICS IMPORT ====
function uiImportIcs() {
  const ui = SpreadsheetApp.getUi();
  const paste = getPasteSheet();
  const fallback = paste.getRange('C3').getValue();
  const resp = ui.prompt('Paste your Canvas ICS URL (or leave blank to use Paste!C3):');
  let url = resp.getResponseText() || fallback;
  if (!url) { ui.alert('No ICS URL provided.'); return; }
  if (!/^https?:\/\//i.test(url)) { // guard if user pasted just the token
    url = 'https://canvas.gmu.edu/feeds/calendars/' + url.replace(/^\/+/, '');
  }
  const ics = UrlFetchApp.fetch(url, {muteHttpExceptions:true}).getContentText();
  importIcsToSheet(ics);
  ui.alert('Canvas ICS import complete.');
}
function importIcsToSheet(icsText) {
  const sh = getSheet();
  const events = parseIcs_(icsText);
  const rows = events.map(e => {
    const { course, title } = extractCourseAndCleanTitle_(e.summary || '');
    return [
      course || '',
      title || (e.summary || ''),
      guessTypeFromSummary(title || e.summary || ''),
      e.dt || '',
      (e.time && e.time !== '00:00') ? e.time : '23:59',
      '',
      '',
      '',
      'Canvas ICS',
      e.url || '',
      'Not started',
      e.desc || '',
      ''
    ];
  });
  if (!rows.length) return;
  writeRowsInBatches_(sh, sh.getLastRow() + 1, rows, 300);
}
// robust ICS parser (handles folded lines)
function parseIcs_(txt) {
  const lines = txt.split(/\r?\n/);
  const out = [];
  let cur = null, lastProp = null;
  const readVal = (line) => {
    const idx = line.indexOf(':');
    return idx >= 0 ? line.slice(idx + 1).trim() : '';
  };
  for (let i = 0; i < lines.length; i++) {
    const ln = lines[i];
    if (ln.startsWith('BEGIN:VEVENT')) { cur = {}; lastProp = null; continue; }
    if (ln.startsWith('END:VEVENT'))   { if (cur) out.push(cur); cur = null; lastProp = null; continue; }
    if (!cur) continue;

    if (/^[ \t]/.test(ln) && lastProp) { // folded continuation line
      cur[lastProp] = (cur[lastProp] || '') + ln.trim();
      continue;
    }

    if (ln.startsWith('SUMMARY'))   { cur.summary = readVal(ln); lastProp = 'summary'; }
    else if (ln.startsWith('DESCRIPTION')) { cur.desc = readVal(ln); lastProp = 'desc'; }
    else if (ln.startsWith('URL'))  { cur.url = readVal(ln); lastProp = 'url'; }
    else if (ln.startsWith('DTSTART')) {
      const v = readVal(ln); // 20250920T235900Z
      const yyyy = v.slice(0,4), mm = v.slice(4,6), dd = v.slice(6,8);
      const hh = v.slice(9,11) || '00', mi = v.slice(11,13) || '00';
      cur.dt = `${yyyy}-${mm}-${dd}`;
      cur.time = `${hh}:${mi}`;
      lastProp = 'dt';
    } else lastProp = null;
  }
  return out;
}

// ==== CLEANERS ====
function cleanTitlesAndFillCourse() {
  const sh = getSheet();
  const body = getBodyRange_(sh);
  if (!body) { SpreadsheetApp.getUi().alert('No rows to process.'); return; }

  const header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const idx = (n) => header.indexOf(n);
  const titleIdx  = idx('Assignment Title');
  const courseIdx = idx('Course');

  const values = body.getValues();
  const bs = 300;
  for (let off = 0; off < values.length; off += bs) {
    const slice = values.slice(off, off + bs);
    for (let r = 0; r < slice.length; r++) {
      const { course, title } = extractCourseAndCleanTitle_(slice[r][titleIdx]);
      if (course) slice[r][courseIdx] = course;
      if (title)  slice[r][titleIdx]  = title;
    }
    sh.getRange(2 + off, 1, slice.length, header.length).setValues(slice);
    SpreadsheetApp.flush();
    Utilities.sleep(20);
  }
  SpreadsheetApp.getUi().alert('Titles cleaned & Course filled.');
}

function autofillAndClean() {
  const sh = getSheet();
  const body = getBodyRange_(sh);
  if (!body) { SpreadsheetApp.getUi().alert('No rows to clean.'); return; }

  const header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const idx = (n) => header.indexOf(n);

  const dueDateIdx  = idx('Due Date');
  const dueTimeIdx  = idx('Due Time');
  const typeIdx     = idx('Type');
  const categoryIdx = idx('Category');
  const pointsIdx   = idx('Points');
  const notesIdx    = idx('Notes');
  const daysLeftIdx = idx('Days Left');

  const values = body.getValues();
  const pointsRegex = /(\d+(?:\.\d+)?)\s*(?:pts?|points?)/i;

  const bs = 300;
  for (let off = 0; off < values.length; off += bs) {
    const slice = values.slice(off, off + bs);
    for (let r = 0; r < slice.length; r++) {
      const t = (slice[r][dueTimeIdx] || '').toString();
      if (!t || t === '0:00' || t === '00:00') slice[r][dueTimeIdx] = '23:59';

      const ty = (slice[r][typeIdx] || '').toString();
      if (ty && !slice[r][categoryIdx]) slice[r][categoryIdx] = ty;

      const notes = (slice[r][notesIdx] || '').toString();
      const pm = notes.match(pointsRegex);
      if (pm && !slice[r][pointsIdx]) slice[r][pointsIdx] = pm[1];

      if (!slice[r][daysLeftIdx]) {
        const sheetRow = 2 + off + r;
        const dueCell = sh.getRange(sheetRow, dueDateIdx + 1).getA1Notation();
        slice[r][daysLeftIdx] = `=IF(${dueCell}="","", ${dueCell} - TODAY())`;
      }
    }
    sh.getRange(2 + off, 1, slice.length, header.length).setValues(slice);
    SpreadsheetApp.flush();
    Utilities.sleep(20);
  }
  SpreadsheetApp.getUi().alert('Autofill complete.');
}

// ==== SYLLABUS (AI, optional) ====
function uiParseSyllabus() {
  const paste = getPasteSheet();
  const course = Browser.inputBox('Course code (optional, e.g., CS 262)');
  const text = paste.getRange('A3').getValue();
  if (!text) { SpreadsheetApp.getUi().alert('Paste your syllabus text into Paste!A3 first.'); return; }
  const items = aiExtractAssignments_(text);
  const sh = getSheet();
  const rows = items.map(it => ([
    course || it.course || '',
    it.title || '',
    it.type || '',
    it.due_date || '',
    it.due_time || '',
    it.points || '',
    it.category || (it.type || ''),
    it.weight || '',
    'Syllabus AI',
    it.link || '',
    'Not started',
    it.notes || '',
    ''
  ]));
  if (rows.length) writeRowsInBatches_(sh, sh.getLastRow()+1, rows, 300);
  SpreadsheetApp.getUi().alert('Syllabus parsing complete.');
}
function aiExtractAssignments_(rawText) {
  const system = "You extract assignments from messy text. Return strict JSON:"+
    "{\"assignments\":[{\"title\":\"\",\"type\":\"\",\"due_date\":\"\",\"due_time\":\"\",\"points\":\"\",\"link\":\"\",\"notes\":\"\"}],"+
    "\"recurring\":[{\"title\":\"\",\"type\":\"\",\"weekday\":\"\",\"start_date\":\"\",\"end_date\":\"\",\"time\":\"\",\"notes\":\"\"}]}";
  const user = "Text:\\n\"\"\""+rawText+"\"\"\"\\n\\nRules:\\n- Detect assignments with dates. If only month/day is given, infer the current academic year."+
               "\\n- type ∈ {Homework, Quiz, Exam, Lab, Project, Reading, Other}\\n- If it says \"quiz every Sunday until Dec 8\", put it in recurring."+
               "\\n- Prefer 24h time like \"23:59\"; if missing, leave empty.\\n- Return JSON only.";
  const json = callOpenAI_(system, user);
  let data = {assignments:[], recurring:[]};
  try { data = JSON.parse(json); } catch(e) {}
  const items = (data.assignments || []).map(a => ({
    title: a.title || '',
    type: a.type || '',
    due_date: a.due_date || '',
    due_time: a.due_time || '',
    points: a.points || '',
    link: a.link || '',
    notes: a.notes || '',
    category: a.type || '',
    weight: ''
  }));
  for (const r of (data.recurring || [])) {
    const expanded = expandWeekly_(r);
    items.push(...expanded.map(dt => ({
      title: r.title || '',
      type: r.type || '',
      due_date: dt,
      due_time: r.time || '',
      points: '',
      link: '',
      notes: r.notes || '',
      category: r.type || '',
      weight: ''
    })));
  }
  return items;
}
function expandWeekly_(rec) {
  try {
    const start = new Date(rec.start_date);
    const end = new Date(rec.end_date);
    if (isNaN(start) || isNaN(end)) return [];
    const dows = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
    const targetDow = dows.indexOf(rec.weekday);
    if (targetDow < 0) return [];
    const dates = [];
    let d = new Date(start);
    while (d.getDay() !== targetDow) d.setDate(d.getDate()+1);
    while (d <= end) {
      const yyyy = d.getFullYear();
      const mm = ('0'+(d.getMonth()+1)).slice(-2);
      const dd = ('0'+d.getDate()).slice(-2);
      dates.push(`${yyyy}-${mm}-${dd}`);
      d.setDate(d.getDate()+7);
    }
    return dates;
  } catch(e) { return []; }
}
function callOpenAI_(system, user) {
  if (!OPENAI_API_KEY || OPENAI_API_KEY === 'YOUR_OPENAI_KEY_HERE') {
    throw new Error('Set OPENAI_API_KEY before using AI features.');
  }
  const res = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + OPENAI_API_KEY },
    payload: JSON.stringify({
      model: MODEL,
      response_format: { type: "json_object" },
      messages: [{role:'system', content:system},{role:'user', content:user}]
    }),
    muteHttpExceptions: true
  });
  const data = JSON.parse(res.getContentText());
  const content = data && data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content;
  if (!content) throw new Error('OpenAI response error:\n' + res.getContentText());
  return content;
}

// ==== CLASSIFY & WEIGHTS ====
function uiClassifyAndWeight() {
  const paste = getPasteSheet();
  const policy = paste.getRange('B3').getValue() || Browser.inputBox('Paste grading policy text (optional)');
  const mapping = policy ? aiPolicyToCategories_(policy) : {};
  applyCategoriesAndWeights_(mapping);
  SpreadsheetApp.getUi().alert('Classification applied.');
}
function aiPolicyToCategories_(text) {
  const system = "Turn grading policy text into JSON mapping of category to percent. Example:\\n{\"Homework\":20,\"Quizzes\":15,\"Midterm\":25,\"Final\":40}";
  const user = "Policy:\\n" + text + "\\nReturn JSON only.";
  try { return JSON.parse(callOpenAI_(system,user)); } catch(e) { return {}; }
}
function applyCategoriesAndWeights_(mapping) {
  const sh = getSheet();
  const body = getBodyRange_(sh);
  if (!body) return;
  const header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const idx = (n) => header.indexOf(n);
  const typeIdx = idx('Type'), catIdx = idx('Category'), wtIdx = idx('Weight %');

  const values = body.getValues();
  for (let r = 0; r < values.length; r++) {
    const type = (values[r][typeIdx]||'').toString();
    let cat = values[r][catIdx] || type;
    let weight = mapping[cat] ?? mapping[type] ?? mapping[type+'s'] ?? '';
    values[r][catIdx] = cat;
    values[r][wtIdx] = weight;
  }
  body.setValues(values);
}

// ==== GOOGLE CALENDAR SYNC ====
function uiAddToCalendar() {
  const cal = CalendarApp.getDefaultCalendar();
  const sh = getSheet();
  const body = getBodyRange_(sh);
  if (!body) { SpreadsheetApp.getUi().alert('No dated rows.'); return; }
  const header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const idx = (n) => header.indexOf(n);
  const titleIdx = idx('Assignment Title'), dateIdx = idx('Due Date'), timeIdx = idx('Due Time'), notesIdx = idx('Notes'), linkIdx = idx('Link');

  const values = body.getValues();
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const date = row[dateIdx];
    if (!date) continue;
    const t = row[timeIdx] ? row[timeIdx] : '23:59';
    const parts = t.split(':');
    const hh = Number(parts[0]||23), mm = Number(parts[1]||59);
    const start = new Date(date); start.setHours(hh, mm, 0, 0);
    const end = new Date(start.getTime()+60*60*1000);
    const desc = (row[notesIdx]||'') + (row[linkIdx] ? `\n${row[linkIdx]}` : '');
    cal.createEvent(row[titleIdx] || 'Assignment', start, end, {description: desc});
  }
  SpreadsheetApp.getUi().alert('Events added to your Google Calendar.');
}

// ==== CANVAS API (points, links) ====
// store per-user domain/token in UserProperties
function uiSetCanvasDomain() {
  const ui = SpreadsheetApp.getUi();
  const def = PropertiesService.getUserProperties().getProperty('CANVAS_DOMAIN') || 'canvas.gmu.edu';
  const resp = ui.prompt('Canvas domain (no protocol):', def, ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const domain = (resp.getResponseText() || def).trim().replace(/^https?:\/\//,'').replace(/\/+$/,'');
  PropertiesService.getUserProperties().setProperty('CANVAS_DOMAIN', domain);
  ui.alert('Saved: ' + domain);
}
function uiSetCanvasToken() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Paste your Canvas API Access Token (Account → Settings → New Access Token):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const token = (resp.getResponseText() || '').trim();
  if (!token) { ui.alert('No token entered.'); return; }
  PropertiesService.getUserProperties().setProperty('CANVAS_TOKEN', token);
  ui.alert('Token saved.');
}
function getCanvasBase_() {
  const domain = PropertiesService.getUserProperties().getProperty('CANVAS_DOMAIN') || 'canvas.gmu.edu';
  return 'https://' + domain + '/api/v1';
}
function getCanvasToken_() {
  const token = PropertiesService.getUserProperties().getProperty('CANVAS_TOKEN');
  if (!token) throw new Error('Canvas token not set. Use "AI + Canvas → Set Canvas API Token".');
  return token;
}
function parseNextLink_(linkHeader) {
  if (!linkHeader) return null;
  for (const part of linkHeader.split(',')) {
    const m = part.match(/<([^>]+)>;\s*rel="next"/i);
    if (m) return m[1];
  }
  return null;
}
function fetchCanvasJsonPaginated_(endpoint, params) {
  const base  = getCanvasBase_();
  const token = getCanvasToken_();
  const query = Object.keys(params || {}).map(k => `${encodeURIComponent(k)}=${encodeURIComponent(params[k])}`).join('&');
  let url = `${base}${endpoint}${query ? ('?' + query) : ''}`;
  let all = [];
  while (url) {
    const res = UrlFetchApp.fetch(url, { method:'get', headers:{ Authorization: 'Bearer ' + token }, muteHttpExceptions:true });
    if (res.getResponseCode() >= 400) throw new Error('Canvas error '+res.getResponseCode()+': '+res.getContentText());
    const page = JSON.parse(res.getContentText());
    if (Array.isArray(page)) all = all.concat(page); else all.push(page);
    url = parseNextLink_(res.getAllHeaders()['Link'] || res.getAllHeaders()['link']);
    Utilities.sleep(100);
  }
  return all;
}
function uiImportFromCanvasApi() {
  const ui = SpreadsheetApp.getUi();
  try { getCanvasToken_(); } catch (e) { ui.alert(e.message); return; }

  const sh = getSheet();
  const courses = fetchCanvasJsonPaginated_('/courses', {
    enrollment_state: 'active',
    per_page: 100
  }).filter(c => !c.access_restricted_by_date);

  if (!courses.length) { ui.alert('No active courses found.'); return; }

  let rows = [];
  for (const c of courses) {
    const courseLabel = (c.course_code || c.name || '').toString().trim();
    const courseId = c.id;
    const assignments = fetchCanvasJsonPaginated_(`/courses/${courseId}/assignments`, { per_page: 100 });
    for (const a of assignments) {
      if (!a || !a.due_at) continue; // skip undated
      const d = new Date(a.due_at);
      const yyyy = d.getFullYear();
      const mm = ('0' + (d.getMonth()+1)).slice(-2);
      const dd = ('0' + d.getDate()).slice(-2);
      const hh = ('0' + d.getHours()).slice(-2);
      const mi = ('0' + d.getMinutes()).slice(-2);
      const title = a.name || '';
      const type  = guessTypeFromSummary(title);
      const points = (a.points_possible != null) ? a.points_possible : '';
      rows.push([
        courseLabel,
        title,
        type,
        `${yyyy}-${mm}-${dd}`,
        `${hh}:${mi}`,
        points,
        '',
        '',
        'Canvas API',
        a.html_url || '',
        'Not started',
        (a.description && stripHtml_(a.description).slice(0, 2000)) || '',
        ''
      ]);
    }
  }
  if (!rows.length) { ui.alert('No dated assignments found via API.'); return; }
  writeRowsInBatches_(sh, sh.getLastRow() + 1, rows, 300);
  autofillAndClean(); // now defined
  ui.alert(`Imported ${rows.length} assignments from Canvas API (with points).`);
}
