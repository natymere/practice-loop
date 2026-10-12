const OWNER_EMAIL = 'fearlessjem17@gmail.com';
const SPREADSHEET_ID = '1hOzK2hN57ZTyq1oJyAZHB0bbTqDN4zLSiRmvwJDYbUQ';
const INTERVALS = [1, 3, 7, 14, 30, 60, 90];
const Q_HEADERS = ['ID', 'URL', 'Canonical URL', 'Title', 'Topic', 'Created at', 'Initial help used', 'Estimated minutes', 'Stage', 'Next review', 'Last result', 'Last reviewed', 'Status', 'Snoozed until', 'Revision', 'Notes', 'Collection', 'Restore location'];
const IMPORT_HEADERS = ['Title', 'URL', 'Topic', 'Collection', 'Result'];
const PAGE_SIZE = 25;
const R_HEADERS = ['Request ID', 'Question ID', 'Reviewed at', 'Result', 'Minutes', 'Mistake note', 'Previous stage', 'New stage', 'Next review', 'Previous revision', 'New revision'];
const D_HEADERS = ['Digest key', 'Attempt ID', 'Started at', 'Channel', 'Status', 'Finished at', 'Detail'];
function assertOwner_() {
  const email = Session.getActiveUser().getEmail().toLowerCase();
  if (!email || email !== OWNER_EMAIL) throw new Error('Sign in with the Google account that owns Practice Loop.');
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}
function book_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}
function rows_(name, headers) {
  const sh = book_().getSheetByName(name);
  if (!sh) throw new Error('Missing database tab: ' + name);
  const all = sh.getDataRange().getValues();
  if (name === 'Questions') {
    const width = all[0].length;
    const known = [15, 16, 18].includes(width) &&
      JSON.stringify(all[0]) === JSON.stringify(Q_HEADERS.slice(0, width));
    if (!known) throw new Error('The Questions columns have changed. Restore the expected headers.');
    const changes = width < Q_HEADERS.length ? [replace_(sh.getSheetId(), 0, Q_HEADERS)] : [];
    all[0] = Q_HEADERS.slice();
    all.slice(1).forEach((row, index) => {
      while (row.length < Q_HEADERS.length) row.push('');
      if (row[12] === 'paused') {
        row[12] = 'bank';
        row[9] = '';
        row[13] = '';
        row[14] = Number(row[14]) + 1;
        changes.push(replace_(sh.getSheetId(), index + 1, row));
      } else if (row[12] === 'deleted' && !row[17]) {
        row[17] = 'active';
        changes.push(replace_(sh.getSheetId(), index + 1, row));
      }
    });
    if (changes.length) commit_(changes);
  }
  if (headers && JSON.stringify(all[0]) !== JSON.stringify(headers)) throw new Error('The ' + name + ' columns have changed. Restore the original headers.');
  const data = all.slice(1);
  if (data.some(r => r[0] === '')) throw new Error('A blank row interrupts the ' + name + ' data. Remove the blank row before continuing.');
  return {
    sheetId: sh.getSheetId(),
    rows: data
  };
}
function cell_(v) {
  return {
    userEnteredValue: typeof v === 'number' ? {
      numberValue: v
    } : typeof v === 'boolean' ? {
      boolValue: v
    } : {
      stringValue: String(v == null ? '' : v)
    }
  };
}
function append_(sheetId, row) {
  return {
    appendCells: {
      sheetId: sheetId,
      rows: [{
        values: row.map(cell_)
      }],
      fields: 'userEnteredValue'
    }
  };
}
function replace_(sheetId, rowIndex, row) {
  return {
    updateCells: {
      start: {
        sheetId: sheetId,
        rowIndex: rowIndex,
        columnIndex: 0
      },
      rows: [{
        values: row.map(cell_)
      }],
      fields: 'userEnteredValue'
    }
  };
}
function commit_(requests) {
  Sheets.Spreadsheets.batchUpdate({
    requests: requests
  }, SPREADSHEET_ID);
}
function settings_() {
  const values = {};
  rows_('Settings').rows.forEach(r => values[r[0]] = r[1]);
  const timezone = String(values.timezone || 'Asia/Singapore');
  try {
    Utilities.formatDate(new Date(), timezone, 'yyyy-MM-dd');
  } catch (_) {
    throw new Error('Invalid timezone in Settings.');
  }
  const reminderHour = Number(values.reminderHour),
    dailyBudget = Number(values.dailyBudget),
    channel = String(values.channel);
  if (!Number.isInteger(reminderHour) || reminderHour < 0 || reminderHour > 23 || !Number.isInteger(dailyBudget) || dailyBudget < 5 || dailyBudget > 240 || ['off', 'email', 'telegram'].indexOf(channel) < 0) throw new Error('Invalid reminder settings.');
  return {
    timezone: timezone,
    reminderHour: reminderHour,
    dailyBudget: dailyBudget,
    channel: channel
  };
}
function day_(now, tz) {
  return Utilities.formatDate(now, tz, 'yyyy-MM-dd');
}
function plusDays_(day, days) {
  const d = new Date(day + 'T12:00:00Z');
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}
function transition_(stage, result, day) {
  if (!Number.isInteger(stage) || stage < 0 || stage >= INTERVALS.length) throw new Error('Invalid review stage.');
  if (['help', 'hard', 'independent'].indexOf(result) < 0) throw new Error('Choose a review result.');
  const nextStage = result === 'help' ? 0 : result === 'independent' ? Math.min(stage + 1, INTERVALS.length - 1) : stage;
  const days = result === 'independent' ? INTERVALS[nextStage] : 1;
  return {
    stage: nextStage,
    due: plusDays_(day, days),
    days: days
  };
}
function text_(v, max, required) {
  const value = String(v == null ? '' : v).trim();
  if (required && !value || value.length > max) throw new Error('Check the required fields and text lengths.');
  return value;
}
function int_(v, min, max) {
  const n = Number(v);
  if (!Number.isInteger(n) || n < min || n > max) throw new Error('Enter a whole number between ' + min + ' and ' + max + '.');
  return n;
}
function canonicalUrl_(input) {
  const raw = text_(input, 2000, true);
  const match = raw.match(/^(https?):\/\/([^/?#\s]+)([^?#\s]*)(\?[^#\s]*)?(#[^\s]*)?$/i);
  if (!match || /[@\\]/.test(match[2]) || /[\u0000-\u0020\\]/.test(raw)) throw new Error('Enter a complete http or https question URL.');
  let scheme = match[1].toLowerCase(),
    host = match[2].toLowerCase(),
    path = match[3] || '/',
    query = match[4] || '',
    hash = match[5] || '';
  if (!/^(?:[a-z0-9-]+\.)+[a-z0-9-]+(?::\d{1,5})?$/.test(host)) throw new Error('Enter a public website URL.');
  host = host.replace(scheme === 'https' ? /:443$/ : /:80$/, '');
  if (/^(www\.)?leetcode\.com$/.test(host)) {
    const problem = path.match(/^\/problems\/([^/]+)/);
    if (problem) return 'https://leetcode.com/problems/' + problem[1].toLowerCase();
  }
  path = path.replace(/\/+$/, '') || '/';
  // Keep question IDs, meaningful query parameters, and hash routes on other sites.
  query = query.slice(1).split('&').filter(Boolean).filter(p => !/^(utm_[^=]*|fbclid|gclid)=/i.test(p)).sort().join('&');
  return scheme + '://' + host + path + (query ? '?' + query : '') + hash;
}
function question_(row, rowIndex) {
  return {
    id: String(row[0]),
    url: String(row[1]),
    canonicalUrl: String(row[2]),
    title: String(row[3]),
    topic: String(row[4]),
    createdAt: String(row[5]),
    initialHelp: row[6] === true || row[6] === 'true',
    minutes: Number(row[7]),
    stage: Number(row[8]),
    due: String(row[9]),
    lastResult: String(row[10]),
    lastReviewed: String(row[11]),
    status: String(row[12]),
    snoozedUntil: String(row[13]),
    revision: Number(row[14]),
    notes: String(row[15] || ''),
    collection: String(row[16] || ''),
    restoreLocation: String(row[17] || ''),
    rowIndex: rowIndex
  };
}
function due_(q, today) {
  return q.status === 'active' && Boolean(q.due) && q.due <= today && (!q.snoozedUntil || q.snoozedUntil <= today);
}
function chooseSession_(questions, today, budget) {
  const due = questions.filter(q => due_(q, today)).sort((a, b) => a.due.localeCompare(b.due) || Number(b.lastResult === 'help') - Number(a.lastResult === 'help') || a.createdAt.localeCompare(b.createdAt));
  const selected = [];
  let minutes = 0;
  for (const q of due) {
    if (!selected.length || minutes + q.minutes <= budget) {
      selected.push(q);
      minutes += q.minutes;
    }
  }
  return {
    due: due,
    selected: selected,
    minutes: minutes
  };
}
function doGet() {
  assertOwner_();
  return HtmlService.createTemplateFromFile('Index').evaluate().setTitle('Practice Loop').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
// Only trusted, static project files are included in the page shell.
function include_(filename) {
  if (['Styles', 'Client'].indexOf(filename) < 0) throw new Error('Unknown page file.');
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
// List responses omit notes/history. Those are fetched only when expanded.
function summary_(q) {
  return {id:q.id, url:q.url, title:q.title, topic:q.topic, collection:q.collection,
    minutes:q.minutes, stage:q.stage, due:q.due, status:q.status,
    snoozedUntil:q.snoozedUntil, revision:q.revision};
}
function scheduleOptions_(today) {
  return INTERVALS.map((days, index) => ({stage:index + 1, days:days, due:plusDays_(today, days)}));
}
function getDashboard() {
  assertOwner_();
  return withLock_(function () {
    const settings = settings_(), today = day_(new Date(), settings.timezone);
    const questions = rows_('Questions', Q_HEADERS).rows.map((r, i) => question_(r, i + 1));
    const plan = chooseSession_(questions, today, settings.dailyBudget);
    return {today:today, settings:settings, stageOptions:scheduleOptions_(today),
      questions:plan.selected.map(summary_), dueCount:plan.due.length,
      sessionMinutes:plan.minutes, practiceCount:questions.filter(q => q.status === 'active').length,
      bankCount:questions.filter(q => q.status === 'bank').length};
  });
}
function getQuestions(input) {
  assertOwner_();
  const view = text_(input.view, 20, true);
  if (!['active', 'bank', 'deleted', 'due'].includes(view)) throw new Error('Choose a Library view.');
  const offset = int_(input.offset == null ? 0 : input.offset, 0, 1000000);
  const search = text_(input.search, 200, false).toLowerCase();
  const topic = text_(input.topic, 100, false), collection = text_(input.collection, 100, false);
  return withLock_(function () {
    const today = day_(new Date(), settings_().timezone);
    const all = rows_('Questions', Q_HEADERS).rows.map((r, i) => question_(r, i + 1));
    const scope = all.filter(q => view === 'due' ? due_(q, today) : q.status === view);
    const questions = scope.filter(q => (!topic || q.topic === topic) &&
      (!collection || q.collection === collection) &&
      [q.title, q.topic, q.url, q.collection].join(' ').toLowerCase().includes(search))
      .sort((a, b) => a.due.localeCompare(b.due) || a.title.localeCompare(b.title) || a.id.localeCompare(b.id));
    const page = questions.slice(offset, offset + PAGE_SIZE);
    return {questions:page.map(summary_), total:questions.length,
      nextOffset:offset + page.length < questions.length ? offset + page.length : null,
      topics:[...new Set(scope.map(q => q.topic).filter(Boolean))].sort(),
      collections:[...new Set(scope.map(q => q.collection).filter(Boolean))].sort()};
  });
}
function getQuestionDetails(input) {
  assertOwner_();
  const id = text_(input.id, 80, true);
  return withLock_(function () {
    const data = rows_('Questions', Q_HEADERS), index = data.rows.findIndex(row => row[0] === id);
    if (index < 0) throw new Error('Question not found.');
    const q = question_(data.rows[index], index + 1);
    const history = rows_('Reviews', R_HEADERS).rows.filter(row => row[1] === id).slice(-20).reverse()
      .map(row => ({at:row[2], result:row[3], note:row[5]}));
    const today = day_(new Date(), settings_().timezone);
    return {question:Object.assign(summary_(q), {notes:q.notes}), history:history,
      stageOptions:scheduleOptions_(today)};
  });
}
function getSettings() {
  assertOwner_();
  return withLock_(function () {
    const deliveries = rows_('Deliveries', D_HEADERS).rows, last = deliveries[deliveries.length - 1];
    const props = PropertiesService.getScriptProperties();
    const importSheet = book_().getSheetByName('Import');
    return {settings:settings_(), sheetUrl:book_().getUrl(),
      importUrl:importSheet ? book_().getUrl() + '#gid=' + importSheet.getSheetId() : '',
      telegramConfigured:!!(props.getProperty('TELEGRAM_BOT_TOKEN') && props.getProperty('TELEGRAM_CHAT_ID')),
      lastDelivery:last ? {at:last[2], status:last[4], detail:last[6]} : null};
  });
}

function importSheet_() {
  let sheet = book_().getSheetByName('Import');
  if (!sheet) sheet = book_().insertSheet('Import');
  const all = sheet.getDataRange().getValues();
  if (all.length === 1 && all[0].every(value => value === '')) {
    commit_([replace_(sheet.getSheetId(), 0, IMPORT_HEADERS)]);
  } else if (JSON.stringify(all[0]) !== JSON.stringify(IMPORT_HEADERS)) {
    throw new Error('Import needs these headers: Title, URL, Topic, Collection, Result.');
  }
  return sheet;
}
// Run once before updating the deployment. Safe to run again after interruption.
function upgrade() {
  assertOwner_();
  return withLock_(function () {
    rows_('Questions', Q_HEADERS);
    const sheet = importSheet_();
    return {ok:true, importUrl:book_().getUrl() + '#gid=' + sheet.getSheetId()};
  });
}
function importQuestions() {
  assertOwner_();
  return withLock_(function () {
    const sheet = importSheet_(), inputRows = sheet.getDataRange().getValues().slice(1);
    const data = rows_('Questions', Q_HEADERS);
    const known = new Set(data.rows.map(row => canonicalUrl_(row[1])));
    const additions = [], results = [];
    const now = new Date().toISOString();
    let imported = 0, duplicates = 0, errors = 0;
    inputRows.forEach(row => {
      if (row.slice(0, 4).every(value => value === '')) { results.push(['']); return; }
      try {
        if (!String(row[0] || '').trim()) throw new Error('Title is required.');
        if (!String(row[1] || '').trim()) throw new Error('URL is required.');
        [['Title', 160], ['URL', 2000], ['Topic', 100], ['Collection', 100]].forEach((field, index) => {
          if (String(row[index] == null ? '' : row[index]).trim().length > field[1])
            throw new Error(field[0] + ' must be ' + field[1] + ' characters or fewer.');
        });
        const title = text_(row[0], 160, true), url = text_(row[1], 2000, true);
        const topic = text_(row[2], 100, false), collection = text_(row[3], 100, false);
        const canonical = canonicalUrl_(url);
        if (known.has(canonical)) { duplicates++; results.push(['Already saved']); return; }
        known.add(canonical);
        additions.push([Utilities.getUuid(), url, canonical, title, topic, now, false, 20,
          0, '', '', '', 'bank', '', 0, '', collection, '']);
        imported++;
        results.push(['Imported']);
      } catch (error) {
        errors++;
        results.push(['Error: ' + error.message]);
      }
    });
    const requests = [];
    if (additions.length) requests.push({appendCells:{sheetId:data.sheetId,
      rows:additions.map(row => ({values:row.map(cell_)})), fields:'userEnteredValue'}});
    if (results.length) requests.push({updateCells:{start:{sheetId:sheet.getSheetId(), rowIndex:1, columnIndex:4},
      rows:results.map(row => ({values:row.map(cell_)})), fields:'userEnteredValue'}});
    // Question rows and their import results commit together. Retries deduplicate.
    if (requests.length) commit_(requests);
    return {imported:imported, duplicates:duplicates, errors:errors};
  });
}
function addQuestion(input) {
  assertOwner_();
  const url = text_(input.url, 2000, true),
    canonical = canonicalUrl_(url),
    title = text_(input.title, 160, true),
    topic = text_(input.topic, 100, false),
    minutes = int_(input.minutes, 5, 180);
  const notes = text_(input.notes, 5000, false);
  if (typeof input.initialHelp !== 'boolean') throw new Error('Record whether you used help.');
  return withLock_(function () {
    const data = rows_('Questions', Q_HEADERS),
      existing = data.rows.find(r => r[2] === canonical);
    if (existing) return {
      duplicate: true,
      id: existing[0],
      title: existing[3],
      deleted: existing[12] === 'deleted',
      status: existing[12]
    };
    const now = new Date(),
      id = Utilities.getUuid(),
      today = day_(now, settings_().timezone);
    const row = [id, url, canonical, title, topic, now.toISOString(), input.initialHelp, minutes, 0, plusDays_(today, 1), '', '', 'active', '', 0, notes, '', ''];
    commit_([append_(data.sheetId, row)]);
    return {
      duplicate: false,
      id: id,
      due: row[9]
    };
  });
}
function recordReview(input) {
  assertOwner_();
  const id = text_(input.id, 80, true),
    requestId = text_(input.requestId, 80, true),
    result = text_(input.result, 20, true),
    note = text_(input.note, 1500, false),
    minutes = int_(input.minutes, 0, 360),
    expected = int_(input.revision, 0, 1000000);
  if (!/^[a-zA-Z0-9_-]{16,80}$/.test(requestId)) throw new Error('Invalid review request. Refresh the page.');
  return withLock_(function () {
    const history = rows_('Reviews', R_HEADERS),
      prior = history.rows.find(r => r[0] === requestId);
    if (prior) {
      if (prior[1] !== id || prior[3] !== result || Number(prior[4]) !== minutes || prior[5] !== note) throw new Error('This request was already used for a different review. Refresh the page.');
      return {
        duplicate: true,
        due: prior[8]
      };
    }
    const data = rows_('Questions', Q_HEADERS),
      index = data.rows.findIndex(r => r[0] === id);
    if (index < 0) throw new Error('Question not found.');
    const row = data.rows[index].slice(),
      q = question_(row, index + 1),
      now = new Date(),
      today = day_(now, settings_().timezone);
    if (q.revision !== expected) throw new Error('This question changed in another tab. Refresh before recording a review.');
    if (!due_(q, today)) throw new Error('This question is not due. A same-day retry does not advance the schedule.');
    const next = transition_(q.stage, result, today);
    row[8] = next.stage;
    row[9] = next.due;
    row[10] = result;
    row[11] = now.toISOString();
    row[13] = '';
    row[14] = q.revision + 1;
    // The schedule and its history entry either both commit, or neither does.
    commit_([replace_(data.sheetId, index + 1, row), append_(history.sheetId, [requestId, id, now.toISOString(), result, minutes, note, q.stage, next.stage, next.due, q.revision, q.revision + 1])]);
    return {
      duplicate: false,
      due: next.due,
      days: next.days
    };
  });
}
function updateQuestion(input) {
  assertOwner_();
  const id = text_(input.id, 80, true),
    expected = int_(input.revision, 0, 1000000),
    action = text_(input.action, 20, true);
  if (['snooze', 'notes', 'delete', 'restore', 'stage', 'activate'].indexOf(action) < 0) throw new Error('Unknown question action.');
  const notes = action === 'notes' ? text_(input.notes, 5000, false) : null;
  const stage = ['stage', 'activate'].includes(action) ? int_(input.stage, 1, 7) - 1 : null;
  return withLock_(function () {
    const data = rows_('Questions', Q_HEADERS),
      i = data.rows.findIndex(r => r[0] === id);
    if (i < 0) throw new Error('Question not found.');
    const row = data.rows[i].slice();
    if (Number(row[14]) !== expected) throw new Error('This question changed. Refresh the page.');
    if (row[12] === 'deleted' && action !== 'restore') {
      throw new Error('This question is deleted. Restore it from Library first.');
    }
    if (action === 'restore' && row[12] !== 'deleted') {
      throw new Error('This question is already in your Library. Refresh the page.');
    }
    // Deletion is reversible. Keep the schedule and history, but stop selection
    // and reminders because only active questions are eligible.
    if (action === 'delete') {
      row[17] = row[12];
      row[12] = 'deleted';
    } else if (action === 'restore') {
      row[12] = row[17] === 'bank' ? 'bank' : 'active';
      row[17] = '';
      row[13] = '';
    } else if (action === 'notes') {
      row[15] = notes;
    } else if (action === 'snooze') {
      if (row[12] !== 'active') throw new Error('Add this question to practice first.');
      row[13] = plusDays_(day_(new Date(), settings_().timezone), 1);
    } else {
      const requiredStatus = action === 'activate' ? 'bank' : 'active';
      if (row[12] !== requiredStatus) throw new Error('This question changed. Refresh before changing its stage.');
      row[12] = 'active';
      row[8] = stage;
      row[9] = plusDays_(day_(new Date(), settings_().timezone), INTERVALS[stage]);
      row[13] = '';
    }
    row[14] = expected + 1;
    commit_([replace_(data.sheetId, i + 1, row)]);
    return {
      ok: true, question:summary_(question_(row, i + 1))
    };
  });
}
function saveSettings(input) {
  assertOwner_();
  const hour = int_(input.reminderHour, 0, 23),
    budget = int_(input.dailyBudget, 5, 240),
    channel = text_(input.channel, 20, true);
  if (['off', 'email', 'telegram'].indexOf(channel) < 0) throw new Error('Unknown reminder channel.');
  if (channel === 'telegram') {
    const p = PropertiesService.getScriptProperties();
    if (!p.getProperty('TELEGRAM_BOT_TOKEN') || !p.getProperty('TELEGRAM_CHAT_ID')) throw new Error('Add Telegram credentials in Apps Script Project Settings first.');
  }
  return withLock_(function () {
    ensureTrigger_();
    const data = rows_('Settings'),
      changes = {
        reminderHour: hour,
        dailyBudget: budget,
        channel: channel
      };
    commit_(data.rows.flatMap((r, i) => Object.prototype.hasOwnProperty.call(changes, r[0]) ? [replace_(data.sheetId, i + 1, [r[0], changes[r[0]], r[2]])] : []));
    return {
      ok: true
    };
  });
}
function ensureTrigger_() {
  const existing = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'sendDailyDigest_');
  if (!existing.length) ScriptApp.newTrigger('sendDailyDigest_').timeBased().everyHours(1).create();
}
function install() {
  assertOwner_();
  return withLock_(function () {
    rows_('Questions', Q_HEADERS);
    rows_('Reviews', R_HEADERS);
    rows_('Deliveries', D_HEADERS);
    settings_();
    book_().setSpreadsheetTimeZone('Asia/Singapore');
    ensureTrigger_();
    const guide = book_().getSheetByName('Guide');
    guide.getRange('B11').setValue('Apps Script installed. Open the private web app to begin.');
    console.log('Practice Loop installed. Reminders stay off until enabled in Settings.');
    return {
      ok: true
    };
  });
}
function sendDailyDigest_() {
  if (Session.getEffectiveUser().getEmail().toLowerCase() !== OWNER_EMAIL) throw new Error('Owner required.');
  return withLock_(function () {
    const cfg = settings_(),
      now = new Date(),
      today = day_(now, cfg.timezone);
    if (cfg.channel === 'off' || Number(Utilities.formatDate(now, cfg.timezone, 'H')) < cfg.reminderHour) return;
    const plan = chooseSession_(rows_('Questions', Q_HEADERS).rows.map((r, i) => question_(r, i + 1)), today, cfg.dailyBudget);
    if (!plan.due.length) return;
    const log = rows_('Deliveries', D_HEADERS),
      key = today; // One digest per date, including a channel switch.
    if (log.rows.some(r => r[0] === key)) return;
    const appUrl = ScriptApp.getService().getUrl();
    if (!appUrl) throw new Error('Deploy the private web app before enabling reminders.');
    const props = PropertiesService.getScriptProperties(),
      token = props.getProperty('TELEGRAM_BOT_TOKEN'),
      chatId = props.getProperty('TELEGRAM_CHAT_ID');
    if (cfg.channel === 'telegram' && (!token || !chatId)) throw new Error('Telegram credentials are missing.');
    const attempt = [key, Utilities.getUuid(), now.toISOString(), cfg.channel, 'attempting', '', 'Delivery may be uncertain if this run stops. No automatic same-day retry.'];
    commit_([append_(log.sheetId, attempt)]);
    const message = 'Practice Loop\n' + plan.due.length + ' question' + (plan.due.length === 1 ? ' is' : 's are') + ' due.\nSuggested session: ' + plan.selected.length + ' question' + (plan.selected.length === 1 ? '' : 's') + ', about ' + plan.minutes + ' minutes.\n\nOpen today\'s practice:\n' + appUrl;
    try {
      if (cfg.channel === 'email') MailApp.sendEmail({
        to: OWNER_EMAIL,
        subject: 'Practice Loop: ' + plan.due.length + ' questions due',
        body: message,
        name: 'Practice Loop'
      });else {
        const response = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
          method: 'post',
          contentType: 'application/json',
          payload: JSON.stringify({
            chat_id: chatId,
            text: message,
            disable_web_page_preview: true
          }),
          muteHttpExceptions: true
        });
        if (response.getResponseCode() !== 200 || !JSON.parse(response.getContentText()).ok) throw new Error('Delivery rejected.');
      }
      attempt[4] = 'sent';
      attempt[5] = new Date().toISOString();
      attempt[6] = 'Provider accepted the message.';
    } catch (_) {
      // Never record the raw fetch error: it may contain the bot token.
      attempt[4] = 'uncertain';
      attempt[5] = new Date().toISOString();
      attempt[6] = 'Check your inbox or Telegram. No automatic same-day retry.';
    }
    commit_([replace_(log.sheetId, log.rows.length + 1, attempt)]);
  });
}
