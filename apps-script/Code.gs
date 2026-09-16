const SHEET_ID = '1ymtxttoGIbpKqHjlByKQoHgs4o058c2-JGYGAhIZHaM';
const REMINDER_EMAIL = 'snjokallinn@snjokallinn.com';
const TZ = 'Atlantic/Reykjavik';

const TEMPLATES = {
  'Tónlistarmyndband': [
    [2, 'Tökur komnar í tölvu'],
    [14, 'Rough cut tilbúið'],
    [21, 'Læst klippa'],
    [28, 'Litgreining og textar'],
    [35, 'Myndband tilbúið']
  ],
  'Portrait-myndataka': [
    [7, 'Myndaval tilbúið'],
    [21, 'Fullunnar myndir afhentar']
  ],
  'Brúðkaup': [
    [2, 'Sneak peeks'],
    [42, 'Fullt gallerí afhent']
  ],
  'Viðburður': [
    [14, 'Fullunnar myndir']
  ],
  'Efni fyrir samfélagsmiðla': [
    [7, 'Fyrsta afhending']
  ]
};

function doGet() {
  return ContentService.createTextOutput('Snjókallinn Verkfæri endpoint is running.');
}

function doPost(e) {
  const expectedKey = PropertiesService.getScriptProperties().getProperty('WRITE_KEY');
  if (!expectedKey || String((e.parameter && e.parameter.key) || '') !== expectedKey) {
    return json_({ ok: false, error: 'Unauthorized' });
  }

  const name = clean_(e.parameter.name);
  const type = clean_(e.parameter.type);
  const size = clean_(e.parameter.size);
  const notes = clean_(e.parameter.notes);
  const projectDate = parseDate_(e.parameter.date);

  if (!name || !TEMPLATES[type]) {
    return json_({ ok: false, error: 'Invalid project data' });
  }

  const projectId = 'P-' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmm') + '-' + Utilities.getUuid().slice(0, 4).toUpperCase();
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const projects = ss.getSheetByName('Verkefni');
  const milestones = ss.getSheetByName('Áfangar');

  projects.appendRow([
    projectId,
    new Date(),
    name,
    type,
    projectDate,
    size,
    notes,
    'Í vinnslu'
  ]);

  const planLines = [];
  TEMPLATES[type].forEach(([days, label]) => {
    const due = addDays_(projectDate, days);
    milestones.appendRow([
      projectId,
      name,
      label,
      due,
      'Ólokið',
      '',
      ''
    ]);
    planLines.push('• ' + formatDate_(due) + ' — ' + label);
  });

  MailApp.sendEmail({
    to: REMINDER_EMAIL,
    subject: 'Nýtt verkefni: ' + name,
    body: [
      'Snjókallinn Verkfæri skráði nýtt verkefni.',
      '',
      name + ' · ' + type,
      'Verkefnisdagur: ' + formatDate_(projectDate),
      'Umfang: ' + size,
      notes ? 'Athugasemdir: ' + notes : '',
      '',
      'Áfangar:',
      ...planLines
    ].filter(Boolean).join('\n')
  });

  return json_({ ok: true, projectId: projectId });
}

function sendDailyProjectReminders() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Áfangar');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const values = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  const today = startOfDay_(new Date());
  const items = [];

  values.forEach((row, i) => {
    const [projectId, projectName, label, dueRaw, status, sentRaw] = row;
    if (!projectId || String(status).toLowerCase() === 'lokið') return;

    const due = startOfDay_(new Date(dueRaw));
    if (isNaN(due.getTime())) return;

    const diff = Math.round((due.getTime() - today.getTime()) / 86400000);
    let marker = '';
    let prefix = '';

    if (diff === 2) {
      marker = '2d';
      prefix = 'Eftir 2 daga';
    } else if (diff === 0) {
      marker = 'due';
      prefix = 'Í dag';
    } else if (diff === -2) {
      marker = 'late2';
      prefix = '2 dögum yfir tíma';
    } else {
      return;
    }

    const sent = String(sentRaw || '').split(',').map(s => s.trim()).filter(Boolean);
    if (sent.includes(marker)) return;

    items.push({
      rowNumber: i + 2,
      projectId,
      projectName,
      label,
      due,
      marker,
      prefix,
      sent
    });
  });

  if (!items.length) return;

  MailApp.sendEmail({
    to: REMINDER_EMAIL,
    subject: 'Snjókallinn · verkefnaáminning (' + items.length + ')',
    body: [
      'Hér er það sem þarf athygli í verkefnunum þínum:',
      '',
      ...items.map(item =>
        item.prefix + ': ' + item.projectName + ' — ' + item.label + ' (' + formatDate_(item.due) + ')'
      )
    ].join('\n')
  });

  items.forEach(item => {
    const newSent = item.sent.concat(item.marker).join(',');
    sheet.getRange(item.rowNumber, 6).setValue(newSent);
    sheet.getRange(item.rowNumber, 7).setValue(new Date());
  });
}

function setupDailyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendDailyProjectReminders')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('sendDailyProjectReminders')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
}

function addDays_(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + Number(days));
  return d;
}

function parseDate_(value) {
  const s = String(value || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return new Date();
  const [y, m, d] = s.split('-').map(Number);
  return new Date(y, m - 1, d, 12, 0, 0);
}

function startOfDay_(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  return d;
}

function formatDate_(date) {
  return Utilities.formatDate(new Date(date), TZ, 'dd.MM.yyyy');
}

function clean_(value) {
  return String(value || '').trim().slice(0, 2000);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
