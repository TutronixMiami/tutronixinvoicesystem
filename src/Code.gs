const TUTRONIX = Object.freeze({
  spreadsheetId: '1FA7OljhKUpDeJJptKrrspOkJz8fjSNoIpA4CEuaWBYQ',
  sheets: {
    dashboard: 'Dashboard',
    schedule: 'Schedule',
    students: 'Students',
    parents: 'Parents',
    tutors: 'Tutors',
    invoices: 'Invoices',
    invoiceSessions: 'Invoice Sessions',
    settings: 'Settings'
  },
  chargeableStatuses: ['Completed', 'Cancelled - Charge', 'No-show - Charge']
});

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Tutronix')
    .addItem('Dashboard', 'openDashboard')
    .addItem('Schedule', 'openSchedule')
    .addSeparator()
    .addItem('Prepare invoices', 'showInvoiceSidebar')
    .addToUi();
}

function openDashboard() {
  activateSheet_(TUTRONIX.sheets.dashboard);
}

function openSchedule() {
  activateSheet_(TUTRONIX.sheets.schedule);
}

function activateSheet_(name) {
  const sheet = getBook_().getSheetByName(name);
  if (sheet) sheet.activate();
}

function getBook_() {
  return SpreadsheetApp.getActiveSpreadsheet() ||
    SpreadsheetApp.openById(TUTRONIX.spreadsheetId);
}

function onEdit(e) {
  if (!e || !e.range || e.range.getRow() < 2) return;
  const sheet = e.range.getSheet();
  const name = sheet.getName();
  const firstRow = e.range.getRow();
  const lastRow = e.range.getLastRow();

  if (name === TUTRONIX.sheets.schedule) {
    for (let row = firstRow; row <= lastRow; row++) syncScheduleRow_(sheet, row, e.range.getColumn(), e.range.getLastColumn());
  } else if (name === TUTRONIX.sheets.students) {
    for (let row = firstRow; row <= lastRow; row++) initializeRecordRow_(sheet, row, 'STU', 2, 8);
  } else if (name === TUTRONIX.sheets.parents) {
    for (let row = firstRow; row <= lastRow; row++) initializeRecordRow_(sheet, row, 'PAR', 2, 8);
  } else if (name === TUTRONIX.sheets.tutors) {
    for (let row = firstRow; row <= lastRow; row++) initializeRecordRow_(sheet, row, 'TUT', 2, 5);
  }
}

function initializeRecordRow_(sheet, row, prefix, nameColumn, activeColumn) {
  if (!sheet.getRange(row, nameColumn).getValue()) return;
  const idCell = sheet.getRange(row, 1);
  if (!idCell.getValue()) idCell.setValue(makeId_(prefix));
  const activeCell = sheet.getRange(row, activeColumn);
  if (!activeCell.getValue()) activeCell.setValue('Yes');
}

function syncScheduleRow_(sheet, row, editedStartColumn, editedEndColumn) {
  const range = sheet.getRange(row, 1, 1, 18);
  const values = range.getValues()[0];
  const date = values[1];
  const start = values[2];
  const end = values[3];
  const student = String(values[5] || '').trim();
  const tutor = String(values[6] || '').trim();
  const hasEntry = date || start || end || student || tutor || values[8];
  const firstEditedColumn = editedStartColumn || 1;
  const lastEditedColumn = editedEndColumn || firstEditedColumn;
  const studentEdited = firstEditedColumn <= 6 && lastEditedColumn >= 6;
  const modeEdited = firstEditedColumn <= 8 && lastEditedColumn >= 8;

  if (!hasEntry) return;
  if (!values[0]) values[0] = makeId_('SES');
  if (!values[9]) values[9] = 'Scheduled';

  if (start instanceof Date && end instanceof Date) {
    values[4] = Math.round((((end - start) / 3600000 + 24) % 24) * 100) / 100;
  } else {
    values[4] = '';
  }

  const studentRecord = student ? findRecord_(TUTRONIX.sheets.students, 2, student) : null;
  if (studentRecord) {
    const parentName = studentRecord[2] || '';
    values[13] = parentName;
    if (!values[6] && studentRecord[5]) values[6] = studentRecord[5];
    if (!values[8] && studentRecord[6]) values[8] = studentRecord[6];

    const parentRecord = parentName ? findRecord_(TUTRONIX.sheets.parents, 2, parentName) : null;
    const suggestedRate = parentRateForMode_(parentRecord, values[7]);
    if ((studentEdited || modeEdited || values[10] === '' || values[10] == null) && suggestedRate !== '') {
      values[10] = suggestedRate;
    }
  }

  const selectedTutor = String(values[6] || '').trim();
  const tutorRecord = selectedTutor ? findRecord_(TUTRONIX.sheets.tutors, 2, selectedTutor) : null;
  if ((values[11] === '' || values[11] == null) && tutorRecord) values[11] = tutorRecord[5] || '';

  const hours = values[14] !== '' && values[14] != null ? Number(values[14]) : Number(values[4] || 0);
  const chargeable = TUTRONIX.chargeableStatuses.indexOf(values[9]) !== -1;
  values[15] = chargeable ? roundMoney_(hours * Number(values[10] || 0)) : 0;
  values[16] = values[9] === 'Completed' ? roundMoney_(hours * Number(values[11] || 0)) : 0;

  range.setValues([values]);
}

function findRecord_(sheetName, keyColumn, key) {
  const sheet = getBook_().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const wanted = String(key).trim().toLowerCase();
  return values.find(row => String(row[keyColumn - 1]).trim().toLowerCase() === wanted) || null;
}

function parentRateForMode_(parentRecord, mode) {
  if (!parentRecord) return '';
  const inPersonRate = parentRecord[5];
  const virtualRate = parentRecord[6];
  const isVirtual = String(mode || '').trim().toLowerCase() === 'virtual';
  if (isVirtual) return virtualRate !== '' && virtualRate != null ? virtualRate : inPersonRate;
  return inPersonRate !== '' && inPersonRate != null ? inPersonRate : virtualRate;
}

function showInvoiceSidebar() {
  const html = HtmlService.createTemplateFromFile('InvoiceSidebar')
    .evaluate()
    .setTitle('Prepare Invoices');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getOutstandingStudents() {
  const sheet = getBook_().getSheetByName(TUTRONIX.sheets.schedule);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 18).getValues();
  const grouped = {};

  rows.forEach(row => {
    const student = String(row[5] || '').trim();
    const status = String(row[9] || '');
    const invoiceId = String(row[17] || '').trim();
    if (!student || invoiceId || TUTRONIX.chargeableStatuses.indexOf(status) === -1) return;

    const parent = String(row[13] || 'Unassigned parent');
    const hours = Number(row[14] !== '' ? row[14] : row[4]) || 0;
    const amount = Number(row[15]) || roundMoney_(hours * Number(row[10] || 0));
    const key = parent + '|' + student;
    if (!grouped[key]) grouped[key] = { parent, student, sessions: 0, hours: 0, amount: 0 };
    grouped[key].sessions++;
    grouped[key].hours += hours;
    grouped[key].amount += amount;
  });

  return Object.keys(grouped).map(key => {
    const item = grouped[key];
    item.hours = Math.round(item.hours * 100) / 100;
    item.amount = roundMoney_(item.amount);
    return item;
  }).sort((a, b) => a.parent.localeCompare(b.parent) || a.student.localeCompare(b.student));
}

function createInvoices(studentKeys) {
  if (!studentKeys || !studentKeys.length) throw new Error('Select at least one student.');
  const selected = new Set(studentKeys);
  const book = getBook_();
  const schedule = book.getSheetByName(TUTRONIX.sheets.schedule);
  const invoices = book.getSheetByName(TUTRONIX.sheets.invoices);
  const lines = book.getSheetByName(TUTRONIX.sheets.invoiceSessions);
  const rows = schedule.getRange(2, 1, Math.max(schedule.getLastRow() - 1, 1), 18).getValues();
  const grouped = {};

  rows.forEach((row, index) => {
    const student = String(row[5] || '').trim();
    const parent = String(row[13] || '').trim();
    const status = String(row[9] || '');
    const key = parent + '|' + student;
    if (!selected.has(key) || row[17] || TUTRONIX.chargeableStatuses.indexOf(status) === -1) return;
    if (!parent) throw new Error('A selected student is missing a parent.');
    if (!grouped[parent]) grouped[parent] = [];
    grouped[parent].push({ row, sheetRow: index + 2 });
  });

  const parents = Object.keys(grouped);
  if (!parents.length) throw new Error('No eligible uninvoiced sessions were found.');

  const created = [];
  parents.forEach(parent => {
    const invoiceId = nextInvoiceId_();
    const sessions = grouped[parent];
    const students = [...new Set(sessions.map(item => item.row[5]))];
    let hours = 0;
    let total = 0;
    const lineRows = [];

    sessions.forEach(item => {
      const row = item.row;
      const duration = Number(row[14] !== '' ? row[14] : row[4]) || 0;
      const amount = Number(row[15]) || roundMoney_(duration * Number(row[10] || 0));
      hours += duration;
      total += amount;
      lineRows.push([invoiceId, row[0], row[5], row[1], row[2], row[3], duration, row[8], row[10], amount]);
      schedule.getRange(item.sheetRow, 18).setValue(invoiceId);
    });

    invoices.appendRow([
      invoiceId,
      parent,
      students.join(', '),
      new Date(),
      Math.round(hours * 100) / 100,
      roundMoney_(total),
      'Created',
      '',
      '',
      '',
      ''
    ]);
    if (lineRows.length) lines.getRange(lines.getLastRow() + 1, 1, lineRows.length, 10).setValues(lineRows);
    created.push(invoiceId);
  });

  return { created, message: created.length + ' invoice' + (created.length === 1 ? '' : 's') + ' created.' };
}

function nextInvoiceId_() {
  const book = getBook_();
  const invoices = book.getSheetByName(TUTRONIX.sheets.invoices);
  const dateKey = Utilities.formatDate(new Date(), book.getSpreadsheetTimeZone(), 'MMddyy');
  let highestSequence = 0;

  if (invoices && invoices.getLastRow() >= 2) {
    const existingIds = invoices.getRange(2, 1, invoices.getLastRow() - 1, 1).getDisplayValues();
    existingIds.forEach(row => {
      const invoiceId = String(row[0] || '').trim();
      if (invoiceId.indexOf(dateKey) !== 0) return;
      const sequence = Number(invoiceId.slice(dateKey.length));
      if (Number.isFinite(sequence)) highestSequence = Math.max(highestSequence, sequence);
    });
  }

  return dateKey + String(highestSequence + 1).padStart(2, '0');
}

function makeId_(prefix) {
  return prefix + '-' + Utilities.getUuid().split('-')[0].toUpperCase();
}

function roundMoney_(value) {
  return Math.round((Number(value) || 0) * 100) / 100;
}
