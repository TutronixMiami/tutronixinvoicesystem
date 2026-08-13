const TUTRONIX = Object.freeze({
  spreadsheetId: '1FA7OljhKUpDeJJptKrrspOkJz8fjSNoIpA4CEuaWBYQ',
  sheets: {
    dashboard: 'Dashboard',
    schedule: 'Schedule',
    sessionHistory: 'Session History',
    students: 'Students',
    parents: 'Parents',
    tutors: 'Tutors',
    invoices: 'Invoices',
    invoiceSessions: 'Invoice Sessions',
    settings: 'Settings',
    availability: 'Availability',
    bookingRequests: 'Booking Requests'
  },
  chargeableStatuses: ['Completed', 'Cancelled - Charge', 'No-show - Charge']
});

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Tutronix')
    .addItem('Dashboard', 'openDashboard')
    .addItem('Schedule', 'openSchedule')
    .addItem('Session history', 'openSessionHistory')
    .addItem('Availability', 'openAvailability')
    .addItem('Booking requests', 'openBookingRequests')
    .addSeparator()
    .addItem('Prepare invoices', 'showInvoiceSidebar')
    .addSeparator()
    .addItem('Approve selected booking', 'approveSelectedBooking')
    .addItem('Decline selected booking', 'declineSelectedBooking')
    .addItem('Refresh parent booking links', 'refreshParentBookingLinks')
    .addItem('Enable automatic booking decisions', 'enableAutomaticBookingDecisions')
    .addToUi();
}

function openDashboard() {
  activateSheet_(TUTRONIX.sheets.dashboard);
}

function openSchedule() {
  activateSheet_(TUTRONIX.sheets.schedule);
}

function openSessionHistory() {
  activateSheet_(TUTRONIX.sheets.sessionHistory);
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
  } else if (name === TUTRONIX.sheets.availability) {
    for (let row = firstRow; row <= lastRow; row++) initializeAvailabilityRow_(sheet, row);
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
  const timeEdited = firstEditedColumn <= 4 && lastEditedColumn >= 3;

  if (!hasEntry) return;
  if (!values[0]) values[0] = makeId_('SES');
  if (!values[9]) values[9] = 'Scheduled';

  if (start instanceof Date && end instanceof Date) {
    values[4] = Math.round((((end - start) / 3600000 + 24) % 24) * 100) / 100;
    if (timeEdited || values[14] === '' || values[14] == null) values[14] = values[4];
  } else {
    values[4] = '';
    if (timeEdited) values[14] = '';
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

    const parent = resolveSessionParent_(row);
    const hours = Number(row[4]) || 0;
    const amount = roundMoney_(hours * Number(row[10] || 0));
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

function resolveSessionParent_(row) {
  const savedParent = String(row[13] || '').trim();
  if (savedParent) return savedParent;
  const student = String(row[5] || '').trim();
  const studentRecord = student ? findRecord_(TUTRONIX.sheets.students, 2, student) : null;
  return studentRecord && String(studentRecord[2] || '').trim()
    ? String(studentRecord[2]).trim()
    : 'Unassigned parent';
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
    const parent = resolveSessionParent_(row);
    const status = String(row[9] || '');
    const key = parent + '|' + student;
    if (!selected.has(key) || row[17] || TUTRONIX.chargeableStatuses.indexOf(status) === -1) return;
    if (parent === 'Unassigned parent') throw new Error(student + ' is missing a parent. Add the parent on the Students tab, then reopen Prepare invoices.');
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
      const duration = Number(row[4]) || 0;
      const amount = roundMoney_(duration * Number(row[10] || 0));
      hours += duration;
      total += amount;
      lineRows.push([invoiceId, row[0], row[5], row[1], row[2], row[3], duration, row[8], row[10], amount]);
      schedule.getRange(item.sheetRow, 18).setNumberFormat('@').setValue(invoiceId);
    });

    const invoiceRow = invoices.getLastRow() + 1;
    invoices.getRange(invoiceRow, 1, 1, 11).setValues([[
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
    ]]);
    invoices.getRange(invoiceRow, 1).setNumberFormat('@').setValue(invoiceId);
    invoices.getRange(invoiceRow, 4).setNumberFormat('mmm d, yyyy');

    if (lineRows.length) {
      const lineStartRow = lines.getLastRow() + 1;
      lines.getRange(lineStartRow, 1, lineRows.length, 10).setValues(lineRows);
      lines.getRange(lineStartRow, 1, lineRows.length, 1)
        .setNumberFormat('@')
        .setValues(lineRows.map(row => [String(row[0])]));
      lines.getRange(lineStartRow, 4, lineRows.length, 1).setNumberFormat('mmm d, yyyy');
      lines.getRange(lineStartRow, 5, lineRows.length, 2).setNumberFormat('h:mm AM/PM');
    }
    const pdf = generateInvoicePdfById_(invoiceId);
    invoices.getRange(invoiceRow, 10).setValue(pdf.url);
    archiveInvoicedSessions_(sessions.map(item => String(item.row[0] || '').trim()), invoiceId);
    created.push(invoiceId);
  });

  return { created, message: created.length + ' invoice' + (created.length === 1 ? '' : 's') + ' created with PDF.' };
}

function archiveInvoicedSessions_(sessionIds, invoiceId) {
  const ids = [...new Set((sessionIds || []).map(id => String(id || '').trim()).filter(Boolean))];
  if (!ids.length) return;

  const book = getBook_();
  const schedule = book.getSheetByName(TUTRONIX.sheets.schedule);
  const history = book.getSheetByName(TUTRONIX.sheets.sessionHistory);
  if (!schedule || !history) throw new Error('Schedule or Session History is unavailable. No sessions were removed.');
  if (schedule.getLastRow() < 2) throw new Error('The invoiced sessions could not be found in Schedule.');

  const values = schedule.getRange(2, 1, schedule.getLastRow() - 1, 18).getValues();
  const idSet = new Set(ids);
  const matches = [];
  values.forEach((row, index) => {
    const sessionId = String(row[0] || '').trim();
    if (!idSet.has(sessionId)) return;
    if (String(row[17] || '').trim() !== String(invoiceId).trim()) {
      throw new Error('Session ' + sessionId + ' is not linked to invoice ' + invoiceId + '. No sessions were removed.');
    }
    matches.push({ sheetRow: index + 2, values: row });
  });

  if (matches.length !== ids.length) {
    throw new Error('Not every invoiced session was found in Schedule. No sessions were removed.');
  }

  const archivedAt = new Date();
  const archiveRows = matches.map(item => item.values.concat([archivedAt]));
  const historyStartRow = history.getLastRow() + 1;
  history.getRange(historyStartRow, 1, archiveRows.length, 19).setValues(archiveRows);

  matches
    .map(item => item.sheetRow)
    .sort((a, b) => b - a)
    .forEach(sheetRow => schedule.deleteRow(sheetRow));
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

function generatePdfForSelectedInvoice() {
  const book = getBook_();
  const sheet = book.getActiveSheet();
  const row = Math.max(sheet.getActiveRange().getRow(), 2);
  if (sheet.getName() !== TUTRONIX.sheets.invoices) {
    throw new Error('Select an invoice row in the Invoices tab first.');
  }
  const invoiceId = String(sheet.getRange(row, 1).getDisplayValue()).trim();
  if (!invoiceId) throw new Error('The selected row does not have an invoice number.');
  const result = generateInvoicePdfById_(invoiceId);
  sheet.getRange(row, 10).setValue(result.url);
  SpreadsheetApp.getUi().alert('Invoice PDF ready', 'The branded PDF for invoice ' + invoiceId + ' has been created and linked.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function generateInvoicePdfById_(invoiceId) {
  const book = getBook_();
  const invoices = book.getSheetByName(TUTRONIX.sheets.invoices);
  const invoiceRows = invoices.getRange(2, 1, Math.max(invoices.getLastRow() - 1, 1), 11).getValues();
  const invoice = invoiceRows.find(row => String(row[0]).trim() === String(invoiceId).trim());
  if (!invoice) throw new Error('Invoice ' + invoiceId + ' was not found.');

  const linesSheet = book.getSheetByName(TUTRONIX.sheets.invoiceSessions);
  const lineRows = linesSheet.getRange(2, 1, Math.max(linesSheet.getLastRow() - 1, 1), 10).getValues()
    .filter(row => String(row[0]).trim() === String(invoiceId).trim());

  const invoiceData = {
    invoiceId: String(invoice[0]),
    parent: String(invoice[1] || ''),
    students: String(invoice[2] || ''),
    createdDate: invoice[3] instanceof Date ? invoice[3] : new Date(),
    totalHours: Number(invoice[4]) || 0,
    totalAmount: Number(invoice[5]) || 0,
    sessions: lineRows
  };
  const pdf = createBrandedInvoicePdf_(invoiceData);
  emailInvoiceCopy_(invoiceData, pdf);
  return pdf;
}

function emailInvoiceCopy_(invoice, pdf) {
  const recipient = 'tutronixmiami@gmail.com';
  const subject = 'Tutronix Invoice ' + invoice.invoiceId + ' - ' + invoice.parent;
  const body = [
    'A Tutronix invoice has been created and is ready for review.',
    '',
    'Invoice: ' + invoice.invoiceId,
    'Parent/Guardian: ' + invoice.parent,
    'Student(s): ' + invoice.students,
    'Total hours: ' + invoice.totalHours.toFixed(2),
    'Amount: $' + invoice.totalAmount.toFixed(2),
    '',
    'The client-ready PDF is attached. This message was sent only to the internal Tutronix email address.'
  ].join('\n');

  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    body: body,
    name: 'Tutronix Method',
    attachments: [DriveApp.getFileById(pdf.id).getBlob()]
  });
}

function createBrandedInvoicePdf_(invoice) {
  const book = getBook_();
  const timeZone = book.getSpreadsheetTimeZone();
  const parentRecord = findRecord_(TUTRONIX.sheets.parents, 2, invoice.parent) || [];
  const settings = getSettings_();
  const businessName = settings['Business Name'] || 'Tutronix';
  const terms = Number(settings['Payment Terms (Days)']) || 14;
  const paymentInstructions = settings['Payment Instructions'] ||
    'Payments can be made by Zelle to tutronixmiami@gmail.com. Payments can also be made by check payable to Tutronix LLC.';
  const dueDate = new Date(invoice.createdDate);
  dueDate.setDate(dueDate.getDate() + terms);

  const forest = '#183D34';
  const cream = '#F5EFE2';
  const coral = '#EF7466';
  const yellow = '#F4C94E';
  const paleBlue = '#C7E3E8';
  const white = '#FFFFFF';
  const doc = DocumentApp.create('Invoice ' + invoice.invoiceId + ' - ' + invoice.parent);
  const body = doc.getBody();
  body.clear();
  body.setMarginTop(32).setMarginBottom(32).setMarginLeft(38).setMarginRight(38);
  body.setAttributes({
    FONT_FAMILY: 'Arial',
    FOREGROUND_COLOR: forest,
    FONT_SIZE: 9
  });

  const brandTable = body.appendTable([['', 'INVOICE']]);
  brandTable.setBorderWidth(0);
  const brandLeft = brandTable.getCell(0, 0);
  try {
    const logoBlob = UrlFetchApp.fetch('https://tutronixmethod.com/tutronix-logo.png').getBlob();
    const image = brandLeft.getChild(0).asParagraph().appendInlineImage(logoBlob);
    image.setWidth(54).setHeight(54);
  } catch (error) {
    brandLeft.setText('TUTRONIX');
  }
  const brandName = brandLeft.appendParagraph('TUTRONIX METHOD');
  brandName.editAsText().setBold(true).setFontSize(11).setForegroundColor(forest);
  const invoiceCell = brandTable.getCell(0, 1);
  invoiceCell.setBackgroundColor(forest);
  const invoiceTitle = invoiceCell.getChild(0).asParagraph();
  invoiceTitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  invoiceTitle.editAsText().setBold(true).setFontSize(24).setForegroundColor(white);
  invoiceCell.appendParagraph('#' + invoice.invoiceId)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .editAsText().setFontSize(11).setForegroundColor(white);

  body.appendParagraph('');
  const meta = body.appendTable([
    ['BILL TO', 'INVOICE DETAILS'],
    [invoice.parent, 'Invoice date: ' + Utilities.formatDate(invoice.createdDate, timeZone, 'MMMM d, yyyy')],
    [String(parentRecord[2] || ''), 'Due date: ' + Utilities.formatDate(dueDate, timeZone, 'MMMM d, yyyy')],
    [String(parentRecord[4] || ''), 'Terms: Net ' + terms]
  ]);
  meta.setBorderWidth(0);
  meta.getRow(0).getCell(0).setBackgroundColor(paleBlue);
  meta.getRow(0).getCell(1).setBackgroundColor(yellow);
  meta.getRow(0).getCell(0).editAsText().setBold(true).setForegroundColor(forest);
  meta.getRow(0).getCell(1).editAsText().setBold(true).setForegroundColor(forest);
  meta.getCell(1, 0).editAsText().setBold(true).setFontSize(11);
  for (let r = 1; r < 4; r++) meta.getCell(r, 1).getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

  body.appendParagraph('');
  const tableData = [['DATE', 'STUDENT', 'TIME', 'SUBJECT', 'HOURS', 'RATE', 'AMOUNT']];
  invoice.sessions.forEach(row => {
    const sessionDate = row[3] instanceof Date ? Utilities.formatDate(row[3], timeZone, 'MMM d, yyyy') : String(row[3] || '');
    const start = row[4] instanceof Date ? Utilities.formatDate(row[4], timeZone, 'h:mm a') : String(row[4] || '');
    const end = row[5] instanceof Date ? Utilities.formatDate(row[5], timeZone, 'h:mm a') : String(row[5] || '');
    tableData.push([
      sessionDate,
      String(row[2] || ''),
      start + ' - ' + end,
      String(row[7] || ''),
      Number(row[6] || 0).toFixed(2),
      '$' + Number(row[8] || 0).toFixed(2),
      '$' + Number(row[9] || 0).toFixed(2)
    ]);
  });
  const sessionsTable = body.appendTable(tableData);
  sessionsTable.setBorderColor('#D7D1C5').setBorderWidth(1);
  const header = sessionsTable.getRow(0);
  for (let c = 0; c < 7; c++) {
    header.getCell(c).setBackgroundColor(forest);
    header.getCell(c).editAsText().setBold(true).setFontSize(8).setForegroundColor(white);
  }
  for (let r = 1; r < sessionsTable.getNumRows(); r++) {
    if (r % 2 === 0) {
      for (let c = 0; c < 7; c++) sessionsTable.getCell(r, c).setBackgroundColor(cream);
    }
    for (let c = 4; c < 7; c++) {
      sessionsTable.getCell(r, c).getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    }
  }

  body.appendParagraph('');
  const totals = body.appendTable([
    ['TOTAL HOURS', invoice.totalHours.toFixed(2)],
    ['AMOUNT DUE', '$' + invoice.totalAmount.toFixed(2)]
  ]);
  totals.setBorderWidth(0);
  totals.getCell(0, 0).editAsText().setBold(true);
  totals.getCell(0, 1).getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  totals.getCell(1, 0).setBackgroundColor(forest).editAsText().setBold(true).setForegroundColor(white).setFontSize(12);
  totals.getCell(1, 1).setBackgroundColor(forest).editAsText().setBold(true).setForegroundColor(white).setFontSize(12);
  totals.getCell(1, 1).getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

  body.appendParagraph('');
  const payment = body.appendTable([['PAYMENT INSTRUCTIONS'], [paymentInstructions]]);
  payment.setBorderColor(forest).setBorderWidth(1);
  payment.getCell(0, 0).setBackgroundColor(cream);
  payment.getCell(0, 0).editAsText().setBold(true).setForegroundColor(forest);
  payment.getCell(1, 0).setBackgroundColor(white);

  body.appendParagraph('');
  const footer = body.appendParagraph('Thank you for choosing ' + businessName + '.');
  footer.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  footer.editAsText().setBold(true).setFontSize(10).setForegroundColor(forest);
  const subfooter = body.appendParagraph('Tutoring + Academic Coaching + Executive Function');
  subfooter.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  subfooter.editAsText().setFontSize(8).setForegroundColor(forest);

  doc.saveAndClose();
  const sourceFile = DriveApp.getFileById(doc.getId());
  const folder = getInvoiceFolder_();
  const pdfFile = folder.createFile(sourceFile.getAs(MimeType.PDF))
    .setName('Tutronix Invoice ' + invoice.invoiceId + ' - ' + invoice.parent + '.pdf');
  sourceFile.setTrashed(true);
  return { id: pdfFile.getId(), url: pdfFile.getUrl(), name: pdfFile.getName() };
}

function getInvoiceFolder_() {
  const propertyKey = 'TUTRONIX_INVOICE_FOLDER_ID';
  const properties = PropertiesService.getScriptProperties();
  const existingId = properties.getProperty(propertyKey);
  if (existingId) {
    try {
      return DriveApp.getFolderById(existingId);
    } catch (error) {}
  }
  const folders = DriveApp.getFoldersByName('Tutronix Invoices');
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('Tutronix Invoices');
  properties.setProperty(propertyKey, folder.getId());
  return folder;
}

function getSettings_() {
  const sheet = getBook_().getSheetByName(TUTRONIX.sheets.settings);
  const result = {};
  if (!sheet || sheet.getLastRow() < 2) return result;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getDisplayValues().forEach(row => {
    const key = String(row[0] || '').trim();
    if (key) result[key] = String(row[1] || '').trim();
  });
  return result;
}

function openAvailability() {
  activateSheet_(TUTRONIX.sheets.availability);
}

function openBookingRequests() {
  activateSheet_(TUTRONIX.sheets.bookingRequests);
}

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('BookingPage');
  template.accessToken = String((e && e.parameter && e.parameter.access) || '');
  return template.evaluate()
    .setTitle('Request a Tutronix Session')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function refreshParentBookingLinks() {
  const sheet = getBook_().getSheetByName(TUTRONIX.sheets.parents);
  const lastRow = sheet.getLastRow();
  const serviceUrl = 'https://script.google.com/macros/s/AKfycbyeyKof4HnGPcrU0t_Hhqie4RveyoQHEqC_M2ktwVimCyDVTUVGhZeV-vD29rDK1lK3/exec';
  if (!serviceUrl) throw new Error('Deploy the project as a web app before generating parent links.');
  if (lastRow < 2) return;

  const values = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
  values.forEach(row => {
    const name = String(row[1] || '').trim();
    const active = String(row[7] || '').trim();
    if (!name || active !== 'Yes') return;
    if (!row[9]) row[9] = Utilities.getUuid().replace(/-/g, '');
    if (!row[11]) row[11] = 'Yes';
    row[10] = serviceUrl + '?access=' + encodeURIComponent(row[9]);
  });
  sheet.getRange(2, 1, values.length, 12).setValues(values);
  sheet.getRange(2, 11, values.length, 1).setWrap(true);
  SpreadsheetApp.getUi().alert('Parent booking links are ready.');
}

function getBookingPortalData(accessToken) {
  const family = getBookingFamily_(accessToken);
  return {
    parent: family.parent,
    students: family.students,
    durations: [
      { minutes: 30, label: '30 minutes' },
      { minutes: 60, label: '1 hour' },
      { minutes: 90, label: '1½ hours' }
    ],
    modes: ['Virtual', 'In-person']
  };
}

function getAvailableBookingSlots(accessToken, durationMinutes, mode) {
  getBookingFamily_(accessToken);
  const duration = Number(durationMinutes);
  if ([30, 60, 90].indexOf(duration) === -1) throw new Error('Choose a valid session length.');
  if (['Virtual', 'In-person'].indexOf(mode) === -1) throw new Error('Choose Virtual or In-person.');

  expireBookingHolds_();
  const book = getBook_();
  const timeZone = book.getSpreadsheetTimeZone();
  const availability = book.getSheetByName(TUTRONIX.sheets.availability);
  if (!availability || availability.getLastRow() < 2) return [];

  const now = new Date();
  const endWindow = new Date(now);
  endWindow.setDate(endWindow.getDate() + 21);
  const rows = availability.getRange(2, 1, availability.getLastRow() - 1, 7).getValues();
  const occupied = getOccupiedBookingIntervals_();
  const bufferMinutes = mode === 'In-person' ? 30 : 0;
  const slots = [];

  rows.forEach(row => {
    const date = row[1];
    const startTime = row[2];
    const endTime = row[3];
    const availableMode = String(row[4] || '');
    const status = String(row[5] || '');
    if (!(date instanceof Date) || !(startTime instanceof Date) || !(endTime instanceof Date)) return;
    if (status !== 'Available') return;
    if (availableMode !== 'Either' && availableMode !== mode) return;

    const blockStart = combineDateAndTime_(date, startTime);
    const blockEnd = combineDateAndTime_(date, endTime);
    if (blockEnd <= now || blockStart > endWindow) return;

    let sessionStart = new Date(blockStart.getTime() + bufferMinutes * 60000);
    const latestStart = new Date(blockEnd.getTime() - (duration + bufferMinutes) * 60000);
    sessionStart = roundUpToHalfHour_(sessionStart);

    while (sessionStart <= latestStart) {
      const sessionEnd = new Date(sessionStart.getTime() + duration * 60000);
      const reservedStart = new Date(sessionStart.getTime() - bufferMinutes * 60000);
      const reservedEnd = new Date(sessionEnd.getTime() + bufferMinutes * 60000);
      const collision = occupied.some(item => intervalsOverlap_(reservedStart, reservedEnd, item.start, item.end));
      if (!collision && sessionStart > now) {
        slots.push({
          startMs: sessionStart.getTime(),
          endMs: sessionEnd.getTime(),
          dateKey: Utilities.formatDate(sessionStart, timeZone, 'yyyy-MM-dd'),
          dateLabel: Utilities.formatDate(sessionStart, timeZone, 'EEEE, MMMM d'),
          timeLabel: Utilities.formatDate(sessionStart, timeZone, 'h:mm a'),
          endLabel: Utilities.formatDate(sessionEnd, timeZone, 'h:mm a')
        });
      }
      sessionStart = new Date(sessionStart.getTime() + 30 * 60000);
    }
  });

  const unique = {};
  slots.forEach(slot => { unique[slot.startMs] = slot; });
  return Object.keys(unique).map(key => unique[key])
    .sort((a, b) => a.startMs - b.startMs);
}

function submitBookingRequest(request) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const family = getBookingFamily_(request.accessToken);
    const student = String(request.student || '').trim();
    if (family.students.indexOf(student) === -1) throw new Error('Choose a valid student.');
    const duration = Number(request.durationMinutes);
    const mode = String(request.mode || '');
    const startMs = Number(request.startMs);
    const available = getAvailableBookingSlots(request.accessToken, duration, mode);
    const selected = available.find(slot => Number(slot.startMs) === startMs);
    if (!selected) throw new Error('That time is no longer available. Please choose another time.');

    const book = getBook_();
    const timeZone = book.getSpreadsheetTimeZone();
    const sheet = book.getSheetByName(TUTRONIX.sheets.bookingRequests);
    const start = new Date(selected.startMs);
    const end = new Date(selected.endMs);
    const dateOnly = new Date(Utilities.formatDate(start, timeZone, 'yyyy/MM/dd'));
    const holdExpires = new Date(Date.now() + 24 * 60 * 60000);
    const requestId = makeId_('REQ');

    sheet.appendRow([
      requestId,
      new Date(),
      family.parent,
      student,
      dateOnly,
      start,
      end,
      duration,
      mode,
      String(request.notes || '').trim(),
      'Pending',
      holdExpires,
      '',
      ''
    ]);
    const row = sheet.getLastRow();
    sheet.getRange(row, 2).setNumberFormat('mmm d, yyyy h:mm AM/PM');
    sheet.getRange(row, 5).setNumberFormat('mmm d, yyyy');
    sheet.getRange(row, 6, 1, 2).setNumberFormat('h:mm AM/PM');
    sheet.getRange(row, 12).setNumberFormat('mmm d, yyyy h:mm AM/PM');

    MailApp.sendEmail({
      to: 'tutronixmiami@gmail.com',
      subject: 'New Tutronix booking request - ' + student,
      body: [
        'A new session request is waiting for review.',
        '',
        'Parent: ' + family.parent,
        'Student: ' + student,
        'Date: ' + selected.dateLabel,
        'Time: ' + selected.timeLabel + ' - ' + selected.endLabel,
        'Length: ' + duration + ' minutes',
        'Mode: ' + mode,
        request.notes ? 'Notes: ' + String(request.notes).trim() : '',
        '',
        'Open the Booking Requests tab to approve or decline it.'
      ].filter(Boolean).join('\n'),
      name: 'Tutronix Booking'
    });

    return {
      requestId,
      message: 'Your request has been sent to Tutronix.',
      dateLabel: selected.dateLabel,
      timeLabel: selected.timeLabel + ' - ' + selected.endLabel
    };
  } finally {
    lock.releaseLock();
  }
}

function approveSelectedBooking() {
  const context = getSelectedBookingRequest_();
  if (context.status !== 'Pending') throw new Error('Only pending requests can be confirmed.');
  confirmBookingRequest_(context);
  SpreadsheetApp.getUi().alert('Booking confirmed and added to Schedule.');
}

function declineSelectedBooking() {
  const context = getSelectedBookingRequest_();
  if (context.status !== 'Pending') throw new Error('Only pending requests can be declined.');
  declineBookingRequest_(context);
  SpreadsheetApp.getUi().alert('Booking declined. The time is available again.');
}

function enableAutomaticBookingDecisions() {
  installBookingDecisionTrigger();
  SpreadsheetApp.getUi().alert('Automatic booking decisions are enabled. Changing Pending to Confirmed or Declined will now run the complete workflow.');
}

function installBookingDecisionTrigger() {
  const handler = 'handleBookingRequestStatusEdit';
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === handler) ScriptApp.deleteTrigger(trigger);
  });
  ScriptApp.newTrigger(handler).forSpreadsheet(TUTRONIX.spreadsheetId).onEdit().create();
  return 'Automatic booking decisions enabled.';
}

function handleBookingRequestStatusEdit(e) {
  if (!e || !e.range || e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== TUTRONIX.sheets.bookingRequests || e.range.getRow() < 2 || e.range.getColumn() !== 11) return;

  const newStatus = String(e.value || '').trim();
  if (newStatus !== 'Confirmed' && newStatus !== 'Declined') return;
  const oldStatus = String(e.oldValue || '').trim();
  const book = e.source || getBook_();
  if (oldStatus !== 'Pending') {
    e.range.setValue(oldStatus || 'Pending');
    book.toast('Only a Pending request can be confirmed or declined.', 'Booking not changed', 6);
    return;
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const context = getBookingRequestByRow_(sheet, e.range.getRow());
    if (newStatus === 'Confirmed') {
      confirmBookingRequest_(context);
      book.toast('Booking confirmed, added to Schedule, and the parent was notified when an email was available.', 'Booking confirmed', 6);
    } else {
      declineBookingRequest_(context);
      book.toast('Booking declined and the time was released.', 'Booking declined', 6);
    }
  } catch (error) {
    const scheduleId = String(sheet.getRange(e.range.getRow(), 13).getDisplayValue() || '').trim();
    if (!scheduleId) sheet.getRange(e.range.getRow(), 11).setValue('Pending');
    book.toast(error.message, 'Booking update needs attention', 8);
  } finally {
    lock.releaseLock();
  }
}

function confirmBookingRequest_(context) {
  if (context.scheduleId) return context.scheduleId;
  if (context.status !== 'Pending' && context.status !== 'Confirmed') throw new Error('Only pending requests can be confirmed.');
  if (context.holdExpires instanceof Date && context.holdExpires < new Date()) {
    context.sheet.getRange(context.row, 11).setValue('Expired');
    throw new Error('This request has expired. Ask the parent to choose a new time.');
  }

  const schedule = getBook_().getSheetByName(TUTRONIX.sheets.schedule);
  const parentRecord = findRecord_(TUTRONIX.sheets.parents, 2, context.parent);
  const rate = parentRateForMode_(parentRecord, context.mode);
  const sessionId = makeId_('SES');
  const durationHours = context.duration / 60;
  schedule.appendRow([
    sessionId, context.date, context.start, context.end, durationHours, context.student, '', context.mode, '',
    'Scheduled', rate, '', '', context.parent, durationHours, 0, 0, ''
  ]);
  const scheduleRow = schedule.getLastRow();
  schedule.getRange(scheduleRow, 2).setNumberFormat('mmm d, yyyy');
  schedule.getRange(scheduleRow, 3, 1, 2).setNumberFormat('h:mm AM/PM');
  schedule.getRange(scheduleRow, 5).setNumberFormat('0.00');
  schedule.getRange(scheduleRow, 11, 1, 2).setNumberFormat('$#,##0.00');

  context.sheet.getRange(context.row, 11).setValue('Confirmed');
  context.sheet.getRange(context.row, 13).setValue(sessionId);
  context.sheet.getRange(context.row, 14).setValue(new Date()).setNumberFormat('mmm d, yyyy h:mm AM/PM');
  emailBookingDecision_(context, true);
  return sessionId;
}

function declineBookingRequest_(context) {
  if (context.scheduleId) throw new Error('This request already has a scheduled session and cannot be declined here.');
  if (context.status !== 'Pending' && context.status !== 'Declined') throw new Error('Only pending requests can be declined.');
  context.sheet.getRange(context.row, 11).setValue('Declined');
  context.sheet.getRange(context.row, 14).setValue(new Date()).setNumberFormat('mmm d, yyyy h:mm AM/PM');
  emailBookingDecision_(context, false);
}

function getSelectedBookingRequest_() {
  const book = getBook_();
  const sheet = book.getActiveSheet();
  if (sheet.getName() !== TUTRONIX.sheets.bookingRequests) {
    throw new Error('Select a row in the Booking Requests tab first.');
  }
  const row = sheet.getActiveRange().getRow();
  if (row < 2) throw new Error('Select a booking request row first.');
  return getBookingRequestByRow_(sheet, row);
}

function getBookingRequestByRow_(sheet, row) {
  const values = sheet.getRange(row, 1, 1, 14).getValues()[0];
  if (!values[0]) throw new Error('The selected row does not contain a request.');
  return {
    sheet, row, requestId: values[0], parent: String(values[2] || ''), student: String(values[3] || ''),
    date: values[4], start: values[5], end: values[6], duration: Number(values[7]) || 0,
    mode: String(values[8] || ''), notes: String(values[9] || ''), status: String(values[10] || ''),
    holdExpires: values[11], scheduleId: String(values[12] || '').trim()
  };
}

function emailBookingDecision_(context, approved) {
  const parentRecord = findRecord_(TUTRONIX.sheets.parents, 2, context.parent) || [];
  const email = String(parentRecord[2] || '').trim();
  if (!email) return;
  const timeZone = getBook_().getSpreadsheetTimeZone();
  MailApp.sendEmail({
    to: email,
    subject: approved ? 'Your Tutronix session is confirmed' : 'Update on your Tutronix session request',
    body: approved
      ? 'Your Tutronix session for ' + context.student + ' is confirmed for ' +
        Utilities.formatDate(context.start, timeZone, 'EEEE, MMMM d') + ', ' +
        Utilities.formatDate(context.start, timeZone, 'h:mm a') + ' - ' +
        Utilities.formatDate(context.end, timeZone, 'h:mm a') + ' (' + context.mode + ').'
      : 'Your requested Tutronix session for ' + context.student + ' was not confirmed. The time has been released. Please use your private booking link to choose another available time.',
    name: 'Tutronix Method'
  });
}

function getBookingFamily_(accessToken) {
  const token = String(accessToken || '').trim();
  if (!token) throw new Error('This booking link is incomplete.');
  const parentSheet = getBook_().getSheetByName(TUTRONIX.sheets.parents);
  const rows = parentSheet.getRange(2, 1, Math.max(parentSheet.getLastRow() - 1, 1), 12).getValues();
  const parentRow = rows.find(row => String(row[9] || '').trim() === token);
  if (!parentRow || String(parentRow[11] || '') !== 'Yes') throw new Error('This booking link is not active.');
  const parent = String(parentRow[1] || '').trim();
  const studentsSheet = getBook_().getSheetByName(TUTRONIX.sheets.students);
  const students = studentsSheet.getRange(2, 1, Math.max(studentsSheet.getLastRow() - 1, 1), studentsSheet.getLastColumn())
    .getValues()
    .filter(row => String(row[2] || '').trim() === parent && String(row[7] || '') === 'Yes')
    .map(row => String(row[1] || '').trim())
    .filter(Boolean);
  return { parent, email: String(parentRow[2] || ''), students };
}

function getOccupiedBookingIntervals_() {
  const book = getBook_();
  const intervals = [];
  const schedule = book.getSheetByName(TUTRONIX.sheets.schedule);
  if (schedule && schedule.getLastRow() >= 2) {
    schedule.getRange(2, 1, schedule.getLastRow() - 1, 18).getValues().forEach(row => {
      const date = row[1], start = row[2], end = row[3];
      const mode = String(row[7] || '');
      const status = String(row[9] || '');
      if (!(date instanceof Date) || !(start instanceof Date) || !(end instanceof Date)) return;
      if (['Cancelled - No Charge', 'Rescheduled'].indexOf(status) !== -1) return;
      const buffer = mode === 'In-person' ? 30 : 0;
      intervals.push({
        start: new Date(combineDateAndTime_(date, start).getTime() - buffer * 60000),
        end: new Date(combineDateAndTime_(date, end).getTime() + buffer * 60000)
      });
    });
  }

  const requests = book.getSheetByName(TUTRONIX.sheets.bookingRequests);
  if (requests && requests.getLastRow() >= 2) {
    requests.getRange(2, 1, requests.getLastRow() - 1, 14).getValues().forEach(row => {
      if (String(row[10] || '') !== 'Pending') return;
      if (row[11] instanceof Date && row[11] < new Date()) return;
      const buffer = String(row[8] || '') === 'In-person' ? 30 : 0;
      intervals.push({
        start: new Date(row[5].getTime() - buffer * 60000),
        end: new Date(row[6].getTime() + buffer * 60000)
      });
    });
  }
  return intervals;
}

function expireBookingHolds_() {
  const sheet = getBook_().getSheetByName(TUTRONIX.sheets.bookingRequests);
  if (!sheet || sheet.getLastRow() < 2) return;
  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14);
  const rows = range.getValues();
  let changed = false;
  rows.forEach(row => {
    if (String(row[10] || '') === 'Pending' && row[11] instanceof Date && row[11] < new Date()) {
      row[10] = 'Expired';
      changed = true;
    }
  });
  if (changed) range.setValues(rows);
}

function combineDateAndTime_(date, time) {
  return new Date(
    date.getFullYear(),
    date.getMonth(),
    date.getDate(),
    time.getHours(),
    time.getMinutes(),
    0,
    0
  );
}

function roundUpToHalfHour_(date) {
  const rounded = new Date(date);
  rounded.setSeconds(0, 0);
  const minutes = rounded.getMinutes();
  if (minutes === 0 || minutes === 30) return rounded;
  rounded.setMinutes(minutes < 30 ? 30 : 60);
  return rounded;
}

function intervalsOverlap_(aStart, aEnd, bStart, bEnd) {
  return aStart < bEnd && bStart < aEnd;
}

function initializeAvailabilityRow_(sheet, row) {
  const values = sheet.getRange(row, 1, 1, 7).getValues()[0];
  if (!values[1] && !values[2] && !values[3]) return;
  if (!values[0]) values[0] = makeId_('AVL');
  if (!values[4]) values[4] = 'Either';
  if (!values[5]) values[5] = 'Available';
  sheet.getRange(row, 1, 1, 7).setValues([values]);
}
