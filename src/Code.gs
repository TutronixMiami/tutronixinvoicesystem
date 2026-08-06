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
    .addItem('Generate PDF for selected invoice', 'generatePdfForSelectedInvoice')
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

function generatePdfForSelectedInvoice() {
  const book = getBook_();
  const sheet = book.getActiveSheet();
  const row = sheet.getActiveRange().getRow();
  if (sheet.getName() !== TUTRONIX.sheets.invoices || row < 2) {
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

  return createBrandedInvoicePdf_({
    invoiceId: String(invoice[0]),
    parent: String(invoice[1] || ''),
    students: String(invoice[2] || ''),
    createdDate: invoice[3] instanceof Date ? invoice[3] : new Date(),
    totalHours: Number(invoice[4]) || 0,
    totalAmount: Number(invoice[5]) || 0,
    sessions: lineRows
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
  totals.getCell(1, 0).setBackgroundColor(coral).editAsText().setBold(true).setForegroundColor(white).setFontSize(12);
  totals.getCell(1, 1).setBackgroundColor(coral).editAsText().setBold(true).setForegroundColor(white).setFontSize(12);
  totals.getCell(1, 1).getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

  body.appendParagraph('');
  const payment = body.appendTable([['PAYMENT INSTRUCTIONS'], [paymentInstructions]]);
  payment.setBorderColor(forest).setBorderWidth(1);
  payment.getCell(0, 0).setBackgroundColor(yellow);
  payment.getCell(0, 0).editAsText().setBold(true).setForegroundColor(forest);
  payment.getCell(1, 0).setBackgroundColor(cream);

  body.appendParagraph('');
  const footer = body.appendParagraph('Thank you for choosing ' + businessName + '.');
  footer.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  footer.editAsText().setBold(true).setFontSize(10).setForegroundColor(forest);
  const subfooter = body.appendParagraph('Tutoring + Academic Coaching + Executive Function');
  subfooter.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  subfooter.editAsText().setFontSize(8).setForegroundColor(coral);

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
