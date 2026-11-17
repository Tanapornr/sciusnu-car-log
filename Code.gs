/** Web app for the vehicle log in the existing Google Sheet. */
const SHEET_ID = '1TSspmOiwVrRD7V3lYWwe5-wn7Ga3qyC3HXVKex09IVk';
const VEHICLE = '40-0210 พิษณุโลก';
const TRIPS = 'บันทึกการใช้รถ';
const FUEL = 'บันทึกเติมน้ำมัน';
const FIRST_ROW = 7;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('บันทึกการใช้รถ 40-0210 พิษณุโลก');
}

function getContext() {
  const book = SpreadsheetApp.openById(SHEET_ID);
  const trips = readRows_(book.getSheetByName(TRIPS), 8);
  const fuel = readRows_(book.getSheetByName(FUEL), 11);
  const latestTrip = lastOdometer_(trips, 5);
  const latestFuel = lastOdometer_(fuel, 2);
  const suggested = Math.max(latestTrip || 0, latestFuel || 0) || null;
  const lastFuel = fuel.length ? fuel[fuel.length - 1] : null;
  const lastEfficiency = [...fuel].reverse().find(row => typeof row[10] === 'number' && row[10] > 0);
  return {
    vehicle: VEHICLE,
    today: Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd'),
    suggestedOdometer: suggested,
    suggestedSource: suggested === null ? '' : suggested === latestTrip ? 'เลขไมล์กลับจากการใช้รถ' : 'เลขไมล์จากการเติมน้ำมัน',
    lastFuelOdometer: lastFuel ? numeric_(lastFuel[2]) : null,
    lastEfficiency: lastEfficiency ? Math.round(lastEfficiency[10] * 100) / 100 : null,
    tripCount: trips.length,
    fuelCount: fuel.length,
    spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${SHEET_ID}/edit`
  };
}

function saveTrip(form) {
  return withLock_(() => {
    const date = date_(form.date);
    const driver = text_(form.driver, 120, 'ชื่อผู้ขับรถ');
    const purpose = text_(form.purpose, 300, 'สถานที่และภารกิจ');
    const start = positive_(form.startOdometer, 'เลขไมล์ออก', true);
    const end = positive_(form.endOdometer, 'เลขไมล์กลับ', true);
    if (end < start) throw new Error('เลขไมล์กลับต้องไม่น้อยกว่าเลขไมล์ออก');
    const note = optionalText_(form.note, 300);
    const book = SpreadsheetApp.openById(SHEET_ID);
    book.setSpreadsheetTimeZone('Asia/Bangkok');
    const sheet = book.getSheetByName(TRIPS);
    const row = nextRow_(sheet);
    ensureRow_(sheet, row);
    sheet.getRange(row, 1, 1, 6).setValues([[date, VEHICLE, driver, purpose, start, end]]);
    sheet.getRange(row, 8).setValue(note);
    SpreadsheetApp.flush();
    return {row, distance: end - start, context: getContext()};
  });
}

function saveFuel(form) {
  return withLock_(() => {
    const date = date_(form.date);
    const odometer = positive_(form.odometer, 'เลขไมล์ขณะเติม', true);
    const liters = positive_(form.liters, 'จำนวนลิตร');
    const total = positive_(form.total, 'ยอดจ่ายจริง');
    const fill = form.fillType;
    if (fill !== 'เต็มถัง' && fill !== 'ไม่เต็มถัง') throw new Error('กรุณาเลือกลักษณะการเติม');
    const receipt = optionalText_(form.receipt, 100);
    const book = SpreadsheetApp.openById(SHEET_ID);
    book.setSpreadsheetTimeZone('Asia/Bangkok');
    const sheet = book.getSheetByName(FUEL);
    const rows = readRows_(sheet, 11);
    if (rows.length && odometer <= numeric_(rows[rows.length - 1][2])) {
      throw new Error('เลขไมล์ต้องมากกว่ารายการเติมน้ำมันก่อนหน้า กรุณาตรวจหน้าปัดและลำดับรายการ');
    }
    const row = nextRow_(sheet);
    ensureRow_(sheet, row);
    const unitPrice = Math.round(total / liters * 10000) / 10000;
    sheet.getRange(row, 1, 1, 8).setValues([[date, VEHICLE, odometer, liters, unitPrice, total, receipt, fill]]);
    SpreadsheetApp.flush();
    return {row, efficiency: sheet.getRange(row, 11).getValue() || null, context: getContext()};
  });
}

function readRows_(sheet, width) {
  if (!sheet) throw new Error('ไม่พบแท็บข้อมูลในชีต');
  const count = sheet.getMaxRows() - FIRST_ROW + 1;
  const values = sheet.getRange(FIRST_ROW, 1, count, width).getValues();
  return values.filter(row => row[0] !== '' && row[0] !== null);
}

function lastOdometer_(rows, index) {
  for (let i = rows.length - 1; i >= 0; i--) {
    const value = numeric_(rows[i][index]);
    if (value !== null) return value;
  }
  return null;
}

function nextRow_(sheet) {
  const values = sheet.getRange(FIRST_ROW, 1, sheet.getMaxRows() - FIRST_ROW + 1, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== '' && values[i][0] !== null) return FIRST_ROW + i + 1;
  }
  return FIRST_ROW;
}

function ensureRow_(sheet, row) {
  if (row <= sheet.getMaxRows()) return;
  const previous = sheet.getMaxRows();
  sheet.insertRowsAfter(previous, 1);
  const calculatedStart = sheet.getName() === TRIPS ? 7 : 9;
  const calculatedWidth = sheet.getName() === TRIPS ? 1 : 3;
  sheet.getRange(previous, calculatedStart, 1, calculatedWidth)
    .copyTo(sheet.getRange(row, calculatedStart, 1, calculatedWidth));
  if (sheet.getName() === FUEL) sheet.getRange(previous, 8).copyTo(sheet.getRange(row, 8), SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
}

function withLock_(action) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('มีผู้บันทึกข้อมูลพร้อมกัน กรุณาลองอีกครั้ง');
  try { return action(); } finally { lock.releaseLock(); }
}

function date_(value) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(value || ''))) throw new Error('กรุณาระบุวันที่');
  const date = new Date(`${value}T12:00:00+07:00`);
  if (Utilities.formatDate(date, 'Asia/Bangkok', 'yyyy-MM-dd') !== value) throw new Error('วันที่ไม่ถูกต้อง');
  return date;
}

function numeric_(value) {
  if (value === '' || value === null || value === undefined) return null;
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

function positive_(value, label, allowZero) {
  const number = numeric_(value);
  if (number === null || (allowZero ? number < 0 : number <= 0)) throw new Error(`${label}ต้องเป็นตัวเลข${allowZero ? 'ไม่ติดลบ' : 'มากกว่า 0'}`);
  return number;
}

function text_(value, max, label) {
  const cleaned = String(value || '').trim();
  if (!cleaned || cleaned.length > max) throw new Error(`กรุณากรอก${label}ไม่เกิน ${max} ตัวอักษร`);
  return safeText_(cleaned);
}

function optionalText_(value, max) {
  const cleaned = String(value || '').trim();
  if (cleaned.length > max) throw new Error(`ข้อความต้องไม่เกิน ${max} ตัวอักษร`);
  return safeText_(cleaned);
}

function safeText_(value) {
  return /^[=+@\-]/.test(value) ? `'${value}` : value;
}
