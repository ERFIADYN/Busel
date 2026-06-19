/**
 * BUMDes Finance Batuatas Liwu
 * Backend Google Apps Script + Google Sheets.
 */

const APP = {
  name: 'BUMDes Finance Batuatas Liwu',
  propertyKey: 'BUMDES_SPREADSHEET_ID',
  sheets: {
    settings: ['key', 'value'],
    units: ['id', 'name', 'category', 'manager', 'status'],
    accounts: ['code', 'name', 'type', 'openingBalance'],
    transactions: [
      'id', 'date', 'reference', 'type', 'category', 'account', 'unit',
      'description', 'amount', 'paymentMethod', 'evidenceUrl', 'status',
      'createdBy', 'createdAt', 'updatedAt'
    ],
    budgets: ['id', 'year', 'category', 'unit', 'amount', 'note'],
    audit: ['timestamp', 'action', 'entity', 'entityId', 'user', 'detail']
  }
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(APP.name)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/** Run once from the Apps Script editor, or let the app initialize on first load. */
function setupBumdesApp() {
  const spreadsheet = getSpreadsheet_();
  ensureSchema_(spreadsheet);
  seedData_(spreadsheet);
  return {
    success: true,
    spreadsheetId: spreadsheet.getId(),
    spreadsheetUrl: spreadsheet.getUrl(),
    message: 'Database BUMDes Finance siap digunakan.'
  };
}

function getBootstrapData() {
  const spreadsheet = getSpreadsheet_();
  ensureSchema_(spreadsheet);
  seedData_(spreadsheet);

  const settings = rowsAsObjects_(spreadsheet.getSheetByName('settings'))
    .reduce(function (acc, row) {
      acc[row.key] = row.value;
      return acc;
    }, {});

  return {
    success: true,
    settings: settings,
    units: rowsAsObjects_(spreadsheet.getSheetByName('units')),
    accounts: rowsAsObjects_(spreadsheet.getSheetByName('accounts')),
    transactions: rowsAsObjects_(spreadsheet.getSheetByName('transactions')),
    budgets: rowsAsObjects_(spreadsheet.getSheetByName('budgets')),
    meta: {
      spreadsheetUrl: spreadsheet.getUrl(),
      currentUser: getUserEmail_(),
      generatedAt: new Date().toISOString()
    }
  };
}

function saveTransaction(payload) {
  payload = payload || {};
  validateTransaction_(payload);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const spreadsheet = getSpreadsheet_();
    const sheet = spreadsheet.getSheetByName('transactions');
    const headers = APP.sheets.transactions;
    const rows = rowsAsObjects_(sheet);
    const now = new Date().toISOString();
    const user = getUserEmail_();
    const id = cleanText_(payload.id) || makeId_('TRX');
    const existingIndex = rows.findIndex(function (item) { return item.id === id; });
    const existing = existingIndex >= 0 ? rows[existingIndex] : {};

    const record = {
      id: id,
      date: normalizeDate_(payload.date),
      reference: cleanText_(payload.reference) || nextReference_(sheet, payload.type),
      type: cleanText_(payload.type),
      category: cleanText_(payload.category),
      account: cleanText_(payload.account),
      unit: cleanText_(payload.unit),
      description: cleanText_(payload.description),
      amount: Number(payload.amount),
      paymentMethod: cleanText_(payload.paymentMethod),
      evidenceUrl: cleanText_(payload.evidenceUrl),
      status: cleanText_(payload.status) || 'Posted',
      createdBy: existing.createdBy || user,
      createdAt: existing.createdAt || now,
      updatedAt: now
    };

    const values = headers.map(function (key) { return record[key]; });
    if (existingIndex >= 0) {
      sheet.getRange(existingIndex + 2, 1, 1, headers.length).setValues([values]);
    } else {
      sheet.appendRow(values);
    }

    appendAudit_(spreadsheet, existingIndex >= 0 ? 'UPDATE' : 'CREATE', 'transaction', id, JSON.stringify({
      reference: record.reference,
      amount: record.amount,
      type: record.type
    }));

    return { success: true, transaction: record, message: 'Transaksi berhasil disimpan.' };
  } finally {
    lock.releaseLock();
  }
}

function voidTransaction(id, reason) {
  id = cleanText_(id);
  if (!id) throw new Error('ID transaksi tidak valid.');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const spreadsheet = getSpreadsheet_();
    const sheet = spreadsheet.getSheetByName('transactions');
    const rows = rowsAsObjects_(sheet);
    const index = rows.findIndex(function (item) { return item.id === id; });
    if (index < 0) throw new Error('Transaksi tidak ditemukan.');

    const statusColumn = APP.sheets.transactions.indexOf('status') + 1;
    const updatedColumn = APP.sheets.transactions.indexOf('updatedAt') + 1;
    sheet.getRange(index + 2, statusColumn).setValue('Void');
    sheet.getRange(index + 2, updatedColumn).setValue(new Date().toISOString());
    appendAudit_(spreadsheet, 'VOID', 'transaction', id, cleanText_(reason) || 'Dibatalkan oleh operator');
    return { success: true, message: 'Transaksi telah dibatalkan.' };
  } finally {
    lock.releaseLock();
  }
}

function saveBudget(payload) {
  payload = payload || {};
  const year = Number(payload.year);
  const amount = Number(payload.amount);
  if (!year || year < 2020 || year > 2100) throw new Error('Tahun anggaran tidak valid.');
  if (!cleanText_(payload.category)) throw new Error('Kategori anggaran wajib dipilih.');
  if (!cleanText_(payload.unit)) throw new Error('Unit usaha wajib dipilih.');
  if (!Number.isFinite(amount) || amount <= 0) throw new Error('Nilai anggaran harus lebih dari nol.');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const spreadsheet = getSpreadsheet_();
    const sheet = spreadsheet.getSheetByName('budgets');
    const rows = rowsAsObjects_(sheet);
    const id = cleanText_(payload.id) || makeId_('BDG');
    const existingIndex = rows.findIndex(function (item) { return item.id === id; });
    const record = {
      id: id,
      year: year,
      category: cleanText_(payload.category),
      unit: cleanText_(payload.unit),
      amount: amount,
      note: cleanText_(payload.note)
    };
    const values = APP.sheets.budgets.map(function (key) { return record[key]; });
    if (existingIndex >= 0) {
      sheet.getRange(existingIndex + 2, 1, 1, values.length).setValues([values]);
    } else {
      sheet.appendRow(values);
    }
    appendAudit_(spreadsheet, existingIndex >= 0 ? 'UPDATE' : 'CREATE', 'budget', id, JSON.stringify(record));
    return { success: true, budget: record, message: 'Anggaran berhasil disimpan.' };
  } finally {
    lock.releaseLock();
  }
}

function saveSettings(payload) {
  payload = payload || {};
  const allowed = ['organizationName', 'villageName', 'districtName', 'regencyName', 'treasurerName', 'currency'];
  const spreadsheet = getSpreadsheet_();
  const sheet = spreadsheet.getSheetByName('settings');
  const rows = rowsAsObjects_(sheet);

  allowed.forEach(function (key) {
    if (typeof payload[key] === 'undefined') return;
    const value = cleanText_(payload[key]);
    const index = rows.findIndex(function (row) { return row.key === key; });
    if (index >= 0) sheet.getRange(index + 2, 2).setValue(value);
    else sheet.appendRow([key, value]);
  });

  appendAudit_(spreadsheet, 'UPDATE', 'settings', 'organization', 'Profil organisasi diperbarui');
  return { success: true, message: 'Pengaturan berhasil diperbarui.' };
}

function getSpreadsheet_() {
  const properties = PropertiesService.getScriptProperties();
  const id = properties.getProperty(APP.propertyKey);
  if (id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (error) {
      properties.deleteProperty(APP.propertyKey);
    }
  }

  const spreadsheet = SpreadsheetApp.create('Database - ' + APP.name);
  properties.setProperty(APP.propertyKey, spreadsheet.getId());
  return spreadsheet;
}

function ensureSchema_(spreadsheet) {
  Object.keys(APP.sheets).forEach(function (name) {
    let sheet = spreadsheet.getSheetByName(name);
    if (!sheet) sheet = spreadsheet.insertSheet(name);
    const headers = APP.sheets[name];
    const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    if (current.join('|') !== headers.join('|')) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#123B5D')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
  });

  const defaultSheet = spreadsheet.getSheetByName('Sheet1');
  if (defaultSheet && spreadsheet.getSheets().length > Object.keys(APP.sheets).length) {
    spreadsheet.deleteSheet(defaultSheet);
  }
}

function seedData_(spreadsheet) {
  const settingsSheet = spreadsheet.getSheetByName('settings');
  if (settingsSheet.getLastRow() <= 1) {
    settingsSheet.getRange(2, 1, 6, 2).setValues([
      ['organizationName', 'BUMDes Batuatas Liwu'],
      ['villageName', 'Desa Batuatas Liwu'],
      ['districtName', 'Kecamatan Batu Atas'],
      ['regencyName', 'Kabupaten Buton Selatan'],
      ['treasurerName', 'Bendahara BUMDes'],
      ['currency', 'IDR']
    ]);
  }

  const unitsSheet = spreadsheet.getSheetByName('units');
  if (unitsSheet.getLastRow() <= 1) {
    unitsSheet.getRange(2, 1, 4, 5).setValues([
      ['UNT-001', 'Perdagangan Desa', 'Perdagangan', 'La Ode Rahman', 'Aktif'],
      ['UNT-002', 'Air Bersih', 'Layanan', 'Wa Ode Nur', 'Aktif'],
      ['UNT-003', 'Sewa Peralatan', 'Penyewaan', 'La Ode Amir', 'Aktif'],
      ['UNT-004', 'Wisata Bahari', 'Pariwisata', 'Wa Ode Mina', 'Aktif']
    ]);
  }

  const accountsSheet = spreadsheet.getSheetByName('accounts');
  if (accountsSheet.getLastRow() <= 1) {
    accountsSheet.getRange(2, 1, 4, 4).setValues([
      ['1101', 'Kas Tunai', 'Kas', 5000000],
      ['1102', 'Bank BRI', 'Bank', 15000000],
      ['1103', 'Bank Sultra', 'Bank', 10000000],
      ['1201', 'Piutang Usaha', 'Piutang', 0]
    ]);
  }

  const budgetsSheet = spreadsheet.getSheetByName('budgets');
  const year = new Date().getFullYear();
  if (budgetsSheet.getLastRow() <= 1) {
    budgetsSheet.getRange(2, 1, 6, 6).setValues([
      [makeId_('BDG'), year, 'Belanja Barang', 'UNT-001', 42000000, 'Pengadaan stok perdagangan'],
      [makeId_('BDG'), year, 'Operasional', 'UNT-001', 18000000, 'Operasional tahunan'],
      [makeId_('BDG'), year, 'Pemeliharaan', 'UNT-002', 24000000, 'Pemeliharaan jaringan air'],
      [makeId_('BDG'), year, 'Operasional', 'UNT-003', 15000000, 'Operasional alat'],
      [makeId_('BDG'), year, 'Promosi', 'UNT-004', 12000000, 'Promosi destinasi'],
      [makeId_('BDG'), year, 'Honorarium', 'UNT-004', 18000000, 'Petugas lapangan']
    ]);
  }

  const txSheet = spreadsheet.getSheetByName('transactions');
  if (txSheet.getLastRow() <= 1) seedTransactions_(txSheet);
}

function seedTransactions_(sheet) {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth();
  const samples = [
    [-1, 'TRX-001', 'Pemasukan', 'Penjualan', '1102', 'UNT-001', 'Penjualan kebutuhan pokok', 8750000, 'Transfer'],
    [-2, 'TRX-002', 'Pengeluaran', 'Belanja Barang', '1102', 'UNT-001', 'Belanja stok perdagangan', 4250000, 'Transfer'],
    [-4, 'TRX-003', 'Pemasukan', 'Pendapatan Layanan', '1101', 'UNT-002', 'Pembayaran layanan air bersih', 6200000, 'Tunai'],
    [-6, 'TRX-004', 'Pengeluaran', 'Pemeliharaan', '1101', 'UNT-002', 'Perbaikan sambungan pipa', 1450000, 'Tunai'],
    [-9, 'TRX-005', 'Pemasukan', 'Pendapatan Sewa', '1102', 'UNT-003', 'Sewa tenda dan kursi', 3750000, 'Transfer'],
    [-12, 'TRX-006', 'Pengeluaran', 'Operasional', '1101', 'UNT-003', 'BBM dan transportasi alat', 875000, 'Tunai'],
    [-16, 'TRX-007', 'Pemasukan', 'Pendapatan Wisata', '1103', 'UNT-004', 'Tiket dan jasa wisata', 4950000, 'QRIS'],
    [-20, 'TRX-008', 'Pengeluaran', 'Promosi', '1103', 'UNT-004', 'Materi promosi digital', 1250000, 'Transfer'],
    [-32, 'TRX-009', 'Pemasukan', 'Penjualan', '1102', 'UNT-001', 'Penjualan bulan sebelumnya', 7100000, 'Transfer'],
    [-35, 'TRX-010', 'Pengeluaran', 'Belanja Barang', '1102', 'UNT-001', 'Belanja stok bulan sebelumnya', 3900000, 'Transfer'],
    [-63, 'TRX-011', 'Pemasukan', 'Pendapatan Layanan', '1101', 'UNT-002', 'Pembayaran layanan air', 5400000, 'Tunai'],
    [-67, 'TRX-012', 'Pengeluaran', 'Pemeliharaan', '1101', 'UNT-002', 'Penggantian meter air', 1100000, 'Tunai']
  ];
  const createdAt = new Date(year, month, 1).toISOString();
  const rows = samples.map(function (sample, index) {
    const date = new Date();
    date.setDate(date.getDate() + sample[0]);
    return [
      makeId_('TRX'), normalizeDate_(date), sample[1], sample[2], sample[3], sample[4],
      sample[5], sample[6], sample[7], sample[8], '', 'Posted', 'system@bumdes.local',
      createdAt, createdAt
    ];
  });
  sheet.getRange(2, 1, rows.length, APP.sheets.transactions.length).setValues(rows);
}

function rowsAsObjects_(sheet) {
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastColumn).getValues();
  return values.filter(function (row) {
    return row.some(function (cell) { return cell !== ''; });
  }).map(function (row) {
    return headers.reduce(function (obj, header, index) {
      let value = row[index];
      if (value instanceof Date) value = normalizeDate_(value);
      obj[header] = value;
      return obj;
    }, {});
  });
}

function validateTransaction_(payload) {
  const allowedTypes = ['Pemasukan', 'Pengeluaran'];
  if (allowedTypes.indexOf(cleanText_(payload.type)) < 0) throw new Error('Jenis transaksi tidak valid.');
  if (!payload.date) throw new Error('Tanggal transaksi wajib diisi.');
  if (!cleanText_(payload.category)) throw new Error('Kategori wajib dipilih.');
  if (!cleanText_(payload.account)) throw new Error('Akun kas/bank wajib dipilih.');
  if (!cleanText_(payload.unit)) throw new Error('Unit usaha wajib dipilih.');
  if (!cleanText_(payload.description)) throw new Error('Uraian transaksi wajib diisi.');
  if (!Number.isFinite(Number(payload.amount)) || Number(payload.amount) <= 0) {
    throw new Error('Nominal transaksi harus lebih dari nol.');
  }
}

function appendAudit_(spreadsheet, action, entity, entityId, detail) {
  spreadsheet.getSheetByName('audit').appendRow([
    new Date().toISOString(), action, entity, entityId, getUserEmail_(), cleanText_(detail)
  ]);
}

function nextReference_(sheet, type) {
  const prefix = type === 'Pemasukan' ? 'BM' : 'BK';
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Jakarta', 'yyyyMMdd');
  const sequence = Math.max(1, sheet.getLastRow());
  return prefix + '-' + stamp + '-' + String(sequence).padStart(3, '0');
}

function normalizeDate_(value) {
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) throw new Error('Format tanggal tidak valid.');
  return Utilities.formatDate(date, Session.getScriptTimeZone() || 'Asia/Jakarta', 'yyyy-MM-dd');
}

function makeId_(prefix) {
  return prefix + '-' + Utilities.getUuid().split('-')[0].toUpperCase();
}

function getUserEmail_() {
  return Session.getActiveUser().getEmail() || 'operator@bumdes.local';
}

function cleanText_(value) {
  if (value === null || typeof value === 'undefined') return '';
  let text = String(value).trim().slice(0, 1000);
  if (/^[=+\-@]/.test(text)) text = "'" + text;
  return text;
}
