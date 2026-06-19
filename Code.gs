/**
 * ============================================================
 *  BUMDES FINANCE - BATUATAS LIWU
 *  Backend Google Apps Script (Spreadsheet sebagai Database)
 * ============================================================
 *
 *  Fitur:
 *   - Login user/password (token-based, hash SHA-256 + salt)
 *   - Multi unit usaha BUMDes
 *   - Catat pemasukan & pengeluaran
 *   - Saldo kas berjalan otomatis
 *   - Laporan & data untuk grafik (laba rugi, kas)
 *   - Manajemen pengguna (khusus admin)
 *
 *  Cara pakai singkat:
 *   1. Buat Spreadsheet baru, salin ID-nya ke SPREADSHEET_ID di bawah
 *      (atau biarkan kosong jika script ini terikat / bound ke Spreadsheet).
 *   2. Jalankan fungsi setupDatabase() sekali dari editor untuk membuat
 *      semua sheet + akun admin default.
 *   3. Deploy > New deployment > Web app.
 * ============================================================
 */

// Jika script TIDAK terikat ke spreadsheet, isi ID spreadsheet di sini.
// Jika dibiarkan kosong dan script bound ke spreadsheet, otomatis pakai aktif.
const SPREADSHEET_ID = '';

const APP_NAME = 'BUMDes Finance — Batuatas Liwu';
const SALT = 'BumdesBatuatasLiwu#2024'; // ganti dengan string rahasia Anda sendiri
const SESSION_HOURS = 8;

const SHEETS = {
  USERS: 'Users',
  UNITS: 'Units',
  TRX: 'Transaksi',
  KATEGORI: 'Kategori'
};

/* ----------------------- Util Spreadsheet ----------------------- */
function _ss() {
  if (SPREADSHEET_ID && SPREADSHEET_ID.trim() !== '') {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) {
    throw new Error('SPREADSHEET_ID belum diisi dan script tidak terikat ke Spreadsheet.');
  }
  return active;
}

function _sheet(name) {
  const ss = _ss();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function _rows(name) {
  const sh = _sheet(name);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return { header: values[0] || [], data: [] };
  const header = values[0];
  const data = values.slice(1).map(function (r) {
    const o = {};
    header.forEach(function (h, i) { o[h] = r[i]; });
    return o;
  });
  return { header: header, data: data };
}

/* ----------------------- Setup / Seeding ----------------------- */
function setupDatabase() {
  const ss = _ss();

  // Users
  const users = _sheet(SHEETS.USERS);
  if (users.getLastRow() === 0) {
    users.appendRow(['Username', 'PasswordHash', 'Nama', 'Role', 'Aktif', 'DibuatPada']);
    users.appendRow(['admin', _hash('admin123'), 'Administrator', 'admin', true, new Date()]);
  }

  // Units
  const units = _sheet(SHEETS.UNITS);
  if (units.getLastRow() === 0) {
    units.appendRow(['KodeUnit', 'NamaUnit', 'Keterangan', 'Aktif', 'DibuatPada']);
    units.appendRow(['U001', 'Unit Simpan Pinjam', 'Layanan keuangan mikro desa', true, new Date()]);
    units.appendRow(['U002', 'Unit Sembako', 'Toko kebutuhan pokok', true, new Date()]);
    units.appendRow(['U003', 'Unit Air Bersih', 'Pengelolaan air bersih desa', true, new Date()]);
  }

  // Kategori
  const kat = _sheet(SHEETS.KATEGORI);
  if (kat.getLastRow() === 0) {
    kat.appendRow(['Kategori', 'Jenis', 'Aktif']);
    [['Penjualan', 'Pemasukan'], ['Iuran/Jasa', 'Pemasukan'], ['Modal/Penyertaan', 'Pemasukan'],
     ['Pendapatan Lain', 'Pemasukan'], ['Pembelian Barang', 'Pengeluaran'], ['Gaji/Honor', 'Pengeluaran'],
     ['Operasional', 'Pengeluaran'], ['Listrik/Air', 'Pengeluaran'], ['Pengeluaran Lain', 'Pengeluaran']]
      .forEach(function (r) { kat.appendRow([r[0], r[1], true]); });
  }

  // Transaksi
  const trx = _sheet(SHEETS.TRX);
  if (trx.getLastRow() === 0) {
    trx.appendRow(['ID', 'Tanggal', 'KodeUnit', 'Jenis', 'Kategori', 'Keterangan', 'Jumlah', 'Operator', 'DibuatPada']);
  }

  return 'Setup selesai. Login: admin / admin123 (segera ganti password).';
}

/* ----------------------- Keamanan ----------------------- */
function _hash(text) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text + SALT, Utilities.Charset.UTF_8);
  return raw.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function _newToken(user) {
  const token = Utilities.getUuid();
  const cache = CacheService.getScriptCache();
  cache.put('tok_' + token, JSON.stringify({ u: user.Username, r: user.Role, n: user.Nama }), SESSION_HOURS * 3600);
  return token;
}

function _auth(token) {
  if (!token) throw new Error('Sesi tidak valid. Silakan login ulang.');
  const cache = CacheService.getScriptCache();
  const raw = cache.get('tok_' + token);
  if (!raw) throw new Error('Sesi berakhir. Silakan login ulang.');
  return JSON.parse(raw);
}

function _requireAdmin(token) {
  const s = _auth(token);
  if (s.r !== 'admin') throw new Error('Akses ditolak. Hanya admin.');
  return s;
}

/* ----------------------- API: Auth ----------------------- */
function login(username, password) {
  const rows = _rows(SHEETS.USERS).data;
  const u = rows.filter(function (x) {
    return String(x.Username).toLowerCase() === String(username).toLowerCase();
  })[0];
  if (!u) return { ok: false, message: 'Username tidak ditemukan.' };
  if (u.Aktif === false) return { ok: false, message: 'Akun nonaktif.' };
  if (String(u.PasswordHash) !== _hash(password)) return { ok: false, message: 'Password salah.' };
  const token = _newToken(u);
  return { ok: true, token: token, user: { username: u.Username, nama: u.Nama, role: u.Role }, appName: APP_NAME };
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('tok_' + token);
  return { ok: true };
}

function changePassword(token, oldPass, newPass) {
  const s = _auth(token);
  const sh = _sheet(SHEETS.USERS);
  const values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]).toLowerCase() === s.u.toLowerCase()) {
      if (String(values[i][1]) !== _hash(oldPass)) return { ok: false, message: 'Password lama salah.' };
      sh.getRange(i + 1, 2).setValue(_hash(newPass));
      return { ok: true, message: 'Password berhasil diubah.' };
    }
  }
  return { ok: false, message: 'User tidak ditemukan.' };
}

/* ----------------------- API: Master data ----------------------- */
function getBootstrap(token) {
  _auth(token);
  return {
    units: _rows(SHEETS.UNITS).data,
    kategori: _rows(SHEETS.KATEGORI).data
  };
}

function listUnits(token) {
  _auth(token);
  return _rows(SHEETS.UNITS).data;
}

function saveUnit(token, unit) {
  _requireAdmin(token);
  const sh = _sheet(SHEETS.UNITS);
  const values = sh.getDataRange().getValues();
  if (unit.KodeUnit) {
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(unit.KodeUnit)) {
        sh.getRange(i + 1, 2).setValue(unit.NamaUnit);
        sh.getRange(i + 1, 3).setValue(unit.Keterangan || '');
        sh.getRange(i + 1, 4).setValue(unit.Aktif !== false);
        return { ok: true, message: 'Unit diperbarui.' };
      }
    }
  }
  const kode = 'U' + ('000' + (values.length)).slice(-3);
  sh.appendRow([kode, unit.NamaUnit, unit.Keterangan || '', true, new Date()]);
  return { ok: true, message: 'Unit ditambahkan.', kode: kode };
}

function deleteUnit(token, kodeUnit) {
  _requireAdmin(token);
  const sh = _sheet(SHEETS.UNITS);
  const values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(kodeUnit)) {
      sh.deleteRow(i + 1);
      return { ok: true, message: 'Unit dihapus.' };
    }
  }
  return { ok: false, message: 'Unit tidak ditemukan.' };
}

/* ----------------------- API: Transaksi ----------------------- */
function listTransaksi(token, filter) {
  _auth(token);
  filter = filter || {};
  var data = _rows(SHEETS.TRX).data.map(function (r) {
    return {
      ID: r.ID,
      Tanggal: _dateStr(r.Tanggal),
      KodeUnit: r.KodeUnit,
      Jenis: r.Jenis,
      Kategori: r.Kategori,
      Keterangan: r.Keterangan,
      Jumlah: Number(r.Jumlah) || 0,
      Operator: r.Operator
    };
  });
  if (filter.unit) data = data.filter(function (r) { return r.KodeUnit === filter.unit; });
  if (filter.jenis) data = data.filter(function (r) { return r.Jenis === filter.jenis; });
  if (filter.dari) data = data.filter(function (r) { return r.Tanggal >= filter.dari; });
  if (filter.sampai) data = data.filter(function (r) { return r.Tanggal <= filter.sampai; });
  data.sort(function (a, b) { return a.Tanggal < b.Tanggal ? 1 : -1; });
  return data;
}

function saveTransaksi(token, trx) {
  const s = _auth(token);
  const sh = _sheet(SHEETS.TRX);
  const jumlah = Number(trx.Jumlah) || 0;
  if (jumlah <= 0) return { ok: false, message: 'Jumlah harus lebih dari 0.' };
  if (!trx.Tanggal) return { ok: false, message: 'Tanggal wajib diisi.' };
  if (!trx.KodeUnit) return { ok: false, message: 'Unit usaha wajib dipilih.' };

  if (trx.ID) {
    const values = sh.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(trx.ID)) {
        sh.getRange(i + 1, 2, 1, 7).setValues([[trx.Tanggal, trx.KodeUnit, trx.Jenis, trx.Kategori || '', trx.Keterangan || '', jumlah, s.u]]);
        return { ok: true, message: 'Transaksi diperbarui.' };
      }
    }
    return { ok: false, message: 'Transaksi tidak ditemukan.' };
  }
  const id = 'TRX' + Date.now();
  sh.appendRow([id, trx.Tanggal, trx.KodeUnit, trx.Jenis, trx.Kategori || '', trx.Keterangan || '', jumlah, s.u, new Date()]);
  return { ok: true, message: 'Transaksi disimpan.', id: id };
}

function deleteTransaksi(token, id) {
  _auth(token);
  const sh = _sheet(SHEETS.TRX);
  const values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      sh.deleteRow(i + 1);
      return { ok: true, message: 'Transaksi dihapus.' };
    }
  }
  return { ok: false, message: 'Transaksi tidak ditemukan.' };
}

/* ----------------------- API: Dashboard / Laporan ----------------------- */
function getDashboard(token, filter) {
  _auth(token);
  const trx = listTransaksi(token, filter);
  const units = {};
  _rows(SHEETS.UNITS).data.forEach(function (u) { units[u.KodeUnit] = u.NamaUnit; });

  var totalIn = 0, totalOut = 0;
  const perBulan = {};   // 'YYYY-MM' -> {in, out}
  const perUnit = {};    // kode -> {in, out}
  const perKategori = {};// kategori -> jumlah (pengeluaran)

  trx.forEach(function (r) {
    const isIn = r.Jenis === 'Pemasukan';
    if (isIn) totalIn += r.Jumlah; else totalOut += r.Jumlah;

    const ym = (r.Tanggal || '').substring(0, 7);
    if (!perBulan[ym]) perBulan[ym] = { in: 0, out: 0 };
    perBulan[ym][isIn ? 'in' : 'out'] += r.Jumlah;

    if (!perUnit[r.KodeUnit]) perUnit[r.KodeUnit] = { in: 0, out: 0 };
    perUnit[r.KodeUnit][isIn ? 'in' : 'out'] += r.Jumlah;

    if (!isIn) {
      const k = r.Kategori || 'Lainnya';
      perKategori[k] = (perKategori[k] || 0) + r.Jumlah;
    }
  });

  const bulanKeys = Object.keys(perBulan).sort();
  const unitKeys = Object.keys(perUnit);

  return {
    totalIn: totalIn,
    totalOut: totalOut,
    saldo: totalIn - totalOut,
    jumlahTransaksi: trx.length,
    timeseries: {
      labels: bulanKeys,
      pemasukan: bulanKeys.map(function (k) { return perBulan[k].in; }),
      pengeluaran: bulanKeys.map(function (k) { return perBulan[k].out; })
    },
    perUnit: {
      labels: unitKeys.map(function (k) { return units[k] || k; }),
      laba: unitKeys.map(function (k) { return perUnit[k].in - perUnit[k].out; }),
      pemasukan: unitKeys.map(function (k) { return perUnit[k].in; }),
      pengeluaran: unitKeys.map(function (k) { return perUnit[k].out; })
    },
    perKategori: {
      labels: Object.keys(perKategori),
      data: Object.keys(perKategori).map(function (k) { return perKategori[k]; })
    }
  };
}

/* Laporan dengan saldo berjalan (urut tanggal naik) */
function getLaporan(token, filter) {
  _auth(token);
  const units = {};
  _rows(SHEETS.UNITS).data.forEach(function (u) { units[u.KodeUnit] = u.NamaUnit; });
  var data = listTransaksi(token, filter);
  data.sort(function (a, b) { return a.Tanggal > b.Tanggal ? 1 : (a.Tanggal < b.Tanggal ? -1 : 0); });
  var saldo = 0;
  const rows = data.map(function (r) {
    saldo += (r.Jenis === 'Pemasukan' ? r.Jumlah : -r.Jumlah);
    return {
      Tanggal: r.Tanggal, Unit: units[r.KodeUnit] || r.KodeUnit, Jenis: r.Jenis,
      Kategori: r.Kategori, Keterangan: r.Keterangan,
      Masuk: r.Jenis === 'Pemasukan' ? r.Jumlah : 0,
      Keluar: r.Jenis === 'Pengeluaran' ? r.Jumlah : 0,
      Saldo: saldo
    };
  });
  return { appName: APP_NAME, rows: rows, saldoAkhir: saldo };
}

/* ----------------------- API: Users (admin) ----------------------- */
function listUsers(token) {
  _requireAdmin(token);
  return _rows(SHEETS.USERS).data.map(function (u) {
    return { Username: u.Username, Nama: u.Nama, Role: u.Role, Aktif: u.Aktif !== false };
  });
}

function saveUser(token, user) {
  _requireAdmin(token);
  const sh = _sheet(SHEETS.USERS);
  const values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]).toLowerCase() === String(user.Username).toLowerCase()) {
      sh.getRange(i + 1, 3).setValue(user.Nama);
      sh.getRange(i + 1, 4).setValue(user.Role);
      sh.getRange(i + 1, 5).setValue(user.Aktif !== false);
      if (user.Password) sh.getRange(i + 1, 2).setValue(_hash(user.Password));
      return { ok: true, message: 'Pengguna diperbarui.' };
    }
  }
  if (!user.Password) return { ok: false, message: 'Password wajib untuk pengguna baru.' };
  sh.appendRow([user.Username, _hash(user.Password), user.Nama, user.Role || 'operator', true, new Date()]);
  return { ok: true, message: 'Pengguna ditambahkan.' };
}

function deleteUser(token, username) {
  const s = _requireAdmin(token);
  if (s.u.toLowerCase() === String(username).toLowerCase()) return { ok: false, message: 'Tidak bisa menghapus akun sendiri.' };
  const sh = _sheet(SHEETS.USERS);
  const values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]).toLowerCase() === String(username).toLowerCase()) {
      sh.deleteRow(i + 1);
      return { ok: true, message: 'Pengguna dihapus.' };
    }
  }
  return { ok: false, message: 'Pengguna tidak ditemukan.' };
}

/* ----------------------- Helpers ----------------------- */
function _dateStr(d) {
  if (!d) return '';
  if (Object.prototype.toString.call(d) === '[object Date]') {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(d).substring(0, 10);
}

/* ----------------------- Web App ----------------------- */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
