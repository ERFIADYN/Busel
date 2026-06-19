# BUMDes Finance — Batuatas Liwu

Aplikasi pencatatan keuangan Badan Usaha Milik Desa (BUMDes) berbasis **Google Apps Script** dengan **Google Spreadsheet sebagai database**.

## Fitur
- 🔐 Login user/password (hash SHA-256 + salt, sesi token 8 jam)
- 🏪 Multi unit usaha BUMDes
- 📝 Catat pemasukan & pengeluaran
- 💰 Saldo kas berjalan otomatis
- 📊 Dashboard + grafik (arus kas, laba/rugi per unit, komposisi pengeluaran)
- 📄 Laporan kas dengan saldo berjalan + Cetak/Export PDF (via print browser)
- 👥 Manajemen pengguna (khusus admin)

## File
| File | Fungsi |
|------|--------|
| `Code.gs` | Backend / logika server + akses Spreadsheet |
| `Index.html` | Antarmuka web (frontend) |
| `appsscript.json` | Manifest deploy web app |

## Cara Deploy (ringkas)
1. Buka https://script.google.com → **New project**.
2. Buat file `Code.gs` (paste isi Code.gs) dan file HTML bernama **`Index`** (paste isi Index.html).
3. (Opsional) Aktifkan manifest: ⚙️ Project Settings → centang "Show appsscript.json", lalu paste isi appsscript.json.
4. Buat Spreadsheet baru → salin ID dari URL → isi ke `SPREADSHEET_ID` di Code.gs.
   (Atau buat project dari Extensions → Apps Script di dalam Spreadsheet agar otomatis terikat.)
5. Jalankan fungsi **`setupDatabase`** sekali (pilih dari dropdown fungsi → Run → izinkan akses).
6. **Deploy → New deployment → Web app** → Execute as: *Me*, Who has access: *Anyone* → Deploy.
7. Buka URL web app. Login awal: **admin / admin123** (segera ganti password).

## Keamanan
- Ganti nilai `SALT` di Code.gs dengan string rahasia Anda sendiri **sebelum** membuat user.
- Segera ganti password admin default.
- Role: `admin` (akses penuh) & `operator` (catat transaksi).
