# 📦 Sistem Stok Faminis Barokah — PWA

Aplikasi manajemen stok terpadu untuk **Faminis Barokah** (Pasar Klewer, Solo). Berjalan sebagai **Progressive Web App (PWA)** — bisa di-install di HP Android maupun iOS seperti aplikasi native, tanpa perlu Play Store atau App Store.

**Live App:** https://arrosyid7.github.io/faminisstock/

---

## 🏗️ Arsitektur Sistem

```
┌─────────────────────────────────────┐
│  HP Kasir / Owner / Master          │
│  Browser (Chrome / Safari)          │
│  GitHub Pages → PWA terinstall      │
└──────────────┬──────────────────────┘
               │ fetch() POST
               ▼
┌─────────────────────────────────────┐
│  Google Apps Script Web App         │
│  (Code.gs — router + business logic)│
└──────────────┬──────────────────────┘
               │ SpreadsheetApp
               ▼
┌─────────────────────────────────────┐
│  Google Spreadsheet                 │
│  (Database: Produk, Stok, Transaksi)│
└─────────────────────────────────────┘
```

| Komponen | Teknologi |
|----------|-----------|
| Frontend | HTML + Tailwind CSS + Vanilla JS |
| Hosting | GitHub Pages (gratis) |
| Backend API | Google Apps Script Web App |
| Database | Google Spreadsheet |
| PWA | manifest.json + service-worker.js |

---

## 🚀 Cara Deploy (Langkah demi Langkah)

### Langkah 1 — Siapkan Google Apps Script

1. Buka [script.google.com](https://script.google.com) → **New Project**
2. Rename project jadi `Faminis Stok System`
3. Copy-paste isi file `Code.gs` ke editor (hapus yang sudah ada dulu)
4. Buat file baru: **File → New → HTML file** → beri nama `Index` → paste isi `Index.html`
5. Buat file baru lagi → nama `TemplatePdf` → paste isi `TemplatePdf.html`
6. Di bagian atas `Code.gs`, ganti jika perlu:
   ```javascript
   const SPREADSHEET_ID = "ID_SPREADSHEET_KAMU";
   const EMAIL_OWNER = "email_owner@gmail.com";
   ```
7. **Jalankan `inisialisasiSistem` sekali:** Klik ▶ Run → pilih fungsi `inisialisasiSistem` → izinkan akses
8. **Deploy sebagai Web App:**
   - Klik **Deploy → New deployment**
   - Type: **Web app**
   - Execute as: **Me (email kamu)**
   - Who has access: **Anyone**
   - Klik **Deploy** → Salin URL yang muncul (bentuknya: `https://script.google.com/macros/s/ABC.../exec`)

### Langkah 2 — Isi URL API di Index.html

Buka `Index.html`, cari baris ini (sekitar baris 451):

```javascript
const URL_API = "https://script.google.com/macros/s/GANTI_DENGAN_URL_KAMU/exec";
```

Ganti string URL-nya dengan URL Web App yang kamu salin di Langkah 1.

> ⚠️ **Penting:** Kalau kamu deploy ulang Apps Script (bukan edit/update versi lama), URL-nya akan BERUBAH. Jangan lupa update di sini juga.

### Langkah 3 — Upload ke GitHub & Aktifkan Pages

1. Pastikan repo sudah ada di GitHub (`github.com/Arrosyid7/faminisstock`)
2. Upload/push semua file ke branch `main`:
   ```bash
   git init
   git add .
   git commit -m "Initial PWA setup"
   git remote add origin https://github.com/Arrosyid7/faminisstock.git
   git push -u origin main
   ```
3. Di GitHub → **Settings → Pages**
4. Source: **GitHub Actions**
5. Tunggu ~1 menit → cek tab **Actions** untuk melihat status deploy
6. Kalau hijau ✅, app sudah live di `https://arrosyid7.github.io/faminisstock/`

---

## 📱 Cara Install di HP

### Android (Chrome)
1. Buka `https://arrosyid7.github.io/faminisstock/` di Chrome
2. Akan muncul banner biru **"Install app ini ke HP"** — klik **Install**
3. Atau: titik tiga (⋮) → **Add to Home Screen**
4. App akan muncul di layar utama seperti aplikasi biasa

### iOS (Safari)
1. Buka URL di **Safari** (bukan Chrome/Firefox)
2. Tap tombol **Share** (kotak dengan panah ke atas) di bagian bawah
3. Scroll ke bawah → pilih **"Add to Home Screen"**
4. Beri nama → tap **Add**

> 💡 **Catatan iOS:** Di iPhone/iPad, tombol Install otomatis tidak tersedia karena keterbatasan iOS Safari. User harus install manual lewat menu Share. Banner "Install" di app hanya muncul di Android Chrome.

---

## 👥 Akun Default

Setelah `inisialisasiSistem` dijalankan, akun berikut tersedia (semua PIN awal: `123456`):

| Username | Peran | Akses |
|----------|-------|-------|
| `master` | Admin Master | Semua fitur (transaksi, label, produk, laporan) |
| `owner` | Pemilik | Laporan & kirim PDF ke email |
| `toko` | Kasir TOKO | Transaksi di lokasi TOKO saja |
| `ruko1` | Kasir RUKO1 | Transaksi di RUKO1 saja |
| `ruko2` | Kasir RUKO2 | Transaksi di RUKO2 saja |
| `ruko3` | Kasir RUKO3 | Transaksi di RUKO3 saja |
| `ruko4` | Kasir RUKO4 | Transaksi di RUKO4 saja |

> 🔐 **Wajib ganti PIN** semua akun setelah pertama kali login! Edit langsung di sheet **Pengguna** di Google Spreadsheet.

---

## 🔧 Fitur Lengkap

| Fitur | Peran yang Bisa Akses |
|-------|----------------------|
| Input transaksi penjualan (ecer/grosir) | Kasir, Master |
| Transfer stok antar lokasi | Kasir, Master |
| Terima barang dari penjahit/supplier | Kasir TOKO, Master |
| Scan barcode via kamera HP | Kasir, Master |
| Generator label barcode (Code128) | Master |
| Kelola produk (tambah/edit) | Master |
| Laporan stok real-time | Owner, Master, Kasir Admin |
| Grafik produk terlaris (harian/bulanan/tahunan) | Owner, Master |
| Laporan keuangan (omset per lokasi) | Owner, Master |
| Kirim laporan PDF ke email owner | Owner, Master |
| Peringatan stok menipis (badge + email otomatis) | Semua |
| Install sebagai PWA di HP | Semua |

---

## 📂 Struktur File

```
faminisstock/
├── Index.html              ← Frontend utama (SPA)
├── manifest.json           ← PWA manifest
├── service-worker.js       ← Offline cache & install
├── offline.html            ← Halaman saat tidak ada internet
├── Code.gs                 ← Backend Apps Script (diupload ke script.google.com)
├── TemplatePdf.html        ← Template laporan PDF (diupload ke Apps Script)
├── icons/
│   ├── icon-192.png        ← Icon PWA 192×192
│   ├── icon-512.png        ← Icon PWA 512×512
│   └── icon-maskable-512.png ← Icon maskable (Android)
├── .github/
│   └── workflows/
│       └── deploy.yml      ← Auto-deploy ke GitHub Pages
└── README.md               ← Dokumentasi ini
```

---

## ❓ Troubleshooting

**Q: Login berhasil tapi data tidak muncul?**
A: Pastikan URL_API di `Index.html` sudah benar. Buka URL API langsung di browser dan lihat apakah ada response JSON.

**Q: Error "Aksi tidak dikenal" atau 404?**
A: Deploy ulang Apps Script — pastikan semua file (`Code.gs`, `Index`, `TemplatePdf`) sudah ada di project yang sama.

**Q: App tidak bisa di-install di HP?**
A: Pastikan diakses lewat HTTPS (GitHub Pages sudah HTTPS). Di Android, gunakan Chrome. Di iOS, gunakan Safari.

**Q: Transaksi gagal dengan "Akun tidak aktif"?**
A: Cek sheet `Pengguna` di Spreadsheet — pastikan kolom `Status` isinya `Aktif` (persis, huruf A kapital).

---

*Dibuat dengan ❤️ untuk Faminis Barokah — Pasar Klewer, Solo*
