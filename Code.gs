// ID Spreadsheet Database.
// PENTING: ganti string di bawah ini dengan ID Spreadsheet database katalog Faminis Barokah kamu.
// Caranya: buka Google Sheets-nya, salin bagian ID dari URL-nya.
// Contoh URL: https://docs.google.com/spreadsheets/d/ID_ADA_DI_BAGIAN_INI/edit
// (Sebelumnya kode ini pakai SpreadsheetApp.getActiveSpreadsheet(), yang HANYA jalan kalau skrip
// ditempel langsung di dalam Spreadsheet lewat Extensions > Apps Script. Kalau skrip berdiri sendiri
// di script.google.com, itu mengembalikan null dan bikin seluruh sistem gagal total. ID yang
// di-hardcode ini aman untuk kedua skenario.)
const SPREADSHEET_ID = "1XswTZzApMzpt508hMEDFQIPWiWX-kK2xUKnGVRqCKkw";

// GANTI dengan email owner yang akan menerima laporan PDF otomatis
const EMAIL_OWNER = "zanuarharun7@gmail.com";

function getDb() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// ==========================================
// FUNGSI SETUP OTOMATIS (JALANKAN SEKALI)
// ==========================================
function inisialisasiSistem() {
  const ss = getDb();

  // Sheet "Produk" -- sheet MASTER PRODUK milik sistem stok ini sendiri, BERDIRI SENDIRI
  // (tidak lagi disambungkan/disamakan dengan sheet katalog website faminisbarokah.my.id yang
  // lama). Dipakai untuk INPUT data produk + menampilkan total stok gabungan otomatis dari
  // semua lokasi (Toko + Ruko1-4).
  pastikanSheetProdukSiap(ss);

  // Sheet-sheet sistem stok -- aman dibuat/diperbarui otomatis. Kalau sheet SUDAH ADA dan sudah
  // ada isinya, data lama TIDAK PERNAH dihapus -- hanya kolom header yang belum ada yang
  // ditambahkan di ujung kanan (lihat loop di bawah).
  const strukturSheetSistem = {
    'Lokasi': ['ID_Lokasi', 'Nama Lokasi', 'Tipe'],
    'Pengguna': ['Username', 'Nama Lengkap', 'PIN', 'Peran', 'ID_Lokasi', 'Status'],
    // Barang_Masuk HANYA untuk restock: dari penjahit sendiri ATAU kulakan dari supplier luar.
    // Kolom Sumber & Nama_Supplier ditambahkan di ujung supaya baris lama (kalau sudah ada) tetap valid.
    'Barang_Masuk': ['Waktu', 'ID_Transaksi', 'SKU', 'ID_Lokasi', 'Qty', 'Tipe_Transaksi', 'Keterangan', 'Pengguna', 'Sumber', 'Nama_Supplier'],
    // Barang_Keluar HANYA untuk transaksi penjualan ecer/grosir (transfer TIDAK lewat sini lagi).
    // Kolom Harga_Satuan menyimpan harga jual PADA SAAT transaksi terjadi (bukan harga sekarang)
    // supaya laporan keuangan tetap akurat historis walau harga produk berubah di kemudian hari.
    'Barang_Keluar': ['Waktu', 'ID_Transaksi', 'SKU', 'ID_Lokasi', 'Qty', 'Tipe_Transaksi', 'Keterangan', 'Pengguna', 'Harga_Satuan'],
    // Transfer sekarang punya sheet SENDIRI, terpisah dari Barang_Masuk/Barang_Keluar --
    // dipecah 2: sisi lokasi yang MENGIRIM (Transfer_Keluar) dan sisi yang MENERIMA (Transfer_Masuk).
    'Transfer_Keluar': ['Waktu', 'ID_Transfer', 'SKU', 'Dari_Lokasi', 'Ke_Lokasi', 'Qty', 'Pengguna'],
    'Transfer_Masuk': ['Waktu', 'ID_Transfer', 'SKU', 'ID_Lokasi', 'Dari_Lokasi', 'Qty', 'Pengguna'],
    // Daftar supplier yang pernah diketik -- tumbuh otomatis, dipakai untuk saran/autocomplete.
    'Supplier': ['Nama_Supplier'],
    // Arsip penjualan bulanan (ringkasan per SKU per lokasi per bulan), diisi otomatis oleh
    // arsipkanBarangKeluarBulanan() tiap awal bulan, lalu Barang_Keluar direset. Kolom Omset_xxx
    // dipakai untuk laporan keuangan tahunan supaya tetap akurat walau Barang_Keluar sudah direset.
    'Data_Tahunan': ['Tahun', 'Bulan', 'ID_Produk', 'Nama Produk', 'ID_Lokasi', 'Qty_Ecer', 'Qty_Grosir', 'Total_Qty', 'Tanggal_Diarsipkan', 'Omset_Ecer', 'Omset_Grosir', 'Total_Omset']
  };

  for (const namaSheet in strukturSheetSistem) {
    let sheet = ss.getSheetByName(namaSheet);
    const headerBaku = strukturSheetSistem[namaSheet];

    if (!sheet) {
      sheet = ss.insertSheet(namaSheet);
      sheet.getRange(1, 1, 1, headerBaku.length).setValues([headerBaku]);
    } else if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headerBaku.length).setValues([headerBaku]);
    } else {
      // Sheet sudah ada & sudah ada isinya -- JANGAN pernah dihapus/ditimpa. Cuma tambahkan
      // kolom header yang belum ada di ujung kanan, supaya data lama tetap aman & tetap valid.
      const lastCol = sheet.getLastColumn();
      const headerSaatIni = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => h.toString().trim());
      headerBaku.forEach(function (h) {
        if (headerSaatIni.indexOf(h) === -1) {
          sheet.getRange(1, sheet.getLastColumn() + 1).setValue(h);
        }
      });
    }
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight("bold").setBackground("#EEDDC8");
  }

  // Data Lokasi awal (hanya kalau sheet baru/kosong, tidak menimpa data yang sudah ada)
  const sheetLokasi = ss.getSheetByName('Lokasi');
  if (sheetLokasi.getLastRow() <= 1) {
    const dataLokasi = [
      ['TOKO', 'Gudang Pusat / Toko Utama', 'Pusat'],
      ['RUKO1', 'Ruko Klewer 1', 'Cabang'],
      ['RUKO2', 'Ruko Klewer 2', 'Cabang'],
      ['RUKO3', 'Ruko Klewer 3', 'Cabang'],
      ['RUKO4', 'Ruko Klewer 4', 'Cabang']
    ];
    sheetLokasi.getRange(2, 1, dataLokasi.length, 3).setValues(dataLokasi);
  }

  // Sheet stok per lokasi (Stok_Gudang, Stok_RUKO1, dst) -- inilah "sumber kebenaran" stok
  // yang dipakai laporan, dan yang otomatis berubah tiap ada transaksi masuk/keluar/transfer.
  // Aman dipanggil berkali-kali: sheet yang sudah ada TIDAK ditimpa, cuma ditambah baris SKU
  // baru yang belum terdaftar.
  pastikanSemuaSheetStokSiap(ss);

  // Pasang/perbarui rumus Stok_Total di sheet Produk (SUMIF otomatis dari semua sheet Stok_xxx)
  pastikanFormulaStokTotal(ss);

  // Data Pengguna awal (hanya kalau sheet baru/kosong)
  const sheetPengguna = ss.getSheetByName('Pengguna');
  if (sheetPengguna.getLastRow() <= 1) {
    const dataPengguna = [
      ['master', 'Administrator Master', '123456', 'master', 'SEMUA', 'Aktif'],
      ['owner', 'Pemilik Faminis Barokah', '123456', 'owner', 'SEMUA', 'Aktif'],
      ['toko', 'Kasir Gudang Pusat', '123456', 'kasir', 'TOKO', 'Aktif'],
      ['ruko1', 'Kasir Ruko 1', '123456', 'kasir', 'RUKO1', 'Aktif'],
      ['ruko2', 'Kasir Ruko 2', '123456', 'kasir', 'RUKO2', 'Aktif'],
      ['ruko3', 'Kasir Ruko 3', '123456', 'kasir', 'RUKO3', 'Aktif'],
      ['ruko4', 'Kasir Ruko 4', '123456', 'kasir', 'RUKO4', 'Aktif']
    ];
    sheetPengguna.getRange(2, 1, dataPengguna.length, 6).setValues(dataPengguna);
  }

  Logger.log("Setup selesai!");
}

/**
 * Memastikan sheet "Produk" siap dipakai sebagai sheet MASTER PRODUK khusus sistem stok ini,
 * BERDIRI SENDIRI (tidak disambungkan ke sheet katalog website lama). Dipakai untuk INPUT data
 * produk (ID, Nama, Kategori, Harga) dan menampilkan Stok_Total (jumlah gabungan dari semua
 * lokasi) secara otomatis lewat rumus SUMIF -- lihat pastikanFormulaStokTotal().
 * - Kalau sheet belum ada -> dibuat baru dengan skema standar di bawah.
 * - Kalau sheet SUDAH ADA -> data yang sudah ada TIDAK PERNAH dihapus, cuma dipastikan kolom
 *   "Stok_Total" dan "Stok_Minimum" tersedia (ditambahkan di ujung kanan kalau belum ada).
 *   Stok_Minimum dipakai sebagai ambang batas peringatan stok menipis per SKU (lihat
 *   ambilSemuaStok() dan tambahTransaksi()).
 */
function pastikanSheetProdukSiap(ss) {
  let sheet = ss.getSheetByName('Produk');
  const headerStandar = ['ID_Produk', 'Nama Produk', 'Kategori', 'Harga Ecer', 'Harga Grosir', 'Stok_Total', 'Stok_Minimum'];

  if (!sheet) {
    sheet = ss.insertSheet('Produk');
    sheet.getRange(1, 1, 1, headerStandar.length).setValues([headerStandar]);
    sheet.getRange(1, 1, 1, headerStandar.length).setFontWeight("bold").setBackground("#EEDDC8");
    return;
  }

  const lastCol = sheet.getLastColumn();
  const headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => h.toString().trim()) : [];
  // Loop generik: tambahkan kolom apapun yang belum ada di ujung kanan, tanpa mengganggu kolom
  // & data yang sudah ada -- aman dipanggil berkali-kali dan aman untuk spreadsheet lama maupun baru.
  ['Stok_Total', 'Stok_Minimum'].forEach(function (namaKolom) {
    const headerSaatIni = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.toString().trim());
    if (headerSaatIni.indexOf(namaKolom) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(namaKolom).setFontWeight("bold").setBackground("#EEDDC8");
    }
  });
}

/**
 * Cari nomor kolom (1-based) berdasarkan nama header persis. Return null kalau tidak ditemukan.
 */
function cariKolom(sheet, namaKolom) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return null;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (let i = 0; i < headers.length; i++) {
    if (headers[i].toString().trim() === namaKolom) return i + 1;
  }
  return null;
}

/**
 * Memasang/memperbarui rumus SUMIF di kolom "Stok_Total" sheet Produk, menjumlahkan stok SKU
 * yang sama dari SEMUA sheet Stok_xxx (Toko + Ruko1-4, otomatis ikut kalau ada lokasi baru).
 * Karena berbentuk RUMUS (bukan angka hasil skrip), angkanya selalu update sendiri secara
 * realtime begitu ada perubahan di sheet Stok_xxx manapun -- tidak perlu tunggu trigger jalan.
 * Aman dipanggil berkali-kali (dipanggil dari inisialisasiSistem(), sinkronSemuaSheetStok(), dan
 * tambahProduk()).
 */
function pastikanFormulaStokTotal(ss) {
  const sheetProduk = ss.getSheetByName('Produk');
  if (!sheetProduk) return;
  const lastRow = sheetProduk.getLastRow();
  if (lastRow < 2) return;

  const kolomStokTotal = cariKolom(sheetProduk, 'Stok_Total');
  if (!kolomStokTotal) return;

  const daftarLokasi = ambilDaftarLokasi();
  const idProdukList = sheetProduk.getRange(2, 1, lastRow - 1, 1).getValues();

  const formulas = idProdukList.map(function (row, idx) {
    const rowNum = idx + 2;
    if (!row[0]) return [''];
    const bagianFormula = daftarLokasi.map(function (lok) {
      const namaSheetStok = dapatkanNamaSheetStok(lok.id);
      return `SUMIF(${namaSheetStok}!$A:$A,$A${rowNum},${namaSheetStok}!$C:$C)`;
    }).join('+');
    return ['=' + (bagianFormula || '0')];
  });

  sheetProduk.getRange(2, kolomStokTotal, formulas.length, 1).setValues(formulas);
}

/**
 * Menambah produk baru ke sheet Produk (khusus Admin Master), lalu langsung memastikan baris
 * stok 0 dibuat di semua lokasi + rumus Stok_Total terpasang untuk produk ini.
 * stokMinimum: ambang batas stok (dijumlah dari semua lokasi) yang memicu peringatan stok
 * menipis -- lihat ambilSemuaStok() dan tambahTransaksi(). Boleh diisi 0 kalau belum mau
 * dipasangi peringatan untuk produk ini.
 *
 * CATATAN: produk TIDAK lagi punya "harga katalog" (Harga Ecer/Harga Grosir) -- harga jual
 * sekarang murni diketik manual oleh admin/kasir di tiap transaksi penjualan (lihat
 * tambahTransaksi()), supaya tidak ada 2 sumber harga yang membingungkan. Kolom "Harga Ecer"/
 * "Harga Grosir" di sheet (kalau masih ada peninggalan versi lama) sengaja dibiarkan apa adanya,
 * tidak diisi/diubah lagi lewat sini.
 *
 * Baris baru dibangun berdasarkan URUTAN HEADER SEBENARNYA di sheet (bukan urutan tetap),
 * supaya tetap aman walau posisi kolom Stok_Total/Stok_Minimum berbeda-beda antar spreadsheet
 * lama & baru (lihat pastikanSheetProdukSiap()).
// ==========================================
// SHEET STOK PER LOKASI (Stok_Gudang, Stok_RUKO1, dst)
// ==========================================
// Arsitektur baru: stok TIDAK lagi dihitung ulang setiap kali dari riwayat Barang_Masuk/
// Barang_Keluar (lambat & tidak bisa dikoreksi manual). Sekarang tiap lokasi punya sheet stok
// sendiri berisi angka stok TERKINI, yang otomatis bertambah/berkurang setiap ada transaksi.
// Sheet Barang_Masuk/Barang_Keluar/Transfer_Masuk/Transfer_Keluar tetap ada & tetap dicatat --
// itu jadi BUKU JURNAL/riwayat transaksi untuk audit dan laporan produk terlaris, bukan sumber
// stok lagi.
//
// Keuntungan tambahan: karena angka stoknya adalah SEL SHEET BIASA, kamu (misalnya saat Stock
// Opname) BOLEH mengoreksi angkanya langsung di sheet Stok_xxx kalau ketahuan selisih -- laporan
// & aplikasi akan otomatis memakai angka terbaru itu.

/**
 * Nama sheet stok untuk 1 ID lokasi. Lokasi TOKO (gudang pusat) sengaja dinamai "Stok_Gudang"
 * biar jelas dibaca, lokasi lain dinamai "Stok_" + ID_Lokasi (mis. Stok_RUKO1, Stok_RUKO2).
 */
function dapatkanNamaSheetStok(idLokasi) {
  return idLokasi === 'TOKO' ? 'Stok_Gudang' : 'Stok_' + idLokasi;
}

function buatSheetStokBaru(ss, namaSheet) {
  const sheet = ss.insertSheet(namaSheet);
  const header = ['ID_Produk', 'Nama Produk', 'Stok'];
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#EEDDC8');
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 220);
  return sheet;
}

/**
 * Memastikan SEMUA lokasi (dari sheet Lokasi) punya sheet stok, dan semua SKU (dari sheet
 * Produk) sudah punya barisnya di tiap sheet stok tsb. TIDAK PERNAH menimpa angka stok yang
 * sudah ada -- SKU yang sudah punya baris dibiarkan apa adanya, cuma SKU baru yang ditambahkan
 * (dengan stok awal 0). Aman dan disarankan dijalankan ulang kapan pun ada produk/lokasi baru
 * (jalankan manual lewat menu Extensions > Apps Script > pilih fungsi ini > Run, atau otomatis
 * lewat inisialisasiSistem()/tambahProduk()).
 */
function pastikanSemuaSheetStokSiap(ss) {
  const dataLokasi = ss.getSheetByName('Lokasi').getDataRange().getValues();
  const dataProduk = ss.getSheetByName('Produk').getDataRange().getValues();

  const daftarSku = [];
  for (let i = 1; i < dataProduk.length; i++) {
    const sku = dataProduk[i][0] ? dataProduk[i][0].toString().trim() : '';
    if (sku) daftarSku.push({ sku: sku, nama: dataProduk[i][1] });
  }

  for (let i = 1; i < dataLokasi.length; i++) {
    const idLokasi = dataLokasi[i][0];
    if (!idLokasi) continue;

    const namaSheet = dapatkanNamaSheetStok(idLokasi);
    let sheet = ss.getSheetByName(namaSheet);
    if (!sheet) sheet = buatSheetStokBaru(ss, namaSheet);

    const dataSheet = sheet.getDataRange().getValues();
    const skuTerdaftar = {};
    for (let j = 1; j < dataSheet.length; j++) {
      if (dataSheet[j][0]) skuTerdaftar[dataSheet[j][0].toString().trim()] = true;
    }

    const barisBaru = [];
    daftarSku.forEach(function (p) {
      if (!skuTerdaftar[p.sku]) barisBaru.push([p.sku, p.nama, 0]);
    });
    if (barisBaru.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, barisBaru.length, 3).setValues(barisBaru);
    }
  }
}

/**
 * Fungsi ini boleh dijalankan manual kapan saja (Extensions > Apps Script > pilih
 * "sinkronSemuaSheetStok" > Run) setelah kamu menambah produk baru atau lokasi/ruko baru,
 * supaya baris stoknya langsung muncul di semua sheet Stok_xxx + rumus Stok_Total terpasang.
 */
function sinkronSemuaSheetStok() {
  const ss = getDb();
  pastikanSemuaSheetStokSiap(ss);
  pastikanFormulaStokTotal(ss);
  Logger.log('Sinkronisasi sheet stok selesai.');
}

function ambilNamaProdukDariSku(sku) {
  const ss = getDb();
  const produk = ss.getSheetByName('Produk').getDataRange().getValues();
  for (let i = 1; i < produk.length; i++) {
    if (produk[i][0] && produk[i][0].toString().trim() === sku) return produk[i][1];
  }
  return sku;
}

/**
 * INTI dari update stok otomatis: menambah/mengurangi angka stok 1 SKU di 1 sheet lokasi.
 * delta positif = stok bertambah (barang masuk / diterima transfer)
 * delta negatif = stok berkurang (penjualan / dikirim transfer)
 * Kalau SKU belum punya baris di sheet lokasi ini (misal produk baru), baris baru otomatis
 * dibuat dengan stok awal = delta.
 * Dipanggil dari dalam tambahTransaksi() dan transferStok(), yang keduanya SUDAH memegang
 * LockService, jadi fungsi ini sengaja tidak mengunci lagi sendiri.
 */
function ubahStokLokasi(idLokasi, sku, delta) {
  const ss = getDb();
  const namaSheet = dapatkanNamaSheetStok(idLokasi);
  let sheet = ss.getSheetByName(namaSheet);
  if (!sheet) sheet = buatSheetStokBaru(ss, namaSheet);

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() === sku) {
      const stokBaru = (parseFloat(data[i][2]) || 0) + delta;
      sheet.getRange(i + 1, 3).setValue(stokBaru);
      return stokBaru;
    }
  }

  // SKU belum terdaftar di sheet stok lokasi ini -> tambahkan baris baru
  const namaProduk = ambilNamaProdukDariSku(sku);
  sheet.appendRow([sku, namaProduk, delta]);
  return delta;
}

// ==========================================
// SUPPLIER (untuk fitur "Terima Barang dari Luar")
// ==========================================
/**
 * Daftar nama supplier yang pernah diinput, dipakai untuk autocomplete di form "Terima Barang
 * dari Luar" -- supaya kasir tinggal pilih dari yang sudah ada, tidak perlu ketik ulang.
 */
function ambilDaftarSupplier() {
  const ss = getDb();
  const sheet = ss.getSheetByName('Supplier');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const hasil = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) hasil.push(data[i][0].toString().trim());
  }
  return hasil;
}

/**
 * Mendaftarkan nama supplier baru ke sheet Supplier kalau belum pernah tercatat (dicek tanpa
 * peduli besar/kecil huruf), supaya daftar autocomplete tumbuh otomatis.
 */
function pastikanSupplierTercatat(ss, namaSupplier) {
  if (!namaSupplier) return;
  let sheet = ss.getSheetByName('Supplier');
  if (!sheet) {
    sheet = ss.insertSheet('Supplier');
    sheet.getRange(1, 1).setValue('Nama_Supplier').setFontWeight('bold').setBackground('#EEDDC8');
  }
  const data = sheet.getDataRange().getValues();
  let sudahAda = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === namaSupplier.trim().toLowerCase()) {
      sudahAda = true;
      break;
    }
  }
  if (!sudahAda) sheet.appendRow([namaSupplier.trim()]);
}

// ==========================================
// 1. ROUTING & ENTRY POINT
// ==========================================
//
// CATATAN PENTING (versi GitHub Pages):
// Sekarang Code.gs ini berperan sebagai API backend saja. Frontend (Index.html) tidak lagi
// wajib dibuka lewat script.google.com -- boleh di-hosting di mana saja (GitHub Pages, dll),
// asalkan dia memanggil URL Web App ini lewat fetch(). Karena itu, doGet/doPost di bawah ini
// dibuat jadi "router" umum: menerima nama aksi + daftar parameter, lalu memanggil fungsi Apps
// Script yang sesuai dan selalu mengembalikan JSON.
//
// doGet('Index') lama TETAP dipertahankan (kalau kamu masih mau buka lewat URL Apps Script
// langsung, misal untuk uji coba cepat dari editor), tapi versi produksi kamu nantinya akan
// dibuka dari GitHub Pages, bukan dari sini.

// Daftar fungsi yang boleh dipanggil dari luar lewat API. HANYA fungsi yang didaftarkan di sini
// yang bisa diakses -- ini juga jadi lapisan pembatas dasar (tidak semua fungsi di file ini
// otomatis "terbuka" ke publik).
const DAFTAR_FUNGSI_API = {
  login: prosesLogin,
  ambilDaftarLokasi: ambilDaftarLokasi,
  ambilDaftarSupplier: ambilDaftarSupplier,
  ambilSemuaStok: ambilSemuaStok,
  tambahTransaksi: tambahTransaksi,
  transferStok: transferStok,
  ambilProdukTerlaris: ambilProdukTerlaris,
  ambilRingkasanPenjualan: ambilRingkasanPenjualan,
  ambilRingkasanBarangMasuk: ambilRingkasanBarangMasuk,
  ambilRingkasanTransfer: ambilRingkasanTransfer,
  ambilLaporanKeuangan: ambilLaporanKeuangan,
  kirimLaporanPDF: kirimLaporanPDF
};

function doGet(e) {
  // Dipanggil sebagai API: /exec?aksi=namaFungsi&params=["a","b"]
  if (e && e.parameter && e.parameter.aksi) {
    return tanganiPermintaanApi(e);
  }
  // Kompatibilitas lama: cek cepat stok lewat ?api=stok
  if (e && e.parameter && e.parameter.api === 'stok') {
    return ContentService.createTextOutput(JSON.stringify(ambilSemuaStok()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // PENTING: file HTML utama harus bernama persis "Index" di Apps Script Editor
  // (case-sensitive), karena createHtmlOutputFromFile('Index') mencarinya dengan nama itu.
  // Ini hanya dipakai kalau kamu buka Web App-nya langsung dari script.google.com.
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Faminis Barokah System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover');
}

function doPost(e) {
  // Dipanggil dari GitHub Pages lewat fetch(URL_API, {method:'POST', body: JSON.stringify({aksi, params})})
  return tanganiPermintaanApi(e);
}

/**
 * Router utama API. Menerima aksi + params dari query string (GET) atau body JSON (POST),
 * memanggil fungsi yang sesuai dari DAFTAR_FUNGSI_API, dan selalu membalas JSON dengan format:
 *   { sukses: true, data: <hasil fungsi> }   -> permintaan berhasil diproses
 *   { sukses: false, pesan: "..." }          -> permintaan gagal / aksi tidak dikenal
 * (Perhatikan: field "sukses" di sini menandakan REQUEST-nya berhasil diproses, bukan berarti
 * transaksinya sukses -- fungsi seperti tambahTransaksi tetap mengembalikan objek {sukses, pesan}
 * miliknya sendiri di dalam "data", jadi cek dua-duanya di sisi klien.)
 */
function tanganiPermintaanApi(e) {
  try {
    let aksi, params;

    if (e.postData && e.postData.contents) {
      const body = JSON.parse(e.postData.contents);
      aksi = body.aksi;
      params = body.params || [];
    } else {
      aksi = e.parameter.aksi;
      params = e.parameter.params ? JSON.parse(e.parameter.params) : [];
    }

    const fungsi = DAFTAR_FUNGSI_API[aksi];
    if (!fungsi) {
      throw new Error('Aksi tidak dikenal: ' + aksi);
    }

    const hasil = fungsi.apply(null, params);
    return keluaranJson({ sukses: true, data: hasil });
  } catch (err) {
    return keluaranJson({ sukses: false, pesan: err.message || 'Terjadi kesalahan di server.' });
  }
}

function keluaranJson(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// 2. OTENTIKASI / LOGIN
// ==========================================

function prosesLogin(username, pin) {
  const ss = getDb();
  const sheet = ss.getSheetByName('Pengguna');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username && data[i][2].toString() === pin) {
      if (data[i][5].toString().toUpperCase() !== 'AKTIF') {
        return { sukses: false, pesan: 'Akun dinonaktifkan. Hubungi Master.' };
      }
      return {
        sukses: true,
        namaLengkap: data[i][1],
        peran: data[i][3],
        id_lokasi: data[i][4],
        username: username
      };
    }
  }
  return { sukses: false, pesan: 'Username atau PIN salah!' };
}

/**
 * Verifikasi bahwa akun masih AKTIF, dicek ulang setiap kali ada transaksi (bukan cuma saat
 * login). Ini mencegah HP yang sudah lama login tetap bisa transaksi meski akunnya baru
 * dinonaktifkan Admin Master.
 */
function pastikanPenggunaAktif(username) {
  const ss = getDb();
  const sheet = ss.getSheetByName('Pengguna');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username) {
      return data[i][5].toString().toUpperCase() === 'AKTIF';
    }
  }
  return false;
}

function ambilPeranPengguna(username) {
  const ss = getDb();
  const sheet = ss.getSheetByName('Pengguna');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username) return data[i][3];
  }
  return null;
}

// ==========================================
// 3. MASTER LOKASI
// ==========================================
function ambilDaftarLokasi() {
  const ss = getDb();
  const sheet = ss.getSheetByName('Lokasi');
  const data = sheet.getDataRange().getValues();
  const hasil = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) hasil.push({ id: data[i][0], nama: data[i][1], tipe: data[i][2] });
  }
  return hasil;
}

// // ==========================================
// 4. INTI: AMBIL STOK
// ==========================================
/**
 * Membaca stok TERKINI langsung dari sheet-sheet Stok_xxx (satu sheet per lokasi).
 * Mengembalikan objek per SKU, termasuk per-lokasi stok rendah berdasarkan Stok_Minimum per SKU dari sheet Produk.
 *   { "SKU-001": { detail: {...}, Total: 12, TOKO: 2, RUKO1: 7,
 *                  stokMinimum: 5, lokasiRendah: [{idLokasi: 'TOKO', stok: 2}], stokRendah: true }, ... }
 */
function ambilSemuaStok() {
  const ss = getDb();
  const sheetProduk = ss.getSheetByName('Produk');
  const produk = sheetProduk.getDataRange().getValues();
  const daftarLokasi = ambilDaftarLokasi();
  const kolomMinimum = cariKolom(sheetProduk, 'Stok_Minimum');

  let petaProduk = {};
  for (let i = 1; i < produk.length; i++) {
    const sku = produk[i][0] ? produk[i][0].toString().trim() : "";
    if (!sku) continue;
    petaProduk[sku] = {
      nama: produk[i][1],
      kategori: produk[i][2],
      h_ecer: produk[i][3],
      h_grosir: produk[i][4],
      stokMinimum: kolomMinimum ? (parseFloat(produk[i][kolomMinimum - 1]) || 0) : 0
    };
  }

  let dataStok = {};

  daftarLokasi.forEach(function (lokasi) {
    const namaSheet = dapatkanNamaSheetStok(lokasi.id);
    const sheet = ss.getSheetByName(namaSheet);
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const sku = data[i][0] ? data[i][0].toString().trim() : "";
      if (!sku) continue;
      const stok = parseFloat(data[i][2]) || 0;

      if (!dataStok[sku]) {
        dataStok[sku] = { detail: petaProduk[sku] || { nama: data[i][1] || "Produk Tidak Dikenal", kategori: "-" }, Total: 0 };
      }
      dataStok[sku][lokasi.id] = stok;
      dataStok[sku].Total += stok;
    }
  });

  // Tandai status stok rendah PER LOKASI berdasarkan Stok_Minimum dari sheet Produk
  Object.keys(dataStok).forEach(function (sku) {
    const minimum = (dataStok[sku].detail && dataStok[sku].detail.stokMinimum) || 0;
    dataStok[sku].stokMinimum = minimum;
    dataStok[sku].lokasiRendah = [];

    if (minimum > 0) {
      daftarLokasi.forEach(function (lokasi) {
        const stokLokasi = dataStok[sku][lokasi.id] || 0;
        if (stokLokasi <= minimum) {
          dataStok[sku].lokasiRendah.push({
            idLokasi: lokasi.id,
            namaLokasi: lokasi.nama,
            stok: stokLokasi
          });
        }
      });
    }
    dataStok[sku].stokRendah = minimum > 0 && (dataStok[sku].lokasiRendah.length > 0 || dataStok[sku].Total <= minimum);
  });

  return dataStok;
}

// ==========================================
// 5. INTI: TAMBAH TRANSAKSI
// ==========================================
function tambahTransaksi(sku, id_lokasi, qty, tipe, keterangan, pengguna, sumber, namaSupplier, hargaManual) {
  if (!pastikanPenggunaAktif(pengguna)) {
    return { sukses: false, pesan: 'Akun tidak aktif. Hubungi Admin Master.' };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getDb();

    const isMasuk = tipe.indexOf('MASUK') !== -1;

    if (isMasuk && tipe !== 'MASUK_RESTOK') {
      throw new Error("Tipe transaksi masuk tidak dikenal. Barang Masuk hanya untuk restock (penjahit/supplier).");
    }
    if (!isMasuk && tipe !== 'KELUAR_ECER' && tipe !== 'KELUAR_GROSIR') {
      throw new Error("Tipe transaksi keluar tidak dikenal. Barang Keluar hanya untuk penjualan ecer/grosir.");
    }
    if (tipe === 'MASUK_RESTOK' && id_lokasi !== 'TOKO' && id_lokasi !== 'SEMUA') {
      throw new Error("Ditolak! Ruko tidak bisa menerima barang baru dari luar secara langsung.");
    }

    let namaSheet = isMasuk ? 'Barang_Masuk' : 'Barang_Keluar';
    const sheet = ss.getSheetByName(namaSheet);
    if (!sheet) throw new Error("Sheet " + namaSheet + " tidak ditemukan!");

    const qtyAbsolut = Math.abs(qty);
    const stokSebelum = ambilSemuaStok();
    const stokLokasiSebelum = (stokSebelum[sku] && stokSebelum[sku][id_lokasi]) || 0;

    if (!isMasuk) {
      if (stokLokasiSebelum < qtyAbsolut) {
        throw new Error(`Stok ${sku} di ${id_lokasi} tidak cukup (tersedia: ${stokLokasiSebelum}, diminta: ${qtyAbsolut}). Cek ulang stok fisik.`);
      }
    }

    const waktu = new Date();
    const idTrx = "TRX-" + new Date().getTime();

    if (namaSheet === 'Barang_Masuk') {
      const sumberFinal = sumber === 'Supplier' ? 'Supplier' : 'Penjahit';
      const namaSupplierFinal = sumberFinal === 'Supplier' ? (namaSupplier || '').toString().trim() : '';
      if (sumberFinal === 'Supplier' && !namaSupplierFinal) {
        throw new Error('Nama supplier wajib diisi kalau sumbernya bukan Penjahit.');
      }
      if (namaSupplierFinal) pastikanSupplierTercatat(ss, namaSupplierFinal);
      sheet.appendRow([waktu, idTrx, sku, id_lokasi, qtyAbsolut, tipe, keterangan, pengguna, sumberFinal, namaSupplierFinal]);
    } else {
      const hargaSatuan = Number(hargaManual) || 0;
      if (hargaSatuan <= 0) {
        throw new Error('Harga jual per pcs wajib diisi (lebih dari 0).');
      }
      sheet.appendRow([waktu, idTrx, sku, id_lokasi, qtyAbsolut, tipe, keterangan, pengguna, hargaSatuan]);
    }

    ubahStokLokasi(id_lokasi, sku, isMasuk ? qtyAbsolut : -qtyAbsolut);
    SpreadsheetApp.flush();

    // Peringatan stok menipis di lokasi transaksi terjadi (berdasarkan Stok_Minimum sheet Produk)
    let peringatanStok = null;
    const detailProdukSemua = ambilDetailProdukSemua(ss);
    const stokMinimum = (detailProdukSemua[sku] && detailProdukSemua[sku].stokMinimum) || 0;
    if (stokMinimum > 0) {
      const stokLokasiSesudah = stokLokasiSebelum + (isMasuk ? qtyAbsolut : -qtyAbsolut);
      if (stokLokasiSesudah <= stokMinimum) {
        const namaProduk = ambilNamaProdukDariSku(sku);
        peringatanStok = `⚠️ Stok ${namaProduk} (${sku}) di ${id_lokasi} tinggal ${stokLokasiSesudah} pcs, sudah di bawah/sama dengan batas minimum (${stokMinimum} pcs). Segera restock!`;
        if (stokLokasiSebelum > stokMinimum) {
          kirimEmailPeringatanStok(namaProduk, sku, stokLokasiSesudah, stokMinimum, id_lokasi);
        }
      }
    }

    return { sukses: true, pesan: 'Transaksi berhasil disimpan!', peringatanStok: peringatanStok };
  } catch (e) {
    return { sukses: false, pesan: e.message || 'Sistem sibuk. Silakan simpan lagi.' };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Kirim email peringatan stok menipis ke owner per lokasi.
 */
function kirimEmailPeringatanStok(namaProduk, sku, stokSekarang, stokMinimum, idLokasi) {
  try {
    MailApp.sendEmail({
      to: EMAIL_OWNER,
      subject: `⚠️ Peringatan Stok Menipis di ${idLokasi || 'Lokasi'}: ${namaProduk}`,
      body: `Halo Owner,\n\nStok produk berikut sudah mencapai atau di bawah batas minimum di lokasi ${idLokasi || '-'}:\n\n` +
            `Produk: ${namaProduk} (${sku})\n` +
            `Lokasi: ${idLokasi || 'Gabungan'}\n` +
            `Stok saat ini di lokasi ini: ${stokSekarang} pcs\n` +
            `Batas minimum (sheet Produk): ${stokMinimum} pcs\n\n` +
            `Segera lakukan restock ke lokasi tersebut.\n\n` +
            `- Sistem Stok Faminis Barokah`
    });
  } catch (e) {
    Logger.log('Gagal kirim email peringatan stok: ' + (e.message || e));
  }
}

// ==========================================
// 6. INTI: TRANSFER STOK (Toko->Ruko atau Ruko->Ruko)
// ==========================================
function transferStok(sku, dariLokasi, keLokasi, qty, pengguna) {
  if (!pastikanPenggunaAktif(pengguna)) {
    return { sukses: false, pesan: 'Akun tidak aktif. Hubungi Admin Master.' };
  }
  if (!dariLokasi || !keLokasi) {
    return { sukses: false, pesan: 'Pilih lokasi asal dan tujuan!' };
  }
  if (dariLokasi === keLokasi) {
    return { sukses: false, pesan: 'Lokasi asal dan tujuan tidak boleh sama!' };
  }
  qty = Math.abs(parseFloat(qty)) || 0;
  if (qty <= 0) {
    return { sukses: false, pesan: 'Jumlah harus lebih dari 0!' };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getDb();

    const stokSaatIni = ambilSemuaStok();
    const stokAsal = (stokSaatIni[sku] && stokSaatIni[sku][dariLokasi]) || 0;
    if (stokAsal < qty) {
      return { sukses: false, pesan: `Stok ${sku} di ${dariLokasi} tidak cukup (tersedia: ${stokAsal}).` };
    }

    const waktu = new Date();
    const idTransfer = "TRF-" + new Date().getTime();

    ss.getSheetByName('Transfer_Keluar').appendRow([waktu, idTransfer, sku, dariLokasi, keLokasi, qty, pengguna]);
    ss.getSheetByName('Transfer_Masuk').appendRow([waktu, idTransfer, sku, keLokasi, dariLokasi, qty, pengguna]);

    ubahStokLokasi(dariLokasi, sku, -qty);
    ubahStokLokasi(keLokasi, sku, qty);

    SpreadsheetApp.flush();

    // Cek peringatan stok menipis di lokasi asal setelah transfer
    let peringatanStok = null;
    const detailProdukSemua = ambilDetailProdukSemua(ss);
    const stokMinimum = (detailProdukSemua[sku] && detailProdukSemua[sku].stokMinimum) || 0;
    if (stokMinimum > 0) {
      const stokAsalSesudah = stokAsal - qty;
      if (stokAsalSesudah <= stokMinimum) {
        const namaProduk = ambilNamaProdukDariSku(sku);
        peringatanStok = `⚠️ Stok ${namaProduk} (${sku}) di ${dariLokasi} tinggal ${stokAsalSesudah} pcs (minimum: ${stokMinimum} pcs).`;
      }
    }

    return { sukses: true, pesan: `Transfer ${qty} pcs ${sku} dari ${dariLokasi} ke ${keLokasi} berhasil!`, peringatanStok: peringatanStok };
  } catch (e) {
    return { sukses: false, pesan: e.message || 'Sistem sibuk. Silakan coba lagi.' };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 7. ANALISIS: PRODUK TERLARIS (Harian / Bulanan / Tahunan)
// ==========================================
/**
 * Menghitung total qty terjual per SKU dari data Barang_Keluar mulai dari batasWaktu (inklusif).
 * batasWaktu = null berarti TANPA filter tanggal (hitung semua baris yang ada di sheet).
 */
function hitungDariBarangKeluar(ss, batasWaktu, tambahRekapFn) {
  const sheetKeluar = ss.getSheetByName('Barang_Keluar');
  const data = sheetKeluar ? sheetKeluar.getDataRange().getValues() : [];
  for (let i = 1; i < data.length; i++) {
    const tipe = data[i][5];
    if (tipe !== 'KELUAR_ECER' && tipe !== 'KELUAR_GROSIR') continue;

    const waktuBaris = new Date(data[i][0]);
    if (isNaN(waktuBaris.getTime())) continue;
    if (batasWaktu && waktuBaris < batasWaktu) continue;

    const sku = data[i][2] ? data[i][2].toString().trim() : "";
    const qty = parseFloat(data[i][4]) || 0;
    tambahRekapFn(sku, qty);
  }
}

/**
 * periode: 'harian' | 'bulanan' | 'tahunan'.
 * PENTING: karena Barang_Keluar diarsipkan & direset tiap awal bulan (lihat
 * arsipkanBarangKeluarBulanan()), perhitungan 'tahunan' MENGGABUNGKAN data yang sudah diarsipkan
 * di Data_Tahunan (bulan-bulan sebelumnya di tahun berjalan) DENGAN data bulan berjalan yang
 * masih ada di Barang_Keluar -- supaya grafik/laporan tahunan tetap lengkap walau sheet
 * Barang_Keluar sendiri sudah beberapa kali direset sepanjang tahun.
 */
function ambilProdukTerlaris(periode) {
  const ss = getDb();
  const produk = ss.getSheetByName('Produk').getDataRange().getValues();

  let petaNama = {};
  for (let i = 1; i < produk.length; i++) {
    const sku = produk[i][0] ? produk[i][0].toString().trim() : "";
    if (sku) petaNama[sku] = produk[i][1];
  }

  let rekap = {};
  function tambahRekap(sku, qty) {
    if (!sku) return;
    if (!rekap[sku]) rekap[sku] = { sku: sku, nama: petaNama[sku] || sku, qty: 0 };
    rekap[sku].qty += qty;
  }

  const sekarang = new Date();

  if (periode === 'tahunan') {
    const sheetTahunan = ss.getSheetByName('Data_Tahunan');
    if (sheetTahunan) {
      const dataTahunan = sheetTahunan.getDataRange().getValues();
      for (let i = 1; i < dataTahunan.length; i++) {
        if (Number(dataTahunan[i][0]) === sekarang.getFullYear()) {
          const sku = dataTahunan[i][2] ? dataTahunan[i][2].toString().trim() : "";
          tambahRekap(sku, parseFloat(dataTahunan[i][7]) || 0);
        }
      }
    }
    // Ditambah data bulan berjalan yang belum sempat diarsipkan (masih di Barang_Keluar)
    hitungDariBarangKeluar(ss, null, tambahRekap);
  } else {
    let batasWaktu;
    if (periode === 'harian') {
      batasWaktu = new Date(sekarang.getFullYear(), sekarang.getMonth(), sekarang.getDate());
    } else { // bulanan
      batasWaktu = new Date(sekarang.getFullYear(), sekarang.getMonth(), 1);
    }
    hitungDariBarangKeluar(ss, batasWaktu, tambahRekap);
  }

  return Object.keys(rekap).map(k => rekap[k]).sort((a, b) => b.qty - a.qty).slice(0, 8);
}

/**
 * Sama seperti ambilProdukTerlaris('bulanan'), TAPI tanpa filter tanggal sama sekali. Dipakai
 * KHUSUS oleh trigger laporan bulanan, yang jalan di awal bulan SEBELUM arsip+reset -- di titik
 * itu Barang_Keluar cuma berisi data bulan yang baru saja selesai, jadi filter "mulai awal bulan
 * INI" justru salah (karena bulan baru saja mulai, hasilnya jadi kosong).
 */
function ambilProdukTerlarisSemuaBarangKeluar() {
  const ss = getDb();
  const produk = ss.getSheetByName('Produk').getDataRange().getValues();
  let petaNama = {};
  for (let i = 1; i < produk.length; i++) {
    const sku = produk[i][0] ? produk[i][0].toString().trim() : "";
    if (sku) petaNama[sku] = produk[i][1];
  }
  let rekap = {};
  function tambahRekap(sku, qty) {
    if (!sku) return;
    if (!rekap[sku]) rekap[sku] = { sku: sku, nama: petaNama[sku] || sku, qty: 0 };
    rekap[sku].qty += qty;
  }
  hitungDariBarangKeluar(ss, null, tambahRekap);
  return Object.keys(rekap).map(k => rekap[k]).sort((a, b) => b.qty - a.qty).slice(0, 8);
}

/**
 * Batas awal periode (harian/bulanan/tahunan), dihitung relatif terhadap tanggalReferensi
 * (default: sekarang). Parameter referensi ini PENTING dipakai saat laporan dipicu trigger
 * bulanan -- di titik itu "sekarang" sudah masuk bulan baru, padahal yang mau dilaporkan
 * adalah bulan yang BARU SAJA SELESAI, jadi referensinya perlu digeser ke bulan lalu.
 */
function tentukanBatasWaktu(periode, tanggalReferensi) {
  const ref = tanggalReferensi || new Date();
  if (periode === 'harian') return new Date(ref.getFullYear(), ref.getMonth(), ref.getDate());
  if (periode === 'tahunan') return new Date(ref.getFullYear(), 0, 1);
  return new Date(ref.getFullYear(), ref.getMonth(), 1); // bulanan
}

function petaNamaProduk(ss) {
  const produk = ss.getSheetByName('Produk').getDataRange().getValues();
  let peta = {};
  for (let i = 1; i < produk.length; i++) {
    const sku = produk[i][0] ? produk[i][0].toString().trim() : '';
    if (sku) peta[sku] = produk[i][1];
  }
  return peta;
}

/**
 * Peta SKU -> {nama, hargaEcer, hargaGrosir, stokMinimum}, dipakai untuk cek ambang batas
 * peringatan stok (tambahTransaksi) dan sebagai FALLBACK laporan keuangan/arsip tahunan untuk
 * baris Barang_Keluar LAMA yang belum punya Harga_Satuan tersimpan sendiri.
 * CATATAN: hargaEcer/hargaGrosir DIBACA APA ADANYA dari kolom "Harga Ecer"/"Harga Grosir" kalau
 * masih ada di sheet (peninggalan versi lama) -- kolom ini TIDAK PERNAH lagi ditulis oleh
 * tambahProduk()/perbaruiProduk(), karena produk sekarang tidak lagi punya harga katalog; harga
 * jual murni diketik manual tiap transaksi (lihat tambahTransaksi()). Nilainya cuma dipakai
 * sebagai jaring pengaman untuk data historis, bukan sumber harga aktif.
 * Stok_Minimum dicari lewat nama header (cariKolom), bukan indeks tetap, supaya aman walau
 * posisi kolomnya berbeda antar spreadsheet lama & baru.
 */
function ambilDetailProdukSemua(ss) {
  const sheet = ss.getSheetByName('Produk');
  const produk = sheet.getDataRange().getValues();
  const kolomMinimum = cariKolom(sheet, 'Stok_Minimum');
  let peta = {};
  for (let i = 1; i < produk.length; i++) {
    const sku = produk[i][0] ? produk[i][0].toString().trim() : '';
    if (sku) {
      peta[sku] = {
        nama: produk[i][1],
        hargaEcer: parseFloat(produk[i][3]) || 0,
        hargaGrosir: parseFloat(produk[i][4]) || 0,
        stokMinimum: kolomMinimum ? (parseFloat(produk[i][kolomMinimum - 1]) || 0) : 0
      };
    }
  }
  return peta;
}

/**
 * Ringkasan PENJUALAN (ecer/grosir) per SKU per lokasi, untuk periode & referensi tanggal
 * tertentu. Dipakai di laporan PDF ("Data Penjualan Bulanan").
 */
function ambilRingkasanPenjualan(periode, tanggalReferensi) {
  const ss = getDb();
  const petaNama = petaNamaProduk(ss);
  const batasWaktu = tentukanBatasWaktu(periode, tanggalReferensi);

  const sheet = ss.getSheetByName('Barang_Keluar');
  const data = sheet ? sheet.getDataRange().getValues() : [];
  let rekap = {};
  for (let i = 1; i < data.length; i++) {
    const waktu = new Date(data[i][0]);
    if (isNaN(waktu.getTime()) || waktu < batasWaktu) continue;
    const sku = data[i][2] ? data[i][2].toString().trim() : '';
    const idLokasi = data[i][3];
    const qty = parseFloat(data[i][4]) || 0;
    const tipe = data[i][5];
    if (!sku || !idLokasi) continue;

    const kunci = sku + '|' + idLokasi;
    if (!rekap[kunci]) rekap[kunci] = { sku: sku, nama: petaNama[sku] || sku, idLokasi: idLokasi, ecer: 0, grosir: 0 };
    if (tipe === 'KELUAR_ECER') rekap[kunci].ecer += qty;
    else if (tipe === 'KELUAR_GROSIR') rekap[kunci].grosir += qty;
  }
  return Object.keys(rekap).map(function (k) {
    const r = rekap[k];
    r.total = r.ecer + r.grosir;
    return r;
  }).sort((a, b) => b.total - a.total);
}

/**
 * Ringkasan BARANG MASUK (restock) per SKU per lokasi per sumber (Penjahit/Supplier), untuk
 * periode & referensi tanggal tertentu. Dipakai di laporan PDF ("Data Barang Masuk Bulanan").
 */
function ambilRingkasanBarangMasuk(periode, tanggalReferensi) {
  const ss = getDb();
  const petaNama = petaNamaProduk(ss);
  const batasWaktu = tentukanBatasWaktu(periode, tanggalReferensi);

  const sheet = ss.getSheetByName('Barang_Masuk');
  const data = sheet ? sheet.getDataRange().getValues() : [];
  let rekap = {};
  for (let i = 1; i < data.length; i++) {
    const waktu = new Date(data[i][0]);
    if (isNaN(waktu.getTime()) || waktu < batasWaktu) continue;
    const sku = data[i][2] ? data[i][2].toString().trim() : '';
    const idLokasi = data[i][3];
    const qty = parseFloat(data[i][4]) || 0;
    const sumber = data[i][8] || 'Penjahit';
    const namaSupplier = data[i][9] || '';
    if (!sku || !idLokasi) continue;

    const kunci = sku + '|' + idLokasi + '|' + sumber + '|' + namaSupplier;
    if (!rekap[kunci]) rekap[kunci] = { sku: sku, nama: petaNama[sku] || sku, idLokasi: idLokasi, sumber: sumber, namaSupplier: namaSupplier, qty: 0 };
    rekap[kunci].qty += qty;
  }
  return Object.keys(rekap).map(k => rekap[k]).sort((a, b) => b.qty - a.qty);
}

/**
 * Ringkasan TRANSFER antar lokasi, untuk periode & referensi tanggal tertentu. Cukup baca dari
 * Transfer_Keluar saja (satu baris = satu transfer, sudah mewakili dua sisi) supaya tidak dobel
 * dengan Transfer_Masuk. Dipakai di laporan PDF ("Data Transfer Bulanan").
 */
function ambilRingkasanTransfer(periode, tanggalReferensi) {
  const ss = getDb();
  const petaNama = petaNamaProduk(ss);
  const batasWaktu = tentukanBatasWaktu(periode, tanggalReferensi);

  const sheet = ss.getSheetByName('Transfer_Keluar');
  const data = sheet ? sheet.getDataRange().getValues() : [];
  let hasil = [];
  for (let i = 1; i < data.length; i++) {
    const waktu = new Date(data[i][0]);
    if (isNaN(waktu.getTime()) || waktu < batasWaktu) continue;
    const sku = data[i][2] ? data[i][2].toString().trim() : '';
    if (!sku) continue;
    hasil.push({
      sku: sku,
      nama: petaNama[sku] || sku,
      dari: data[i][3],
      ke: data[i][4],
      qty: parseFloat(data[i][5]) || 0
    });
  }
  return hasil.sort((a, b) => b.qty - a.qty);
}

/**
 * Laporan KEUANGAN penjualan (omset) per lokasi, untuk periode & referensi tanggal tertentu.
 * Beda dengan ambilRingkasanPenjualan (yang per SKU): ini fokus ke NILAI RUPIAH per lokasi,
 * dipakai untuk "Laporan Keuangan" yang terpisah dari "Laporan Stok".
 * Return: { perLokasi: [{idLokasi, namaLokasi, qtyEcer, qtyGrosir, omsetEcer, omsetGrosir, omsetTotal}],
 *           grandTotal: {omsetEcer, omsetGrosir, omsetTotal} }
 */
function ambilLaporanKeuangan(periode, tanggalReferensi) {
  const ss = getDb();
  const daftarLokasi = ambilDaftarLokasi();
  const detailProduk = ambilDetailProdukSemua(ss);

  let rekapLokasi = {};
  daftarLokasi.forEach(function (l) {
    rekapLokasi[l.id] = { idLokasi: l.id, namaLokasi: l.nama, qtyEcer: 0, qtyGrosir: 0, omsetEcer: 0, omsetGrosir: 0 };
  });

  function tambahDariBaris(row) {
    const idLokasi = row[3];
    const qty = parseFloat(row[4]) || 0;
    const tipe = row[5];
    const sku = row[2] ? row[2].toString().trim() : '';

    // Harga_Satuan (kolom index 8) -- fallback ke harga produk saat ini untuk baris lama
    // sebelum kolom ini ada.
    let harga = parseFloat(row[8]) || 0;
    if (!harga) {
      const d = detailProduk[sku];
      harga = d ? (tipe === 'KELUAR_GROSIR' ? d.hargaGrosir : d.hargaEcer) : 0;
    }
    const nilai = qty * harga;

    if (!rekapLokasi[idLokasi]) rekapLokasi[idLokasi] = { idLokasi: idLokasi, namaLokasi: idLokasi, qtyEcer: 0, qtyGrosir: 0, omsetEcer: 0, omsetGrosir: 0 };
    if (tipe === 'KELUAR_ECER') { rekapLokasi[idLokasi].qtyEcer += qty; rekapLokasi[idLokasi].omsetEcer += nilai; }
    else if (tipe === 'KELUAR_GROSIR') { rekapLokasi[idLokasi].qtyGrosir += qty; rekapLokasi[idLokasi].omsetGrosir += nilai; }
  }

  if (periode === 'tahunan') {
    // Bulan-bulan yang sudah diarsipkan tahun ini -- pakai Omset_xxx yang sudah tersimpan,
    // TIDAK dihitung ulang dari harga sekarang (supaya akurat historis).
    const sheetTahunan = ss.getSheetByName('Data_Tahunan');
    if (sheetTahunan) {
      const dataTahunan = sheetTahunan.getDataRange().getValues();
      const sekarang = new Date();
      for (let i = 1; i < dataTahunan.length; i++) {
        if (Number(dataTahunan[i][0]) !== sekarang.getFullYear()) continue;
        const idLokasi = dataTahunan[i][4];
        if (!rekapLokasi[idLokasi]) rekapLokasi[idLokasi] = { idLokasi: idLokasi, namaLokasi: idLokasi, qtyEcer: 0, qtyGrosir: 0, omsetEcer: 0, omsetGrosir: 0 };
        rekapLokasi[idLokasi].qtyEcer += parseFloat(dataTahunan[i][5]) || 0;
        rekapLokasi[idLokasi].qtyGrosir += parseFloat(dataTahunan[i][6]) || 0;
        rekapLokasi[idLokasi].omsetEcer += parseFloat(dataTahunan[i][9]) || 0;
        rekapLokasi[idLokasi].omsetGrosir += parseFloat(dataTahunan[i][10]) || 0;
      }
    }
    // Ditambah bulan berjalan yang belum sempat diarsipkan (masih di Barang_Keluar)
    const sheetKeluar = ss.getSheetByName('Barang_Keluar');
    const data = sheetKeluar ? sheetKeluar.getDataRange().getValues() : [];
    for (let i = 1; i < data.length; i++) {
      const tipe = data[i][5];
      if (tipe !== 'KELUAR_ECER' && tipe !== 'KELUAR_GROSIR') continue;
      tambahDariBaris(data[i]);
    }
  } else {
    const batasWaktu = tentukanBatasWaktu(periode, tanggalReferensi);
    const sheetKeluar = ss.getSheetByName('Barang_Keluar');
    const data = sheetKeluar ? sheetKeluar.getDataRange().getValues() : [];
    for (let i = 1; i < data.length; i++) {
      const tipe = data[i][5];
      if (tipe !== 'KELUAR_ECER' && tipe !== 'KELUAR_GROSIR') continue;
      const waktu = new Date(data[i][0]);
      if (isNaN(waktu.getTime()) || waktu < batasWaktu) continue;
      tambahDariBaris(data[i]);
    }
  }

  const perLokasi = Object.keys(rekapLokasi).map(function (k) {
    const r = rekapLokasi[k];
    r.omsetTotal = r.omsetEcer + r.omsetGrosir;
    return r;
  });

  const grandTotal = {
    omsetEcer: perLokasi.reduce((s, r) => s + r.omsetEcer, 0),
    omsetGrosir: perLokasi.reduce((s, r) => s + r.omsetGrosir, 0),
    omsetTotal: perLokasi.reduce((s, r) => s + r.omsetTotal, 0)
  };

  return { perLokasi: perLokasi, grandTotal: grandTotal };
}

// ==========================================
// 7b. ARSIP BULANAN: Barang_Keluar -> Data_Tahunan, lalu Barang_Keluar direset
// ==========================================
/**
 * Meringkas seluruh isi Barang_Keluar (penjualan ecer/grosir) per SKU per lokasi per bulan,
 * menambahkannya ke Data_Tahunan, lalu MENGOSONGKAN Barang_Keluar (sisa header) supaya siap
 * diisi bulan berikutnya. Aman dijalankan berkali-kali -- kalau Barang_Keluar sedang kosong,
 * fungsi ini tidak melakukan apa-apa selain melapor demikian.
 */
function arsipkanBarangKeluarBulanan() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const ss = getDb();
    const sheetKeluar = ss.getSheetByName('Barang_Keluar');
    const data = sheetKeluar.getDataRange().getValues();
    if (data.length <= 1) {
      return { sukses: true, pesan: 'Tidak ada data Barang Keluar untuk diarsipkan.' };
    }

    const produk = ss.getSheetByName('Produk').getDataRange().getValues();
    let petaNama = {};
    for (let i = 1; i < produk.length; i++) {
      const sku = produk[i][0] ? produk[i][0].toString().trim() : '';
      if (sku) petaNama[sku] = produk[i][1];
    }

    const detailProduk = ambilDetailProdukSemua(ss);

    // Ringkas per Tahun-Bulan-SKU-Lokasi
    let rekap = {};
    for (let i = 1; i < data.length; i++) {
      const waktu = new Date(data[i][0]);
      if (isNaN(waktu.getTime())) continue;
      const sku = data[i][2] ? data[i][2].toString().trim() : '';
      const idLokasi = data[i][3];
      const qty = parseFloat(data[i][4]) || 0;
      const tipe = data[i][5];
      if (!sku || !idLokasi) continue;

      // Harga_Satuan (kolom index 8) -- kalau kosong (baris lama sebelum kolom ini ada),
      // fallback ke harga produk saat ini supaya arsip tetap terisi meski kurang akurat historis.
      let harga = parseFloat(data[i][8]) || 0;
      if (!harga) {
        const d = detailProduk[sku];
        harga = d ? (tipe === 'KELUAR_GROSIR' ? d.hargaGrosir : d.hargaEcer) : 0;
      }
      const nilai = qty * harga;

      const tahun = waktu.getFullYear();
      const bulan = waktu.getMonth() + 1;
      const kunci = tahun + '|' + bulan + '|' + sku + '|' + idLokasi;

      if (!rekap[kunci]) {
        rekap[kunci] = { tahun: tahun, bulan: bulan, sku: sku, idLokasi: idLokasi, ecer: 0, grosir: 0, omsetEcer: 0, omsetGrosir: 0 };
      }
      if (tipe === 'KELUAR_ECER') { rekap[kunci].ecer += qty; rekap[kunci].omsetEcer += nilai; }
      else if (tipe === 'KELUAR_GROSIR') { rekap[kunci].grosir += qty; rekap[kunci].omsetGrosir += nilai; }
    }

    const sheetTahunan = ss.getSheetByName('Data_Tahunan');
    const waktuArsip = new Date();
    const barisBaru = Object.keys(rekap).map(function (k) {
      const r = rekap[k];
      const total = r.ecer + r.grosir;
      const totalOmset = r.omsetEcer + r.omsetGrosir;
      return [r.tahun, r.bulan, r.sku, petaNama[r.sku] || r.sku, r.idLokasi, r.ecer, r.grosir, total, waktuArsip, r.omsetEcer, r.omsetGrosir, totalOmset];
    });

    if (barisBaru.length > 0) {
      sheetTahunan.getRange(sheetTahunan.getLastRow() + 1, 1, barisBaru.length, 12).setValues(barisBaru);
    }

    // Reset Barang_Keluar: hapus semua baris DATA, sisakan header
    const lastRow = sheetKeluar.getLastRow();
    if (lastRow > 1) {
      sheetKeluar.getRange(2, 1, lastRow - 1, sheetKeluar.getLastColumn()).clearContent();
    }

    SpreadsheetApp.flush();
    return { sukses: true, pesan: `Arsip berhasil: ${barisBaru.length} baris ringkasan dipindah ke Data_Tahunan, Barang_Keluar direset.` };
  } catch (e) {
    return { sukses: false, pesan: e.message || 'Gagal mengarsipkan data.' };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 8. DASHBOARD GRAFIK DI SPREADSHEET
// ==========================================
function perbaruiDashboardSheet() {
  const ss = getDb();
  let sheet = ss.getSheetByName('Dashboard');
  if (!sheet) {
    sheet = ss.insertSheet('Dashboard');
  } else {
    sheet.getCharts().forEach(c => sheet.removeChart(c));
    sheet.clear();
  }

  sheet.getRange(1, 1).setValue('Dashboard Produk Terlaris - Faminis Barokah').setFontWeight('bold').setFontSize(14);
  sheet.getRange(2, 1).setValue('Diperbarui otomatis: ' + new Date().toLocaleString('id-ID'));

  const daftarPeriode = [['harian', 'Harian'], ['bulanan', 'Bulanan'], ['tahunan', 'Tahunan']];
  let startRow = 4;

  daftarPeriode.forEach(function (p) {
    const periode = p[0], label = p[1];
    const data = ambilProdukTerlaris(periode);

    sheet.getRange(startRow, 1).setValue(`Produk Terlaris (${label})`).setFontWeight('bold');
    sheet.getRange(startRow + 1, 1, 1, 2).setValues([['Nama Produk', 'Qty Terjual']]).setFontWeight('bold').setBackground('#EEDDC8');

    const rows = data.map(d => [d.nama, d.qty]);
    if (rows.length > 0) {
      sheet.getRange(startRow + 2, 1, rows.length, 2).setValues(rows);

      const chart = sheet.newChart()
        .setChartType(Charts.ChartType.BAR)
        .addRange(sheet.getRange(startRow + 2, 1, rows.length, 2))
        .setPosition(startRow, 4, 0, 0)
        .setOption('title', `Produk Terlaris - ${label}`)
        .setOption('width', 500)
        .setOption('height', 260)
        .build();
      sheet.insertChart(chart);
    }
    startRow += Math.max(rows.length, 1) + 16;
  });

  sheet.autoResizeColumns(1, 2);
}

// ==========================================
// 9. LAPORAN PDF OTOMATIS (dengan grafik)
// ==========================================
function buatGrafikPNG(dataProduk, judul) {
  if (!dataProduk || dataProduk.length === 0) return null;
  const dataTable = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, 'Produk')
    .addColumn(Charts.ColumnType.NUMBER, 'Qty');
  dataProduk.forEach(d => dataTable.addRow([d.nama, d.qty]));

  const chart = Charts.newBarChart()
    .setDataTable(dataTable)
    .setTitle(judul)
    .setDimensions(520, 300)
    .setColors(['#8B5A2B'])
    .build();

  return chart.getAs('image/png');
}

/**
 * Kirim laporan PDF ke email owner. periode: 'harian' | 'bulanan' | 'tahunan' (default 'bulanan')
 * produkTerlarisOverride (opsional): kalau diisi, dipakai langsung tanpa hitung ulang -- dipakai
 * oleh triggerLaporanBulanan() supaya laporan memakai data bulan yang BARU SELESAI, bukan bulan
 * yang baru saja mulai (lihat ambilProdukTerlarisSemuaBarangKeluar()).
 * tanggalReferensi (opsional): tanggal acuan untuk menghitung batas awal periode. Default hari
 * ini -- tapi WAJIB digeser ke bulan lalu saat dipanggil dari trigger bulanan (lihat
 * triggerLaporanBulanan()), supaya ringkasan penjualan/masuk/transfer tidak kosong.
 */
function kirimLaporanPDF(periode, produkTerlarisOverride, tanggalReferensi) {
  periode = periode || 'bulanan';
  const ref = tanggalReferensi ? new Date(tanggalReferensi) : new Date();

  const dataStok = ambilSemuaStok();
  const daftarLokasi = ambilDaftarLokasi();
  const produkTerlaris = produkTerlarisOverride || ambilProdukTerlaris(periode);
  const ringkasanPenjualan = ambilRingkasanPenjualan(periode, ref);
  const ringkasanBarangMasuk = ambilRingkasanBarangMasuk(periode, ref);
  const ringkasanTransfer = ambilRingkasanTransfer(periode, ref);
  const laporanKeuangan = ambilLaporanKeuangan(periode, ref);
  const grafikBlob = buatGrafikPNG(produkTerlaris, `Produk Terlaris (${periode})`);
  const grafikBase64 = grafikBlob ? Utilities.base64Encode(grafikBlob.getBytes()) : null;
  const grafikOmsetData = laporanKeuangan.perLokasi.map(r => ({ nama: r.namaLokasi, qty: Math.round(r.omsetTotal) }));
  const grafikOmsetBlob = buatGrafikPNG(grafikOmsetData, `Omset Penjualan per Lokasi (${periode})`);
  const grafikOmsetBase64 = grafikOmsetBlob ? Utilities.base64Encode(grafikOmsetBlob.getBytes()) : null;

  const template = HtmlService.createTemplateFromFile('TemplatePdf');
  template.data = dataStok;
  template.daftarLokasi = daftarLokasi;
  template.produkTerlaris = produkTerlaris;
  template.ringkasanPenjualan = ringkasanPenjualan;
  template.ringkasanBarangMasuk = ringkasanBarangMasuk;
  template.ringkasanTransfer = ringkasanTransfer;
  template.laporanKeuangan = laporanKeuangan;
  template.grafikBase64 = grafikBase64;
  template.grafikOmsetBase64 = grafikOmsetBase64;
  template.periode = periode;

  const htmlContent = template.evaluate().getContent();
  const blob = Utilities.newBlob(htmlContent, 'text/html', 'Laporan.html');
  const pdf = blob.getAs('application/pdf');
  pdf.setName(`Laporan_Stok_Faminis_${periode}_${new Date().toLocaleDateString('id-ID')}.pdf`);

  MailApp.sendEmail({
    to: EMAIL_OWNER,
    subject: `Laporan Stok Faminis Barokah (${periode.charAt(0).toUpperCase() + periode.slice(1)})`,
    body: "Halo Owner, terlampir adalah laporan stok Faminis Barokah.",
    attachments: [pdf]
  });

  return { sukses: true, pesan: "Laporan berhasil dikirim ke email owner!" };
}

function triggerLaporanBulanan() {
  // Urutan WAJIB seperti ini: kirim laporan bulan yang baru saja selesai DULU (pakai data asli
  // sebelum direset), baru SETELAH itu arsipkan & reset Barang_Keluar untuk bulan berikutnya.
  // tanggalReferensi digeser ke bulan lalu supaya ringkasan penjualan/masuk/transfer memakai
  // batas tanggal bulan yang benar (bukan bulan baru yang baru saja mulai).
  const sekarang = new Date();
  const referensiBulanLalu = new Date(sekarang.getFullYear(), sekarang.getMonth() - 1, 1);

  const terlarisBulanLalu = ambilProdukTerlarisSemuaBarangKeluar();
  kirimLaporanPDF('bulanan', terlarisBulanLalu, referensiBulanLalu);
  arsipkanBarangKeluarBulanan();
}
function triggerLaporanTahunan() {
  kirimLaporanPDF('tahunan');
}

/**
 * Mendaftarkan trigger otomatis: laporan bulanan (sekaligus arsip+reset Barang_Keluar), laporan
 * tahunan, dan pembaruan Dashboard grafik tiap jam. JALANKAN FUNGSI INI SEKALI SAJA di editor.
 */
function aturTriggerOtomatis() {
  const semuaTrigger = ScriptApp.getProjectTriggers();
  const namaFungsi = ['triggerLaporanBulanan', 'triggerLaporanTahunan', 'perbaruiDashboardSheet'];
  semuaTrigger.forEach(t => {
    if (namaFungsi.indexOf(t.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Laporan Bulanan + Arsip: tanggal 1 jam 06:00 pagi
  ScriptApp.newTrigger('triggerLaporanBulanan').timeBased().onMonthDay(1).atHour(6).create();

  // Laporan Tahunan: 31 Desember jam 23:00 malam
  ScriptApp.newTrigger('triggerLaporanTahunan').timeBased().onMonthDay(31).atHour(23).create();

  // Dashboard grafik di spreadsheet: refresh tiap jam
  ScriptApp.newTrigger('perbaruiDashboardSheet').timeBased().everyHours(1).create();

  Logger.log('Trigger otomatis berhasil diatur!');
}

// ==========================================
// 10. RESET DATA (JALANKAN MANUAL DARI EDITOR -- SENGAJA TIDAK ADA DI APP/API)
// ==========================================
// ⚠️ PERINGATAN: fungsi ini MENGHAPUS PERMANEN seluruh riwayat transaksi, SELURUH DAFTAR PRODUK,
// dan SELURUH BARIS di semua sheet Stok_xxx. TIDAK BISA DIBATALKAN setelah dijalankan. SEBELUM
// menjalankan ini:
//   1. BACKUP dulu -- di spreadsheet ini: File > Buat salinan, atau File > Download > Microsoft Excel.
//   2. Pastikan kamu benar-benar mau hapus SEMUA produk yang sudah didaftarkan juga (bukan cuma
//      riwayat transaksinya) -- kalau sudah dijalankan, produk harus didaftarkan ULANG dari nol
//      lewat halaman "Kelola Produk" (atau langsung di sheet Produk) sebelum bisa transaksi lagi.
//
// Yang DIHAPUS (baris datanya saja, header baris 1 tetap ada): Barang_Masuk, Barang_Keluar,
// Transfer_Keluar, Transfer_Masuk, Data_Tahunan (arsip tahunan), Supplier (daftar nama supplier),
// Produk (SELURUH daftar produk beserta Stok_Minimum-nya), dan SEMUA sheet Stok_xxx (Stok_Gudang,
// Stok_RUKO1, dst -- baris SKU-nya ikut terhapus total, bukan cuma angkanya di-nol-kan, karena
// daftar produknya sendiri sudah tidak ada lagi).
// Yang TIDAK DISENTUH SAMA SEKALI: Lokasi (Toko/Gudang & Ruko 1-4 tetap terdaftar) dan Pengguna
// (semua akun & PIN tetap aktif seperti sebelumnya).
//
// CARA PAKAI: buka fungsi ini di Apps Script Editor, pilih "resetTransaksiStokDanProduk" dari
// dropdown fungsi di toolbar, lalu klik ▶️ Run. Sengaja TIDAK didaftarkan di DAFTAR_FUNGSI_API
// supaya TIDAK BISA dipicu lewat aplikasi/HP oleh siapapun (kasir, admin, bahkan master) -- cuma
// bisa dijalankan manual langsung dari editor ini, oleh yang pegang akses Spreadsheet/Apps Script.
function resetTransaksiStokDanProduk() {
  const ss = getDb();
  let ringkasan = [];

  // 1) Kosongkan riwayat transaksi + SELURUH daftar produk (baris data saja, header tetap ada)
  const sheetDihapusBarisnya = ['Barang_Masuk', 'Barang_Keluar', 'Transfer_Keluar', 'Transfer_Masuk', 'Data_Tahunan', 'Supplier', 'Produk'];
  sheetDihapusBarisnya.forEach(function (nama) {
    const sheet = ss.getSheetByName(nama);
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
      ringkasan.push(`${nama}: ${lastRow - 1} baris dihapus`);
    } else {
      ringkasan.push(`${nama}: sudah kosong, tidak ada yang dihapus`);
    }
  });

  // 2) Kosongkan TOTAL semua sheet Stok_xxx (baris SKU ikut terhapus, bukan cuma angkanya
  // di-nol-kan) -- karena daftar produknya sendiri baru saja dihapus habis di langkah 1.
  const daftarLokasi = ambilDaftarLokasi();
  daftarLokasi.forEach(function (lok) {
    const namaSheetStok = dapatkanNamaSheetStok(lok.id);
    const sheet = ss.getSheetByName(namaSheetStok);
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
      ringkasan.push(`${namaSheetStok}: ${lastRow - 1} baris dihapus`);
    } else {
      ringkasan.push(`${namaSheetStok}: sudah kosong, tidak ada yang dihapus`);
    }
  });

  // 3) Refresh Dashboard grafik (kalau sheet-nya ada) supaya langsung ikut kosong, tidak perlu
  // menunggu trigger per-jam.
  try { perbaruiDashboardSheet(); } catch (e) { /* Dashboard opsional, abaikan kalau error/belum ada */ }

  const pesan = 'RESET SELESAI:\n' + ringkasan.join('\n') +
    '\n\nCatatan: daftar produk sekarang KOSONG TOTAL -- daftarkan produk baru langsung di sheet Produk di Google Spreadsheet.';
  Logger.log(pesan);
  return { sukses: true, ringkasan: ringkasan };
}