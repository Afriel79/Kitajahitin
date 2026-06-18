/**
 * ============================================================
 * KITAJAHITIN ERP v3.1 - Backend (Code.gs) — FINAL MERGED
 * PT Bersama Ravisa Sejahtera
 * Sistem ERP Manufaktur Garmen + Integrasi Website Katalog
 * ============================================================
 *
 * Sheet ERP (existing):
 *   UserRoles, Buku Besar Kas, Detail Transaksi, BebanOperasional,
 *   Stok_Barang, LaporanPembelian, OutputProduksi, OutputProduksiDetail,
 *   LaporanFinishing, JadwalProduksi, SampleTracking, CuttingTracking,
 *   PackingList, SaldoAwal
 *
 * Sheet BARU (database katalog, ada di dalam ERP):
 *   Data Pesanan (Katalog Kitajahitin)  — 13 kolom
 *   Data Pelanggan (Katalog Kitajahitin) — 7 kolom
 *
 * Alur:
 *   Website Katalog → doPost → fungsi katalog
 *                             → auto-bridge ke Buku Besar Kas & Detail Transaksi ERP
 *   Website ERP     → doGet  → dashboard ERP
 *                   → google.script.run → fungsi ERP
 * ============================================================
 */

// ============================================================
// KONFIGURASI GLOBAL
// ============================================================

// ====== CARI SEKTOR INI DAN GANTI MENJADI: ======
var SS_ID = "1B7tWeaAGBG_H8lw5pTzLBgbEL4fxyEfFfXDilOrHAXY";

function getSheet_(name) {
  return SpreadsheetApp.openById(SS_ID).getSheetByName(name);
}

// ============================================================
// KONFIGURASI KATALOG & ONGKIR
// ============================================================

var RAJAONGKIR_API_KEY  = "MASUKKAN_API_KEY_ANDA_DISINI";
var RAJAONGKIR_BASE_URL = "https://api.rajaongkir.com/starter";
var ORIGIN_CITY_ID      = "114"; // Depok
var BERAT_PER_PCS_GRAM  = 300;

var ADMIN_WA = "6281770714551";

var ORIGIN_ADDRESS   = "Jl. H. M. Tohir Blok Hm Tohir No.50, RT.2/RW.2, Pondok Cina, Kecamatan Beji, Kota Depok, Jawa Barat 16424";
var ORIGIN_LAT       = -6.3728;
var ORIGIN_LNG       = 106.8286;

var LALAMOVE_BASE_FARE   = 8000;
var LALAMOVE_BASE_KM     = 3;
var LALAMOVE_PER_KM_FARE = 3500;

var GOOGLE_MAPS_API_KEY = "MASUKKAN_GOOGLE_MAPS_API_KEY_ANDA";

// ============================================================
// 1. doGet — Entry Point Web App (ERP Dashboard)
// ============================================================

/**
 * Handler utama HTTP GET — menampilkan dashboard ERP
 * @param {Object} e - Event parameter
 * @return {HtmlOutput}
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Kitajahitin | Enterprise System v3.1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================================
// 2. doPost — UNIFIED ROUTER (ERP internal + Katalog eksternal)
// ============================================================

/**
 * Handler utama HTTP POST — satu pintu masuk untuk semua request
 * Dari website katalog (eksternal) maupun dari ERP internal.
 *
 * Format payload:
 * { "action": "namaFungsi", "arguments": [...args] }
 *
 * @param {Object} e - Event parameter
 * @return {TextOutput} JSON response
 */
function doPost(e) {
  var result = { success: false, message: "Unknown error" };
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, message: "Tidak ada data POST yang diterima." }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var payload = JSON.parse(e.postData.contents);
    var action  = payload.action;
    var args    = payload.arguments || [];

    // ── Daftar action yang diizinkan dipanggil via HTTP POST ──
    // Mencakup: fungsi katalog (dari website eksternal) +
    //           fungsi ERP yang bisa dipanggil via API
    var allowedActions = {
      // ---- KATALOG: Autentikasi Pelanggan ----
      "loginCustomer":               loginCustomer,
      "registerCustomer":            registerCustomer,

      // ---- KATALOG: Pesanan ----
      "saveOrderToSheet":            saveOrderToSheet,
      "cancelOrder":                 cancelOrder,
      "checkStatusOrder":            checkStatusOrder,
      "getCustomerOrders":           getCustomerOrders,

      // ---- KATALOG: Sample ----
      "saveRequestSample":           saveRequestSample,
      "saveRequestSampleWithImage":  saveRequestSampleWithImage,

      // ---- KATALOG: Draft ----
      "saveDraft":                   saveDraft,

      // ---- KATALOG: Dokumen ----
      "generateQuotationPdf":        generateQuotationPdf,
      "regenerateInvoicePdf":        regenerateInvoicePdf,

      // ---- KATALOG: Ongkir ----
      "getProvinces":                getProvinces,
      "getCities":                   getCities,
      "checkOngkir":                 checkOngkir,
      "getDeliveryEstimate":         getDeliveryEstimate,

      // ---- ERP: Data Katalog (untuk halaman monitor di ERP) ----
      "getKatalogPesanan":           getKatalogPesanan,
      "getKatalogPelanggan":         getKatalogPelanggan,
      "updateStatusKatalogPesanan":  updateStatusKatalogPesanan
    };

    if (allowedActions[action]) {
      var data = allowedActions[action].apply(null, args);
      result = { success: true, data: data };
    } else {
      result = { success: false, message: "Aksi '" + action + "' tidak terdaftar." };
    }

  } catch (err) {
    result = { success: false, message: "Server Error: " + err.toString() };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// 3. AUTENTIKASI ERP
// ============================================================

function validateLogin(username, password) {
  try {
    var sheet = getSheet_('UserRoles');
    if (!sheet) return { success: false, message: 'Sheet UserRoles tidak ditemukan' };

    var data    = sheet.getDataRange().getValues();
    var headers = data[0];
    var colU    = headers.indexOf('Username');
    var colP    = headers.indexOf('Password');
    var colR    = headers.indexOf('Role');
    var colN    = headers.indexOf('Nama');

    if (colU === -1 || colP === -1 || colR === -1 || colN === -1) {
      return { success: false, message: 'Struktur kolom UserRoles tidak valid' };
    }

    for (var i = 1; i < data.length; i++) {
      if (data[i][colU].toString().trim().toLowerCase() === username.toString().trim().toLowerCase() &&
          data[i][colP].toString().trim() === password.toString().trim()) {
        return {
          success:  true,
          role:     data[i][colR].toString().trim(),
          nama:     data[i][colN].toString().trim(),
          username: data[i][colU].toString().trim()
        };
      }
    }
    return { success: false, message: 'Username atau password salah' };
  } catch (err) {
    return { success: false, message: 'Kesalahan sistem: ' + err.toString() };
  }
}

function changePassword(username, oldPassword, newPassword) {
  try {
    var sheet = getSheet_('UserRoles');
    if (!sheet) return { success: false, message: 'Sheet UserRoles tidak ditemukan' };

    var data    = sheet.getDataRange().getValues();
    var headers = data[0];
    var colU    = headers.indexOf('Username');
    var colP    = headers.indexOf('Password');

    for (var i = 1; i < data.length; i++) {
      if (data[i][colU].toString().trim().toLowerCase() === username.toString().trim().toLowerCase()) {
        if (data[i][colP].toString().trim() !== oldPassword.toString().trim()) {
          return { success: false, message: 'Password lama tidak cocok' };
        }
        if (!newPassword || newPassword.toString().trim().length < 4) {
          return { success: false, message: 'Password baru minimal 4 karakter' };
        }
        sheet.getRange(i + 1, colP + 1).setValue(newPassword.toString().trim());
        return { success: true, message: 'Password berhasil diubah' };
      }
    }
    return { success: false, message: 'Username tidak ditemukan' };
  } catch (err) {
    return { success: false, message: 'Kesalahan: ' + err.toString() };
  }
}

// ============================================================
// 4. DATA RETRIEVAL ERP
// ============================================================

function getInitialData() {
  try {
    var result = { transactions: [], bebanOps: [], stokBarang: [], pembelian: [], saldoAwal: [] };

    // Buku Besar Kas (200 baris terakhir, dibalik)
    var sheetBBK = getSheet_('Buku Besar Kas');
    if (sheetBBK && sheetBBK.getLastRow() > 1) {
      var allData  = sheetBBK.getDataRange().getValues();
      var headers  = allData[0];
      var startRow = Math.max(1, allData.length - 200);
      var sliced   = allData.slice(startRow);
      sliced.reverse();
      for (var i = 0; i < sliced.length; i++) {
        var obj = {};
        for (var j = 0; j < headers.length; j++) {
          var val = sliced[i][j];
          obj[headers[j]] = val instanceof Date
            ? Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd') : val;
        }
        result.transactions.push(obj);
      }
    }

    // BebanOperasional
    var sheetBeban = getSheet_('BebanOperasional');
    if (sheetBeban && sheetBeban.getLastRow() > 1) {
      var allBeban = sheetBeban.getDataRange().getValues();
      var headersB = allBeban[0];
      for (var i = 1; i < allBeban.length; i++) {
        var obj = {};
        for (var j = 0; j < headersB.length; j++) {
          var val = allBeban[i][j];
          obj[headersB[j]] = val instanceof Date
            ? Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd') : val;
        }
        result.bebanOps.push(obj);
      }
    }

    // Stok_Barang
    var sheetStok = getSheet_('Stok_Barang');
    if (sheetStok && sheetStok.getLastRow() > 1) {
      var allStok = sheetStok.getDataRange().getValues();
      var headersS = allStok[0];
      for (var i = 1; i < allStok.length; i++) {
        var obj = {};
        for (var j = 0; j < headersS.length; j++) {
          var val = allStok[i][j];
          obj[headersS[j]] = val instanceof Date
            ? Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd') : val;
        }
        result.stokBarang.push(obj);
      }
    }

    // LaporanPembelian
    var sheetPembelian = getSheet_('LaporanPembelian');
    if (sheetPembelian && sheetPembelian.getLastRow() > 1) {
      var allPembelian = sheetPembelian.getDataRange().getValues();
      var headersP = allPembelian[0];
      for (var i = 1; i < allPembelian.length; i++) {
        var obj = {};
        for (var j = 0; j < headersP.length; j++) {
          var val = allPembelian[i][j];
          obj[headersP[j]] = val instanceof Date
            ? Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd') : val;
        }
        result.pembelian.push(obj);
      }
    }

    // SaldoAwal
    var sheetSaldo = getSheet_('SaldoAwal');
    if (sheetSaldo && sheetSaldo.getLastRow() > 1) {
      var allSaldo = sheetSaldo.getDataRange().getValues();
      var headersSa = allSaldo[0];
      for (var i = 1; i < allSaldo.length; i++) {
        var obj = {};
        for (var j = 0; j < headersSa.length; j++) {
          var val = allSaldo[i][j];
          obj[headersSa[j]] = val instanceof Date
            ? Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd') : val;
        }
        result.saldoAwal.push(obj);
      }
    }

    return result;
  } catch (err) {
    Logger.log('Error getInitialData: ' + err.toString());
    return { transactions: [], bebanOps: [], stokBarang: [], pembelian: [], saldoAwal: [], error: err.toString() };
  }
}

function getFilteredTransactions(keyword) {
  try {
    var sheet = getSheet_('Buku Besar Kas');
    if (!sheet || sheet.getLastRow() <= 1) return [];

    var data       = sheet.getDataRange().getValues();
    var headers    = data[0];
    var results    = [];
    var colInvoice = headers.indexOf('InvoiceID');
    var colKlien   = headers.indexOf('NamaKlien');
    var colProduk  = headers.indexOf('Produk');
    var search     = keyword.toString().toLowerCase().trim();

    for (var i = 1; i < data.length; i++) {
      var inv   = data[i][colInvoice] ? data[i][colInvoice].toString().toLowerCase() : '';
      var klien = data[i][colKlien]   ? data[i][colKlien].toString().toLowerCase()   : '';
      var prod  = data[i][colProduk]  ? data[i][colProduk].toString().toLowerCase()  : '';
      if (inv.indexOf(search) !== -1 || klien.indexOf(search) !== -1 || prod.indexOf(search) !== -1) {
        var obj = {};
        for (var j = 0; j < headers.length; j++) {
          var val = data[i][j];
          obj[headers[j]] = val instanceof Date
            ? Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd') : val;
        }
        results.push(obj);
      }
    }
    results.reverse();
    return results;
  } catch (err) {
    Logger.log('Error getFilteredTransactions: ' + err.toString());
    return [];
  }
}

function getDetailTransaksi(invoiceId) {
  try {
    var sheet = getSheet_('Detail Transaksi');
    if (!sheet || sheet.getLastRow() <= 1) return [];

    var data       = sheet.getDataRange().getValues();
    var headers    = data[0];
    var colInvoice = headers.indexOf('InvoiceID');
    var results    = [];

    for (var i = 1; i < data.length; i++) {
      if (data[i][colInvoice].toString().trim() === invoiceId.toString().trim()) {
        var obj = {};
        for (var j = 0; j < headers.length; j++) {
          var val = data[i][j];
          obj[headers[j]] = val instanceof Date
            ? Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd') : val;
        }
        results.push(obj);
      }
    }
    return results;
  } catch (err) {
    Logger.log('Error getDetailTransaksi: ' + err.toString());
    return [];
  }
}

function getPeringatanStok() {
  try {
    var sheet = getSheet_('Stok_Barang');
    if (!sheet || sheet.getLastRow() <= 1) return [];

    var data        = sheet.getDataRange().getValues();
    var headers     = data[0];
    var colStokAkhir = headers.indexOf('StokAkhir');
    var results     = [];

    for (var i = 1; i < data.length; i++) {
      if ((parseFloat(data[i][colStokAkhir]) || 0) <= 5) {
        var obj = {};
        for (var j = 0; j < headers.length; j++) {
          var val = data[i][j];
          obj[headers[j]] = val instanceof Date
            ? Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd') : val;
        }
        results.push(obj);
      }
    }
    return results;
  } catch (err) {
    Logger.log('Error getPeringatanStok: ' + err.toString());
    return [];
  }
}

// ============================================================
// 5. TRANSACTION CRUD ERP
// ============================================================

function generateInvoiceId() {
  try {
    var today   = new Date();
    var dateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyyMMdd');
    var prefix  = 'INV-' + dateStr + '-';
    var sheet   = getSheet_('Buku Besar Kas');
    var lastNum = 0;

    if (sheet && sheet.getLastRow() > 1) {
      var data       = sheet.getDataRange().getValues();
      var colInvoice = data[0].indexOf('InvoiceID');
      if (colInvoice !== -1) {
        for (var i = 1; i < data.length; i++) {
          var id  = data[i][colInvoice].toString();
          var num = parseInt(id.replace(prefix, ''), 10);
          if (id.indexOf(prefix) === 0 && !isNaN(num) && num > lastNum) lastNum = num;
        }
      }
    }
    return prefix + ('0000' + (lastNum + 1)).slice(-4);
  } catch (err) {
    return 'INV-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd') + '-0001';
  }
}

function saveTransaction(data) {
  try {
    var invoiceId = generateInvoiceId();
    var sheet     = getSheet_('Buku Besar Kas');
    if (!sheet) return { success: false, invoiceId: '', message: 'Sheet Buku Besar Kas tidak ditemukan' };

    var headers       = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return h.toString().trim(); });
    var headerUpdated = false;

    if (headers.indexOf('ModalPO') === -1) {
      sheet.getRange(1, headers.length + 1).setValue('ModalPO');
      headers.push('ModalPO');
      headerUpdated = true;
    }
    if (headers.indexOf('Keuntungan') === -1) {
      sheet.getRange(1, headers.length + 1).setValue('Keuntungan');
      headers.push('Keuntungan');
      headerUpdated = true;
    }
    if (headerUpdated) SpreadsheetApp.flush();

    data = sanitizeData(data);

    var mapping = {
      'InvoiceID':         invoiceId,
      'Tanggal':           data.tanggal || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      'NamaKlien':         data.namaKlien        || '',
      'Produk':            data.produk           || '',
      'Qty':               parseFloat(data.qty)  || 0,
      'HargaSatuan':       parseFloat(data.hargaSatuan) || 0,
      'Diskon':            parseFloat(data.diskon) || 0,
      'Ongkir':            parseFloat(data.ongkir) || 0,
      'GrandTotal':        parseFloat(data.grandTotal) || 0,
      'DP':                parseFloat(data.dp) || 0,
      'SisaTagihan':       parseFloat(data.sisaTagihan) || 0,
      'StatusPembayaran':  data.statusPembayaran || 'Belum Bayar',
      'Proses':            data.proses           || '',
      'SumberKlien':       data.sumberKlien      || '',
      'Catatan':           data.catatan          || '',
      'Periode':           data.periode          || '',
      'WhatsApp':          data.whatsapp         || '',
      'Instansi':          data.instansi         || '',
      'KategoriProduk':    data.kategoriProduk   || '',
      'KeteranganTambahan':data.keteranganTambahan || '',
      'Deadline':          data.deadline         || '',
      'ModalPO':           parseFloat(data.modalPO) || 0,
      'Keuntungan':        parseFloat(data.keuntungan) || 0,
      'TahapProduksi':     data.tahapProduksi    || 'Antrean',
      'PeriodeKas':        data.periodeKas       || '',
      'InputDate':         Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
    };

    var newRow = headers.map(function(h) { return mapping.hasOwnProperty(h) ? mapping[h] : ''; });
    sheet.appendRow(newRow);

    // Simpan detail items
    if (data.items && data.items.length > 0) {
      var detailSheet = getSheet_('Detail Transaksi');
      if (detailSheet) {
        var detailHeaders = detailSheet.getRange(1, 1, 1, detailSheet.getLastColumn()).getValues()[0];
        for (var k = 0; k < data.items.length; k++) {
          var item = data.items[k];
          var detailMap = {
            'InvoiceID':     invoiceId,
            'Tanggal':       data.tanggal || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
            'NamaKlien':     data.namaKlien || '',
            'NamaProduk':    item.namaProduk || item.produk || '',
            'Komponen':      item.komponen  || '',
            'Ukuran':        item.ukuran || item.size || '',
            'KategoriProduk':item.kategoriProduk || item.komponen || '',
            'Warna':         item.warna    || '',
            'Size':          item.size     || '',
            'Qty':           parseFloat(item.qty) || 0,
            'HargaSatuan':   parseFloat(item.hargaSatuan) || 0,
            'Subtotal':      parseFloat(item.subtotal) || ((parseFloat(item.qty)||0) * (parseFloat(item.hargaSatuan)||0)),
            'Keterangan':    item.keterangan || ''
          };
          detailSheet.appendRow(detailHeaders.map(function(h) {
            return detailMap.hasOwnProperty(h.toString().trim()) ? detailMap[h.toString().trim()] : '';
          }));
        }
      }
    }

    return { success: true, invoiceId: invoiceId, message: 'Transaksi berhasil disimpan' };
  } catch (err) {
    Logger.log('Error saveTransaction: ' + err.toString());
    return { success: false, invoiceId: '', message: 'Gagal menyimpan: ' + err.toString() };
  }
}

function updateTransaction(data) {
  try {
    var sheet = getSheet_('Buku Besar Kas');
    if (!sheet) return { success: false, message: 'Sheet Buku Besar Kas tidak ditemukan' };

    var allData = sheet.getDataRange().getValues();
    var headers = allData[0].map(function(h) { return h.toString().trim(); });

    var headerUpdated = false;
    if (headers.indexOf('ModalPO') === -1) { sheet.getRange(1, headers.length + 1).setValue('ModalPO'); headers.push('ModalPO'); headerUpdated = true; }
    if (headers.indexOf('Keuntungan') === -1) { sheet.getRange(1, headers.length + 1).setValue('Keuntungan'); headers.push('Keuntungan'); headerUpdated = true; }
    if (headerUpdated) { SpreadsheetApp.flush(); allData = sheet.getDataRange().getValues(); headers = allData[0].map(function(h) { return h.toString().trim(); }); }

    var colInvoice    = headers.indexOf('InvoiceID');
    var targetInvoice = (data.invoiceId || data.InvoiceID || '').toString().trim();
    if (!targetInvoice) return { success: false, message: 'InvoiceID tidak boleh kosong' };

    data = sanitizeData(data);

    for (var i = 1; i < allData.length; i++) {
      if (allData[i][colInvoice].toString().trim() === targetInvoice) {
        var fields = {
          'NamaKlien': data.namaKlien, 'Produk': data.produk,
          'Qty':           data.qty           !== undefined ? parseFloat(data.qty)           || 0 : undefined,
          'HargaSatuan':   data.hargaSatuan   !== undefined ? parseFloat(data.hargaSatuan)   || 0 : undefined,
          'Diskon':        data.diskon        !== undefined ? parseFloat(data.diskon)        || 0 : undefined,
          'Ongkir':        data.ongkir        !== undefined ? parseFloat(data.ongkir)        || 0 : undefined,
          'GrandTotal':    data.grandTotal    !== undefined ? parseFloat(data.grandTotal)    || 0 : undefined,
          'DP':            data.dp            !== undefined ? parseFloat(data.dp)            || 0 : undefined,
          'SisaTagihan':   data.sisaTagihan   !== undefined ? parseFloat(data.sisaTagihan)   || 0 : undefined,
          'StatusPembayaran': data.statusPembayaran, 'Proses': data.proses,
          'SumberKlien': data.sumberKlien, 'Catatan': data.catatan, 'Periode': data.periode,
          'WhatsApp': data.whatsapp, 'Instansi': data.instansi, 'KategoriProduk': data.kategoriProduk,
          'KeteranganTambahan': data.keteranganTambahan, 'Deadline': data.deadline,
          'ModalPO':    data.modalPO    !== undefined ? parseFloat(data.modalPO)    || 0 : undefined,
          'Keuntungan': data.keuntungan !== undefined ? parseFloat(data.keuntungan) || 0 : undefined,
          'TahapProduksi': data.tahapProduksi, 'PeriodeKas': data.periodeKas
        };
        for (var key in fields) {
          if (fields[key] !== undefined && fields[key] !== null) {
            var col = headers.indexOf(key);
            if (col !== -1) sheet.getRange(i + 1, col + 1).setValue(fields[key]);
          }
        }
        return { success: true, message: 'Transaksi ' + targetInvoice + ' berhasil diupdate' };
      }
    }
    return { success: false, message: 'InvoiceID ' + targetInvoice + ' tidak ditemukan' };
  } catch (err) {
    Logger.log('Error updateTransaction: ' + err.toString());
    return { success: false, message: 'Gagal update: ' + err.toString() };
  }
}

function updatePeriode(invoiceId, newPeriode) {
  try {
    var sheet = getSheet_('Buku Besar Kas');
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };

    var data       = sheet.getDataRange().getValues();
    var headers    = data[0];
    var colInvoice = headers.indexOf('InvoiceID');
    var colPeriode = headers.indexOf('Periode');

    for (var i = 1; i < data.length; i++) {
      if (data[i][colInvoice].toString().trim() === invoiceId.toString().trim()) {
        sheet.getRange(i + 1, colPeriode + 1).setValue(newPeriode);
        return { success: true, message: 'Periode berhasil diupdate' };
      }
    }
    return { success: false, message: 'InvoiceID tidak ditemukan' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function deleteTransaction(invoiceId, searchFields) {
  try {
    var targetId      = invoiceId ? invoiceId.toString().trim() : '';
    var isPlaceholder = !targetId || targetId.startsWith('M-');

    if (isPlaceholder && searchFields) {
      var sheet = getSheet_('Buku Besar Kas');
      if (sheet) {
        var data      = sheet.getDataRange().getValues();
        var headers   = data[0];
        var colInv    = headers.indexOf('InvoiceID');
        var colTgl    = headers.indexOf('Tanggal');
        var colKlien  = headers.indexOf('NamaKlien');
        var colGrand  = headers.indexOf('GrandTotal');
        for (var i = data.length - 1; i >= 1; i--) {
          var rowTgl   = colTgl !== -1 ? (data[i][colTgl] instanceof Date ? Utilities.formatDate(data[i][colTgl], Session.getScriptTimeZone(), 'yyyy-MM-dd') : data[i][colTgl].toString().trim()) : '';
          var rowKlien = colKlien !== -1 ? data[i][colKlien].toString().trim() : '';
          var rowGrand = colGrand !== -1 ? parseFloat(data[i][colGrand]) || 0 : 0;
          if (rowTgl === searchFields.tanggal && rowKlien === searchFields.namaKlien && Math.abs(rowGrand - (parseFloat(searchFields.grandTotal)||0)) < 0.01) {
            targetId = colInv !== -1 ? data[i][colInv].toString().trim() : '';
            break;
          }
        }
      }
    }

    if (targetId && !targetId.startsWith('M-')) {
      var detailSheet = getSheet_('Detail Transaksi');
      if (detailSheet && detailSheet.getLastRow() > 1) {
        var detailData = detailSheet.getDataRange().getValues();
        var colDInv    = detailData[0].indexOf('InvoiceID');
        if (colDInv !== -1) {
          for (var d = detailData.length - 1; d >= 1; d--) {
            if (detailData[d][colDInv].toString().trim() === targetId) detailSheet.deleteRow(d + 1);
          }
        }
      }
      var sheet2 = getSheet_('Buku Besar Kas');
      if (sheet2) {
        var data2    = sheet2.getDataRange().getValues();
        var colInv2  = data2[0].indexOf('InvoiceID');
        if (colInv2 !== -1) {
          for (var i = data2.length - 1; i >= 1; i--) {
            if (data2[i][colInv2].toString().trim() === targetId) {
              sheet2.deleteRow(i + 1);
              return { success: true, message: 'Transaksi ' + targetId + ' berhasil dihapus' };
            }
          }
        }
      }
    }

    if (isPlaceholder && searchFields) {
      var sheet3  = getSheet_('Buku Besar Kas');
      if (sheet3) {
        var data3   = sheet3.getDataRange().getValues();
        var headers3 = data3[0];
        var colTgl3  = headers3.indexOf('Tanggal');
        var colK3    = headers3.indexOf('NamaKlien');
        var colG3    = headers3.indexOf('GrandTotal');
        for (var i = data3.length - 1; i >= 1; i--) {
          var tgl3   = colTgl3 !== -1 ? (data3[i][colTgl3] instanceof Date ? Utilities.formatDate(data3[i][colTgl3], Session.getScriptTimeZone(), 'yyyy-MM-dd') : data3[i][colTgl3].toString().trim()) : '';
          var k3     = colK3   !== -1 ? data3[i][colK3].toString().trim() : '';
          var g3     = colG3   !== -1 ? parseFloat(data3[i][colG3]) || 0 : 0;
          if (tgl3 === searchFields.tanggal && k3 === searchFields.namaKlien && Math.abs(g3 - (parseFloat(searchFields.grandTotal)||0)) < 0.01) {
            sheet3.deleteRow(i + 1);
            return { success: true, message: 'Transaksi berhasil dihapus' };
          }
        }
      }
    }
    return { success: false, message: 'Transaksi tidak ditemukan' };
  } catch (err) {
    Logger.log('Error deleteTransaction: ' + err.toString());
    return { success: false, message: err.toString() };
  }
}

// ============================================================
// 6. BEBAN OPERASIONAL
// ============================================================

function saveBebanOps(data) {
  try {
    var sheet = getSheet_('BebanOperasional');
    if (!sheet) return { success: false, id: '', message: 'Sheet BebanOperasional tidak ditemukan' };

    var id      = generateSequentialId('BebanOperasional', 'OPS');
    data        = sanitizeData(data);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var mapping = {
      'ID': id,
      'Tanggal':    data.tanggal    || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      'Kategori':   data.kategori   || '',
      'SubKategori':data.subKategori|| '',
      'Keterangan': data.keterangan || '',
      'Nominal':    parseFloat(data.nominal) || 0,
      'InputBy':    data.inputBy    || '',
      'InputDate':  Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
    };
    sheet.appendRow(headers.map(function(h) { return mapping.hasOwnProperty(h.toString().trim()) ? mapping[h.toString().trim()] : ''; }));
    return { success: true, id: id, message: 'Beban operasional berhasil disimpan' };
  } catch (err) {
    return { success: false, id: '', message: err.toString() };
  }
}

function deleteBebanOps(id, searchFields) {
  return deleteRowById_('BebanOperasional', 'ID', id, searchFields);
}

// ============================================================
// 7. GUDANG / STOK
// ============================================================

function saveBarangGudang(data) {
  try {
    var sheet = getSheet_('Stok_Barang');
    if (!sheet) return { success: false, id: '', message: 'Sheet Stok_Barang tidak ditemukan' };

    var id          = generateSequentialId('Stok_Barang', 'STK');
    data            = sanitizeData(data);
    var stokMasuk   = parseFloat(data.stokMasuk) || 0;
    var hargaSatuan = parseFloat(data.hargaSatuan) || 0;
    var stokAkhir   = stokMasuk;
    var headers     = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var mapping = {
      'ID': id, 'Tanggal': Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      'NamaBarang': data.namaBarang || '', 'Kategori': data.kategori || '', 'Satuan': data.satuan || '',
      'StokMasuk': stokMasuk, 'StokKeluar': 0, 'StokAkhir': stokAkhir,
      'HargaSatuan': hargaSatuan, 'TotalNilai': stokAkhir * hargaSatuan,
      'Keterangan': data.keterangan || '',
      'InputDate': Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
    };
    sheet.appendRow(headers.map(function(h) { return mapping.hasOwnProperty(h.toString().trim()) ? mapping[h.toString().trim()] : ''; }));
    return { success: true, id: id, message: 'Barang gudang berhasil disimpan' };
  } catch (err) {
    return { success: false, id: '', message: err.toString() };
  }
}

function updateBarangGudang(data) {
  try {
    var sheet = getSheet_('Stok_Barang');
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };

    var allData = sheet.getDataRange().getValues();
    var headers = allData[0];
    var colId   = headers.indexOf('ID');
    var colMasuk = headers.indexOf('StokMasuk');
    var colKeluar = headers.indexOf('StokKeluar');
    var colAkhir  = headers.indexOf('StokAkhir');
    var colHarga  = headers.indexOf('HargaSatuan');
    var colTotal  = headers.indexOf('TotalNilai');
    var targetId  = (data.id || '').toString().trim();
    var targetField = (data.field || '').toString().trim();
    var targetValue = data.value;

    for (var i = 1; i < allData.length; i++) {
      if (allData[i][colId].toString().trim() === targetId) {
        var colTarget = headers.indexOf(targetField);
        if (colTarget !== -1) {
          var cellValue = ['StokMasuk','StokKeluar','HargaSatuan'].indexOf(targetField) !== -1
            ? parseFloat(targetValue) || 0 : targetValue;
          sheet.getRange(i + 1, colTarget + 1).setValue(cellValue);
        }
        var updatedRow = sheet.getRange(i + 1, 1, 1, headers.length).getValues()[0];
        var masuk  = parseFloat(updatedRow[colMasuk])  || 0;
        var keluar = parseFloat(updatedRow[colKeluar]) || 0;
        var harga  = parseFloat(updatedRow[colHarga])  || 0;
        var akhir  = masuk - keluar;
        if (colAkhir !== -1) sheet.getRange(i + 1, colAkhir + 1).setValue(akhir);
        if (colTotal  !== -1) sheet.getRange(i + 1, colTotal  + 1).setValue(akhir * harga);
        return { success: true, message: 'Barang gudang berhasil diupdate' };
      }
    }
    return { success: false, message: 'ID tidak ditemukan' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ============================================================
// 8. LAPORAN PEMBELIAN
// ============================================================

function saveLaporanPembelian(data) {
  try {
    var sheet = getSheet_('LaporanPembelian');
    if (!sheet) return { success: false, id: '', message: 'Sheet LaporanPembelian tidak ditemukan' };

    var id      = generateSequentialId('LaporanPembelian', 'PB');
    data        = sanitizeData(data);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (headers.indexOf('NoHP') === -1) { sheet.getRange(1, headers.length + 1).setValue('NoHP'); headers.push('NoHP'); }

    var mapping = {
      'ID': id,
      'Tanggal':          data.tanggal || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      'NamaPemasok':      data.namaPemasok || data.namaSupplier || '',
      'NamaSupplier':     data.namaPemasok || data.namaSupplier || '',
      'NamaBarang':       data.namaBarang  || '',
      'Kategori':         data.kategori    || '',
      'Qty':              parseFloat(data.qty) || 0,
      'Satuan':           data.satuan      || '',
      'HargaSatuan':      parseFloat(data.hargaSatuan) || 0,
      'TotalHarga':       parseFloat(data.totalHarga) || ((parseFloat(data.qty)||0) * (parseFloat(data.hargaSatuan)||0)),
      'MetodePembayaran': data.metodePembayaran || '',
      'StatusPembayaran': data.statusPembayaran || '',
      'Keterangan':       data.keterangan || '',
      'NoHP':             data.noHP       || '',
      'InputBy':          data.inputBy    || '',
      'InputDate':        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
    };
    sheet.appendRow(headers.map(function(h) { return mapping.hasOwnProperty(h.toString().trim()) ? mapping[h.toString().trim()] : ''; }));
    return { success: true, id: id, message: 'Laporan pembelian berhasil disimpan' };
  } catch (err) {
    return { success: false, id: '', message: err.toString() };
  }
}

function deletePembelian(id, searchFields) {
  return deleteRowById_('LaporanPembelian', 'ID', id, searchFields);
}

// ============================================================
// 9. OUTPUT PRODUKSI
// ============================================================

function getOutputProduksiData(bulan) {
  try {
    var result = { summary: [], detail: [] };

    var sheetSum = getSheet_('OutputProduksi');
    if (sheetSum && sheetSum.getLastRow() > 1) {
      var dataSum = sheetSum.getDataRange().getValues();
      var headSum = dataSum[0];
      var colTgl  = headSum.indexOf('Tanggal');
      for (var i = 1; i < dataSum.length; i++) {
        var tVal = dataSum[i][colTgl];
        var tStr = tVal instanceof Date ? Utilities.formatDate(tVal, Session.getScriptTimeZone(), 'yyyy-MM') : tVal.toString().substring(0,7);
        if (!bulan || tStr === bulan) {
          var obj = {};
          for (var j = 0; j < headSum.length; j++) {
            var v = dataSum[i][j];
            obj[headSum[j]] = v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : v;
          }
          result.summary.push(obj);
        }
      }
    }

    var sheetDet = getSheet_('OutputProduksiDetail');
    if (sheetDet && sheetDet.getLastRow() > 1) {
      var dataDet = sheetDet.getDataRange().getValues();
      var headDet = dataDet[0];
      var colTgl2 = headDet.indexOf('Tanggal');
      for (var i = 1; i < dataDet.length; i++) {
        var tVal2 = dataDet[i][colTgl2];
        var tStr2 = tVal2 instanceof Date ? Utilities.formatDate(tVal2, Session.getScriptTimeZone(), 'yyyy-MM') : tVal2.toString().substring(0,7);
        if (!bulan || tStr2 === bulan) {
          var obj2 = {};
          for (var j = 0; j < headDet.length; j++) {
            var v2 = dataDet[i][j];
            obj2[headDet[j]] = v2 instanceof Date ? Utilities.formatDate(v2, Session.getScriptTimeZone(), 'yyyy-MM-dd') : v2;
          }
          result.detail.push(obj2);
        }
      }
    }
    return result;
  } catch (err) {
    return { summary: [], detail: [], error: err.toString() };
  }
}

function saveOutputProduksi(data) {
  try {
    var sheet = getSheet_('OutputProduksi');
    if (!sheet) return { success: false, id: '', message: 'Sheet OutputProduksi tidak ditemukan' };

    var id      = generateSequentialId('OutputProduksi', 'OP');
    data        = sanitizeData(data);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var mapping = {
      'ID': id, 'Tanggal': data.tanggal || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      'Hari': data.hari || '', 'Pola': parseFloat(data.pola)||0, 'Sampel': parseFloat(data.sampel)||0,
      'Cutting': parseFloat(data.cutting)||0, 'Vislin': parseFloat(data.vislin)||0,
      'PasangVislin': parseFloat(data.pasangVislin)||0, 'CMT': parseFloat(data.cmt)||0,
      'LineApp': parseFloat(data.lineApp)||0, 'LineHijab': parseFloat(data.lineHijab)||0,
      'OutputCMT': parseFloat(data.outputCMT)||0, 'OutputLineApp': parseFloat(data.outputLineApp)||0,
      'OutputLineHijab': parseFloat(data.outputLineHijab)||0,
      'QcACMT': parseFloat(data.qcACMT)||0, 'QcAApp': parseFloat(data.qcAApp)||0,
      'QcAHijab': parseFloat(data.qcAHijab)||0, 'QcBCMT': parseFloat(data.qcBCMT)||0,
      'QcBApp': parseFloat(data.qcBApp)||0, 'QcBHijab': parseFloat(data.qcBHijab)||0,
      'LubangKancing': parseFloat(data.lubangKancing)||0, 'PasangKancing': parseFloat(data.pasangKancing)||0,
      'JahitSom': parseFloat(data.jahitSom)||0, 'BB': parseFloat(data.bb)||0,
      'Steam': parseFloat(data.steam)||0, 'Packing': parseFloat(data.packing)||0,
      'InputBy': data.inputBy || '',
      'InputDate': Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
    };
    sheet.appendRow(headers.map(function(h) { return mapping.hasOwnProperty(h.toString().trim()) ? mapping[h.toString().trim()] : ''; }));
    return { success: true, id: id, message: 'Output produksi berhasil disimpan' };
  } catch (err) {
    return { success: false, id: '', message: err.toString() };
  }
}

function saveOutputProduksiDetail(data) {
  try {
    var sheet = getSheet_('OutputProduksiDetail');
    if (!sheet) return { success: false, id: '', message: 'Sheet tidak ditemukan' };

    var id      = generateSequentialId('OutputProduksiDetail', 'OPD');
    data        = sanitizeData(data);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var mapping = {
      'ID': id, 'Tanggal': data.tanggal || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      'Hari': data.hari||'', 'Department': data.department||'', 'Brand': data.brand||'',
      'Model': data.model||'', 'Size': data.size||'', 'Warna': data.warna||'',
      'JenisKomponen': data.jenisKomponen||'', 'JumlahPO': parseFloat(data.jumlahPO)||0,
      'Cutting': parseFloat(data.cutting)||0, 'Vislin': parseFloat(data.vislin)||0,
      'PasangVislin': parseFloat(data.pasangVislin)||0, 'CMT': parseFloat(data.cmt)||0,
      'LineApp': parseFloat(data.lineApp)||0, 'LineHijab': parseFloat(data.lineHijab)||0,
      'OutputCMT': parseFloat(data.outputCMT)||0, 'OutputLineApp': parseFloat(data.outputLineApp)||0,
      'OutputLineHijab': parseFloat(data.outputLineHijab)||0, 'QcA': parseFloat(data.qcA)||0,
      'QcB': parseFloat(data.qcB)||0, 'LubangKancing': parseFloat(data.lubangKancing)||0,
      'PasangKancing': parseFloat(data.pasangKancing)||0, 'KancingManual': parseFloat(data.kancingManual)||0,
      'LogoPlat': parseFloat(data.logoPlat)||0, 'NgeSom': parseFloat(data.ngeSom)||0,
      'Dll': data.dll||'', 'BB': parseFloat(data.bb)||0, 'Steam': parseFloat(data.steam)||0,
      'Packing': parseFloat(data.packing)||0, 'InputBy': data.inputBy||'',
      'InputDate': Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
    };
    sheet.appendRow(headers.map(function(h) { return mapping.hasOwnProperty(h.toString().trim()) ? mapping[h.toString().trim()] : ''; }));
    return { success: true, id: id, message: 'Detail output produksi berhasil disimpan' };
  } catch (err) {
    return { success: false, id: '', message: err.toString() };
  }
}

function deleteOutputProduksi(id)       { return deleteRowById_('OutputProduksi', 'ID', id); }
function deleteOutputProduksiDetail(id) { return deleteRowById_('OutputProduksiDetail', 'ID', id); }

// ============================================================
// 10. LAPORAN FINISHING
// ============================================================

function getLaporanFinishingData(brand) {
  try {
    var sheet = getSheet_('LaporanFinishing');
    if (!sheet || sheet.getLastRow() <= 1) return [];
    var data     = sheet.getDataRange().getValues();
    var headers  = data[0];
    var colBrand = headers.indexOf('Brand');
    var results  = [];
    for (var i = 1; i < data.length; i++) {
      if (brand && brand.trim() !== '' && colBrand !== -1 && data[i][colBrand].toString().trim().toLowerCase() !== brand.trim().toLowerCase()) continue;
      var obj = {};
      for (var j = 0; j < headers.length; j++) {
        var val = data[i][j];
        obj[headers[j]] = val instanceof Date ? Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd') : val;
      }
      results.push(obj);
    }
    return results;
  } catch (err) { return []; }
}

// ============================================================
// FUNGSI: SIMPAN LAPORAN FINISHING
// ============================================================
function saveLaporanFinishing(data) {
  try {
    var sheet = getSheet_('LaporanFinishing');
    if (!sheet) return { success: false, id: '', message: 'Sheet tidak ditemukan' };

    var id      = generateSequentialId('LaporanFinishing', 'FIN');
    data        = sanitizeData(data);
    var op      = parseFloat(data.outputProduksi)||0, bb = parseFloat(data.bbPlakat)||0;
    var qc      = parseFloat(data.qc)||0, steam = parseFloat(data.steam)||0;
    var packing = parseFloat(data.packing)||0, kirim = parseFloat(data.pengiriman)||0;
    var qtyPO   = parseFloat(data.qtyPO)||0;
    var total   = kirim||packing||steam||qc||bb||op;
    var selisih = total - qtyPO;
    var status  = (selisih >= 0 && kirim > 0) ? 'Selesai' : 'Proses';
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var mapping = {
      'ID': id, 'Tanggal': Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      'Brand': data.brand||'', 'NamaProduk': data.namaProduk||'', 'Komponen': data.komponen||'',
      'Warna': data.warna||'', 'QtyPO': qtyPO, 'ActualCutting': parseFloat(data.actualCutting)||0,
      'OutputProduksi': op, 'BBPlakat': bb, 'QC': qc, 'Steam': steam, 'Packing': packing,
      'Pengiriman': kirim, 'TotalOutput': total, 'Selisih': selisih, 'Status': status,
      'InputBy': data.inputBy||'',
      'InputDate': Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
    };
    sheet.appendRow(headers.map(function(h) { return mapping.hasOwnProperty(h.toString().trim()) ? mapping[h.toString().trim()] : ''; }));
    return { success: true, id: id, message: 'Laporan finishing berhasil disimpan' };
  } catch (err) {
    return { success: false, id: '', message: err.toString() };
  }
}

function updateLaporanFinishing(data) {
  try {
    var sheet = getSheet_('LaporanFinishing');
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };

    var allData  = sheet.getDataRange().getValues();
    var headers  = allData[0];
    var colId    = headers.indexOf('ID');
    var targetId = (data.id || data.ID || '').toString().trim();
    if (!targetId) return { success: false, message: 'ID tidak boleh kosong' };
    data = sanitizeData(data);

    for (var i = 1; i < allData.length; i++) {
      if (allData[i][colId].toString().trim() === targetId) {
        var fields = {
          'Brand': data.brand, 'NamaProduk': data.namaProduk, 'Komponen': data.komponen, 'Warna': data.warna,
          'QtyPO':          data.qtyPO          !== undefined ? parseFloat(data.qtyPO)          || 0 : undefined,
          'ActualCutting':  data.actualCutting  !== undefined ? parseFloat(data.actualCutting)  || 0 : undefined,
          'OutputProduksi': data.outputProduksi !== undefined ? parseFloat(data.outputProduksi) || 0 : undefined,
          'BBPlakat':       data.bbPlakat       !== undefined ? parseFloat(data.bbPlakat)       || 0 : undefined,
          'QC':             data.qc             !== undefined ? parseFloat(data.qc)             || 0 : undefined,
          'Steam':          data.steam          !== undefined ? parseFloat(data.steam)          || 0 : undefined,
          'Packing':        data.packing        !== undefined ? parseFloat(data.packing)        || 0 : undefined,
          'Pengiriman':     data.pengiriman     !== undefined ? parseFloat(data.pengiriman)     || 0 : undefined
        };
        for (var key in fields) {
          if (fields[key] !== undefined) { var col = headers.indexOf(key); if (col !== -1) sheet.getRange(i+1, col+1).setValue(fields[key]); }
        }
        var row   = sheet.getRange(i + 1, 1, 1, headers.length).getValues()[0];
        var colsF = ['OutputProduksi','BBPlakat','QC','Steam','Packing','Pengiriman','QtyPO','TotalOutput','Selisih','Status'].map(function(n){ return headers.indexOf(n); });
        var op2   = parseFloat(row[colsF[0]])||0, bb2 = parseFloat(row[colsF[1]])||0;
        var qc2   = parseFloat(row[colsF[2]])||0, st2 = parseFloat(row[colsF[3]])||0;
        var pk2   = parseFloat(row[colsF[4]])||0, ki2 = parseFloat(row[colsF[5]])||0;
        var po2   = parseFloat(row[colsF[6]])||0;
        var tot2  = ki2||pk2||st2||qc2||bb2||op2;
        var sel2  = tot2 - po2;
        var sta2  = (sel2 >= 0 && ki2 > 0) ? 'Selesai' : 'Proses';
        if (colsF[7] !== -1) sheet.getRange(i+1, colsF[7]+1).setValue(tot2);
        if (colsF[8] !== -1) sheet.getRange(i+1, colsF[8]+1).setValue(sel2);
        if (colsF[9] !== -1) sheet.getRange(i+1, colsF[9]+1).setValue(sta2);
        return { success: true, message: 'Laporan finishing berhasil diupdate' };
      }
    }
    return { success: false, message: 'ID tidak ditemukan' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function deleteLaporanFinishing(id) { return deleteRowById_('LaporanFinishing', 'ID', id); }

// ============================================================
// 11. JADWAL PRODUKSI
// ============================================================

function getJadwalProduksiData(bulan) {
  return {
    jadwal:   getFilteredByMonth_('JadwalProduksi', bulan),
    samples:  getFilteredByMonth_('SampleTracking', bulan),
    cuttings: getFilteredByMonth_('CuttingTracking', bulan)
  };
}

function saveJadwalProduksi(data) {
  try {
    var sheet = getSheet_('JadwalProduksi');
    if (!sheet) return { success: false, id: '', message: 'Sheet tidak ditemukan' };
    var id      = generateSequentialId('JadwalProduksi', 'JP');
    data        = sanitizeData(data);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var mapping = {
      'ID': id, 'Tanggal': data.tanggal||Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      'Brand': data.brand||'', 'Model': data.model||'', 'Komponen': data.komponen||'',
      'Warna': data.warna||'', 'Size': data.size||'', 'QtyPO': parseFloat(data.qtyPO)||0,
      'TargetHarian': parseFloat(data.targetHarian)||0, 'TanggalMulai': data.tanggalMulai||'',
      'TanggalSelesai': data.tanggalSelesai||'', 'Status': data.status||'Pending',
      'Prioritas': data.prioritas||'Normal', 'Keterangan': data.keterangan||'',
      'InputBy': data.inputBy||'',
      'InputDate': Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
    };
    sheet.appendRow(headers.map(function(h){ return mapping.hasOwnProperty(h.toString().trim()) ? mapping[h.toString().trim()] : ''; }));
    return { success: true, id: id, message: 'Jadwal produksi berhasil disimpan' };
  } catch (err) { return { success: false, id: '', message: err.toString() }; }
}

function saveSampleTracking(data) {
  try {
    var sheet = getSheet_('SampleTracking');
    if (!sheet) return { success: false, id: '', message: 'Sheet tidak ditemukan' };
    var id      = generateSequentialId('SampleTracking', 'SMP');
    data        = sanitizeData(data);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var mapping = {
      'ID': id, 'Tanggal': data.tanggal||Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      'Brand': data.brand||'', 'Model': data.model||'', 'Komponen': data.komponen||'',
      'Warna': data.warna||'', 'Size': data.size||'', 'Qty': parseFloat(data.qty)||0,
      'StatusSampel': data.statusSampel||'Pending', 'TanggalMulai': data.tanggalMulai||'',
      'TanggalSelesai': data.tanggalSelesai||'', 'Approval': data.approval||'',
      'Catatan': data.catatan||'', 'InputBy': data.inputBy||'',
      'InputDate': Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
    };
    sheet.appendRow(headers.map(function(h){ return mapping.hasOwnProperty(h.toString().trim()) ? mapping[h.toString().trim()] : ''; }));
    return { success: true, id: id, message: 'Sample tracking berhasil disimpan' };
  } catch (err) { return { success: false, id: '', message: err.toString() }; }
}

function saveCuttingTracking(data) {
  try {
    var sheet = getSheet_('CuttingTracking');
    if (!sheet) return { success: false, id: '', message: 'Sheet tidak ditemukan' };
    var id      = generateSequentialId('CuttingTracking', 'CUT');
    data        = sanitizeData(data);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var mapping = {
      'ID': id, 'Tanggal': data.tanggal||Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      'Brand': data.brand||'', 'Model': data.model||'', 'Komponen': data.komponen||'',
      'Warna': data.warna||'', 'Size': data.size||'', 'Qty': parseFloat(data.qty)||0,
      'JenisBahan': data.jenisBahan||'', 'KebutuhanKain': parseFloat(data.kebutuhanKain)||0,
      'StatusCutting': data.statusCutting||'Pending', 'TanggalCutting': data.tanggalCutting||'',
      'HasilCutting': parseFloat(data.hasilCutting)||0, 'Catatan': data.catatan||'',
      'InputBy': data.inputBy||'',
      'InputDate': Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
    };
    sheet.appendRow(headers.map(function(h){ return mapping.hasOwnProperty(h.toString().trim()) ? mapping[h.toString().trim()] : ''; }));
    return { success: true, id: id, message: 'Cutting tracking berhasil disimpan' };
  } catch (err) { return { success: false, id: '', message: err.toString() }; }
}

function deleteJadwalProduksi(id)  { return deleteRowById_('JadwalProduksi', 'ID', id); }
function deleteSampleTracking(id)  { return deleteRowById_('SampleTracking', 'ID', id); }
function deleteCuttingTracking(id) { return deleteRowById_('CuttingTracking', 'ID', id); }

// ============================================================
// 12. PACKING LIST
// ============================================================

function getPackingListData(brand, status) {
  try {
    var sheet = getSheet_('PackingList');
    if (!sheet || sheet.getLastRow() <= 1) return [];
    var data     = sheet.getDataRange().getValues();
    var headers  = data[0];
    var colB     = headers.indexOf('Brand');
    var colS     = headers.indexOf('Status');
    var results  = [];
    for (var i = 1; i < data.length; i++) {
      if (brand  && brand.trim()  && colB !== -1 && data[i][colB].toString().trim().toLowerCase() !== brand.trim().toLowerCase())  continue;
      if (status && status.trim() && colS !== -1 && data[i][colS].toString().trim().toLowerCase() !== status.trim().toLowerCase()) continue;
      var obj = {};
      for (var j = 0; j < headers.length; j++) {
        var val = data[i][j];
        obj[headers[j]] = val instanceof Date ? Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd') : val;
      }
      results.push(obj);
    }
    return results;
  } catch (err) { return []; }
}

function savePackingList(data) {
  try {
    var sheet = getSheet_('PackingList');
    if (!sheet) return { success: false, id: '', message: 'Sheet tidak ditemukan' };
    var id      = generateSequentialId('PackingList', 'PCK');
    data        = sanitizeData(data);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var mapping = {
      'ID': id, 'Tanggal': data.tanggal||Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      'Brand': data.brand||'', 'Model': data.model||'', 'Komponen': data.komponen||'',
      'Warna': data.warna||'', 'Size': data.size||'', 'Qty': parseFloat(data.qty)||0,
      'NomorKarton': data.nomorKarton||'', 'IsiPerKarton': parseFloat(data.isiPerKarton)||0,
      'JumlahKarton': parseFloat(data.jumlahKarton)||0, 'BeratPerKarton': parseFloat(data.beratPerKarton)||0,
      'TotalBerat': parseFloat(data.totalBerat)||0, 'Dimensi': data.dimensi||'',
      'Tujuan': data.tujuan||'', 'NomorPO': data.nomorPO||'', 'TanggalKirim': data.tanggalKirim||'',
      'Ekspedisi': data.ekspedisi||'', 'NoResi': data.noResi||'', 'Status': data.status||'Pending',
      'Keterangan': data.keterangan||'', 'InputBy': data.inputBy||'',
      'InputDate': Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
    };
    sheet.appendRow(headers.map(function(h){ return mapping.hasOwnProperty(h.toString().trim()) ? mapping[h.toString().trim()] : ''; }));
    return { success: true, id: id, message: 'Packing list berhasil disimpan' };
  } catch (err) { return { success: false, id: '', message: err.toString() }; }
}

function deletePackingList(id) { return deleteRowById_('PackingList', 'ID', id); }

// ============================================================
// 13. LAPORAN KEUANGAN
// ============================================================

function getLaporanKeuangan(tahun, bulan) {
  try {
    var tahunInt  = parseInt(tahun) || new Date().getFullYear();
    var bulanInt  = parseInt(bulan) || (new Date().getMonth() + 1);
    var bulanStr  = ('0' + bulanInt).slice(-2);
    var periodeKey = tahunInt + '-' + bulanStr;

    // Saldo Awal
    var saldoAwalValues = {};
    var sheetSaldo = getSheet_('SaldoAwal');
    if (sheetSaldo && sheetSaldo.getLastRow() > 1) {
      var saldoData = sheetSaldo.getDataRange().getValues();
      var colNama = saldoData[0].indexOf('NamaAkun');
      var colNilai = saldoData[0].indexOf('Nilai');
      if (colNama !== -1 && colNilai !== -1) {
        for (var s = 1; s < saldoData.length; s++) {
          saldoAwalValues[saldoData[s][colNama].toString().trim()] = parseFloat(saldoData[s][colNilai]) || 0;
        }
      }
    }
    var saldoAwalKas = saldoAwalValues['SaldoAwalKas'] || 0;
    var piutangKaryawan = saldoAwalValues['PiutangKaryawan'] || 0;
    var persediaanBB = saldoAwalValues['PersediaanBahanBaku'] || 0;
    var persediaanBJ = saldoAwalValues['PersediaanBarangJadi'] || 0;
    var mesinJahit = saldoAwalValues['MesinJahit'] || 0;
    var mesinObras = saldoAwalValues['MesinObras'] || 0;
    var peralatanLainnya = saldoAwalValues['PeralatanLainnya'] || 0;
    var akumPenyusutan = saldoAwalValues['AkumulasiPenyusutan'] || 0;
    var hutangBank = saldoAwalValues['HutangBank'] || 0;
    var hutangPinjaman = saldoAwalValues['HutangPinjaman'] || 0;
    var modalAwal = saldoAwalValues['ModalAwal'] || 0;
    var labaDitahan = saldoAwalValues['LabaDitahan'] || 0;
    var prive = saldoAwalValues['Prive'] || 0;
    var pendapatanLainnya = saldoAwalValues['PendapatanLainnya'] || 0;

    // Buku Besar Kas
    var pendapatanUsaha = 0, dpMasuk = 0, totalPengeluaran = 0, piutangUsaha = 0;
    var sheetBBK = getSheet_('Buku Besar Kas');
    if (sheetBBK && sheetBBK.getLastRow() > 1) {
      var bbkData = sheetBBK.getDataRange().getValues();
      var bbkH    = bbkData[0];
      var colTgl  = bbkH.indexOf('Tanggal'), colKlien = bbkH.indexOf('NamaKlien');
      var colGT   = bbkH.indexOf('GrandTotal'), colDP = bbkH.indexOf('DP');
      var colSisa = bbkH.indexOf('SisaTagihan'), colSP = bbkH.indexOf('StatusPembayaran');
      for (var b = 1; b < bbkData.length; b++) {
        var tBBK = bbkData[b][colTgl];
        var pBBK = tBBK instanceof Date ? Utilities.formatDate(tBBK, Session.getScriptTimeZone(), 'yyyy-MM') : tBBK.toString().substring(0,7);
        if (pBBK !== periodeKey) continue;
        var klien = bbkData[b][colKlien] ? bbkData[b][colKlien].toString().trim() : '';
        var gt    = colGT !== -1 ? parseFloat(bbkData[b][colGT]) || 0 : 0;
        var dp    = colDP !== -1 ? parseFloat(bbkData[b][colDP]) || 0 : 0;
        var sisa  = colSisa !== -1 ? parseFloat(bbkData[b][colSisa]) || 0 : 0;
        var sp    = colSP  !== -1 ? bbkData[b][colSP].toString().trim().toLowerCase() : '';
        if (klien !== '') { pendapatanUsaha += gt; dpMasuk += dp; } else { totalPengeluaran += gt; }
        if (sp === 'dp' || sp === 'piutang') piutangUsaha += sisa;
      }
    }

    // Beban Operasional
    var opexCat = ['Operasional Harian','Operasional Produksi','Gaji & Infaq','Bidang Sosial','Marketing','Admin Bank','Sewa Tempat','Konsumsi','Listrik & Wifi','Transportasi','Operasional Kantor','Service & Perawatan','Obat & P3K'];
    var hppBreakdown = {}, opexBreakdown = {}, totalBebanOps = 0;
    opexCat.forEach(function(c) { opexBreakdown[c] = 0; });

    var sheetBeban = getSheet_('BebanOperasional');
    if (sheetBeban && sheetBeban.getLastRow() > 1) {
      var bData = sheetBeban.getDataRange().getValues();
      var bH    = bData[0];
      var colBT = bH.indexOf('Tanggal'), colBK = bH.indexOf('Kategori'), colBN = bH.indexOf('Nominal');
      for (var be = 1; be < bData.length; be++) {
        var tB = bData[be][colBT];
        var pB = tB instanceof Date ? Utilities.formatDate(tB, Session.getScriptTimeZone(), 'yyyy-MM') : tB.toString().substring(0,7);
        if (pB !== periodeKey) continue;
        var kat = colBK !== -1 ? bData[be][colBK].toString().trim() : '';
        var nom = colBN !== -1 ? parseFloat(bData[be][colBN]) || 0 : 0;
        totalBebanOps += nom;
        if (hppBreakdown.hasOwnProperty(kat)) hppBreakdown[kat] += nom;
        else if (opexBreakdown.hasOwnProperty(kat)) opexBreakdown[kat] += nom;
        else { if (!opexBreakdown['Lainnya']) opexBreakdown['Lainnya'] = 0; opexBreakdown['Lainnya'] += nom; }
      }
    }

    // Pembelian
    var totalPembelian = 0;
    var sheetPembelian = getSheet_('LaporanPembelian');
    if (sheetPembelian && sheetPembelian.getLastRow() > 1) {
      var pData = sheetPembelian.getDataRange().getValues();
      var pH    = pData[0];
      var colPT = pH.indexOf('Tanggal'), colPH = pH.indexOf('TotalHarga');
      for (var p = 1; p < pData.length; p++) {
        var tP = pData[p][colPT];
        var pP = tP instanceof Date ? Utilities.formatDate(tP, Session.getScriptTimeZone(), 'yyyy-MM') : tP.toString().substring(0,7);
        if (pP === periodeKey) totalPembelian += colPH !== -1 ? parseFloat(pData[p][colPH]) || 0 : 0;
      }
    }
    hppBreakdown['Bahan Kain'] = (hppBreakdown['Bahan Kain'] || 0) + totalPembelian;

    var totalHPP  = Object.keys(hppBreakdown).reduce(function(s,k){ return s + hppBreakdown[k]; }, 0);
    var totalOpex = Object.keys(opexBreakdown).reduce(function(s,k){ return s + opexBreakdown[k]; }, 0);
    var labaKotor = pendapatanUsaha + pendapatanLainnya - totalHPP;
    var labaUsaha = labaKotor - totalOpex;
    var kasBank   = saldoAwalKas + dpMasuk - totalPengeluaran - totalBebanOps;
    var totalHL   = kasBank + piutangUsaha + piutangKaryawan + persediaanBB + persediaanBJ;
    var totalHT   = mesinJahit + mesinObras + peralatanLainnya - akumPenyusutan;
    var totalAset = totalHL + totalHT;
    var totalKewajiban = hutangBank + hutangPinjaman;
    var totalModal = modalAwal + labaDitahan + labaUsaha - prive;

    return {
      periode: { tahun: tahunInt, bulan: bulanInt, label: periodeKey },
      labaRugi: { pendapatanUsaha: pendapatanUsaha, pendapatanLainnya: pendapatanLainnya, totalPendapatan: pendapatanUsaha+pendapatanLainnya, hpp: hppBreakdown, totalHPP: totalHPP, totalPembelianBahan: totalPembelian, labaKotor: labaKotor, opex: opexBreakdown, totalOpex: totalOpex, labaUsaha: labaUsaha },
      neraca: {
        hartaLancar: { kasBank: kasBank, piutangUsaha: piutangUsaha, piutangKaryawan: piutangKaryawan, persediaanBahanBaku: persediaanBB, persediaanBarangJadi: persediaanBJ, total: totalHL },
        hartaTetap:  { mesinJahit: mesinJahit, mesinObras: mesinObras, peralatanLainnya: peralatanLainnya, akumulasiPenyusutan: akumPenyusutan, total: totalHT },
        totalAset: totalAset,
        kewajiban: { hutangBank: hutangBank, hutangPinjaman: hutangPinjaman, total: totalKewajiban },
        modal: { modalAwal: modalAwal, labaDitahan: labaDitahan, labaUsaha: labaUsaha, prive: prive, total: totalModal },
        totalKewajibanModal: totalKewajiban + totalModal
      },
      summary: { dpMasuk: dpMasuk, totalPengeluaran: totalPengeluaran, totalBebanOps: totalBebanOps, saldoAwalKas: saldoAwalKas },
      saldoAwal: saldoAwalValues
    };
  } catch (err) {
    Logger.log('Error getLaporanKeuangan: ' + err.toString());
    return { error: err.toString() };
  }
}

function updateSaldoAwal(data) {
  try {
    var sheet = getSheet_('SaldoAwal');
    if (!sheet) return { success: false, message: 'Sheet SaldoAwal tidak ditemukan' };
    var allData  = sheet.getDataRange().getValues();
    var headers  = allData[0];
    var colNama  = headers.indexOf('NamaAkun');
    var colNilai = headers.indexOf('Nilai');
    if (colNama === -1 || colNilai === -1) return { success: false, message: 'Struktur kolom tidak valid' };
    if (!data || !Array.isArray(data)) return { success: false, message: 'Data harus berupa array' };

    var updatedCount = 0;
    for (var d = 0; d < data.length; d++) {
      var target = (data[d].namaAkun || '').toString().trim();
      var nilai  = parseFloat(data[d].nilai) || 0;
      if (!target) continue;
      var found = false;
      for (var i = 1; i < allData.length; i++) {
        if (allData[i][colNama].toString().trim() === target) {
          sheet.getRange(i + 1, colNilai + 1).setValue(nilai);
          found = true; updatedCount++; break;
        }
      }
      if (!found) {
        var newRow = [];
        for (var h = 0; h < headers.length; h++) newRow.push(h === colNama ? target : h === colNilai ? nilai : '');
        sheet.appendRow(newRow); updatedCount++;
      }
    }
    return { success: true, message: updatedCount + ' akun berhasil diupdate' };
  } catch (err) { return { success: false, message: err.toString() }; }
}

// ============================================================
// 14. FUNGSI getAllData (Frontend ERP)
// ============================================================

function getAllData() {
  try {
    var base   = getInitialData();
    var result = {
      transaksi: [], belanjaVendor: [], bebanOps: [], gudang: [], pembelian: [],
      jadwalHarian: [], sampleTracking: [], cuttingTracking: [],
      outputSummary: [], outputDetail: [], finishing: [], packingList: [],
      saldoAwal: {}
    };

    // Transaksi
    result.transaksi = base.transactions.map(function(t) {
      return {
        id: t.InvoiceID||'', tanggal: t.Tanggal||'', namaKlien: t.NamaKlien||'',
        produk: t.Produk||'', qty: t.Qty||0, hargaSatuan: t.HargaSatuan||0,
        diskon: t.Diskon||0, ongkir: t.Ongkir||0, totalHarga: t.TotalHarga||0,
        grandTotal: Number(t.GrandTotal||0), dp: Number(t.DP||0),
        sisaTagihan: Number(t.SisaTagihan||0), statusPembayaran: t.StatusPembayaran||'',
        proses: t.Proses||'', sumber: t.SumberKlien||'', catatan: t.Catatan||'',
        periode: t.Periode||'', whatsapp: t.WhatsApp||'', instansi: t.Instansi||'',
        kategoriProduk: t.KategoriProduk||'', keteranganTambahan: t.KeteranganTambahan||'',
        deadline: t.Deadline||'', modalPO: Number(t.ModalPO||0),
        keuntungan: Number(t.Keuntungan||0), tahapProduksi: t.TahapProduksi||'',
        periodeKas: t.PeriodeKas||''
      };
    });

    // Attach detail items
    try {
      var detailSheet = getSheet_('Detail Transaksi');
      if (detailSheet && detailSheet.getLastRow() > 1) {
        var dData = detailSheet.getDataRange().getValues();
        var dHead = dData[0];
        var itemsByInv = {};
        for (var di = 1; di < dData.length; di++) {
          var dObj = {};
          for (var dj = 0; dj < dHead.length; dj++) {
            var dv = dData[di][dj];
            dObj[dHead[dj]] = dv instanceof Date ? Utilities.formatDate(dv, Session.getScriptTimeZone(), 'yyyy-MM-dd') : dv;
          }
          var invId = (dObj.InvoiceID||'').toString().trim();
          if (invId) {
            if (!itemsByInv[invId]) itemsByInv[invId] = [];
            itemsByInv[invId].push({ namaProduk: dObj.NamaProduk||'', ukuran: dObj.Size||dObj.Ukuran||'', kategoriProduk: dObj.KategoriProduk||dObj.Komponen||'', qty: Number(dObj.Qty||0), hargaSatuan: Number(dObj.HargaSatuan||0), subtotal: Number(dObj.Subtotal||0) });
          }
        }
        result.transaksi = result.transaksi.map(function(t) { t.items = itemsByInv[t.id] || []; return t; });
      }
    } catch(e) { Logger.log('Error loading detail items: ' + e.toString()); }

    // BebanOps
    result.bebanOps = base.bebanOps.map(function(b) {
      return { id: b.ID||'', tanggal: b.Tanggal||'', kategori: b.Kategori||'', subKategori: b.SubKategori||'', keterangan: b.Keterangan||'', nominal: Number(b.Nominal||0), inputBy: b.InputBy||'' };
    });

    // Gudang
    result.gudang = base.stokBarang.map(function(g) {
      return { id: g.ID||'', namaBarang: g.NamaBarang||'', kategori: g.Kategori||'', satuan: g.Satuan||'', stokMasuk: Number(g.StokMasuk||0), stokKeluar: Number(g.StokKeluar||0), stokAkhir: Number(g.StokAkhir||0), harga: Number(g.HargaSatuan||0), totalNilai: Number(g.TotalNilai||0), keterangan: g.Keterangan||'' };
    });

    // Pembelian
    result.pembelian = base.pembelian.map(function(p) {
      return { id: p.ID||'', tanggal: p.Tanggal||'', pemasok: p.NamaPemasok||'', barang: p.NamaBarang||'', qty: Number(p.Qty||0), satuan: p.Satuan||'', harga: Number(p.HargaSatuan||0), totalHarga: Number(p.TotalHarga||0), keterangan: p.Keterangan||'', noHP: p.NoHP||'' };
    });
    result.belanjaVendor = result.pembelian.map(function(p) {
      return { id: p.id, tanggal: p.tanggal, namaPemasok: p.pemasok, namaBarang: p.barang, qty: p.qty, satuan: p.satuan, hargaSatuan: p.harga, totalHarga: p.totalHarga, keterangan: p.keterangan, noHP: p.noHP };
    });

    // Jadwal
    var jadwalSheets = getJadwalProduksiData('');
    if (jadwalSheets) {
      result.jadwalHarian  = jadwalSheets.jadwal   || [];
      result.sampleTracking  = jadwalSheets.samples  || [];
      result.cuttingTracking = jadwalSheets.cuttings || [];
    }

    // Output Produksi
    var outputSheets = getOutputProduksiData('');
    if (outputSheets) { result.outputSummary = outputSheets.summary || []; result.outputDetail = outputSheets.detail || []; }

    // Finishing
    result.finishing = (getLaporanFinishingData('') || []).map(function(f) {
      return { id: f.ID||'', brand: f.Brand||'', produk: f.NamaProduk||'', komponen: f.Komponen||'', warna: f.Warna||'', po: Number(f.QtyPO||0), actCutting: Number(f.ActualCutting||0), output: Number(f.OutputProduksi||0), bbPlakat: Number(f.BBPlakat||0), qc: Number(f.QC||0), steam: Number(f.Steam||0), packing: Number(f.Packing||0), kirim: Number(f.Pengiriman||0), total: Number(f.TotalOutput||0), selisih: Number(f.Selisih||0), status: f.Status||'' };
    });

    // Packing List
    result.packingList = (getPackingListData('','') || []).map(function(p) {
      return { id: p.ID||'', tanggal: p.Tanggal||'', brand: p.Brand||'', noPO: p.NoPO||'', item: p.NamaItem||'', warna: p.Warna||'', size: p.Size||'', qty: Number(p.QtyPcs||p.Qty||0), carton: Number(p.JumlahKarton||0), berat: Number(p.TotalBerat||0), noKirim: p.NomorPO||'', tujuan: p.Tujuan||'', status: p.Status||'', keterangan: p.Keterangan||'' };
    });

    // Saldo Awal
    var saObj = {};
    (base.saldoAwal || []).forEach(function(s) {
      var nama  = (s.NamaAkun||s.namaAkun||'').toString().trim();
      var nilai = Number(s.Nilai||s.nilai||0);
      if (nama.indexOf('Kas') !== -1) saObj.kasBank = nilai;
      else if (nama.indexOf('Piutang Karyawan') !== -1) saObj.piutangKaryawan = nilai;
      else if (nama.indexOf('Persediaan Bahan') !== -1) saObj.persediaanBB = nilai;
      else if (nama.indexOf('Persediaan Barang') !== -1) saObj.persediaanBJ = nilai;
      else if (nama.indexOf('Mesin Jahit') !== -1) saObj.mesinJahit = nilai;
      else if (nama.indexOf('Mesin Obras') !== -1) saObj.mesinObras = nilai;
      else if (nama.indexOf('Peralatan') !== -1) saObj.peralatan = nilai;
      else if (nama.indexOf('Penyusutan') !== -1) saObj.penyusutan = nilai;
      else if (nama.indexOf('Hutang Bank') !== -1) saObj.hutangBank = nilai;
      else if (nama.indexOf('Hutang Pinjaman') !== -1) saObj.hutangPinjaman = nilai;
      else if (nama.indexOf('Modal Awal') !== -1) saObj.modalAwal = nilai;
      else if (nama.indexOf('Laba Ditahan') !== -1) saObj.labaDitahan = nilai;
      else if (nama.indexOf('Prive') !== -1) saObj.prive = nilai;
      else if (nama.indexOf('Pendapatan Lainnya') !== -1) saObj.pendapatanLainnya = nilai;
    });
    result.saldoAwal = saObj;

    return result;
  } catch (err) {
    Logger.log('Error getAllData: ' + err.toString());
    return { transaksi:[], belanjaVendor:[], bebanOps:[], gudang:[], pembelian:[], jadwalHarian:[], sampleTracking:[], cuttingTracking:[], outputSummary:[], outputDetail:[], finishing:[], packingList:[], saldoAwal:{}, error: err.toString() };
  }
}

// ============================================================
// 15. ALIAS FRONTEND ERP (bridge penamaan index.html → Code.gs)
// ============================================================

function addTransaksi(data) {
  if (data.sumber && !data.sumberKlien) data.sumberKlien = data.sumber;
  if (data.modalPO !== undefined && data.grandTotal !== undefined) data.keuntungan = Number(data.grandTotal||0) - Number(data.modalPO||0);
  if (data.items && data.items.length > 0 && !data.produk) {
    data.produk = data.items.map(function(it){ return it.namaProduk; }).join(', ');
    if (!data.qty) data.qty = data.items.reduce(function(s,it){ return s + (Number(it.qty)||0); }, 0);
    if (!data.hargaSatuan && data.items.length === 1) data.hargaSatuan = data.items[0].hargaSatuan;
  }
  var result = saveTransaction(data);
  return { success: result.success, id: result.invoiceId, message: result.message };
}

function updateTransaksi(invoiceIdOrIndex, data) {
  if (data && typeof data === 'object') data.invoiceId = data.invoiceId || data.id || invoiceIdOrIndex;
  return updateTransaction(data);
}

function updateTransaksiStatus(indexOrId, newStatus, searchFields) {
  try {
    var sheet = getSheet_('Buku Besar Kas');
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };

    var allData  = sheet.getDataRange().getValues();
    var headers  = allData[0];
    var colInv   = headers.indexOf('InvoiceID');
    var colSP    = headers.indexOf('StatusPembayaran');
    var colDP    = headers.indexOf('DP');
    var colGrand = headers.indexOf('GrandTotal');
    var colSisa  = headers.indexOf('SisaTagihan');
    if (colSP === -1) return { success: false, message: 'Kolom StatusPembayaran tidak ditemukan' };

    var targetRow = -1;
    if (typeof indexOrId === 'number' && indexOrId >= 0) {
      var sheetRow = allData.length - 1 - indexOrId;
      if (sheetRow >= 1 && sheetRow < allData.length) targetRow = sheetRow;
    }
    if (targetRow === -1) {
      var searchId = indexOrId.toString().trim();
      for (var i = 1; i < allData.length; i++) {
        if (colInv !== -1 && allData[i][colInv].toString().trim() === searchId) { targetRow = i; break; }
      }
    }
    if (targetRow === -1 && searchFields) {
      var colTgl  = headers.indexOf('Tanggal'), colKlien = headers.indexOf('NamaKlien');
      for (var i = allData.length - 1; i >= 1; i--) {
        var tgl = colTgl !== -1 ? (allData[i][colTgl] instanceof Date ? Utilities.formatDate(allData[i][colTgl], Session.getScriptTimeZone(), 'yyyy-MM-dd') : allData[i][colTgl].toString().trim()) : '';
        var kli = colKlien !== -1 ? allData[i][colKlien].toString().trim() : '';
        var grd = colGrand !== -1 ? parseFloat(allData[i][colGrand]) || 0 : 0;
        if (tgl === searchFields.tanggal && kli === searchFields.namaKlien && Math.abs(grd - (parseFloat(searchFields.grandTotal)||0)) < 0.01) { targetRow = i; break; }
      }
    }
    if (targetRow === -1) return { success: false, message: 'Transaksi tidak ditemukan' };

    sheet.getRange(targetRow + 1, colSP + 1).setValue(newStatus);
    if (newStatus === 'Lunas') {
      if (colSisa  !== -1) sheet.getRange(targetRow + 1, colSisa  + 1).setValue(0);
      if (colDP !== -1 && colGrand !== -1) sheet.getRange(targetRow + 1, colDP + 1).setValue(parseFloat(allData[targetRow][colGrand]) || 0);
    }
    return { success: true, message: 'Status berhasil diperbarui' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function addBelanjaVendor(data) {
  return saveLaporanPembelian({ tanggal: data.tanggal, namaPemasok: data.namaPemasok||data.pemasok||'', namaBarang: data.namaBarang||data.barang||'', qty: data.qty, satuan: data.satuan, hargaSatuan: data.hargaSatuan||data.harga, totalHarga: data.totalHarga||(Number(data.qty||0)*Number(data.hargaSatuan||data.harga||0)), keterangan: data.keterangan||'', noHP: data.noHP||'' });
}
function addBebanOps(data)     { return saveBebanOps(data); }
function addGudangItem(data)   { return saveBarangGudang({ namaBarang: data.namaBarang||data.nama||'', kategori: data.kategori, satuan: data.satuan, stokMasuk: data.stokMasuk, hargaSatuan: data.hargaSatuan||data.harga, keterangan: data.keterangan||data.ket||'' }); }
function addPembelian(data)    { return saveLaporanPembelian({ tanggal: data.tanggal, namaPemasok: data.pemasok||data.namaPemasok||'', namaBarang: data.barang||data.namaBarang||'', qty: data.qty, satuan: data.satuan, hargaSatuan: data.harga||data.hargaSatuan, totalHarga: Number(data.qty||0)*Number(data.harga||data.hargaSatuan||0), keterangan: data.keterangan||'' }); }
function addJadwalHarian(data) { return saveJadwalProduksi(data); }
function addSampleTracking(data)  { return saveSampleTracking(data); }
function addCuttingTracking(data) { return saveCuttingTracking(data); }
function addOutputSummary(data)   { return saveOutputProduksi(data); }
function addOutputDetail(data)    { return saveOutputProduksiDetail(data); }
function addFinishing(data)       { return saveLaporanFinishing({ brand: data.brand, namaProduk: data.produk||data.namaProduk||'', komponen: data.komponen, warna: data.warna, qtyPO: data.po||data.qtyPO, actualCutting: data.actCutting||data.actualCutting, outputProduksi: data.output||data.outputProduksi, bbPlakat: data.bbPlakat||data.bb, qc: data.qc, steam: data.steam, packing: data.packing, pengiriman: data.kirim||data.pengiriman, inputBy: data.inputBy||'' }); }
function addPackingList(data)     { return savePackingList({ tanggal: data.tanggal, brand: data.brand, nomorPO: data.noPO, nomorKarton: data.nomorKarton||'', namaItem: data.item||data.namaItem||'', warna: data.warna, size: data.size, qty: data.qty||data.qtyPcs, jumlahKarton: data.carton||data.jumlahKarton, totalBerat: data.berat||data.totalBerat, noPengiriman: data.noKirim||'', tujuan: data.tujuan, status: data.status, keterangan: data.keterangan||data.ket||'', inputBy: data.inputBy||'' }); }
function deleteJadwalData(type, id) {
  if (type === 'harian')   return deleteJadwalProduksi(id);
  if (type === 'sample')   return deleteSampleTracking(id);
  if (type === 'cutting')  return deleteCuttingTracking(id);
  return { success: false, message: 'Tipe tidak dikenali' };
}
function deleteOutputDetail(id)  { return deleteOutputProduksiDetail(id); }
function deleteFinishing(id)     { return deleteLaporanFinishing(id); }
function deletePacking(id)       { return deletePackingList(id); }
function deleteGudangItem(indexOrId) {
  var result = deleteRowById_('Stok_Barang', 'ID', indexOrId);
  if (result.success) return result;
  if (typeof indexOrId === 'number') {
    var sheet = getSheet_('Stok_Barang');
    if (sheet && indexOrId >= 0 && indexOrId < sheet.getLastRow() - 1) {
      sheet.deleteRow(indexOrId + 2);
      return { success: true, message: 'Barang berhasil dihapus' };
    }
  }
  return { success: false, message: 'Barang tidak ditemukan' };
}
function updateGudangStok(indexOrId, data) {
  var mappedData = { id: data.id || data.ID || indexOrId, stokMasuk: data.stokMasuk, stokKeluar: data.stokKeluar };
  if (typeof indexOrId === 'number') {
    var sheet = getSheet_('Stok_Barang');
    if (sheet && sheet.getLastRow() > 1) {
      var allData = sheet.getDataRange().getValues();
      var headers = allData[0];
      var colMasuk = headers.indexOf('StokMasuk'), colKeluar = headers.indexOf('StokKeluar');
      var colAkhir = headers.indexOf('StokAkhir'), colHarga  = headers.indexOf('HargaSatuan');
      var colTotal = headers.indexOf('TotalNilai'), colUpdate = headers.indexOf('TanggalUpdate');
      var rowIdx = indexOrId + 1;
      if (rowIdx < allData.length) {
        var masuk  = data.stokMasuk  !== undefined ? parseFloat(data.stokMasuk)  || 0 : parseFloat(allData[rowIdx][colMasuk])  || 0;
        var keluar = data.stokKeluar !== undefined ? parseFloat(data.stokKeluar) || 0 : parseFloat(allData[rowIdx][colKeluar]) || 0;
        var akhir  = masuk - keluar;
        var harga  = parseFloat(allData[rowIdx][colHarga]) || 0;
        if (colMasuk  !== -1) sheet.getRange(rowIdx + 1, colMasuk  + 1).setValue(masuk);
        if (colKeluar !== -1) sheet.getRange(rowIdx + 1, colKeluar + 1).setValue(keluar);
        if (colAkhir  !== -1) sheet.getRange(rowIdx + 1, colAkhir  + 1).setValue(akhir);
        if (colTotal  !== -1) sheet.getRange(rowIdx + 1, colTotal  + 1).setValue(akhir * harga);
        if (colUpdate !== -1) sheet.getRange(rowIdx + 1, colUpdate + 1).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
        return { success: true, message: 'Stok berhasil diperbarui' };
      }
    }
  }
  return updateBarangGudang(mappedData);
}

// ============================================================
// 16. ═══════════════════════════════════════════════════════
//     BAGIAN KATALOG — 2 SHEET BARU DI DATABASE ERP
//     Sheet: "Data Pesanan (Katalog Kitajahitin)"   [13 kolom]
//     Sheet: "Data Pelanggan (Katalog Kitajahitin)" [7 kolom]
// ═══════════════════════════════════════════════════════════
// ============================================================

// ── Helper: Get/Create Sheet Pesanan Katalog ──
function getOrCreatePesananSheet() {
  var ss        = SpreadsheetApp.openById(SS_ID);
  var sheetName = "Data Pesanan (Katalog Kitajahitin)";
  var sheet     = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, 13).setValues([[
      "Timestamp","Order ID","Nama","WhatsApp","Email",
      "Kode Produk / Kategori","Qty","Total Tagihan",
      "Catatan / Ukuran","Status Pembayaran","Status Pesanan",
      "Alamat","Tipe"
    ]]);
    sheet.getRange(1, 1, 1, 13).setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    Logger.log("✓ Sheet Pesanan Katalog dibuat.");
  }
  return sheet;
}

function getOrCreatePelangganSheet() {
  var ss        = SpreadsheetApp.openById(SS_ID);
  var sheetName = "Data Pelanggan (Katalog Kitajahitin)";
  var sheet     = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, 7).setValues([[
      "Timestamp","Customer ID","Nama","WhatsApp","Email","PIN Hash","Alamat Default"
    ]]);
    sheet.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    Logger.log("✓ Sheet Pelanggan Katalog dibuat.");
  }
  return sheet;
}

// ── Ambil data pesanan katalog (untuk halaman monitor ERP) ──
function getKatalogPesanan(tipe) {
  try {
    var sheet = getOrCreatePesananSheet();
    if (sheet.getLastRow() <= 1) return [];

    var data    = sheet.getDataRange().getValues();
    var results = [];

    for (var i = 1; i < data.length; i++) {
      var rowTipe = data[i][12] ? data[i][12].toString() : "";
      if (tipe && tipe.trim() !== "" && rowTipe !== tipe) continue;

      var ts      = data[i][0];
      var dateStr = ts instanceof Date ? Utilities.formatDate(ts, "GMT+7", "d MMM yyyy, HH:mm") : ts.toString();
      var total   = data[i][7];

      results.push({
        timestamp:     dateStr,
        orderId:       data[i][1].toString(),
        nama:          data[i][2].toString(),
        wa:            data[i][3].toString().replace(/'/g, ''),
        email:         data[i][4].toString(),
        kodeProduk:    data[i][5].toString(),
        qty:           data[i][6],
        totalTagihan:  typeof total === 'number' ? total : 0,
        totalStr:      typeof total === 'number' ? "Rp " + Number(total).toLocaleString("id-ID") : total.toString(),
        catatan:       data[i][8].toString(),
        statusBayar:   data[i][9].toString(),
        statusPesanan: data[i][10].toString(),
        alamat:        data[i][11] ? data[i][11].toString() : "",
        tipe:          rowTipe,
        rowIndex:      i + 1
      });
    }
    results.reverse();
    return results;
  } catch (err) {
    Logger.log("getKatalogPesanan error: " + err.toString());
    return [];
  }
}

// ── Ambil data pelanggan katalog (untuk halaman monitor ERP) ──
function getKatalogPelanggan() {
  try {
    var sheet = getOrCreatePelangganSheet();
    if (sheet.getLastRow() <= 1) return [];

    var data    = sheet.getDataRange().getValues();
    var results = [];

    for (var i = 1; i < data.length; i++) {
      var ts      = data[i][0];
      var dateStr = ts instanceof Date ? Utilities.formatDate(ts, "GMT+7", "d MMM yyyy") : ts.toString();
      results.push({
        timestamp:  dateStr,
        customerId: data[i][1].toString(),
        nama:       data[i][2].toString(),
        wa:         data[i][3].toString().replace(/'/g, ''),
        email:      data[i][4].toString(),
        alamat:     data[i][6] ? data[i][6].toString() : ""
      });
    }
    results.reverse();
    return results;
  } catch (err) {
    Logger.log("getKatalogPelanggan error: " + err.toString());
    return [];
  }
}

// ── Update status pesanan katalog (dari halaman monitor ERP) ──
function updateStatusKatalogPesanan(rowIndex, statusBayar, statusPesanan) {
  try {
    var sheet = getOrCreatePesananSheet();
    if (rowIndex < 2 || rowIndex > sheet.getLastRow()) {
      return { success: false, message: "Row index tidak valid." };
    }
    if (statusBayar)   sheet.getRange(rowIndex, 10).setValue(statusBayar);
    if (statusPesanan) sheet.getRange(rowIndex, 11).setValue(statusPesanan);

    // Jika sudah Lunas, update juga di Buku Besar Kas (sync ERP)
    if (statusBayar === "Lunas") {
      var orderId = sheet.getRange(rowIndex, 2).getValue().toString().trim();
      if (orderId) {
        var sheetBBK = getSheet_('Buku Besar Kas');
        if (sheetBBK) {
          var bbkData  = sheetBBK.getDataRange().getValues();
          var bbkH     = bbkData[0];
          var colInv   = bbkH.indexOf('InvoiceID');
          var colSP    = bbkH.indexOf('StatusPembayaran');
          var colDP    = bbkH.indexOf('DP');
          var colGT    = bbkH.indexOf('GrandTotal');
          var colSisa  = bbkH.indexOf('SisaTagihan');
          for (var i = 1; i < bbkData.length; i++) {
            if (colInv !== -1 && bbkData[i][colInv].toString().trim() === orderId) {
              if (colSP   !== -1) sheetBBK.getRange(i+1, colSP   + 1).setValue('Lunas');
              if (colSisa !== -1) sheetBBK.getRange(i+1, colSisa + 1).setValue(0);
              if (colDP   !== -1 && colGT !== -1) sheetBBK.getRange(i+1, colDP + 1).setValue(parseFloat(bbkData[i][colGT]) || 0);
              Logger.log("✓ Sync Lunas ke Buku Besar Kas: " + orderId);
              break;
            }
          }
        }
      }
    }
    return { success: true, message: "Status berhasil diperbarui." };
  } catch (err) {
    Logger.log("updateStatusKatalogPesanan error: " + err.toString());
    return { success: false, message: err.toString() };
  }
}

// ============================================================
// 17. FUNGSI KATALOG — Autentikasi Pelanggan
// ============================================================

function registerKatalogCustomer(regData) {
  return registerCustomer(regData);
}

function loginKatalogCustomer(wa, pin) {
  return loginCustomer(wa, pin);
}

// ============================================================
// 22. FUNGSI KATALOG — Ongkir (RajaOngkir & Lalamove)
// ============================================================

function getProvinces() {
  try {
    var res = UrlFetchApp.fetch(RAJAONGKIR_BASE_URL + "/province", { method: "GET", headers: { "key": RAJAONGKIR_API_KEY }, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return { success: false, message: "API tidak merespons." };
    var json = JSON.parse(res.getContentText());
    return json.rajaongkir && json.rajaongkir.status.code === 200 ? { success: true, data: json.rajaongkir.results } : { success: false, message: json.rajaongkir.status.description };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function getCities(provinceId) {
  try {
    var res = UrlFetchApp.fetch(RAJAONGKIR_BASE_URL + "/city?province=" + provinceId, { method: "GET", headers: { "key": RAJAONGKIR_API_KEY }, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return { success: false, message: "API tidak merespons." };
    var json = JSON.parse(res.getContentText());
    return json.rajaongkir && json.rajaongkir.status.code === 200 ? { success: true, data: json.rajaongkir.results } : { success: false, message: json.rajaongkir.status.description };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function checkOngkir(destinationCityId, totalQty, courier) {
  try {
    var berat   = Math.max(totalQty * BERAT_PER_PCS_GRAM, 1000);
    var payload = "origin=" + ORIGIN_CITY_ID + "&destination=" + destinationCityId + "&weight=" + berat + "&courier=" + courier;
    var res     = UrlFetchApp.fetch(RAJAONGKIR_BASE_URL + "/cost", { method: "POST", payload: payload, headers: { "key": RAJAONGKIR_API_KEY }, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return { success: false, message: "API tidak merespons (HTTP " + res.getResponseCode() + ")." };
    var json    = JSON.parse(res.getContentText());
    if (json.rajaongkir && json.rajaongkir.status.code === 200) {
      var results = json.rajaongkir.results;
      if (!results || results.length === 0) return { success: false, message: "Kurir tidak tersedia." };
      var services = [];
      results[0].costs.forEach(function(c) {
        services.push({ courier: courier.toUpperCase(), service: c.service, description: c.description, cost: c.cost[0].value, etd: c.cost[0].etd ? c.cost[0].etd.replace("HARI","").trim() + " Hari" : "N/A" });
      });
      return { success: true, data: services, beratGram: berat };
    }
    return { success: false, message: json.rajaongkir.status.description };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function getDeliveryEstimate(destinationAddress) {
  if (!destinationAddress || destinationAddress.trim().length < 10) return { success: false, message: "Alamat terlalu singkat." };
  try {
    var geoUrl  = "https://maps.googleapis.com/maps/api/geocode/json?address=" + encodeURIComponent(destinationAddress.trim()) + "&key=" + GOOGLE_MAPS_API_KEY + "&region=id";
    var geoRes  = UrlFetchApp.fetch(geoUrl, { muteHttpExceptions: true });
    if (geoRes.getResponseCode() !== 200) return { success: false, message: "Google Maps tidak merespons." };
    var geoJson = JSON.parse(geoRes.getContentText());
    if (geoJson.status !== "OK" || !geoJson.results || geoJson.results.length === 0) return { success: false, message: "Alamat tidak ditemukan." };

    var destLat = geoJson.results[0].geometry.location.lat;
    var destLng = geoJson.results[0].geometry.location.lng;
    var fmtAddr = geoJson.results[0].formatted_address;

    var matUrl  = "https://maps.googleapis.com/maps/api/distancematrix/json?origins=" + ORIGIN_LAT + "," + ORIGIN_LNG + "&destinations=" + destLat + "," + destLng + "&mode=driving&language=id&key=" + GOOGLE_MAPS_API_KEY;
    var matRes  = UrlFetchApp.fetch(matUrl, { muteHttpExceptions: true });
    if (matRes.getResponseCode() !== 200) return { success: false, message: "Gagal menghitung jarak." };
    var matJson = JSON.parse(matRes.getContentText());
    if (matJson.status !== "OK") return { success: false, message: "Gagal menghitung jarak." };

    var element = matJson.rows[0].elements[0];
    if (element.status !== "OK") return { success: false, message: "Rute tidak ditemukan." };

    var distKm   = Math.round(element.distance.value / 100) / 10;
    var durText  = element.duration.text;
    var fare     = distKm <= LALAMOVE_BASE_KM ? LALAMOVE_BASE_FARE : LALAMOVE_BASE_FARE + Math.ceil(distKm - LALAMOVE_BASE_KM) * LALAMOVE_PER_KM_FARE;
    fare = Math.ceil(fare / 500) * 500;

    return { success: true, distanceKm: distKm, durationText: durText, fare: fare, fareText: "Rp " + fare.toLocaleString("id-ID"), formattedAddress: fmtAddr };
  } catch (err) {
    Logger.log("getDeliveryEstimate error: " + err.toString());
    return { success: false, message: "Estimasi Lalamove tidak tersedia." };
  }
}

// ============================================================
// 23. HELPER FUNCTIONS INTERNAL
// ============================================================

function generateSequentialId(sheetName, prefix) {
  try {
    var today      = new Date();
    var dateStr    = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyyMMdd');
    var fullPrefix = prefix + '-' + dateStr + '-';
    var sheet      = getSheet_(sheetName);
    var lastNum    = 0;
    if (sheet && sheet.getLastRow() > 1) {
      var data   = sheet.getDataRange().getValues();
      var colId  = data[0].indexOf('ID');
      if (colId !== -1) {
        for (var i = 1; i < data.length; i++) {
          var id  = data[i][colId].toString();
          var num = parseInt(id.replace(fullPrefix, ''), 10);
          if (id.indexOf(fullPrefix) === 0 && !isNaN(num) && num > lastNum) lastNum = num;
        }
      }
    }
    return fullPrefix + ('000' + (lastNum + 1)).slice(-3);
  } catch (err) {
    return prefix + '-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd') + '-001';
  }
}

function sanitizeData(data) {
  if (!data || typeof data !== 'object') return {};
  var sanitized = {};
  for (var key in data) {
    if (!data.hasOwnProperty(key)) continue;
    var val = data[key];
    if (val === null || val === undefined) { sanitized[key] = ''; }
    else if (Array.isArray(val)) {
      sanitized[key] = val.map(function(v) {
        return typeof v === 'object' && v !== null ? sanitizeData(v) : typeof v === 'string' ? v.trim().replace(/<[^>]*>/g,'') : v;
      });
    } else if (typeof val === 'string') { sanitized[key] = val.trim().replace(/<[^>]*>/g,''); }
    else if (typeof val === 'object')   { sanitized[key] = sanitizeData(val); }
    else                                { sanitized[key] = val; }
  }
  return sanitized;
}

function formatRupiah(number) {
  try {
    var num        = parseFloat(number) || 0;
    var isNegative = num < 0;
    num = Math.abs(num);
    var formatted = Math.round(num).toString();
    var result = '', count = 0;
    for (var i = formatted.length - 1; i >= 0; i--) {
      result = formatted[i] + result; count++;
      if (count % 3 === 0 && i > 0) result = '.' + result;
    }
    return (isNegative ? '-Rp ' : 'Rp ') + result;
  } catch (e) { return 'Rp 0'; }
}

function deleteRowById_(sheetName, idColumn, idValue, searchFields) {
  try {
    var sheet = getSheet_(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'Data kosong di ' + sheetName };

    var data     = sheet.getDataRange().getValues();
    var headers  = data[0];
    var colId    = headers.indexOf(idColumn);
    if (colId === -1) return { success: false, message: 'Kolom ' + idColumn + ' tidak ditemukan' };

    var targetId      = idValue ? idValue.toString().trim() : '';
    var isPlaceholder = !targetId || targetId.startsWith('M-') || targetId.startsWith('K-') || targetId.startsWith('O-') || targetId.startsWith('PB-') || targetId.startsWith('OPS-');

    if (targetId && !isPlaceholder) {
      for (var i = data.length - 1; i >= 1; i--) {
        if (data[i][colId].toString().trim() === targetId) {
          sheet.deleteRow(i + 1);
          return { success: true, message: 'Data ' + targetId + ' berhasil dihapus dari ' + sheetName };
        }
      }
    }

    if (searchFields && typeof searchFields === 'object') {
      var colTgl  = headers.indexOf('Tanggal');
      var searchT = searchFields.tanggal || '';
      for (var i = data.length - 1; i >= 1; i--) {
        var rowTgl = colTgl !== -1 ? (data[i][colTgl] instanceof Date ? Utilities.formatDate(data[i][colTgl], Session.getScriptTimeZone(), 'yyyy-MM-dd') : data[i][colTgl].toString().trim()) : '';
        if (rowTgl !== searchT) continue;

        var match = false;
        if (sheetName === 'LaporanPembelian') {
          var colP = headers.indexOf('NamaPemasok'); if (colP === -1) colP = headers.indexOf('NamaSupplier');
          var colB = headers.indexOf('NamaBarang'), colTH = headers.indexOf('TotalHarga');
          var sP   = searchFields.namaPemasok||searchFields.pemasok||'', sB = searchFields.namaBarang||searchFields.barang||'';
          var sTH  = parseFloat(searchFields.totalHarga) || 0;
          if (data[i][colP].toString().trim() === sP && data[i][colB].toString().trim() === sB && Math.abs((parseFloat(data[i][colTH])||0) - sTH) < 0.01) match = true;
        } else if (sheetName === 'BebanOperasional') {
          var colK = headers.indexOf('Kategori'), colSK = headers.indexOf('SubKategori'), colN = headers.indexOf('Nominal');
          var sK   = searchFields.kategori||'', sSK = searchFields.subKategori||'', sN = parseFloat(searchFields.nominal)||0;
          if (data[i][colK].toString().trim() === sK && data[i][colSK].toString().trim() === sSK && Math.abs((parseFloat(data[i][colN])||0) - sN) < 0.01) match = true;
        }
        if (match) { sheet.deleteRow(i + 1); return { success: true, message: 'Data berhasil dihapus dari ' + sheetName }; }
      }
    }
    return { success: false, message: 'Data tidak ditemukan di ' + sheetName };
  } catch (err) {
    Logger.log('Error deleteRowById_: ' + err.toString());
    return { success: false, message: err.toString() };
  }
}

function getFilteredByMonth_(sheetName, bulan) {
  try {
    var sheet = getSheet_(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) return [];
    var data    = sheet.getDataRange().getValues();
    var headers = data[0];
    var colTgl  = headers.indexOf('Tanggal');
    var results = [];
    for (var i = 1; i < data.length; i++) {
      if (bulan && bulan.trim() && colTgl !== -1) {
        var tVal = data[i][colTgl];
        var tStr = tVal instanceof Date ? Utilities.formatDate(tVal, Session.getScriptTimeZone(), 'yyyy-MM') : tVal.toString().substring(0,7);
        if (tStr !== bulan.trim()) continue;
      }
      var obj = {};
      for (var j = 0; j < headers.length; j++) {
        var val = data[i][j];
        obj[headers[j]] = val instanceof Date ? Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd') : val;
      }
      results.push(obj);
    }
    return results;
  } catch (err) { return []; }
}

function _getOrCreateFolder(folderName) {
  var folders = DriveApp.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
}

// ============================================================
// 24. SETUP — Jalankan SEKALI dari Apps Script Editor
// ============================================================

/**
 * Jalankan fungsi ini sekali dari menu Run di Apps Script Editor
 * untuk membuat 2 sheet baru katalog secara otomatis.
 */
function setupKatalogSheets() {
  getOrCreatePesananSheet();
  getOrCreatePelangganSheet();
  Logger.log("✓ Setup selesai. 2 sheet katalog siap digunakan.");
  return "Setup selesai. 2 sheet katalog siap digunakan.";
}
