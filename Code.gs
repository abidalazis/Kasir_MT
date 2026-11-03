// Code.gs (Google Apps Script POS Backend - Improved Version v2)

var SHEET_DB = "Database";
var SHEET_TX = "Transactions";
var SHEET_DEBT = "Hutang";
var SHEET_SUPPLIER = "Supplier_Purchases";

/**
 * Serves the HTML page as a web app.
 */
function doGet() {
  try {
    initSheets(); // Called ONCE here, not in every function
  } catch (e) {
    Logger.log("Inisialisasi sheet gagal: " + e.toString());
  }
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('POS Toko Pertanian')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Helper to get script timezone safely.
 */
function getTimeZoneSafe() {
  var tz = "GMT+7";
  try {
    tz = Session.getScriptTimeZone() || SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || "GMT+7";
  } catch (e) {}
  return tz;
}

/**
 * Helper to get or create sheet with default headers.
 */
function getOrCreateSheet(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}

/**
 * Verifies and initializes crucial sheet structures.
 */
function initSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  getOrCreateSheet(SHEET_TX, ["ID Transaksi", "Tanggal", "Pelanggan", "Total Belanja", "Bayar", "Kembali", "Sisa Hutang", "Status", "Detail Barang", "Items JSON"]);
  getOrCreateSheet(SHEET_DEBT, ["ID Hutang", "ID Transaksi", "Tanggal", "Pelanggan", "Jumlah Hutang", "Sisa Hutang", "Status", "Riwayat Pembayaran"]);
  getOrCreateSheet(SHEET_SUPPLIER, ["ID Pembelian", "Tanggal", "Supplier", "Kode Produk", "Nama Produk", "Qty", "Harga Beli", "Total", "Keterangan"]);

  var dbSheet = ss.getSheetByName(SHEET_DB);
  if (dbSheet) {
    var lastCol = dbSheet.getLastColumn();
    var lastRow = dbSheet.getLastRow();
    if (lastRow > 0 && lastCol === 6) {
      dbSheet.getRange(1, 7).setValue("Stok");
      if (lastRow > 1) {
        var zeros = [];
        for (var i = 2; i <= lastRow; i++) zeros.push([0]);
        dbSheet.getRange(2, 7, lastRow - 1, 1).setValues(zeros);
      }
    }
  } else {
    getOrCreateSheet(SHEET_DB, ["No", "Kode", "Nama", "Modal", "Jual", "Kategori", "Stok"]);
  }
}

/**
 * Retrieve list of all products from the Database sheet.
 */
function getProducts() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_DB);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    var values = sheet.getRange(2, 1, lastRow - 1, Math.max(7, sheet.getLastColumn())).getValues();
    var products = [];
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      products.push({
        index: i + 2,
        no: row[0],
        kode: String(row[1] || '').trim(),
        nama: String(row[2] || '').trim(),
        modal: Number(row[3]) || 0,
        jual: Number(row[4]) || 0,
        kategori: String(row[5] || '').trim(),
        stok: Number(row[6]) || 0
      });
    }
    return products;
  } catch (e) {
    Logger.log("Error getProducts: " + e.toString());
    throw new Error("Gagal mengambil data produk: " + e.message);
  }
}

/**
 * Retrieve active debt entries from the Hutang sheet.
 */
function getDebts() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_DEBT);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    var values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    var debts = [];
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      debts.push({
        rowIndex: i + 2,
        idHutang: String(row[0]),
        idTransaksi: String(row[1]),
        tanggal: String(row[2]),
        pelanggan: String(row[3]),
        jumlahHutang: Number(row[4]) || 0,
        sisaHutang: Number(row[5]) || 0,
        status: String(row[6]),
        riwayat: String(row[7] || '')
      });
    }
    return debts;
  } catch (e) {
    Logger.log("Error getDebts: " + e.toString());
    throw new Error("Gagal mengambil data hutang: " + e.message);
  }
}

/**
 * Retrieve transaction history with optional date range filter.
 * startDate and endDate are strings in "yyyy-MM-dd" format.
 * No hard limit on number of rows returned.
 */
function getTransactions(startDate, endDate) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_TX);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    var tz = getTimeZoneSafe();
    var values = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    var transactions = [];

    for (var i = values.length - 1; i >= 0; i--) {
      var row = values[i];
      var dateVal = row[1];
      var dateStr = (dateVal instanceof Date)
        ? Utilities.formatDate(dateVal, tz, "yyyy-MM-dd")
        : String(dateVal).substring(0, 10);

      if (startDate && dateStr < startDate) continue;
      if (endDate && dateStr > endDate) continue;

      transactions.push({
        idTransaksi: String(row[0]),
        tanggal: String(row[1]),
        pelanggan: String(row[2]),
        total: Number(row[3]) || 0,
        bayar: Number(row[4]) || 0,
        kembali: Number(row[5]) || 0,
        sisaHutang: Number(row[6]) || 0,
        status: String(row[7]),
        detail: String(row[8]),
        itemsJson: String(row[9])
      });
    }
    return transactions;
  } catch (e) {
    Logger.log("Error getTransactions: " + e.toString());
    throw new Error("Gagal mengambil data transaksi: " + e.message);
  }
}

/**
 * Get sales + profit report data for a given date range.
 */
function getReportData(startDate, endDate) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var txSheet = ss.getSheetByName(SHEET_TX);
    var lastRow = txSheet.getLastRow();
    var tz = getTimeZoneSafe();

    var totalPenjualan = 0, totalModal = 0, totalTransaksi = 0, totalHutangBaru = 0;

    if (lastRow >= 2) {
      var values = txSheet.getRange(2, 1, lastRow - 1, 10).getValues();
      for (var i = 0; i < values.length; i++) {
        var row = values[i];
        var dateVal = row[1];
        var dateStr = (dateVal instanceof Date)
          ? Utilities.formatDate(dateVal, tz, "yyyy-MM-dd")
          : String(dateVal).substring(0, 10);

        if (startDate && dateStr < startDate) continue;
        if (endDate && dateStr > endDate) continue;

        totalPenjualan += Number(row[3]) || 0;
        totalTransaksi++;
        if (String(row[7]) === 'HUTANG') totalHutangBaru += Number(row[6]) || 0;

        // Calculate modal from items JSON
        var itemsJson = String(row[9] || '');
        if (itemsJson && itemsJson !== 'undefined') {
          try {
            var items = JSON.parse(itemsJson);
            for (var j = 0; j < items.length; j++) {
              totalModal += (Number(items[j].modal) || 0) * (Number(items[j].qty) || 0);
            }
          } catch (e2) {}
        }
      }
    }

    return {
      totalPenjualan: totalPenjualan,
      totalModal: totalModal,
      labaBersih: totalPenjualan - totalModal,
      totalTransaksi: totalTransaksi,
      totalHutangBaru: totalHutangBaru
    };
  } catch (e) {
    Logger.log("Error getReportData: " + e.toString());
    return { totalPenjualan: 0, totalModal: 0, labaBersih: 0, totalTransaksi: 0, totalHutangBaru: 0 };
  }
}

/**
 * Fetches quick summary stats for dashboard.
 */
function getSummaryData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tz = getTimeZoneSafe();
    var todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

    var txSheet = ss.getSheetByName(SHEET_TX);
    var txRows = txSheet.getLastRow();
    var totalPenjualanHariIni = 0;
    if (txRows >= 2) {
      var txValues = txSheet.getRange(2, 2, txRows - 1, 3).getValues();
      for (var i = 0; i < txValues.length; i++) {
        var dv = txValues[i][0];
        var ds = (dv instanceof Date) ? Utilities.formatDate(dv, tz, "yyyy-MM-dd") : String(dv).split('T')[0];
        if (ds.substring(0, 10) === todayStr) totalPenjualanHariIni += Number(txValues[i][2]) || 0;
      }
    }

    var debtSheet = ss.getSheetByName(SHEET_DEBT);
    var debtRows = debtSheet.getLastRow();
    var totalHutangAktif = 0;
    if (debtRows >= 2) {
      var debtValues = debtSheet.getRange(2, 6, debtRows - 1, 2).getValues();
      for (var i = 0; i < debtValues.length; i++) {
        if (String(debtValues[i][1]) === "BELUM LUNAS") totalHutangAktif += Number(debtValues[i][0]) || 0;
      }
    }

    return { totalSalesToday: totalPenjualanHariIni, totalActiveDebt: totalHutangAktif };
  } catch (e) {
    Logger.log("Error getSummaryData: " + e.toString());
    return { totalSalesToday: 0, totalActiveDebt: 0 };
  }
}

/**
 * Retrieve supplier purchase history with optional date range.
 */
function getSupplierPurchases(startDate, endDate) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_SUPPLIER);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    var tz = getTimeZoneSafe();
    var values = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
    var purchases = [];

    for (var i = values.length - 1; i >= 0; i--) {
      var row = values[i];
      var dateVal = row[1];
      var dateStr = (dateVal instanceof Date)
        ? Utilities.formatDate(dateVal, tz, "yyyy-MM-dd")
        : String(dateVal).substring(0, 10);

      if (startDate && dateStr < startDate) continue;
      if (endDate && dateStr > endDate) continue;

      purchases.push({
        idPembelian: String(row[0]),
        tanggal: String(row[1]),
        supplier: String(row[2]),
        kode: String(row[3]),
        nama: String(row[4]),
        qty: Number(row[5]) || 0,
        hargaBeli: Number(row[6]) || 0,
        total: Number(row[7]) || 0,
        keterangan: String(row[8] || '')
      });
    }
    return purchases;
  } catch (e) {
    Logger.log("Error getSupplierPurchases: " + e.toString());
    throw new Error("Gagal mengambil data pembelian supplier: " + e.message);
  }
}

/**
 * Register sale transaction (JSON payload).
 * Now validates stock BEFORE saving.
 */
function saveSale(txDataJson) {
  try {
    var txData = JSON.parse(txDataJson);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // --- PRE-CHECK: Validate stock before writing ---
    var dbSheet = ss.getSheetByName(SHEET_DB);
    var lastDbRow = dbSheet.getLastRow();
    if (lastDbRow >= 2) {
      var dbCodes = dbSheet.getRange(2, 2, lastDbRow - 1, 1).getValues();
      var dbStoks = dbSheet.getRange(2, 7, lastDbRow - 1, 1).getValues();
      var codeStokMap = {};
      for (var r = 0; r < dbCodes.length; r++) {
        var c = String(dbCodes[r][0] || '').trim();
        if (c) codeStokMap[c] = Number(dbStoks[r][0]) || 0;
      }
      for (var k = 0; k < txData.items.length; k++) {
        var item = txData.items[k];
        var tc = String(item.kode).trim();
        if (codeStokMap.hasOwnProperty(tc)) {
          var avail = codeStokMap[tc];
          if (avail <= 0) {
            return { success: false, message: "Stok " + item.nama + " sudah HABIS (stok = 0)." };
          }
          if (item.qty > avail) {
            return { success: false, message: "Stok " + item.nama + " tidak cukup. Tersedia: " + avail + ", diminta: " + item.qty };
          }
        }
      }
    }

    var tz = getTimeZoneSafe();
    var date = new Date();
    var dateStr = Utilities.formatDate(date, tz, "yyyy-MM-dd HH:mm:ss");
    var timePart = Utilities.formatDate(date, tz, "yyMMdd-HHmmss");
    var randomPart = Math.floor(100 + Math.random() * 900);
    var txId = "TR-" + timePart + "-" + randomPart;

    var summaryParts = [];
    for (var i = 0; i < txData.items.length; i++) {
      var it = txData.items[i];
      summaryParts.push(it.nama + " (" + it.qty + ")");
    }

    // 1. Write to Transactions
    var txSheet = ss.getSheetByName(SHEET_TX);
    txSheet.appendRow([
      txId, dateStr, txData.pelanggan || "Umum",
      Number(txData.total) || 0, Number(txData.bayar) || 0,
      Number(txData.kembali) || 0, Number(txData.sisaHutang) || 0,
      txData.status, summaryParts.join(", "), JSON.stringify(txData.items)
    ]);

    // 2. Write to debts if unpaid
    if (txData.status === "HUTANG" && txData.sisaHutang > 0) {
      var debtSheet = ss.getSheetByName(SHEET_DEBT);
      debtSheet.appendRow([
        "HT-" + timePart + "-" + randomPart, txId, dateStr, txData.pelanggan,
        Number(txData.sisaHutang) || 0, Number(txData.sisaHutang) || 0, "BELUM LUNAS", ""
      ]);
    }

    // 3. Update stock in Database sheet (batch read/write)
    if (lastDbRow >= 2) {
      var codeRange2 = dbSheet.getRange(2, 2, lastDbRow - 1, 1).getValues();
      var stockRange2 = dbSheet.getRange(2, 7, lastDbRow - 1, 1).getValues();
      var codeToIdx = {};
      for (var r = 0; r < codeRange2.length; r++) {
        var cd = String(codeRange2[r][0] || '').trim();
        if (cd) codeToIdx[cd] = r;
      }
      var updated = false;
      for (var k = 0; k < txData.items.length; k++) {
        var si = txData.items[k];
        var sc = String(si.kode).trim();
        if (codeToIdx.hasOwnProperty(sc)) {
          var lr = codeToIdx[sc];
          stockRange2[lr][0] = Math.max(0, (Number(stockRange2[lr][0]) || 0) - Number(si.qty));
          updated = true;
        }
      }
      if (updated) dbSheet.getRange(2, 7, lastDbRow - 1, 1).setValues(stockRange2);
    }

    return { success: true, txId: txId, dateStr: dateStr };
  } catch (e) {
    Logger.log("Error saveSale: " + e.toString());
    return { success: false, message: "Gagal menyimpan transaksi: " + e.toString() };
  }
}

/**
 * Save supplier purchase details (JSON payload).
 * Only updates modal price if a new price is provided.
 */
function saveSupplierInput(purchDataJson) {
  try {
    var purchData = JSON.parse(purchDataJson);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var tz = getTimeZoneSafe();
    var date = new Date();
    var dateStr = Utilities.formatDate(date, tz, "yyyy-MM-dd HH:mm:ss");
    var purchId = "SL-" + Utilities.formatDate(date, tz, "yyMMdd-HHmmss");

    var qty = Number(purchData.qty) || 0;
    var price = Number(purchData.hargaBeli) || 0;

    var supplierSheet = ss.getSheetByName(SHEET_SUPPLIER);
    supplierSheet.appendRow([purchId, dateStr, purchData.supplier, purchData.kode,
      purchData.nama, qty, price, qty * price, purchData.keterangan || ""]);

    // Update stock and optionally modal price in Database
    var dbSheet = ss.getSheetByName(SHEET_DB);
    var lastRow = dbSheet.getLastRow();
    if (lastRow >= 2) {
      var codeValues = dbSheet.getRange(2, 2, lastRow - 1, 1).getValues();
      var targetCode = String(purchData.kode).trim();
      var foundIndex = -1;
      for (var r = 0; r < codeValues.length; r++) {
        if (String(codeValues[r][0]).trim() === targetCode) { foundIndex = r + 2; break; }
      }
      if (foundIndex !== -1) {
        var currentStok = Number(dbSheet.getRange(foundIndex, 7).getValue()) || 0;
        dbSheet.getRange(foundIndex, 7).setValue(currentStok + qty);
        // Only overwrite modal price if a new price > 0 is given
        if (price > 0) dbSheet.getRange(foundIndex, 4).setValue(price);
      }
    }

    return { success: true };
  } catch (e) {
    Logger.log("Error saveSupplierInput: " + e.toString());
    return { success: false, message: "Gagal mencatat barang supplier: " + e.toString() };
  }
}

/**
 * Handle debt payments.
 */
function payDebt(debtId, paymentAmount, note) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var debtSheet = ss.getSheetByName(SHEET_DEBT);
    var lastRow = debtSheet.getLastRow();
    if (lastRow < 2) return { success: false, message: "Belum ada catatan hutang." };

    var values = debtSheet.getRange(2, 1, lastRow - 1, 8).getValues();
    var foundRow = -1, txId = "", oldSisa = 0;
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][0]).trim() === String(debtId).trim()) {
        foundRow = i + 2; txId = String(values[i][1]); oldSisa = Number(values[i][5]) || 0; break;
      }
    }
    if (foundRow === -1) return { success: false, message: "Hutang ID " + debtId + " tidak ditemukan." };

    paymentAmount = Number(paymentAmount) || 0;
    var newSisa = Math.max(0, oldSisa - paymentAmount);
    var status = newSisa <= 0 ? "LUNAS" : "BELUM LUNAS";

    var tz = getTimeZoneSafe();
    var dateStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    var existingLog = String(debtSheet.getRange(foundRow, 8).getValue() || '');
    var newLog = (existingLog ? existingLog + "; " : "") +
      dateStr + ": Rp " + paymentAmount.toLocaleString('id-ID') + " (" + (note || 'Bayar Cicilan') + ")";

    debtSheet.getRange(foundRow, 6).setValue(newSisa);
    debtSheet.getRange(foundRow, 7).setValue(status);
    debtSheet.getRange(foundRow, 8).setValue(newLog);

    // Also update the parent transaction record
    var txSheet = ss.getSheetByName(SHEET_TX);
    var txLastRow = txSheet.getLastRow();
    if (txLastRow >= 2) {
      var txValues = txSheet.getRange(2, 1, txLastRow - 1, 1).getValues();
      for (var k = 0; k < txValues.length; k++) {
        if (String(txValues[k][0]).trim() === txId.trim()) {
          var tfr = k + 2;
          txSheet.getRange(tfr, 5).setValue((Number(txSheet.getRange(tfr, 5).getValue()) || 0) + paymentAmount);
          txSheet.getRange(tfr, 7).setValue(Math.max(0, (Number(txSheet.getRange(tfr, 7).getValue()) || 0) - paymentAmount));
          txSheet.getRange(tfr, 8).setValue(newSisa <= 0 ? "LUNAS" : "HUTANG");
          break;
        }
      }
    }

    return { success: true, sisaHutang: newSisa };
  } catch (e) {
    Logger.log("Error payDebt: " + e.toString());
    return { success: false, message: "Gagal membayar hutang: " + e.toString() };
  }
}

/**
 * Update a single product row in the Database sheet.
 */
function updateProduct(productJson) {
  try {
    var product = JSON.parse(productJson);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_DB);
    var rowIndex = Number(product.rowIndex);
    if (!rowIndex || rowIndex < 2) return { success: false, message: "Row index tidak valid." };

    sheet.getRange(rowIndex, 2).setValue(String(product.kode).trim());
    sheet.getRange(rowIndex, 3).setValue(String(product.nama).trim());
    sheet.getRange(rowIndex, 4).setValue(Number(product.modal) || 0);
    sheet.getRange(rowIndex, 5).setValue(Number(product.jual) || 0);
    sheet.getRange(rowIndex, 6).setValue(String(product.kategori || '').trim());
    sheet.getRange(rowIndex, 7).setValue(Number(product.stok) || 0);

    return { success: true };
  } catch (e) {
    Logger.log("Error updateProduct: " + e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * Add a new product to Database sheet.
 */
function addProduct(productJson) {
  try {
    var product = JSON.parse(productJson);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_DB);
    var lastRow = sheet.getLastRow();

    if (lastRow >= 2) {
      var existingCodes = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
      for (var i = 0; i < existingCodes.length; i++) {
        if (String(existingCodes[i][0]).trim().toUpperCase() === String(product.kode).trim().toUpperCase()) {
          return { success: false, message: "Kode produk " + product.kode + " sudah ada!" };
        }
      }
    }

    var newNo = lastRow > 1 ? lastRow : 1;
    sheet.appendRow([
      newNo,
      String(product.kode).trim().toUpperCase(),
      String(product.nama).trim().toUpperCase(),
      Number(product.modal) || 0,
      Number(product.jual) || 0,
      String(product.kategori || '').trim().toUpperCase(),
      Number(product.stok) || 0
    ]);

    return { success: true };
  } catch (e) {
    Logger.log("Error addProduct: " + e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * Delete a product row by its kode, then renumber the No column.
 */
function deleteProduct(kode) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_DB);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: "Tidak ada produk." };

    var codeValues = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    var foundRow = -1;
    for (var i = 0; i < codeValues.length; i++) {
      if (String(codeValues[i][0]).trim() === String(kode).trim()) {
        foundRow = i + 2; break;
      }
    }
    if (foundRow === -1) return { success: false, message: "Produk kode " + kode + " tidak ditemukan." };

    sheet.deleteRow(foundRow);

    // Renumber the No column after deletion
    var newLastRow = sheet.getLastRow();
    if (newLastRow >= 2) {
      var noValues = [];
      for (var n = 1; n <= newLastRow - 1; n++) noValues.push([n]);
      sheet.getRange(2, 1, newLastRow - 1, 1).setValues(noValues);
    }

    return { success: true };
  } catch (e) {
    Logger.log("Error deleteProduct: " + e.toString());
    return { success: false, message: e.toString() };
  }
}
