/**
 * === BACKEND (Code.gs) ===
 * Script untuk Google Apps Script
 */

// Konstanta Nama Sheet
const SHEET_ITEMS = 'MasterBarang';
const SHEET_INBOUND = 'BarangMasuk';
const SHEET_OUTBOUND = 'BarangKeluar';

function doGet(e) {
  // Jika ada request GET dengan parameter ?action=getData, kembalikan JSON
  if (e && e.parameter && e.parameter.action === 'getData') {
    const data = getAllData();
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Default (jika URL diakses langsung di browser): Tampilkan UI HTML
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('SI MBAH ICPH - Inventory')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Menerima request POST untuk menyimpan data (REST API dari Vercel/luar)
 */
function doPost(e) {
  try {
    // Membaca payload request (dikirim sebagai text/plain untuk bypass CORS Preflight)
    const dataStr = e.postData.contents;
    saveDataState(dataStr);
    
    return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Setup Database awal. 
 * Jalankan fungsi ini sekali dari editor Apps Script sebelum deploy.
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Setup Master Barang
  let sheetItems = ss.getSheetByName(SHEET_ITEMS);
  if (!sheetItems) {
    sheetItems = ss.insertSheet(SHEET_ITEMS);
    sheetItems.appendRow(['id', 'nama', 'kategori', 'unit', 'satuan', 'hargaBeli', 'stokAwal', 'minStok', 'createdAt']);
    sheetItems.getRange('A1:I1').setFontWeight('bold').setBackground('#0d6efd').setFontColor('white');
  }

  // Setup Barang Masuk
  let sheetInbound = ss.getSheetByName(SHEET_INBOUND);
  if (!sheetInbound) {
    sheetInbound = ss.insertSheet(SHEET_INBOUND);
    sheetInbound.appendRow(['id', 'itemId', 'jumlah', 'tanggal', 'keterangan']);
    sheetInbound.getRange('A1:E1').setFontWeight('bold').setBackground('#198754').setFontColor('white');
  }

  // Setup Barang Keluar
  let sheetOutbound = ss.getSheetByName(SHEET_OUTBOUND);
  if (!sheetOutbound) {
    sheetOutbound = ss.insertSheet(SHEET_OUTBOUND);
    sheetOutbound.appendRow(['id', 'itemId', 'subUnit', 'jumlah', 'tanggal', 'keterangan']);
    sheetOutbound.getRange('A1:F1').setFontWeight('bold').setBackground('#dc3545').setFontColor('white');
  }
}

/**
 * Mengambil semua data dari spreadsheet untuk frontend
 */
function getAllData() {
  return {
    items: getSheetDataAsObject(SHEET_ITEMS),
    inbound: getSheetDataAsObject(SHEET_INBOUND),
    outbound: getSheetDataAsObject(SHEET_OUTBOUND)
  };
}

/**
 * Menyimpan seluruh state data dari frontend (Simulasi DB Sederhana)
 * Pada produksi skala besar, sebaiknya pisahkan per sheet/tindakan (saveItem, saveInbound, dll)
 * Namun untuk MVP SPA ini, kita simpan seluruh state (overwrite) untuk sinkronisasi mudah.
 */
function saveDataState(dataStr) {
  const data = JSON.parse(dataStr);
  
  writeArrayToSheet(SHEET_ITEMS, data.items, ['id', 'nama', 'kategori', 'unit', 'satuan', 'hargaBeli', 'stokAwal', 'minStok', 'createdAt']);
  writeArrayToSheet(SHEET_INBOUND, data.inbound, ['id', 'itemId', 'jumlah', 'tanggal', 'keterangan']);
  writeArrayToSheet(SHEET_OUTBOUND, data.outbound, ['id', 'itemId', 'subUnit', 'jumlah', 'tanggal', 'keterangan']);
  
  return true;
}

// --- HELPER FUNCTIONS ---

function getSheetDataAsObject(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const rows = data.slice(1);
  
  return rows.map(row => {
    let obj = {};
    headers.forEach((header, i) => {
      obj[header] = row[i];
    });
    return obj;
  });
}

function writeArrayToSheet(sheetName, arrayData, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;
  
  // Clear existing data except header
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }
  
  if (arrayData && arrayData.length > 0) {
    const rows = arrayData.map(obj => {
      return headers.map(header => obj[header] || '');
    });
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
}