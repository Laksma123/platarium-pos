// --- CONFIGURATION ---
const SHEET_ID = '1zuZF0Hazyj7OeoxBZ9gKO7MZojv2R9LaoqBK8ycf-YQ';

// [2026-01-17] Perbaikan doGet untuk memastikan scriptUrl terkirim dengan benar
function doGet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const configSheet = ss.getSheetByName('Config');
  const rawUrl = configSheet.getRange("B5").getValue(); 
  
  function getDirectImgLink(url) {
    if (!url) return "";
    const match = url.match(/[-\w]{25,}/);
    if (match) {
      let fileId = match[0];
      return "https://lh3.googleusercontent.com/d/" + fileId;
    }
    return url;
  }

  const template = HtmlService.createTemplateFromFile('Index');
  template.bgImageUrl = getDirectImgLink(rawUrl); 
  
  // Memastikan URL yang dikirim adalah URL Web App yang aktif
  let url = ScriptApp.getService().getUrl();
  // Jika URL berakhir dengan /dev, kita biarkan saja agar testing tetap di /dev
  template.scriptUrl = url; 
  
  return template.evaluate()
    .setTitle('Platarium POS Professional')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

// [UPDATE] Modifikasi loginUser untuk mengambil Nama Kasir
function loginUser(password) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const configSheet = ss.getSheetByName('Config');
  const configValues = configSheet.getRange("B1:B8").getValues();
  
  const adminPass = String(configValues[0][0]); 
  const kasirPass = String(configValues[1][0]); 
  const storeName = configValues[2][0];         
  const loginGif  = configValues[4][0];         
  
  // Cuplikan di dalam loginUser (Code.gs)
  const cashierList = [
    configValues[5][0], // Sel B6
    configValues[6][0], // Sel B7
    configValues[7][0]  // Sel B8
  ].filter(name => name !== "" && name !== null); // Hanya ambil yang tidak kosong

  let role = null;
  if (String(password) === adminPass) role = 'ADMIN';
  else if (String(password) === kasirPass) role = 'CASHIER';

  if (role) {
    return { 
      status: 'success', role: role, 
      config: { StoreName: storeName, LoginGif: loginGif, Cashiers: cashierList } 
    };
  }
  return { status: 'error' };
}

// [BARU] Fungsi Validasi PIN Kasir
function verifyCashierPIN(name, pin) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const configSheet = ss.getSheetByName('Config');
  const data = configSheet.getRange("B6:C8").getValues(); // Ambil Nama (B) dan PIN (C)
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === name && String(data[i][1]) === String(pin)) {
      return { status: 'success' };
    }
  }
  return { status: 'error', message: 'PIN Salah!' };
}

// Tambahan fungsi untuk memastikan data dibaca real-time dari Sheet
function getProducts() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Products');
  // Memaksa pengambilan data terbaru (bypass cache)
  SpreadsheetApp.flush(); 
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; // Jika hanya ada header
  
  return data.slice(1).map(row => ({
    id: row[0],
    name: row[1],
    price: Number(row[2]),
    category: row[3],
    img: row[4]
  }));
}

/**
 * [2026-01-18] Fungsi Tambah Produk Lengkap
 * Mendukung ID Manual/Otomatis & Upload Gambar sekaligus
 */
function addProduct(id, name, price, category, base64Image) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Products');
    
    // 1. Logika Penentuan ID
    let finalId = id.trim();
    if (!finalId) {
      const lastId = sheet.getLastRow() > 1 ? sheet.getRange(sheet.getLastRow(), 1).getValue() : "P000";
      finalId = "P" + (parseInt(lastId.replace(/[^\d]/g, "")) + 1).toString().padStart(3, '0');
    }

    // 2. Cek duplikasi ID
    const existingIds = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues().flat();
    if (id.trim() !== "" && existingIds.includes(finalId)) {
      return { status: 'error', message: "ID Produk '" + finalId + "' sudah ada!" };
    }

    // 3. Penanganan Gambar (Opsional)
    let imageUrl = "";
    if (base64Image) {
      // Menggunakan logika yang mirip dengan uploadProductImage namun langsung return URL
      const folderName = "Platarium_Products";
      let folder;
      const folders = DriveApp.getFoldersByName(folderName);
      folder = folders.hasNext() ? folders.next() : DriveApp.getRootFolder().createFolder(folderName);

      const splitData = base64Image.split(',');
      const contentType = splitData[0].substring(5, splitData[0].indexOf(';'));
      const bytes = Utilities.base64Decode(splitData[1]);
      const blob = Utilities.newBlob(bytes, contentType, "img_" + finalId + ".jpg");
      
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      imageUrl = "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w800";
    }

    // 4. Simpan ke Sheet
    sheet.appendRow([finalId, name, price, category, imageUrl]);

    // [BARU] Copy format dari baris di atasnya agar seragam (Font, Alignment, dll)
    const lastRow = sheet.getLastRow();
    if (lastRow > 2) { 
      // Ambil format dari baris sebelumnya (kolom A sampai E)
      const sourceRange = sheet.getRange(lastRow - 1, 1, 1, 5);
      const targetRange = sheet.getRange(lastRow, 1, 1, 5);
      sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    }

    SpreadsheetApp.flush();
    
    return { status: 'success', id: finalId };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

// [2026-01-17] Perbaikan fungsi upload dengan penanganan scope yang lebih kuat
function uploadProductImage(productId, base64Data, fileName) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Products');
    const data = sheet.getDataRange().getValues();
    
    // 1. Memastikan Izin Folder (Memaksa pendeteksian scope full Drive)
    let folder;
    const folderName = "Platarium_Products";
    const folders = DriveApp.getFoldersByName(folderName);
    
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      // Pastikan akun memiliki izin untuk membuat folder di root
      folder = DriveApp.getRootFolder().createFolder(folderName);
    }

    // 2. Decode Data Base64
    const splitData = base64Data.split(',');
    const contentType = splitData[0].substring(5, splitData[0].indexOf(';'));
    const bytes = Utilities.base64Decode(splitData[1]);
    const blob = Utilities.newBlob(bytes, contentType, fileName);
    
    // 3. Simpan File dan Set Publik
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const fileId = file.getId();
    // [PERUBAHAN DISINI] Gunakan link thumbnail dengan parameter ukuran (sz=w800 artinya lebar 800px)
    // Link ini jauh lebih stabil untuk ditampilkan di web app dibanding link /uc?id=
    const directUrl = "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w800";

    // 4. Update Spreadsheet
    let success = false;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(productId)) {
        sheet.getRange(i + 1, 5).setValue(directUrl); 
        success = true;
        break;
      }
    }

    if (!success) throw new Error("ID Produk tidak ditemukan.");

    SpreadsheetApp.flush(); // Memastikan data tertulis sebelum refresh di frontend
    return { status: 'success', newUrl: directUrl };
  } catch (e) {
    console.error("Upload Error: " + e.toString());
    return { status: 'error', message: "Izin Drive ditolak atau masalah teknis: " + e.message };
  }
}

function deleteProduct(productId) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Products');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(productId)) {
        // Hapus file gambar di Drive jika ada
        let currentImgUrl = data[i][4];
        if (currentImgUrl && currentImgUrl.includes("id=")) {
          try {
            let fileId = currentImgUrl.split("id=")[1];
            DriveApp.getFileById(fileId).setTrashed(true);
          } catch(err) { console.log("Gagal hapus file Drive: " + err); }
        }
        sheet.deleteRow(i + 1); // Hapus baris di sheet
        return { status: 'success' };
      }
    }
    return { status: 'error', message: 'ID tidak ditemukan' };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

// Fungsi untuk menghapus gambar produk saja
function deleteProductImage(productId) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Products');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(productId)) {
      let url = data[i][4];
      if (url && url.includes("id=")) {
        try {
          let fileId = url.split("id=")[1];
          DriveApp.getFileById(fileId).setTrashed(true);
        } catch(e) {}
      }
      sheet.getRange(i + 1, 5).setValue(""); // Kosongkan kolom gambar
      return { status: 'success' };
    }
  }
}

// Gunakan fungsi include jika ingin memisahkan CSS/JS ke file lain
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// [2026-01-17] Perbaikan Fungsi Pancingan Otorisasi - Timpa fungsi pancingOtorisasi yang lama
function pancingOtorisasi() {
  // Memanggil DriveApp secara eksplisit untuk memaksa scope https://www.googleapis.com/auth/drive
  try {
    const root = DriveApp.getRootFolder();
    const folderName = "Platarium_Products";
    const folders = DriveApp.getFoldersByName(folderName);
    
    if (folders.hasNext()) {
      Logger.log("Folder sudah ada: " + folders.next().getName());
    } else {
      // Baris ini krusial untuk memicu izin 'createFolder'
      const newFolder = DriveApp.createFolder(folderName);
      Logger.log("Folder berhasil dibuat untuk pancingan izin.");
      // Hapus folder pancingan jika ingin bersih
      newFolder.setTrashed(true);
    }
  } catch (e) {
    Logger.log("Error Otorisasi: " + e.toString());
  }
}

// Fungsi pembantu untuk dipanggil dari Frontend jika scriptUrl di atas gagal
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * [2026-01-19] Perbaikan Fungsi Absensi - Simpan Foto Bukti ke Drive
 */
function recordAttendance(staffName, type, base64Image) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('Attendance');
    
    // 1. Penanganan Foto Bukti (Disimpan ke folder Staff_Attendance_Proofs)
    let proofUrl = "";
    if (base64Image) {
      const folderName = "Attendance_Proofs";
      let folder = DriveApp.getFoldersByName(folderName).hasNext() ? 
                   DriveApp.getFoldersByName(folderName).next() : 
                   DriveApp.getRootFolder().createFolder(folderName);
      
      const blob = Utilities.newBlob(Utilities.base64Decode(base64Image.split(',')[1]), "image/jpeg", "PROOF_" + staffName + "_" + Date.now() + ".jpg");
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      proofUrl = "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w200";
    }

    const now = new Date();
    const dateStr = Utilities.formatDate(now, "GMT+7", "dd/MM/yyyy");
    const timeStr = Utilities.formatDate(now, "GMT+7", "HH:mm:ss");
    
    const data = sheet.getDataRange().getValues();
    let rowToUpdate = -1;

    for (let i = 1; i < data.length; i++) {
      const rowDate = (data[i][0] instanceof Date) ? Utilities.formatDate(data[i][0], "GMT+7", "dd/MM/yyyy") : "";
      if (data[i][1] === staffName && rowDate === dateStr) {
        rowToUpdate = i + 1;
        break;
      }
    }

    if (type === 'IN') {
      if (rowToUpdate !== -1) return { status: 'error', message: 'Sudah Check-In!' };
      // Simpan Foto di Kolom F (Index 6)
      sheet.appendRow([now, staffName, timeStr, '-', 'Hadir', proofUrl]);
    } else {
      if (rowToUpdate === -1) return { status: 'error', message: 'Belum Check-In!' };
      sheet.getRange(rowToUpdate, 4).setValue(timeStr);
      // Update foto saat checkout jika ingin mengganti bukti visual terbaru
      sheet.getRange(rowToUpdate, 6).setValue(proofUrl);
    }

    SpreadsheetApp.flush();
    return { status: 'success', time: timeStr };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

/**
 * [2026-01-19] Perbaikan Pengambilan Log Absensi - Anti Stuck
 */
function getAttendanceLogs() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('Attendance');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    // Kita olah data menjadi objek bersih agar tidak terjadi error mapping di frontend
    return data.slice(1).reverse().map(function(row) {
      // Perbaikan: Pastikan jam dikirim sebagai string agar tidak jadi tahun 1899
      return {
        tgl: (row[0] instanceof Date) ? Utilities.formatDate(row[0], "GMT+7", "dd/MM/yyyy") : String(row[0]),
        nama: String(row[1]),
        in: String(row[2]),
        out: String(row[3]),
        status: String(row[4]),
        img: String(row[5] || "")
      };
    });
  } catch (e) {
    console.error("Error Logs: " + e.toString());
    return [];
  }
}

/**
 * [2026-01-19] Fungsi Simpan Foto Master Staf (Daftarkan Kasir)
 */
function registerStaffFace(staffName, base64Data) {
  try {
    const folderName = "Staff_Faces_Master";
    let folder = DriveApp.getFoldersByName(folderName).hasNext() ? 
                 DriveApp.getFoldersByName(folderName).next() : 
                 DriveApp.getRootFolder().createFolder(folderName);

    // Hapus foto lama jika ada (Sistem Update/Ganti Foto)
    const oldFiles = folder.getFilesByName("MASTER_" + staffName + ".jpg");
    while (oldFiles.hasNext()) { oldFiles.next().setTrashed(true); }

    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data.split(',')[1]), "image/jpeg", "MASTER_" + staffName + ".jpg");
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const faceUrl = "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w400";
    
    // Update ke Sheet Config kolom D (D6:D8)
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const configSheet = ss.getSheetByName('Config');
    const staffNames = configSheet.getRange("B6:B8").getValues().flat();
    
    for (let i = 0; i < staffNames.length; i++) {
      if (staffNames[i] === staffName) {
        configSheet.getRange(6 + i, 4).setValue(faceUrl);
        break;
      }
    }
    SpreadsheetApp.flush();
    return { status: 'success', url: faceUrl };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

// Ambil URL master wajah dari Sheet Config
function getStaffMasterFaces() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Config');
  const names = sheet.getRange("B6:B8").getValues().flat();
  const urls = sheet.getRange("D6:D8").getValues().flat();
  
  let map = {};
  names.forEach((name, i) => {
    if (name) map[name] = urls[i] || "";
  });
  return map;
}

/**
 * Menghapus Foto Master Kasir dari Sheet Config
 */
function deleteMasterFaceRecord(staffName) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const configSheet = ss.getSheetByName('Config');
    const staffNames = configSheet.getRange("B6:B8").getValues().flat();
    
    for (let i = 0; i < staffNames.length; i++) {
      if (staffNames[i] === staffName) {
        configSheet.getRange(6 + i, 4).setValue(""); // Kosongkan kolom D
        return { status: 'success' };
      }
    }
    return { status: 'error', message: 'Nama tidak ditemukan' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

/**
 * [2026-01-19] Fungsi untuk menghapus baris absensi tertentu
 * Berdasarkan Nama dan Tanggal (Mencegah salah hapus)
 */
function deleteAttendanceRecord(nama, tgl) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Attendance');
    if (!sheet) throw new Error("Sheet 'Attendance' tidak ditemukan.");

    const data = sheet.getDataRange().getValues();
    let rowDeleted = false;
    
    // Iterasi mundur untuk menjaga indeks baris saat menghapus
    for (let i = data.length - 1; i >= 1; i--) {
      const rowDate = (data[i][0] instanceof Date) ?
        Utilities.formatDate(data[i][0], "GMT+7", "dd/MM/yyyy") : String(data[i][0]);
      
      if (String(data[i][1]) === String(nama) && rowDate === String(tgl)) {
        sheet.deleteRow(i + 1);
        rowDeleted = true;
      }
    }

    if (rowDeleted) {
      SpreadsheetApp.flush();
      return { status: 'success' };
    } else {
      return { status: 'error', message: 'Data tidak ditemukan di database.' };
    }
  } catch (e) {
    console.error("Delete Attendance Error: " + e.toString());
    return { status: 'error', message: e.toString() };
  }
}

/**
 * Mengambil semua data transaksi beserta detail itemnya
 */
function getAllTransactionData() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const transSheet = ss.getSheetByName('Transactions');
    const detailSheet = ss.getSheetByName('TransactionDetails');
    
    if (!transSheet || !detailSheet) return { transactions: [], summary: {} };

    const transData = transSheet.getDataRange().getValues();
    const detailData = detailSheet.getDataRange().getValues();

    // Mapping details berdasarkan Order ID
    const detailsMap = {};
    for (let i = 1; i < detailData.length; i++) {
      const oid = detailData[i][0];
      if (!detailsMap[oid]) detailsMap[oid] = [];
      detailsMap[oid].push({
        name: detailData[i][2],
        qty: detailData[i][3],
        price: detailData[i][4],
        subtotal: detailData[i][5]
      });
    }

    // Olah data transaksi utama
    const transactions = transData.slice(1).reverse().map(row => {
      const orderId = row[1];
      return {
        timestamp: row[0] instanceof Date ? Utilities.formatDate(row[0], "GMT+7", "dd/MM/yyyy HH:mm") : String(row[0]),
        orderId: orderId,
        total: row[2],
        method: row[3],
        cashier: row[6],
        items: detailsMap[orderId] || []
      };
    });

    return transactions;
  } catch (e) {
    console.error("Error Get All Data: " + e.toString());
    return [];
  }
}

/**
 * Menghapus transaksi dari kedua sheet berdasarkan Order ID
 */
function deleteTransactionRecord(orderId) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const transSheet = ss.getSheetByName('Transactions');
    const detailSheet = ss.getSheetByName('TransactionDetails');

    // Hapus di Sheet Transactions
    const tData = transSheet.getDataRange().getValues();
    for (let i = tData.length - 1; i >= 1; i--) {
      if (tData[i][1] === orderId) transSheet.deleteRow(i + 1);
    }

    // Hapus di Sheet TransactionDetails
    const dData = detailSheet.getDataRange().getValues();
    for (let i = dData.length - 1; i >= 1; i--) {
      if (dData[i][0] === orderId) detailSheet.deleteRow(i + 1);
    }

    SpreadsheetApp.flush();
    return { status: 'success' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

/**
 * [STRICT & CACHED] Mengambil harga pasar murni dari Gemini AI.
 * Dilengkapi CacheService untuk menghemat kuota API (20 RPD).
 */
function getMaterialsData() {
  const CACHE_KEY = "market_prices_data";
  const cache = CacheService.getScriptCache();
  
  try {
    // 1. Cek apakah ada data di Cache (berlaku selama 1 jam)
    const cached = cache.get(CACHE_KEY);
    if (cached) return JSON.parse(cached);

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Materials');
    if (!sheet) return { error: "SHEET_MISSING" };

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    const dataRange = sheet.getRange(2, 2, lastRow - 1, 2).getValues();
    const materialList = dataRange
      .filter(row => row[0] !== "")
      .map(row => ({ name: row[0], unit: row[1] || "unit" }));

    if (materialList.length === 0) return [];

    const GEMINI_API_KEY = "AIzaSyDrOjeNni8JIdAgyUrgNAy1AtBGdVJjgAI"; 
    // Menggunakan endpoint model terbaru sesuai pilihanmu
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-lite:generateContent?key=${GEMINI_API_KEY}`;

    const prompt = `Berikan estimasi harga pasar terbaru (IDR) di Indonesia untuk:
    ${materialList.map(m => `- ${m.name} (per ${m.unit})`).join("\n")}
    
    SYARAT:
    1. Konversi harga ke satuan ${materialList.map(m => m.unit).join(", ")} secara akurat.
    2. Respon WAJIB JSON murni: {"Nama": harga_angka}`;

    const payload = { "contents": [{ "parts": [{ "text": prompt }] }] };
    const options = {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };

    const response = UrlFetchApp.fetch(url, options);
    const resJson = JSON.parse(response.getContentText());
    
    let livePrices = {};
    const aiText = resJson.candidates[0].content.parts[0].text;
    livePrices = JSON.parse(aiText.replace(/```json|```/g, ""));

    const result = materialList.map(m => ({
      name: m.name,
      price: livePrices[m.name] || 0,
      unit: m.unit
    }));

    // 2. Simpan hasil ke Cache selama 3600 detik (1 jam)
    cache.put(CACHE_KEY, JSON.stringify(result), 3600);

    return result;

  } catch (e) {
    return { error: "AI_OFFLINE" };
  }
}

function clearMarketCache() {
  CacheService.getScriptCache().remove("market_prices_data");
  return { status: 'success' };
}

/**
 * Mengambil ID Transaksi urut berikutnya (ORD-XXX)
 */
function getNextOrderId() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const transSheet = ss.getSheetByName('Transactions');
  if (!transSheet) return "ORD-1";
  
  const lastRow = transSheet.getLastRow();
  if (lastRow < 2) return "ORD-1";
  
  const lastId = transSheet.getRange(lastRow, 2).getValue(); // Kolom B
  const lastNum = parseInt(String(lastId).replace("ORD-", "")) || 0;
  return "ORD-" + (lastNum + 1);
}

/**
 * [FIXED] Simpan transaksi & Picu pengurangan stok otomatis
 */
function saveTransaction(data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let transSheet = ss.getSheetByName('Transactions');
    if (!transSheet) {
      transSheet = ss.insertSheet('Transactions');
      transSheet.appendRow(['Timestamp', 'Order ID', 'Total', 'Metode', 'Bayar', 'Kembalian', 'Kasir']);
      transSheet.getRange("1:1").setFontWeight("bold").setBackground("#f3f3f3");
    }

    let detailSheet = ss.getSheetByName('TransactionDetails');
    if (!detailSheet) {
      detailSheet = ss.insertSheet('TransactionDetails');
      detailSheet.appendRow(['Order ID', 'Product ID', 'Name', 'Qty', 'Price', 'Subtotal']);
      detailSheet.getRange("1:1").setFontWeight("bold").setBackground("#f3f3f3");
    }

    const timestamp = new Date();
    const orderId = data.manualId || ("ORD-" + timestamp.getTime());

    // Simpan Header
    transSheet.appendRow([timestamp, orderId, data.total, data.method, data.cash, data.change, data.cashier]);

    // Simpan Detail
    data.items.forEach(item => {
      detailSheet.appendRow([orderId, item.id, item.name, item.qty, item.price, (item.price * item.qty)]);
    });

    // --- LOGIKA PENGURANGAN STOK (SANGAT KRUSIAL) ---
    deductInventory(data.items);
    
    SpreadsheetApp.flush();
    return { status: 'success', orderId: orderId };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

/**
 * [OPTIMIZED] Logika Pengurangan Stok Berdasarkan Recipes
 * Memastikan ID P001 di-match secara akurat dengan Recipes & Materials
 */
function deductInventory(items) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const recipeSheet = ss.getSheetByName('Recipes');
  const materialSheet = ss.getSheetByName('Materials');
  
  if (!recipeSheet || !materialSheet) return;

  const recipes = recipeSheet.getDataRange().getValues();
  const materials = materialSheet.getDataRange().getValues();
  
  // 1. Map Lokasi Stok di Sheet Materials agar update cepat
  let materialMap = {};
  for (let i = 1; i < materials.length; i++) {
    let mId = String(materials[i][0]).trim(); // Kolom A: ID Material
    if (mId) {
      materialMap[mId] = {
        row: i + 1,
        currentStock: Number(materials[i][3]) || 0 // Kolom D: CurrentStock
      };
    }
  }

  // 2. Iterasi Item yang terjual
  items.forEach(item => {
    const soldProdId = String(item.id).trim(); // Contoh: "P001"
    
    // Cari resep untuk Product ID ini
    for (let j = 1; j < recipes.length; j++) {
      const recipeProdId = String(recipes[j][0]).trim(); // Kolom A di Recipes
      const matIdInRecipe = String(recipes[j][1]).trim(); // Kolom B di Recipes
      const qtyNeeded = Number(recipes[j][2]) || 0;      // Kolom C di Recipes

      if (recipeProdId === soldProdId) {
        const totalDeduction = qtyNeeded * item.qty;
        
        // Jika material ada di gudang, kurangi stoknya
        if (materialMap[matIdInRecipe]) {
          materialMap[matIdInRecipe].currentStock -= totalDeduction;
          
          // Update langsung ke cell untuk real-time sync
          materialSheet.getRange(materialMap[matIdInRecipe].row, 4)
                       .setValue(materialMap[matIdInRecipe].currentStock);
        }
      }
    }
  });
}

/**
 * Mengambil semua data dari sheet Materials untuk Tab Stock
 */
function getMaterialsStock() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Materials');
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  return data.slice(1).map(row => ({
    id: row[0],
    name: row[1],
    unit: row[2],
    stock: Number(row[3]) || 0
  }));
}

/**
 * Update stock manual dari UI (Icon Pensil)
 */
function updateMaterialStockManual(id, newStock) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Materials');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.getRange(i + 1, 4).setValue(newStock); // Kolom D
        SpreadsheetApp.flush();
        return { status: 'success' };
      }
    }
    return { status: 'error', message: 'ID Material tidak ditemukan' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

// Letakkan setelah kode terakhir Anda di Code.gs

// [BARU] Router Utama untuk Menerima Request dari GitHub Pages (Frontend)
function doPost(e) {
  try {
    const request = JSON.parse(e.postData.contents);
    const action = request.action;
    const args = request.args || [];
    
    let result;
    
    // Mendistribusikan request ke fungsi-fungsi yang sudah ada
    switch(action) {
      case 'loginUser': result = loginUser(args[0]); break;
      case 'verifyCashierPIN': result = verifyCashierPIN(args[0], args[1]); break;
      case 'getProducts': result = getProducts(); break;
      case 'addProduct': result = addProduct(args[0], args[1], args[2], args[3], args[4]); break;
      case 'uploadProductImage': result = uploadProductImage(args[0], args[1], args[2]); break;
      case 'deleteProduct': result = deleteProduct(args[0]); break;
      case 'deleteProductImage': result = deleteProductImage(args[0]); break;
      case 'recordAttendance': result = recordAttendance(args[0], args[1], args[2]); break;
      case 'getAttendanceLogs': result = getAttendanceLogs(); break;
      case 'registerStaffFace': result = registerStaffFace(args[0], args[1]); break;
      case 'getStaffMasterFaces': result = getStaffMasterFaces(); break;
      case 'deleteMasterFaceRecord': result = deleteMasterFaceRecord(args[0]); break;
      case 'deleteAttendanceRecord': result = deleteAttendanceRecord(args[0], args[1]); break;
      case 'getAllTransactionData': result = getAllTransactionData(); break;
      case 'deleteTransactionRecord': result = deleteTransactionRecord(args[0]); break;
      case 'getMaterialsData': result = getMaterialsData(); break;
      case 'clearMarketCache': result = clearMarketCache(); break;
      case 'getNextOrderId': result = getNextOrderId(); break;
      case 'saveTransaction': result = saveTransaction(args[0]); break;
      case 'getMaterialsStock': result = getMaterialsStock(); break;
      case 'updateMaterialStockManual': result = updateMaterialStockManual(args[0], args[1]); break;
      default: 
        throw new Error("Fungsi '" + action + "' tidak terdaftar di router backend.");
    }
    
    return ContentService.createTextOutput(JSON.stringify(result))
          .setMimeType(ContentService.MimeType.JSON);
          
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
          .setMimeType(ContentService.MimeType.JSON);
  }
}