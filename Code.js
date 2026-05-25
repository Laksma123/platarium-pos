// --- CONFIGURATION ---
const SHEET_ID = '1zuZF0Hazyj7OeoxBZ9gKO7MZojv2R9LaoqBK8ycf-YQ';

function doGet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const configSheet = ss.getSheetByName('Config');
  const logoRaw = configSheet.getRange("B4").getValue();
  const wallRaw = configSheet.getRange("B5").getValue(); 
  
  function getDirectImgLink(url) {
    if (!url) return "";
    const match = url.match(/[-\w]{25,}/);
    if (match) {

      return "https://drive.google.com/thumbnail?id=" + match[0] + "&sz=w2560";
    }
    return url;
  }

  const template = HtmlService.createTemplateFromFile('Index');
  template.logoUrl = getDirectImgLink(logoRaw);
  template.bgImageUrl = getDirectImgLink(wallRaw);
  template.scriptUrl = ScriptApp.getService().getUrl();
  
  return template.evaluate()
    .setTitle('Platarium POS Professional')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

function loginUser(password) {
  // --- KILL SWITCH SUBSCRIPTION ---
  const EXPIRY_DATE = new Date("2028-05-21T23:59:59"); // Format YYYY-MM-DD
  if (new Date() > EXPIRY_DATE) {
    return { 
      status: 'error', 
      message: 'Masa aktif langganan POS telah habis. Hubungi Developer.' 
    };
  }
  // --------------------------------
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const configSheet = ss.getSheetByName('Config');
  // Ambil data B1 hingga C25
  const configValues = configSheet.getRange("B1:C25").getValues();
  
  const ownerPass = String(configValues[0][0]);   // B1 (Owner)
  const monitorPass = String(configValues[1][0]); // B2 (Order Monitor / Kitchen / Barista)
  const storeName = configValues[2][0];           // B3
  const loginGif  = configValues[4][0];           // B5
  
  // Cek Status Multi Owner & Tax
  const props = PropertiesService.getScriptProperties();
  const isMultiOwnerOn = props.getProperty('MULTI_OWNER') === 'true';
  const isTaxOn = props.getProperty('TAX_ACTIVE') === 'true';
  const taxPercent = Number(props.getProperty('TAX_PERCENT')) || 0;

  // 1. Cek Login Owner (Hanya jika Multi Owner OFF)
  if (!isMultiOwnerOn && String(password) === ownerPass) {
    return { 
      status: 'success', 
      role: 'ADMIN', 
      name: 'Owner',
      isPassDone: true, // Owner tunggal di-bypass
      config: { StoreName: storeName, LoginGif: loginGif, TaxActive: isTaxOn, TaxPercent: taxPercent } 
    };
  }

  // 1.5 Cek Login Global Order Monitor (Cell B2)
  if (password !== "" && String(password) === monitorPass) {
      return { 
        status: 'success', 
        role: 'MONITOR', 
        name: 'Order Monitor',
        config: { StoreName: storeName, LoginGif: loginGif, TaxActive: isTaxOn, TaxPercent: taxPercent } 
      };
  }

  // 2. Cek Staff
  let foundName = null;
  for (let i = 5; i <= 24; i++) {
    const staffName = configValues[i][0];
    const staffPin = configValues[i][1];
    if (staffName && String(staffPin) === String(password)) {
      foundName = staffName;
      break;
    }
  }

  if (foundName) {
    const staffSheet = ss.getSheetByName('Staff');
    let role = 'CASHIER'; 
    let isPassDone = true; // Default bypass untuk role selain Admin
    
    if (staffSheet) {
      const staffData = staffSheet.getDataRange().getValues();
      for (let i = 1; i < staffData.length; i++) {
        if (staffData[i][1] === foundName) {
          role = staffData[i][2].toUpperCase();
          if (role === 'ADMIN') {
              // Cek Kolom F (Indeks ke 5)
              isPassDone = (staffData[i][5] === 'Done'); 
          }
          break;
        }
      }
    }

    if (['ADMIN', 'CASHIER', 'KASIR', 'KITCHEN', 'BARISTA', 'MONITOR'].includes(role)) {
      let finalRole = role === 'ADMIN' ? 'ADMIN' : (role === 'KASIR' ? 'CASHIER' : role);
      return { 
        status: 'success', 
        role: finalRole, 
        name: foundName,
        isPassDone: isPassDone, // Kembalikan status ke frontend
        config: { StoreName: storeName, LoginGif: loginGif, TaxActive: isTaxOn, TaxPercent: taxPercent } 
      };
    }
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
    posKerja: row[4], // Kolom E Sekarang adalah Pos Kerja
    img: row[5]       // Kolom F Sekarang adalah Image Thumbnail
  }));
}

/**
 * [2026-01-18] Fungsi Tambah Produk Lengkap (Dengan Recipe Engine)
 * Mendukung ID Manual/Otomatis, Upload Gambar, dan Penulisan Resep
 */
function addProduct(id, name, price, category, posKerja, base64Image, recipeBasket) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Products');
    
    // 1. Logika Penentuan ID
    let finalId = id.trim();
    if (!finalId) {
      const lastId = sheet.getLastRow() > 1 ? sheet.getRange(sheet.getLastRow(), 1).getValue() : "P000";
      finalId = "P" + (parseInt(lastId.replace(/[^\d]/g, "")) + 1).toString().padStart(3, '0');
    }

    // 2. Cek duplikasi ID Produk
    const existingIds = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues().flat();
    if (id.trim() !== "" && existingIds.includes(finalId)) {
      return { status: 'error', message: "ID Produk '" + finalId + "' sudah ada!" };
    }

    // 3. Penanganan Gambar (Opsional)
    let imageUrl = "";
    if (base64Image) {
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

    // 4. Simpan ke Sheet Products
    sheet.appendRow([finalId, name, price, category, posKerja, imageUrl]);

    // Copy format dari baris di atasnya agar seragam (Font, Alignment, dll)
    const lastRow = sheet.getLastRow();
    if (lastRow > 2) { 
      const sourceRange = sheet.getRange(lastRow - 1, 1, 1, 6); // Ubah bentangan ke kolom 6
      const targetRange = sheet.getRange(lastRow, 1, 1, 6);
      sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    }

    // 5. PROSES RESEP (Jika ada data dikirim dari UI)
    if (recipeBasket && recipeBasket.length > 0) {
      let recipeSheet = ss.getSheetByName('Recipes');
      if (!recipeSheet) {
        recipeSheet = ss.insertSheet('Recipes');
        recipeSheet.appendRow(['Product ID', 'Material ID', 'Qty', 'Unit']);
        recipeSheet.getRange("1:1").setFontWeight("bold").setBackground("#f3f3f3");
      }

      const materialSheet = ss.getSheetByName('Materials');
      let existingMaterials = [];
      if (materialSheet && materialSheet.getLastRow() > 1) {
        existingMaterials = materialSheet.getRange(2, 1, materialSheet.getLastRow() - 1, 1).getValues().flat();
      }

      recipeBasket.forEach(m => {
        let matId = m.id;
        
        // A. Jika Bahan Baku benar-benar baru
        if (m.isNew) {
          // Format Otomatis: M-NAMABAHAN (Maks 8 Karakter, tanpa spasi, huruf besar)
          let baseMatId = "M-" + m.name.toUpperCase().replace(/[^A-Z0-9]/g, '').substring(0, 8);
          matId = baseMatId;
          let counter = 1;
          
          // Cek kalau ID M-NAMABAHAN sudah ada, tambahkan angka (contoh: M-GULA, M-GULA1)
          while (existingMaterials.includes(matId)) {
            matId = baseMatId + counter;
            counter++;
          }
          
          if (materialSheet) {
            // Tulis bahan baru ke Sheet Materials dengan Stok awal = 0
            materialSheet.appendRow([matId, m.name, m.unit, 0]); 
            existingMaterials.push(matId); // Update referensi lokal agar tidak dobel ID
          }
        }
        
        // B. Tulis ke Sheet Recipes
        recipeSheet.appendRow([finalId, matId, m.qty, m.unit]);
      });
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
        sheet.getRange(i + 1, 6).setValue(directUrl); // Simpan ke Kolom F (ke-6)
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
        let currentImgUrl = data[i][5]; // Baca gambar dari Kolom F (Indeks 5)
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
      let url = data[i][5]; // Baca gambar dari Kolom F (Indeks 5)
      if (url && url.includes("id=")) {
        try {
          let fileId = url.split("id=")[1];
          DriveApp.getFileById(fileId).setTrashed(true);
        } catch(e) {}
      }
      sheet.getRange(i + 1, 6).setValue(""); // Kosongkan Kolom F (ke-6)
      return { status: 'success' };
    }
  }
}

function updateProductPrice(productId, newPrice) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Products');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(productId)) {
        sheet.getRange(i + 1, 3).setValue(newPrice); // Kolom C (Indeks 2) adalah Price
        SpreadsheetApp.flush();
        return { status: 'success' };
      }
    }
    return { status: 'error', message: 'ID Produk tidak ditemukan' };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

function uploadConfigImage(type, base64Data, fileName) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Config');
    
    let folder;
    const folderName = "Platarium_Config";
    const folders = DriveApp.getFoldersByName(folderName);
    
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.getRootFolder().createFolder(folderName);
    }

    const splitData = base64Data.split(',');
    const contentType = splitData[0].substring(5, splitData[0].indexOf(';'));
    const bytes = Utilities.base64Decode(splitData[1]);
    const blob = Utilities.newBlob(bytes, contentType, fileName);
    
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const directUrl = "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w2560";
    
    // Baris 4 untuk LOGO (B4), Baris 5 untuk WALLPAPER (B5)
    const row = type === 'LOGO' ? 4 : 5;
    
    // Hapus gambar lama jika ada agar Drive tidak penuh
    const oldUrl = sheet.getRange(row, 2).getValue();
    if(oldUrl && oldUrl.includes("id=")) {
        try {
            const oldId = oldUrl.match(/[-\w]{25,}/)[0];
            DriveApp.getFileById(oldId).setTrashed(true);
        } catch(e) { console.log("Gagal hapus file lama: " + e); }
    }
    
    sheet.getRange(row, 2).setValue(directUrl);
    SpreadsheetApp.flush(); 
    return { status: 'success', newUrl: directUrl };
  } catch (e) {
    return { status: 'error', message: "Gagal upload: " + e.message };
  }
}

// [TAMBAHKAN] Fungsi Baru: Letakkan di bawah uploadConfigImage
function deleteConfigImage(type) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Config');
    const row = type === 'LOGO' ? 4 : 5;
    const currentUrl = sheet.getRange(row, 2).getValue();
    
    if (currentUrl && currentUrl.includes("id=")) {
      try {
        const fileId = currentUrl.match(/[-\w]{25,}/)[0];
        DriveApp.getFileById(fileId).setTrashed(true);
      } catch(e) {}
    }
    sheet.getRange(row, 2).setValue("");
    SpreadsheetApp.flush();
    return { status: 'success' };
  } catch(e) {
    return { status: 'error', message: e.toString() };
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

function normalizeDateStr(cellValue) {
  if (cellValue instanceof Date) {
    return Utilities.formatDate(cellValue, "GMT+7", "dd/MM/yyyy");
  }
  var raw = String(cellValue).trim();
  if (!raw) return "";
  
  var datePart = raw.split(" ")[0]; // Ambil bagian tanggal saja jika ada jamnya
  
  // Jika format yyyy-MM-dd
  if (datePart.includes("-")) {
    var parts = datePart.split("-");
    if (parts[0].length === 4) {
      return parts[2].padStart(2, '0') + "/" + parts[1].padStart(2, '0') + "/" + parts[0];
    }
  }
  
  // Jika format d/m/yyyy atau m/d/yyyy
  if (datePart.includes("/")) {
    var parts = datePart.split("/");
    var d = parts[0].padStart(2, '0');
    var m = parts[1].padStart(2, '0');
    var y = parts[2];
    if (d.length === 4) { // yyyy/mm/dd
      return parts[2].padStart(2, '0') + "/" + parts[1].padStart(2, '0') + "/" + d;
    }
    return d + "/" + m + "/" + y;
  }
  return datePart;
}

function recordAttendance(staffName, type, base64Image) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('Attendance');
    
    // Validasi eksistensi sheet secara eksplisit
    if (!sheet) return { status: 'error', message: 'Sheet bernama "Attendance" tidak ditemukan di Google Sheet Anda!' };
    
    let proofUrl = "";
    if (base64Image && base64Image.includes(',')) {
      try {
        const folderName = "Attendance_Proofs";
        let folder;
        const folders = DriveApp.getFoldersByName(folderName);
        if (folders.hasNext()) {
          folder = folders.next();
        } else {
          folder = DriveApp.getRootFolder().createFolder(folderName);
        }
        const blob = Utilities.newBlob(Utilities.base64Decode(base64Image.split(',')[1]), "image/jpeg", "PROOF_" + staffName + "_" + Date.now() + ".jpg");
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        proofUrl = "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w200";
      } catch (driveErr) {
        console.error("Gagal menyimpan foto ke Drive: " + driveErr.toString());
      }
    }

    const now = new Date();
    const dateStr = normalizeDateStr(now);
    const timeStr = Utilities.formatDate(now, "GMT+7", "HH:mm:ss");
    
    const data = sheet.getDataRange().getValues();
    let rowToUpdate = -1;
    const targetStaff = String(staffName).trim().toLowerCase();
    
    for (let i = 1; i < data.length; i++) {
      const rowName = String(data[i][1]).trim().toLowerCase();
      const rowDate = normalizeDateStr(data[i][0]);

      if (rowName === targetStaff && rowDate === dateStr) {
        rowToUpdate = i + 1;
      }
    }

    if (type === 'IN') {
      if (rowToUpdate !== -1) return { status: 'error', message: 'Staf sudah melakukan Check-In hari ini!' };
      sheet.appendRow([now, staffName, timeStr, '-', 'Hadir', proofUrl]);
    } else {
      if (rowToUpdate === -1) return { status: 'error', message: 'Staf belum melakukan Check-In hari ini!' };
      sheet.getRange(rowToUpdate, 4).setValue(timeStr);
      sheet.getRange(rowToUpdate, 6).setValue(proofUrl);
    }

    SpreadsheetApp.flush();
    return { status: 'success', time: timeStr };
  } catch (e) {
    return { status: 'error', message: 'Crash Server: ' + e.toString() };
  }
}

function getAttendanceLogs() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('Attendance');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    return data.slice(1).reverse().map(function(row) {
      // FIX EKSTRAKSI TANGGAL UNTUK UI
      let dateStr = "";
      if (row[0] instanceof Date) {
        dateStr = Utilities.formatDate(row[0], "GMT+7", "dd/MM/yyyy");
      } else {
        dateStr = String(row[0]).split(" ")[0];
      }

      return {
        tgl: dateStr,
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

function registerStaffFace(staffName, base64Data, descriptorJson) {
  try {
    const folderName = "Staff_Faces_Master";
    let folder = DriveApp.getFoldersByName(folderName).hasNext() ? 
                 DriveApp.getFoldersByName(folderName).next() : 
                 DriveApp.getRootFolder().createFolder(folderName);

    const oldFiles = folder.getFilesByName("MASTER_" + staffName + ".jpg");
    while (oldFiles.hasNext()) { oldFiles.next().setTrashed(true); }

    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data.split(',')[1]), "image/jpeg", "MASTER_" + staffName + ".jpg");
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const faceUrl = "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w400";

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const configSheet = ss.getSheetByName('Config');
    const staffNames = configSheet.getRange("B6:B8").getValues().flat();
    
    for (let i = 0; i < staffNames.length; i++) {
      if (staffNames[i] === staffName) {
        configSheet.getRange(6 + i, 4).setValue(faceUrl); // Simpan URL Foto di Kolom D
        configSheet.getRange(6 + i, 5).setValue(descriptorJson || ""); // Simpan Descriptor di Kolom E
        break;
      }
    }
    SpreadsheetApp.flush();
    return { status: 'success', url: faceUrl };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

function getStaffMasterFaces() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Config');
  const names = sheet.getRange("B6:B8").getValues().flat();
  const urls = sheet.getRange("D6:D8").getValues().flat();
  const descriptors = sheet.getRange("E6:E8").getValues().flat(); 
  
  let map = {};
  names.forEach((name, i) => {
    if (name) {
      map[name] = {
        url: urls[i] || "",
        descriptor: descriptors[i] || ""
      };
    }
  });
  return map;
}

function deleteMasterFaceRecord(staffName) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const configSheet = ss.getSheetByName('Config');
    const staffNames = configSheet.getRange("B6:B8").getValues().flat();
    
    for (let i = 0; i < staffNames.length; i++) {
      if (staffNames[i] === staffName) {
        // Ambil URL Master saat ini
        const currentUrl = configSheet.getRange(6 + i, 4).getValue();
        
        // Hapus file di Google Drive agar tidak menumpuk
        if (currentUrl && currentUrl.includes("id=")) {
          try {
            const fileId = currentUrl.split("id=")[1].split("&")[0];
            DriveApp.getFileById(fileId).setTrashed(true);
          } catch (err) {
            console.log("Gagal hapus foto master di Drive: " + err);
          }
        }

        configSheet.getRange(6 + i, 4).setValue(""); // Kosongkan sel D (URL Foto)
        configSheet.getRange(6 + i, 5).setValue(""); // Kosongkan sel E (Descriptor Biometrik)
        
        return { status: 'success' };
      }
    }
    return { status: 'error', message: 'Nama tidak ditemukan' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

function deleteAttendanceRecord(nama, tgl) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Attendance');
    if (!sheet) return { status: 'error', message: 'Sheet bernama "Attendance" tidak ditemukan!' };

    const data = sheet.getDataRange().getValues();
    let rowDeleted = false;
    const targetNama = String(nama).trim().toLowerCase();
    const targetTgl = normalizeDateStr(tgl);
    
    for (let i = data.length - 1; i >= 1; i--) {
      const rowName = String(data[i][1]).trim().toLowerCase();
      const rowDate = normalizeDateStr(data[i][0]);
      
      if (rowName === targetNama && rowDate === targetTgl) {
        const imgUrl = String(data[i][5]);
        if (imgUrl && imgUrl.includes("id=")) {
          try {
            const fileId = imgUrl.split("id=")[1].split("&")[0];
            DriveApp.getFileById(fileId).setTrashed(true);
          } catch (err) {}
        }
        sheet.deleteRow(i + 1);
        rowDeleted = true;
      }
    }

    if (rowDeleted) {
      SpreadsheetApp.flush();
      return { status: 'success' };
    } else {
      return { status: 'error', message: 'Data gagal dihapus karena tidak cocok di database. Nama: "' + nama + '", Tgl format: "' + targetTgl + '"' };
    }
  } catch (e) {
    return { status: 'error', message: 'Crash Server: ' + e.toString() };
  }
}

function getAllTransactionData() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const transSheet = ss.getSheetByName('Transactions');
    const detailSheet = ss.getSheetByName('TransactionDetails');
    
    if (!transSheet || !detailSheet) return { transactions: [], summary: {} };

    const transData = transSheet.getDataRange().getValues();
    const detailData = detailSheet.getDataRange().getValues();

    const detailsMap = {};
    for (let i = 1; i < detailData.length; i++) {
      const oid = detailData[i][0];
      if (!detailsMap[oid]) detailsMap[oid] = [];
      detailsMap[oid].push({
        id: detailData[i][1], // Pastikan ID ditarik untuk filtering
        name: detailData[i][2],
        qty: detailData[i][3],
        price: detailData[i][4],
        subtotal: detailData[i][5]
      });
    }

    const transactions = transData.slice(1).reverse().map(row => {
      const orderId = row[1];
      return {
        timestamp: row[0] instanceof Date ? Utilities.formatDate(row[0], "GMT+7", "dd/MM/yyyy HH:mm") : String(row[0]),
        rawTime: row[0] instanceof Date ? row[0].getTime() : new Date(row[0]).getTime(),
        orderId: orderId,
        total: row[2],
        method: row[3],
        cashier: row[6],
        status: row[7] || 'Normal',
        kitchenStatus: row[8] || 'QUEUE',
        barStatus: row[9] || 'QUEUE',
        priceNotes: row[10] || '', // Mengambil data catatan perubahan harga Add-On dari Kolom K (Indeks ke-10)
        addonNotes: row[11] || '', // [BARU] Mengambil data catatan addon dari Kolom L
        items: detailsMap[orderId] || []
      };
    });
    return transactions;
  } catch (e) {
    console.error("Error school Get All Data: " + e.toString());
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
 * [STRICT & CACHED] Mengambil harga pasar dari Gemini AI (Mode Gratis).
 * Fitur Search Grounding dimatikan agar tidak terkena error Billing/Quota.
 */
function getMaterialsData() {
  const CACHE_KEY = "market_prices_standard";
  const cache = CacheService.getScriptCache();
  
  try {
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

    const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!GEMINI_API_KEY) return { error: "API_KEY_MISSING" };
    
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-flash-lite:generateContent?key=${GEMINI_API_KEY}`;

    const now = new Date();
    const months = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    const currentMonth = months[now.getMonth()];
    const currentYear = now.getFullYear();

    const promptText = `Berikan estimasi harga pasar e-commerce (Tokopedia/Shopee) BULAN ${currentMonth} TAHUN ${currentYear} di Indonesia untuk daftar bahan berikut:
    ${materialList.map(m => `- ${m.name} (satuan: per ${m.unit})`).join("\n")}
    
    ATURAN KETAT:
    1. JANGAN LAKUKAN PEMBULATAN SAMA SEKALI. Tulis angka aslinya secara spesifik (contoh: 19450, 49900).
    2. Jika satu bahan memiliki beberapa jenis spesifik di pasaran (misal pencarian "Biji Kopi" memunculkan harga spesifik untuk Arabica, Robusta, dll), berikan MAKSIMAL 3 jenis varian beserta harganya.
    3. Jika tidak ada harga yang masuk akal, kembalikan array kosong [].
    4. Respon WAJIB berupa JSON object murni tanpa markdown \`\`\`json. Format persis seperti ini:
    {
      "Nama Bahan Sesuai Daftar": [
         {"variant": "Nama Jenis/Varian (atau ulangi nama bahan jika tidak ada varian)", "price": 19450}
      ]
    }`;

    // Payload dikembalikan ke mode standar tanpa "tools" googleSearch
    const payload = { 
      "contents": [{ "parts": [{ "text": promptText }] }]
    };
    
    const options = {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    const resJson = JSON.parse(responseText);
    
    if (responseCode !== 200) {
      console.error("Gemini API Error: " + responseText);
      return { error: resJson.error ? resJson.error.message : "AI_API_ERROR" };
    }
    
    const aiText = resJson.candidates[0].content.parts[0].text;
    const cleanJsonText = aiText.replace(/```json/g, "").replace(/```/g, "").trim();
    const livePrices = JSON.parse(cleanJsonText);

    const result = materialList.map(m => ({
      name: m.name,
      unit: m.unit,
      variants: livePrices[m.name] || []
    }));

    cache.put(CACHE_KEY, JSON.stringify(result), 10800);

    return result;

  } catch (e) {
    console.error("AI Parsing/Execution Error: " + e.toString());
    return { error: "AI_PARSE_ERROR" };
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
  const lastNum = parseInt(String(lastId).replace(/^(ORD|QRIS)-/, "")) || 0;
  return "ORD-" + (lastNum + 1);
}

function saveTransaction(data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let transSheet = ss.getSheetByName('Transactions');
    if (!transSheet) {
      transSheet = ss.insertSheet('Transactions');
      transSheet.appendRow(['Timestamp', 'Order ID', 'Total', 'Metode', 'Bayar', 'Kembalian', 'Kasir', 'Status', 'Kitchen Status', 'Bar Status', 'Price Notes', 'Addon Notes']);
      transSheet.getRange("1:1").setFontWeight("bold").setBackground("#f3f3f3");
    }

    let detailSheet = ss.getSheetByName('TransactionDetails');
    if (!detailSheet) {
      detailSheet = ss.insertSheet('TransactionDetails');
      detailSheet.appendRow(['Order ID', 'Product ID', 'Name', 'Qty', 'Price', 'Subtotal']);
      detailSheet.getRange("1:1").setFontWeight("bold").setBackground("#f3f3f3");
    }

    const timestamp = new Date();
    const orderPrefix = data.method === 'QRIS' ? 'QRIS-' : 'ORD-';
    const orderId = data.manualId || (orderPrefix + timestamp.getTime());
    
    // Default Status Orderan Baru = QUEUE untuk Kitchen dan Bar
    // Menambahkan catatan perubahan harga (data.priceNotes) ke elemen ke-11 (Kolom K) dan addon notes ke Kolom L
    transSheet.appendRow([timestamp, orderId, data.total, data.method, data.cash, data.change, data.cashier, 'Normal', 'QUEUE', 'QUEUE', data.priceNotes || '', data.addonNotes || '']);
    
    data.items.forEach(item => {
      detailSheet.appendRow([orderId, item.id, item.name, item.qty, item.price, (item.price * item.qty)]);
    });
    
    deductInventory(data.items);
    SpreadsheetApp.flush();
    return { status: 'success', orderId: orderId };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

function updateMonitorStatusBackend(orderId, type, newStatus) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const transSheet = ss.getSheetByName('Transactions');
    if (!transSheet) return { status: 'error', message: 'Sheet Transactions tidak ditemukan' };
    
    const data = transSheet.getDataRange().getValues();
    // Iterasi dari bawah ke atas untuk efisiensi, karena order baru ada di bawah
    for(let i = data.length - 1; i >= 1; i--) {
      if(data[i][1] === orderId) { // Kolom B adalah Order ID
        // Kolom I untuk Kitchen (indeks 9), Kolom J untuk Bar (indeks 10)
        const col = type === 'KITCHEN' ? 9 : 10; 
        transSheet.getRange(i + 1, col).setValue(newStatus);
        SpreadsheetApp.flush(); // Memastikan perubahan langsung ditulis
        return { status: 'success' };
      }
    }
    return { status: 'error', message: 'Order tidak ditemukan' };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

function deductInventory(items) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const recipeSheet = ss.getSheetByName('Recipes');
  const materialSheet = ss.getSheetByName('Materials');
  
  if (!recipeSheet || !materialSheet) return;

  const recipes = recipeSheet.getDataRange().getValues();
  const materials = materialSheet.getDataRange().getValues();
  
  const shotSheet = ss.getSheetByName('Shots');
  let shotsMap = {};
  if (shotSheet) {
    const shotData = shotSheet.getDataRange().getValues();
    for(let i = 1; i < shotData.length; i++) {
      shotsMap[shotData[i][0]] = Number(shotData[i][1]);
    }
  }
  
  let materialMap = {};
  for (let i = 1; i < materials.length; i++) {
    let mId = String(materials[i][0]).trim();
    if (mId) {
      materialMap[mId] = {
        row: i + 1,
        unit: String(materials[i][2]).trim().toLowerCase(),
        currentStock: Number(materials[i][3]) || 0 
      };
    }
  }

  items.forEach(item => {
    const soldProdId = String(item.id).trim(); 
    
    for (let j = 1; j < recipes.length; j++) {
      const recipeProdId = String(recipes[j][0]).trim(); 
      const matIdInRecipe = String(recipes[j][1]).trim(); 
      let qtyNeeded = Number(recipes[j][2]) || 0;      
      let recipeUnit = String(recipes[j][3]).trim().toLowerCase();

      if (recipeProdId === soldProdId && materialMap[matIdInRecipe]) {
        let baseUnit = materialMap[matIdInRecipe].unit;

        if (recipeUnit === 'shot') {
           const shotGram = shotsMap[matIdInRecipe];
           if (!shotGram) continue; // PENTING: Jangan potong stok jika shot belum diatur sama sekali
           
           qtyNeeded = qtyNeeded * shotGram;
           recipeUnit = (baseUnit === 'l' || baseUnit === 'ml') ? 'ml' : 'gr';
        }

        if ((recipeUnit === 'gr' && baseUnit === 'kg') || (recipeUnit === 'ml' && baseUnit === 'l')) {
            qtyNeeded = qtyNeeded / 1000;
        } else if ((recipeUnit === 'kg' && baseUnit === 'gr') || (recipeUnit === 'l' && baseUnit === 'ml')) {
            qtyNeeded = qtyNeeded * 1000;
        }

        const totalDeduction = qtyNeeded * item.qty;
        materialMap[matIdInRecipe].currentStock -= totalDeduction;
        
        materialSheet.getRange(materialMap[matIdInRecipe].row, 4)
                     .setValue(materialMap[matIdInRecipe].currentStock);
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
    stock: Number(row[3]) || 0,
    price: Number(row[4]) || 0  // Kolom E
  }));
}

function updateMaterialStockAndPrice(id, newStockBase, newPricePerBase, newUnit) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Materials');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        // SET REPLACE (Bug Fix: Jangan mengakumulasikan lagi input replace dengan currentStock lama)
        sheet.getRange(i + 1, 4).setValue(newStockBase); 
        sheet.getRange(i + 1, 5).setValue(newPricePerBase); 
        
        // Simpan konversi unit dasar terbaru jika berubah
        if (newUnit) {
          sheet.getRange(i + 1, 3).setValue(newUnit);
        }
        
        SpreadsheetApp.flush();
        return { status: 'success' };
      }
    }
    return { status: 'error', message: 'ID Material tidak ditemukan' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

// Fungsi Baru Hapus Bahan Master Database Gudang
function deleteMaterialRecord(id) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Materials');
    const data = sheet.getDataRange().getValues();
    
    // Looping reverse untuk keamanan delete baris
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === id) {
        sheet.deleteRow(i + 1);
        SpreadsheetApp.flush();
        return { status: 'success' };
      }
    }
    return { status: 'error', message: 'ID Material tidak ditemukan' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

// Fungsi Management "Atur Shot" Config (Disimpan di sheet 'Shots' agar rapi)
function initShotsSheet(ss) {
  let sheet = ss.getSheetByName('Shots');
  if (!sheet) {
    sheet = ss.insertSheet('Shots');
    sheet.appendRow(['Material ID', 'Grams per Shot']);
    sheet.getRange("1:1").setFontWeight("bold").setBackground("#f3f3f3");
  }
  return sheet;
}

function saveShotConfig(matId, grams) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = initShotsSheet(ss);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === matId) {
        sheet.getRange(i + 1, 2).setValue(grams);
        SpreadsheetApp.flush();
        return { status: 'success' };
      }
    }
    sheet.appendRow([matId, grams]);
    SpreadsheetApp.flush();
    return { status: 'success' };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

function deleteShotConfig(matId) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = initShotsSheet(ss);
    const data = sheet.getDataRange().getValues();
    
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === matId) {
        sheet.deleteRow(i + 1);
        SpreadsheetApp.flush();
        return { status: 'success' };
      }
    }
    return { status: 'success' };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

function getSystemConfigsAndHPP() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  const shotSheet = initShotsSheet(ss);
  const shotData = shotSheet.getDataRange().getValues();
  let shotsMap = {};
  for(let i = 1; i < shotData.length; i++) {
    shotsMap[shotData[i][0]] = Number(shotData[i][1]);
  }

  const matSheet = ss.getSheetByName('Materials');
  const matData = matSheet ? matSheet.getDataRange().getValues() : [];
  let matPricesMap = {};
  let matUnitsMap = {};
  for(let i = 1; i < matData.length; i++) {
    matPricesMap[matData[i][0]] = Number(matData[i][4]) || 0; 
    matUnitsMap[matData[i][0]] = String(matData[i][2] || "").trim().toLowerCase();
  }

  const recSheet = ss.getSheetByName('Recipes');
  const recData = recSheet ? recSheet.getDataRange().getValues() : [];
  let hppMap = {};
  
  for(let i = 1; i < recData.length; i++) {
    const pId = recData[i][0];
    const mId = recData[i][1];
    let qty = Number(recData[i][2]) || 0;
    let unit = String(recData[i][3]).toLowerCase();
    let baseUnit = matUnitsMap[mId] || "";
    let cost = 0;

    if (unit === 'shot') {
      const shotGram = shotsMap[mId]; 
      if (!shotGram) {
         cost = 0; // Shot tidak diatur, abaikan cost untuk bahan ini (Set ke 0 HPP)
      } else {
         qty = qty * shotGram;
         unit = (baseUnit === 'l' || baseUnit === 'ml') ? 'ml' : 'gr';
         
         if ((unit === 'gr' && baseUnit === 'kg') || (unit === 'ml' && baseUnit === 'l')) {
             qty = qty / 1000;
         } else if ((unit === 'kg' && baseUnit === 'gr') || (unit === 'l' && baseUnit === 'ml')) {
             qty = qty * 1000;
         }
         cost = qty * (matPricesMap[mId] || 0);
      }
    } else {
      if ((unit === 'gr' && baseUnit === 'kg') || (unit === 'ml' && baseUnit === 'l')) {
          qty = qty / 1000;
      } else if ((unit === 'kg' && baseUnit === 'gr') || (unit === 'l' && baseUnit === 'ml')) {
          qty = qty * 1000;
      }
      cost = qty * (matPricesMap[mId] || 0);
    }

    hppMap[pId] = (hppMap[pId] || 0) + cost;
  }

  return { shots: shotsMap, hppMap: hppMap };
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

// ==========================================
// FUNGSI BACKEND RECIPE ENGINE & BULK EDIT
// ==========================================

function getProductRecipe(productId) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const recipeSheet = ss.getSheetByName('Recipes');
  const materialSheet = ss.getSheetByName('Materials');

  if (!recipeSheet || !materialSheet) return [];

  const materials = materialSheet.getDataRange().getValues();
  let matMap = {};
  for(let i = 1; i < materials.length; i++) {
    matMap[materials[i][0]] = { name: materials[i][1], unit: materials[i][2] };
  }

  const recipes = recipeSheet.getDataRange().getValues();
  let result = [];
  for(let i = 1; i < recipes.length; i++) {
    if(recipes[i][0] === productId) {
      const mId = recipes[i][1];
      result.push({
        id: mId,
        name: matMap[mId] ? matMap[mId].name : "Bahan Tidak Dikenal",
        qty: recipes[i][2],
        unit: recipes[i][3] || (matMap[mId] ? matMap[mId].unit : ""),
        isNew: false
      });
    }
  }
  return result;
}

function updateProductRecipe(productId, recipeBasket) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let recipeSheet = ss.getSheetByName('Recipes');
    let materialSheet = ss.getSheetByName('Materials');

    if (!recipeSheet) {
      recipeSheet = ss.insertSheet('Recipes');
      recipeSheet.appendRow(['Product ID', 'Material ID', 'Qty', 'Unit']);
    }

    // 1. Hapus resep lama untuk produk ini
    const data = recipeSheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === productId) {
        recipeSheet.deleteRow(i + 1);
      }
    }

    // 2. Masukkan resep baru hasil edit
    let existingMaterials = [];
    if (materialSheet && materialSheet.getLastRow() > 1) {
      existingMaterials = materialSheet.getRange(2, 1, materialSheet.getLastRow() - 1, 1).getValues().flat();
    }

    recipeBasket.forEach(m => {
      let matId = m.id;
      if (m.isNew) {
        let baseMatId = "M-" + m.name.toUpperCase().replace(/[^A-Z0-9]/g, '').substring(0, 8);
        matId = baseMatId;
        let counter = 1;
        while (existingMaterials.includes(matId)) {
          matId = baseMatId + counter;
          counter++;
        }
        if (materialSheet) {
          materialSheet.appendRow([matId, m.name, m.unit, 0]);
          existingMaterials.push(matId);
        }
      }
      recipeSheet.appendRow([productId, matId, m.qty, m.unit]);
    });

    SpreadsheetApp.flush();
    return { status: 'success' };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

function getProductsByMaterial(matId) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const recipeSheet = ss.getSheetByName('Recipes');
  const productSheet = ss.getSheetByName('Products');
  
  if(!recipeSheet || !productSheet) return [];

  const pData = productSheet.getDataRange().getValues();
  let pMap = {};
  for(let i = 1; i < pData.length; i++) {
    pMap[pData[i][0]] = pData[i][1];
  }

  const rData = recipeSheet.getDataRange().getValues();
  let result = [];
  for(let i = 1; i < rData.length; i++) {
    if(rData[i][1] === matId) {
      const pId = rData[i][0];
      result.push({ id: pId, name: pMap[pId] || "Menu Dihapus (" + pId + ")" });
    }
  }
  return result;
}

function executeBulkRecipeAction(matId, action, newValue, affectedProducts, newUnit) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const recipeSheet = ss.getSheetByName('Recipes');
    if(!recipeSheet) throw new Error("Sheet Recipes tidak ditemukan.");

    const rData = recipeSheet.getDataRange().getValues();

    // Iterasi mundur agar indeks baris tidak rusak saat ada row yang dihapus
    for(let i = rData.length - 1; i >= 1; i--) {
      const pId = rData[i][0];
      const mId = rData[i][1];

      if(mId === matId && affectedProducts.includes(pId)) {
        if(action === 'DELETE') {
          recipeSheet.deleteRow(i + 1);
        } else if (action === 'EDIT') {
          recipeSheet.getRange(i + 1, 3).setValue(newValue);
          // Jika ada pergantian unit (Misal Kg ke gr, input unit barunya ke tabel Recipes)
          if(newUnit) recipeSheet.getRange(i + 1, 4).setValue(newUnit);
        }
      }
    }
    SpreadsheetApp.flush();
    return { status: 'success' };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

// [FUNGSI BARU] Letakkan fungsi ini sebelum fungsi doPost() di Code.js
function updateTransactionStatus(orderId, newStatus) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const transSheet = ss.getSheetByName('Transactions');
    const tData = transSheet.getDataRange().getValues();

    for (let i = tData.length - 1; i >= 1; i--) {
      if (tData[i][1] === orderId) {
        // Otomatis update kolom H (ke-8)
        transSheet.getRange(i + 1, 8).setValue(newStatus); 
        SpreadsheetApp.flush();
        return { status: 'success' };
      }
    }
    return { status: 'error', message: 'Transaksi tidak ditemukan' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

// [BARU] Fungsi untuk mengambil gambar Logo dan Wallpaper saat halaman web dimuat (Khusus eksternal host)
function getAppConfig() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const configSheet = ss.getSheetByName('Config');
    return { 
      status: 'success', 
      logo: configSheet.getRange("B4").getValue(), 
      wallpaper: configSheet.getRange("B5").getValue() 
    };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

// ==========================================
// FUNGSI MANNING & SCHEDULING
// ==========================================

function initStaffSheets(ss) {
  let staffSheet = ss.getSheetByName('Staff');
  if (!staffSheet) {
    staffSheet = ss.insertSheet('Staff');
    staffSheet.appendRow(['ID', 'Name', 'Role', 'BaseSalary', 'Status']);
    staffSheet.getRange("1:1").setFontWeight("bold").setBackground("#f3f3f3");
  }
  
  let scheduleSheet = ss.getSheetByName('Schedules');
  if (!scheduleSheet) {
    scheduleSheet = ss.insertSheet('Schedules');
    scheduleSheet.appendRow(['SchedID', 'StaffName', 'Date', 'Start', 'End']);
    scheduleSheet.getRange("1:1").setFontWeight("bold").setBackground("#f3f3f3");
  }
  return { staffSheet, scheduleSheet };
}

function getManningData() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheets = initStaffSheets(ss);
    
    const staffData = sheets.staffSheet.getDataRange().getValues();
    const schedData = sheets.scheduleSheet.getDataRange().getValues();
    
    let staffList = [];
    if (staffData.length > 1) {
      staffList = staffData.slice(1).map(r => ({
        id: r[0], name: r[1], role: r[2], salary: Number(r[3]) || 0, status: r[4]
      }));
    }
    
    let schedList = [];
    if (schedData.length > 1) {
      schedList = schedData.slice(1).map(r => {
        let dateVal = r[2];
        let startVal = r[3];
        let endVal = r[4];

        // Format ulang objek Date yang bocor dari GSheet agar tidak merusak UI JSON
        if (dateVal instanceof Date) dateVal = Utilities.formatDate(dateVal, "GMT+7", "dd/MM/yyyy");
        if (startVal instanceof Date) startVal = Utilities.formatDate(startVal, "GMT+7", "HH:mm");
        if (endVal instanceof Date) endVal = Utilities.formatDate(endVal, "GMT+7", "HH:mm");

        return {
          id: r[0], staffName: r[1], date: dateVal, start: startVal, end: endVal
        };
      });
    }
    
    // Ambil data foto biometrik untuk sinkronisasi avatar profile
    const configSheet = ss.getSheetByName('Config');
    const names = configSheet.getRange("B6:B15").getValues().flat();
    const urls = configSheet.getRange("D6:D15").getValues().flat();
    let faceMap = {};
    for(let i=0; i<names.length; i++) {
      if(names[i] && urls[i]) faceMap[names[i]] = urls[i];
    }
    
    return { status: 'success', staff: staffList, schedules: schedList, faces: faceMap };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

function saveStaff(id, name, role, salary, pin) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheets = initStaffSheets(ss);
    const staffSheet = sheets.staffSheet;
    
    let finalId = id;
    let oldName = ""; 
    
    if (!finalId) {
      finalId = "STF-" + Date.now();
      // Tambahkan string kosong '' untuk Kolom F agar tidak error saat di baca Admin baru
      staffSheet.appendRow([finalId, name, role, salary, 'Active', '']);
    } else {
      const data = staffSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === finalId) {
          oldName = data[i][1];
          staffSheet.getRange(i + 1, 2).setValue(name);
          staffSheet.getRange(i + 1, 3).setValue(role);
          staffSheet.getRange(i + 1, 4).setValue(salary);
          break;
        }
      }
    }
    
    const updatedCashiers = syncConfigStaff(ss, staffSheet, oldName, name, role, pin);
    SpreadsheetApp.flush();
    return { status: 'success', cashiers: updatedCashiers };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

function deleteStaff(id) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const staffSheet = ss.getSheetByName('Staff');
    const data = staffSheet.getDataRange().getValues();
    
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === id) {
        staffSheet.deleteRow(i + 1);
        break;
      }
    }
    
    // Sinkronisasi otomatis ke Config setelah penghapusan
    const updatedCashiers = syncConfigStaff(ss, staffSheet, null, null, null, null);
    SpreadsheetApp.flush();
    return { status: 'success', cashiers: updatedCashiers };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

function updateStaffStatus(id, newStatus) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const staffSheet = ss.getSheetByName('Staff');
    const data = staffSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        staffSheet.getRange(i + 1, 5).setValue(newStatus); // Update Kolom Status (E)
        break;
      }
    }
    
    // Sinkronisasi otomatis ke Config agar Kasir yang Inactive hilang dari dropdown
    const updatedCashiers = syncConfigStaff(ss, staffSheet, null, null, null, null);
    SpreadsheetApp.flush();
    return { status: 'success', cashiers: updatedCashiers, newStatus: newStatus };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

function syncConfigStaff(ss, staffSheet, oldName, newName, role, newPin) {
  const configSheet = ss.getSheetByName('Config');
  if (!configSheet) return [];
  const staffData = staffSheet.getDataRange().getValues();
  
  // 1. Ambil semua staff dari sheet Staff beserta Status-nya
  let allStaff = [];
  for (let i = 1; i < staffData.length; i++) {
    allStaff.push({
       name: staffData[i][1],
       role: staffData[i][2],
       status: staffData[i][4] // Ambil status Active/Inactive dari Kolom E
    });
  }
  
  // 2. Baca data Config saat ini (B6:E25) untuk menyelamatkan PIN & Biometrik lama
  const configRange = configSheet.getRange("B6:E25");
  const configValues = configRange.getValues();
  
  let configMap = {};
  for (let i = 0; i < configValues.length; i++) {
    let cName = configValues[i][0];
    if (cName) {
      configMap[cName] = {
        pin: configValues[i][1],
        photo: configValues[i][2],
        descriptor: configValues[i][3]
      };
    }
  }
  
  if (oldName && oldName !== newName && configMap[oldName]) {
     configMap[newName] = configMap[oldName];
     delete configMap[oldName];
  }
  
  configRange.clearContent();
  let cashiers = [];
  
  // 4. Tulis ulang dengan data terbaru dan aman
  for (let i = 0; i < allStaff.length; i++) {
    let sName = allStaff[i].name;
    let sRole = allStaff[i].role.toUpperCase();
    let sStatus = allStaff[i].status;
    
    let sPin = "";
    let sPhoto = "";
    let sDesc = "";
    
    if (configMap[sName]) {
        sPin = configMap[sName].pin;
        sPhoto = configMap[sName].photo;
        sDesc = configMap[sName].descriptor;
    }
    
    if (sName === newName) {
       // [UPDATE] Izinkan ADMIN untuk menyimpan password di sheet Config
       if (sRole === 'CASHIER' || sRole === 'KASIR' || sRole === 'ADMIN') {
          sPin = newPin ? newPin : sPin;
       } else {
          sPin = "";
       }
    } else {
       if (sRole !== 'CASHIER' && sRole !== 'KASIR' && sRole !== 'ADMIN') {
          sPin = "";
       }
    }
    
    // HANYA MASUKKAN KE DROPDOWN POS JIKA ROLE KASIR DAN STATUS ACTIVE (Admin tersembunyi)
    if ((sRole === 'CASHIER' || sRole === 'KASIR') && sStatus === 'Active') {
       cashiers.push(sName);
    }
    
    configSheet.getRange(6 + i, 2).setValue(sName);
    configSheet.getRange(6 + i, 3).setValue(sPin); 
    configSheet.getRange(6 + i, 4).setValue(sPhoto);
    configSheet.getRange(6 + i, 5).setValue(sDesc);
  }
  
  return cashiers;
}

function syncConfigCashiers(ss, staffSheet) {
  const configSheet = ss.getSheetByName('Config');
  if (!configSheet) return [];
  
  const staffData = staffSheet.getDataRange().getValues();
  let cashiers = [];
  for (let i = 1; i < staffData.length; i++) {
    if (staffData[i][2].toUpperCase() === 'CASHIER' || staffData[i][2].toUpperCase() === 'KASIR') {
      cashiers.push(staffData[i][1]); // Ambil Nama
    }
  }
  
  // Clear existing cashiers in Config (B6:B15 to be safe)
  configSheet.getRange("B6:B15").clearContent();
  
  // Write new cashiers
  for (let i = 0; i < cashiers.length; i++) {
    configSheet.getRange(6 + i, 2).setValue(cashiers[i]);
  }
  
  return cashiers; // Kembalikan array terbaru untuk render frontend realtime
}

function saveSchedule(staffName, date, start, end, existingId) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheets = initStaffSheets(ss);
    
    // Normalisasi format tanggal dd/MM/yyyy
    let formattedDate = date;
    if (date.includes('-')) {
      const dateParts = date.split('-');
      formattedDate = `${dateParts[2]}/${dateParts[1]}/${dateParts[0]}`;
    }

    if (existingId) {
        // Mode Edit
        const data = sheets.scheduleSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
            if (data[i][0] === existingId) {
                sheets.scheduleSheet.getRange(i + 1, 3).setValue(formattedDate);
                sheets.scheduleSheet.getRange(i + 1, 4).setValue(start);
                sheets.scheduleSheet.getRange(i + 1, 5).setValue(end);
                break;
            }
        }
    } else {
        // Mode Buat Baru
        const schedId = "SCH-" + Date.now();
        sheets.scheduleSheet.appendRow([schedId, staffName, formattedDate, start, end]);
    }
    
    SpreadsheetApp.flush();
    return { status: 'success' };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

function deleteSchedule(schedId) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const scheduleSheet = ss.getSheetByName('Schedules');
    const data = scheduleSheet.getDataRange().getValues();
    
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === schedId) {
        scheduleSheet.deleteRow(i + 1);
        break;
      }
    }
    SpreadsheetApp.flush();
    return { status: 'success' };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

// ==========================================
// MODIFIKASI FUNGSI addAuditLog (TELEGRAM INJECTED)
// ==========================================
function addAuditLog(user, action, detail, device) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('AuditLog');
    
    if (!sheet) {
      sheet = ss.insertSheet('AuditLog');
      sheet.appendRow(['Timestamp', 'User', 'Action', 'Detail', 'Device']);
      sheet.getRange("1:1").setFontWeight("bold").setBackground("#f3f3f3");
    }

    const now = new Date();
    const timestampStr = Utilities.formatDate(now, "GMT+7", "dd/MM/yyyy HH:mm:ss");
    
    sheet.insertRowAfter(1);
    sheet.getRange(2, 1, 1, 5).setValues([[timestampStr, user, action, detail, device]]);
    SpreadsheetApp.flush();

    // ---> INJEKSI TELEGRAM CERDAS <---
    if (action.toUpperCase().includes("TRANSAKSI")) {
        // Jika ini adalah log transaksi, kirim ke grup Transaksi
        const tgText = `<b>💸 LOG TRANSAKSI</b>\n` +
                       `<b>Waktu:</b> ${timestampStr}\n` +
                       `<b>User:</b> ${user}\n` +
                       `<b>Detail:</b> ${detail}\n` +
                       `<b>Device:</b> ${device}`;
        sendTelegramMessage('TRANSACTION', tgText);
    } else {
        // Untuk semua log lainnya, kirim ke grup Audit
        const tgText = `<b>📋 AUDIT LOG BARU</b>\n` +
                       `<b>Waktu:</b> ${timestampStr}\n` +
                       `<b>User:</b> ${user}\n` +
                       `<b>Aksi:</b> ${action}\n` +
                       `<b>Detail:</b> ${detail}\n` +
                       `<b>Device:</b> ${device}`;
        sendTelegramMessage('AUDIT', tgText);
    }

    return { status: 'success' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

// ==========================================
// FUNGSI SETTING & TELEGRAM BOT KONEKSI
// ==========================================
function getSettingConfig(userName) {
  try {
    const props = PropertiesService.getScriptProperties();
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const configSheet = ss.getSheetByName('Config');
    const ownerPass = String(configSheet.getRange("B1").getValue());

    // Membaca konfigurasi Telegram dan Link Youtube dari Sheet Config secara realtime
    const tgToken = String(configSheet.getRange("B31").getValue());
    const chatAudit = String(configSheet.getRange("B32").getValue());
    const chatTrans = String(configSheet.getRange("B33").getValue());
    const chatSched = String(configSheet.getRange("B34").getValue());
    const ytLink = String(configSheet.getRange("B35").getValue());
    const chatOwner = String(configSheet.getRange("B36").getValue());

    return {
      status: 'success',
      active: props.getProperty('TG_ACTIVE') === 'true',
      token: tgToken,
      chatAudit: chatAudit,
      chatTrans: chatTrans,
      chatSchedule: chatSched,
      chatOwner: chatOwner,
      ytLink: ytLink,
      multiOwner: props.getProperty('MULTI_OWNER') === 'true',
      taxActive: props.getProperty('TAX_ACTIVE') === 'true',
      taxPercent: Number(props.getProperty('TAX_PERCENT')) || 0,
      ownerPass: userName === 'Owner' ? ownerPass : '' 
    };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

function saveSettingConfig(config) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('TG_ACTIVE', config.active ? 'true' : 'false');
    props.setProperty('MULTI_OWNER', config.multiOwner ? 'true' : 'false');
    props.setProperty('TAX_ACTIVE', config.taxActive ? 'true' : 'false');
    props.setProperty('TAX_PERCENT', config.taxPercent || 0);

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const configSheet = ss.getSheetByName('Config');
    
    // Menyimpan Token dan ID Chat Telegram, beserta Link Youtube langsung ke Sheet Config
    configSheet.getRange("B31").setValue(config.token || '');
    configSheet.getRange("B32").setValue(config.chatAudit || '');
    configSheet.getRange("B33").setValue(config.chatTrans || '');
    configSheet.getRange("B34").setValue(config.chatSchedule || '');
    configSheet.getRange("B36").setValue(config.chatOwner || '');
    
    if (config.ytLink !== undefined) {
      configSheet.getRange("B35").setValue(config.ytLink || '');
    }

    if (config.ownerPass !== undefined && config.ownerPass.trim() !== "" && config.isOwner) {
      configSheet.getRange("B1").setValue(config.ownerPass);
    }
    
    SpreadsheetApp.flush();
    return { status: 'success' };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

function sendTelegramMessage(type, text) {
  try {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty('TG_ACTIVE') !== 'true') return; // Bypass jika toggle OFF
    
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const configSheet = ss.getSheetByName('Config');
    
    // Ambil Token & ID Chat dari Sheet Config
    const token = String(configSheet.getRange("B31").getValue()).trim();
    if (!token) return;

    let chatId = '';
    if (type === 'AUDIT') chatId = String(configSheet.getRange("B32").getValue()).trim();
    else if (type === 'TRANSACTION') chatId = String(configSheet.getRange("B33").getValue()).trim();

    if (!chatId) return;

    const url = "https://api.telegram.org/bot" + token + "/sendMessage";
    const payload = {
      chat_id: chatId,
      text: text,
      parse_mode: 'HTML' 
    };

    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    UrlFetchApp.fetch(url, options);
  } catch(e) {
    console.error("Gagal kirim telegram: " + e.toString());
  }
}

function sendTelegramMessage(type, text) {
  try {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty('TG_ACTIVE') !== 'true') return; // Bypass jika toggle OFF
    
    const token = props.getProperty('TG_TOKEN');
    if (!token) return;

    let chatId = '';
    if (type === 'AUDIT') chatId = props.getProperty('TG_CHAT_AUDIT');
    else if (type === 'TRANSACTION') chatId = props.getProperty('TG_CHAT_TRANS');

    if (!chatId) return;

    const url = "https://api.telegram.org/bot" + token + "/sendMessage";
    const payload = {
      chat_id: chatId,
      text: text,
      parse_mode: 'HTML' 
    };

    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    UrlFetchApp.fetch(url, options);
  } catch(e) {
    console.error("Gagal kirim telegram: " + e.toString());
  }
}

// ==========================================
// FUNGSI UPDATE NAMA TOKO & OPERATIONAL COST
// ==========================================

function updateStoreName(newName) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const configSheet = ss.getSheetByName('Config');
    configSheet.getRange("B3").setValue(newName);
    return { status: 'success', newName: newName };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

function initOpCostSheet(ss) {
  let sheet = ss.getSheetByName('OperationalCosts');
  if (!sheet) {
    sheet = ss.insertSheet('OperationalCosts');
    sheet.appendRow(['ID', 'Name', 'Amount']);
    sheet.getRange("1:1").setFontWeight("bold").setBackground("#f3f3f3");
  }
  return sheet;
}

function sendScheduleToTelegram(base64Data, captionMsg) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const configSheet = ss.getSheetByName('Config');
    
    // Ambil Token & ID Chat Schedule dari Sheet Config
    const token = String(configSheet.getRange("B31").getValue()).trim();
    const chatId = String(configSheet.getRange("B34").getValue()).trim();
    
    if (!token || !chatId) {
      return { status: 'error', message: 'Token Bot atau ID Chat belum diatur di Sheet Config (B31 / B34).' };
    }

    // Pisahkan header base64 dari data aslinya
    const splitData = base64Data.split(',');
    const bytes = Utilities.base64Decode(splitData[1]);
    const blob = Utilities.newBlob(bytes, 'image/jpeg', 'Jadwal_Staff.jpg');

    const url = "https://api.telegram.org/bot" + token + "/sendPhoto";
    const payload = {
      chat_id: chatId,
      photo: blob,
      caption: captionMsg,
      parse_mode: 'HTML'
    };

    const options = {
      method: "post",
      payload: payload,
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    
    if(result.ok) {
       return { status: 'success' };
    } else {
       return { status: 'error', message: result.description };
    }
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

function sendLaporanToTelegram(base64Data, captionMsg) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const configSheet = ss.getSheetByName('Config');
    
    // Ambil Token & ID Chat Owner dari Sheet Config
    const token = String(configSheet.getRange("B31").getValue()).trim();
    const chatId = String(configSheet.getRange("B36").getValue()).trim();
    
    if (!token || !chatId) {
      return { status: 'error', message: 'Token Bot atau ID Chat Owner belum diatur di Sheet Config (B31 / B36).' };
    }

    // Pisahkan header base64 dari data aslinya
    const splitData = base64Data.split(',');
    const bytes = Utilities.base64Decode(splitData[1]);
    const blob = Utilities.newBlob(bytes, 'image/jpeg', 'Laporan_Keuangan.jpg');

    const url = "https://api.telegram.org/bot" + token + "/sendPhoto";
    const payload = {
      chat_id: chatId,
      photo: blob,
      caption: captionMsg,
      parse_mode: 'HTML'
    };

    const options = {
      method: "post",
      payload: payload,
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    
    if(result.ok) {
       return { status: 'success' };
    } else {
       return { status: 'error', message: result.description };
    }
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

function getOperationalData() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const staffSheet = ss.getSheetByName('Staff');
    const opSheet = initOpCostSheet(ss);
    const props = PropertiesService.getScriptProperties();
    const isMultiOwnerOn = props.getProperty('MULTI_OWNER') === 'true';

    let totalSalary = 0;
    let adminShares = []; // Menyimpan data persentase saham admin

    if (staffSheet) {
      const staffData = staffSheet.getDataRange().getValues();
      for (let i = 1; i < staffData.length; i++) {
        if (staffData[i][4] === 'Active') {
          const role = String(staffData[i][2]).toUpperCase();
          const value = Number(staffData[i][3]) || 0; // Salary atau Persentase
          const name = staffData[i][1];

          if (isMultiOwnerOn && role === 'ADMIN') {
            // Jika Multi Owner ON, Admin tidak masuk total gaji, tapi masuk daftar saham
            adminShares.push({ name: name, percentage: value });
          } else {
            // Selain itu (Staff biasa, atau Admin jika Multi Owner OFF), masuk hitungan gaji
            totalSalary += value;
          }
        }
      }
    }

    const opData = opSheet.getDataRange().getValues();
    let customCosts = [];
    for (let i = 1; i < opData.length; i++) {
      customCosts.push({
        id: opData[i][0],
        name: opData[i][1],
        amount: Number(opData[i][2]) || 0
      });
    }

    return { 
      status: 'success', 
      totalSalary: totalSalary, 
      customCosts: customCosts,
      adminShares: adminShares, // Kirim data saham ke frontend
      isMultiOwner: isMultiOwnerOn
    };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

function saveOpCost(id, name, amount) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = initOpCostSheet(ss);
    
    if (id) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === id) {
          sheet.getRange(i + 1, 2).setValue(name);
          sheet.getRange(i + 1, 3).setValue(amount);
          return { status: 'success' };
        }
      }
    } else {
      sheet.appendRow(['OPC-' + Date.now(), name, amount]);
    }
    return { status: 'success' };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

function deleteOpCost(id) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('OperationalCosts');
    if (!sheet) return { status: 'error', message: 'Sheet tidak ditemukan' };
    
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === id) {
        sheet.deleteRow(i + 1);
        return { status: 'success' };
      }
    }
    return { status: 'error', message: 'ID Cost tidak ditemukan' };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

function forceChangeAdminPass(name, newPass) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const configSheet = ss.getSheetByName('Config');
    const staffSheet = ss.getSheetByName('Staff');
    
    // 1. Update Password di Config
    const configData = configSheet.getRange("B6:C25").getValues();
    for (let i = 0; i < configData.length; i++) {
      if (configData[i][0] === name) {
        configSheet.getRange(6 + i, 3).setValue(newPass); 
        break;
      }
    }
    
    // 2. Beri stempel 'Done' di sheet Staff (Kolom F)
    if (staffSheet) {
      const staffData = staffSheet.getDataRange().getValues();
      for (let i = 1; i < staffData.length; i++) {
        if (staffData[i][1] === name) {
          staffSheet.getRange(i + 1, 6).setValue('Done');
          break;
        }
      }
    }
    SpreadsheetApp.flush();
    return { status: 'success' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

function saveBatchSchedule(weekDatesArr, newSchedules) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const scheduleSheet = ss.getSheetByName('Schedules');
    if (!scheduleSheet) return { status: 'error', message: 'Sheet Schedules tidak ditemukan' };

    const data = scheduleSheet.getDataRange().getValues();
    
    // Looping reverse untuk menghapus jadwal yang beririsan dengan tanggal minggu yang sedang di-edit
    let rowsToDelete = [];
    for (let i = data.length - 1; i >= 1; i--) {
      let rowDate = data[i][2];
      if (rowDate instanceof Date) {
        rowDate = Utilities.formatDate(rowDate, "GMT+7", "dd/MM/yyyy");
      } else {
        rowDate = String(rowDate).trim();
      }
      
      if (weekDatesArr.includes(rowDate)) {
        rowsToDelete.push(i + 1); // 1-based indexing for sheets
      }
    }
    
    // Eksekusi hapus baris
    rowsToDelete.forEach(r => scheduleSheet.deleteRow(r));
    
    // Insert batch jadwal baru hasil drag & drop
    if (newSchedules && newSchedules.length > 0) {
      let rowsToInsert = [];
      newSchedules.forEach((s, idx) => {
        rowsToInsert.push(["SCH-" + Date.now() + "-" + idx, s.staffName, s.date, s.start, s.end]);
      });
      scheduleSheet.getRange(scheduleSheet.getLastRow() + 1, 1, rowsToInsert.length, 5).setValues(rowsToInsert);
    }
    
    SpreadsheetApp.flush();
    return { status: 'success' };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

function doPost(e) {
  try {
    const request = JSON.parse(e.postData.contents);
    const action = request.action;
    const args = request.args || [];
    
    let result;
    switch(action) {
      case 'loginUser': result = loginUser(args[0]); break;
      case 'verifyCashierPIN': result = verifyCashierPIN(args[0], args[1]); break;
      case 'getProducts': result = getProducts(); break;
      case 'addProduct': result = addProduct(args[0], args[1], args[2], args[3], args[4], args[5], args[6]); break;
      case 'uploadProductImage': result = uploadProductImage(args[0], args[1], args[2]); break;
      case 'deleteProduct': result = deleteProduct(args[0]); break;
      case 'deleteProductImage': result = deleteProductImage(args[0]); break;
      case 'recordAttendance': result = recordAttendance(args[0], args[1], args[2]); break;
      case 'getAttendanceLogs': result = getAttendanceLogs(); break;
      case 'registerStaffFace': result = registerStaffFace(args[0], args[1], args[2]); break;
      case 'getStaffMasterFaces': result = getStaffMasterFaces(); break;
      case 'deleteMasterFaceRecord': result = deleteMasterFaceRecord(args[0]); break;
      case 'deleteAttendanceRecord': result = deleteAttendanceRecord(args[0], args[1]); break;
      case 'getAllTransactionData': result = getAllTransactionData(); break;
      case 'deleteTransactionRecord': result = deleteTransactionRecord(args[0]); break;
      case 'updateTransactionStatus': result = updateTransactionStatus(args[0], args[1]); break;
      case 'getMaterialsData': result = getMaterialsData(); break;
      case 'clearMarketCache': result = clearMarketCache(); break;
      case 'getNextOrderId': result = getNextOrderId(); break;
      case 'saveTransaction': result = saveTransaction(args[0]); break;
      case 'getMaterialsStock': result = getMaterialsStock(); break;
      case 'updateMaterialStockManual': result = updateMaterialStockManual(args[0], args[1]); break;
      case 'getProductRecipe': result = getProductRecipe(args[0]); break;
      case 'updateProductRecipe': result = updateProductRecipe(args[0], args[1]); break;
      case 'getProductsByMaterial': result = getProductsByMaterial(args[0]); break;
      case 'executeBulkRecipeAction': result = executeBulkRecipeAction(args[0], args[1], args[2], args[3]); break;
      case 'uploadConfigImage': result = uploadConfigImage(args[0], args[1], args[2]); break;
      case 'deleteConfigImage': result = deleteConfigImage(args[0]); break;
      case 'getAppConfig': result = getAppConfig(); break;
      case 'updateMaterialStockAndPrice': result = updateMaterialStockAndPrice(args[0], args[1], args[2], args[3]); break;
      case 'deleteMaterialRecord': result = deleteMaterialRecord(args[0]); break;
      case 'saveShotConfig': result = saveShotConfig(args[0], args[1]); break;
      case 'deleteShotConfig': result = deleteShotConfig(args[0]); break;
      case 'getSystemConfigsAndHPP': result = getSystemConfigsAndHPP(); break;
      case 'getManningData': result = getManningData(); break;
      case 'saveStaff': result = saveStaff(args[0], args[1], args[2], args[3], args[4]); break;
      case 'deleteStaff': result = deleteStaff(args[0]); break;
      case 'updateStaffStatus': result = updateStaffStatus(args[0], args[1]); break;
      case 'saveSchedule': result = saveSchedule(args[0], args[1], args[2], args[3], args[4]); break;
      case 'saveBatchSchedule': result = saveBatchSchedule(args[0], args[1]); break;
      case 'deleteSchedule': result = deleteSchedule(args[0]); break;
      case 'addAuditLog': result = addAuditLog(args[0], args[1], args[2], args[3]); break;
      case 'getSettingConfig': result = getSettingConfig(args[0]); break;
      case 'saveSettingConfig': result = saveSettingConfig(args[0]); break;
      case 'sendTelegramMessageBackend': result = sendTelegramMessageBackend(args[0], args[1]); break;
      case 'sendScheduleToTelegram': result = sendScheduleToTelegram(args[0], args[1]); break;
      case 'sendLaporanToTelegram': result = sendLaporanToTelegram(args[0], args[1]); break;
      case 'updateStoreName': result = updateStoreName(args[0]); break;
      case 'getOperationalData': result = getOperationalData(); break;
      case 'saveOpCost': result = saveOpCost(args[0], args[1], args[2]); break;
      case 'deleteOpCost': result = deleteOpCost(args[0]); break;
      case 'updateProductPrice': result = updateProductPrice(args[0], args[1]); break;
      case 'forceChangeAdminPass': result = forceChangeAdminPass(args[0], args[1]); break;
      default: 
        if (action === 'updateMonitorStatusBackend') result = updateMonitorStatusBackend(args[0], args[1], args[2]); else
        throw new Error("Fungsi '" + action + "' tidak terdaftar di router backend.");
    }
    
    return ContentService.createTextOutput(JSON.stringify(result))
          .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
          .setMimeType(ContentService.MimeType.JSON);
  }
}