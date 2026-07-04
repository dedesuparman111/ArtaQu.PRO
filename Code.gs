/* Rahaza Digital Premium Webapps Builder by ddsuparman */

// ID Spreadsheet yang ditargetkan — SATU-SATUNYA sumber data.
// index.html tidak lagi punya field/kirim spreadsheetId (hanya URL Web App),
// jadi ID inilah yang selalu dipakai oleh resolveSpreadsheetId() di bawah.
const SPREADSHEET_ID = "1ch-TruFPooKccRMJW0fuVRCuOujdq-eNaYiPmMBQ3lk";

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ArtaQu - Pro')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/**
 * Handle External POST Requests (API/Webhook) - Mendukung PWA Standalone Mode
 */
function doPost(e) {
  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    }
    
    const action = payload.action;
    const args = payload.arguments || [];
    const spreadsheetId = payload.spreadsheetId || '';

    let result;
    if (action === "loginUser") {
      result = loginUser(args[0], args[1], spreadsheetId);
    } else if (action === "getAppData") {
      result = getAppData(spreadsheetId);
    } else if (action === "getDashboardData") {
      result = getDashboardData(spreadsheetId);
    } else if (action === "getTransactions") {
      result = getTransactions(spreadsheetId);
    } else if (action === "getInstallments") {
      result = getInstallments(spreadsheetId);
    } else if (action === "addTransaction") {
      result = addTransaction(args[0], spreadsheetId);
    } else if (action === "deleteTransaction") {
      result = deleteTransaction(args[0], spreadsheetId);
    } else if (action === "updateTransaction") {
      result = updateTransaction(args[0], spreadsheetId);
    } else if (action === "addInstallment") {
      result = addInstallment(args[0], spreadsheetId);
    } else if (action === "deleteInstallment") {
      result = deleteInstallment(args[0], spreadsheetId);
    } else if (action === "updateInstallment") {
      result = updateInstallment(args[0], spreadsheetId);
    } else if (action === "resetDatabase") {
      result = resetDatabase(spreadsheetId);
    } else if (action === "setupDatabase") {
      result = setupDatabase(spreadsheetId);
    } else if (action === "updateUser") {
      result = updateUser(args[0], args[1], args[2], spreadsheetId);
    } else {
      // Fallback jika memanggil webhook luar dengan payload format addTransaction
      if (action === "addTransaction" && payload.data) {
        result = addTransaction(payload.data, spreadsheetId);
      } else {
        result = createResponse(true, null, "POST request diterima dengan baik.");
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Menangani permintaan OPTIONS preflight (CORS) jika diperlukan browser
function doOptions(e) {
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * Standardized API Response Contract
 */
function createResponse(success, data, message) {
  return { "success": success, "data": data, "message": message };
}

/**
 * Helper: Safely serialize a spreadsheet row value.
 */
function serializeValue(val) {
  if (val instanceof Date) return val.toISOString();
  if (typeof val === 'number') return val;
  return String(val).trim();
}

/**
 * ── UPDATED: Selalu fallback ke SPREADSHEET_ID bawaan, tidak throw lagi ──
 * index.html sudah disederhanakan — menu Pengaturan cuma minta URL Web App,
 * field Spreadsheet ID sudah tidak ada. Akibatnya client TIDAK PERNAH
 * mengirim spreadsheetId, dan error "Spreadsheet ID belum dikonfigurasi..."
 * selalu muncul (termasuk saat login). Fix: kalau kosong, otomatis pakai
 * SPREADSHEET_ID konstan. Kalau suatu saat ada pemanggil yang tetap mengirim
 * ID spesifik, ID itu tetap dipakai (override tetap didukung).
 */
function resolveSpreadsheetId(spreadsheetId) {
  const cleaned = spreadsheetId ? String(spreadsheetId).trim() : '';
  if (cleaned === '' || cleaned === 'SPREADSHEET ID') {
    return SPREADSHEET_ID;
  }
  return cleaned;
}

/**
 * ── FIX: helper untuk membuka spreadsheet ──
 * Selalu pakai resolveSpreadsheetId agar konsisten dengan ID yang dipilih user.
 */
function openSS(spreadsheetId) {
  return SpreadsheetApp.openById(resolveSpreadsheetId(spreadsheetId));
}

/**
 * ── UPDATED: setupDatabase sekarang lebih robust ──
 * Membuat sheet USERS, TRANSACTIONS, INSTALLMENTS kalau belum ada.
 * Juga menambahkan default user (admin/admin123) kalau USERS sheet baru.
 * Dipanggil otomatis saat pertama kali user setup koneksi database.
 */
function setupDatabase(spreadsheetId) {
  try {
    const ss = openSS(spreadsheetId);

    // Sheet USERS
    let usersSheet = ss.getSheetByName("USERS");
    if (!usersSheet) {
      usersSheet = ss.insertSheet("USERS", 0); // Insert di posisi pertama
      usersSheet.appendRow(["id", "username", "password", "role", "status"]);
      usersSheet.appendRow(["U-001", "admin", "admin123", "Admin", "Active"]);
      usersSheet.getRange("1:1").setFontWeight("bold").setBackground("#e2e8f0");
    }

    // Sheet TRANSACTIONS
    let trxSheet = ss.getSheetByName("TRANSACTIONS");
    if (!trxSheet) {
      trxSheet = ss.insertSheet("TRANSACTIONS");
      trxSheet.appendRow(["id", "date", "type", "category", "amount", "description", "status", "timestamp"]);
      trxSheet.getRange("1:1").setFontWeight("bold").setBackground("#e2e8f0");
    }

    // Sheet INSTALLMENTS
    let instSheet = ss.getSheetByName("INSTALLMENTS");
    if (!instSheet) {
      instSheet = ss.insertSheet("INSTALLMENTS");
      instSheet.appendRow(["id", "name", "creditor", "total_amount", "paid_amount", "remaining", "start_date", "due_date", "status", "description", "timestamp"]);
      instSheet.getRange("1:1").setFontWeight("bold").setBackground("#e2e8f0");
    }

    return createResponse(true, null, "Database siap digunakan. Sheets USERS, TRANSACTIONS, dan INSTALLMENTS telah dibuat.");
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

/**
 * Versi manual — jalankan SEKALI langsung dari Apps Script Editor (bukan lewat
 * Web App) untuk provisioning spreadsheet baru yang sedang di-bind ke project
 * ini. Bedanya dengan setupDatabase(): pakai getActiveSpreadsheet() + alert()
 * karena memang didesain dijalankan manual oleh developer, bukan dipanggil
 * dari doGet()/doPost().
 */
function manualSetupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = setupDatabase(ss.getId());
  SpreadsheetApp.getUi().alert(result.success ? result.message : ("Gagal: " + result.message));
}

/**
 * Reset Database: Mengosongkan data transaksi dan cicilan (meninggalkan header).
 */
function resetDatabase(spreadsheetId) {
  try {
    const ss = openSS(spreadsheetId);
    
    const sheetTrx = ss.getSheetByName("TRANSACTIONS");
    if (sheetTrx) {
      sheetTrx.clear();
      sheetTrx.appendRow(["id", "date", "type", "category", "amount", "description", "status", "timestamp"]);
    }
    
    const sheetInst = ss.getSheetByName("INSTALLMENTS");
    if (sheetInst) {
      sheetInst.clear();
      sheetInst.appendRow(["id", "name", "creditor", "total_amount", "paid_amount", "remaining", "start_date", "due_date", "status", "description", "timestamp"]);
    }
    
    return createResponse(true, null, "Database transaksi dan cicilan berhasil dikosongkan.");
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

/* ================================================================
   AUTH — LOGIN / LOGOUT / UPDATE USER
   ================================================================ */

/**
 * Login: cek username & password di sheet USERS.
 * Mengembalikan data user (tanpa password) jika berhasil.
 */
function loginUser(username, password, spreadsheetId) {
  try {
    const ss = openSS(spreadsheetId);
    
    // Auto-setup database jika USERS belum terbentuk
    let sheet = ss.getSheetByName("USERS");
    if (!sheet) {
      setupDatabase(spreadsheetId);
      sheet = ss.getSheetByName("USERS");
    }

    const data = sheet.getDataRange().getValues();
    // Header: id | username | password | role | status
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[1]).trim() === String(username).trim() &&
          String(row[2]).trim() === String(password).trim()) {

        if (String(row[4]).trim() !== 'Active') {
          return createResponse(false, null, "Akun Anda tidak aktif. Hubungi administrator.");
        }
        return createResponse(true, {
          id:       String(row[0]),
          username: String(row[1]),
          role:     String(row[3]),
          status:   String(row[4])
        }, "Login berhasil.");
      }
    }
    return createResponse(false, null, "Username atau password salah.");
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

/**
 * ── NEW: Update User Account ──
 * Mengubah username dan/atau password user berdasarkan ID.
 * Dipanggil dari frontend saat user update profil di modal Pengaturan.
 */
function updateUser(userId, username, password, spreadsheetId) {
  try {
    const ss = openSS(spreadsheetId);
    
    let sheet = ss.getSheetByName("USERS");
    if (!sheet) {
      return createResponse(false, null, "Sheet USERS tidak ditemukan.");
    }

    const data = sheet.getDataRange().getValues();
    // Header: id | username | password | role | status
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(userId).trim()) {
        const rowNum = i + 1;
        
        // Validasi: username jangan kosong
        if (!username || String(username).trim() === '') {
          return createResponse(false, null, "Username tidak boleh kosong.");
        }
        
        // Validasi: password jangan kosong
        if (!password || String(password).trim() === '') {
          return createResponse(false, null, "Password tidak boleh kosong.");
        }
        
        // Update username dan password di sheet
        sheet.getRange(rowNum, 2, 1, 2).setValues([[String(username).trim(), String(password).trim()]]);
        
        return createResponse(true, 
          { id: userId, username: String(username).trim() }, 
          "Akun berhasil diperbarui. Username dan password telah diubah."
        );
      }
    }
    
    return createResponse(false, null, "User ID tidak ditemukan.");
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

/* ================================================================
   COMBINED LOADER — dipakai App.refreshData() di frontend
   Menggabungkan getDashboardData + getTransactions + getInstallments
   jadi SATU request: 1x SpreadsheetApp.openById(), 1x baca TRANSACTIONS,
   1x baca INSTALLMENTS. Sebelumnya 3 fungsi terpisah di atas membuka
   spreadsheet & membaca sheet yang sama berkali-kali → inilah sumber
   utama sinkronisasi/refresh terasa lambat.
   ================================================================ */

function getAppData(spreadsheetId) {
  try {
    const ss = openSS(spreadsheetId);

    // ── TRANSACTIONS: dibaca sekali, dipakai untuk daftar transaksi + agregat dashboard ──
    let trxSheet = ss.getSheetByName("TRANSACTIONS");
    if (!trxSheet) { setupDatabase(spreadsheetId); trxSheet = ss.getSheetByName("TRANSACTIONS"); }

    const trxData    = trxSheet.getDataRange().getValues();
    const trxHeaders = trxData[0];
    const transactions = [];
    let totalIncome = 0, totalExpense = 0, totalReceivable = 0, totalInstallmentPaid = 0;

    for (let i = trxData.length - 1; i > 0; i--) {
      const row = trxData[i];
      if (!row[0] || String(row[0]).trim() === '') continue;

      const type   = String(row[2]).trim();
      const amount = parseFloat(row[4]) || 0;
      if      (type === "Pendapatan")  totalIncome   += amount;
      else if (type === "Pengeluaran") totalExpense  += amount;
      else if (type === "Piutang")     totalReceivable += amount;
      else if (type === "Cicilan")     totalInstallmentPaid += amount;

      const trx = {};
      for (let j = 0; j < trxHeaders.length; j++) trx[trxHeaders[j]] = serializeValue(row[j]);
      trx['amount'] = amount;

      const rawDate = row[1];
      if (rawDate instanceof Date) {
        trx['date'] = rawDate.toISOString();
      } else if (rawDate && String(rawDate).trim() !== '') {
        const parsed = new Date(String(rawDate).trim());
        trx['date'] = isNaN(parsed.getTime()) ? String(rawDate).trim() : parsed.toISOString();
      } else {
        trx['date'] = '';
      }
      transactions.push(trx);
    }

    // ── INSTALLMENTS: dibaca sekali, dipakai untuk daftar cicilan + sisa outstanding ──
    let instSheet = ss.getSheetByName("INSTALLMENTS");
    if (!instSheet) { setupDatabase(spreadsheetId); instSheet = ss.getSheetByName("INSTALLMENTS"); }

    const instData    = instSheet.getDataRange().getValues();
    const instHeaders = instData[0];
    const installments = [];
    let totalInstallmentOutstanding = 0;

    for (let i = instData.length - 1; i > 0; i--) {
      const row = instData[i];
      if (!row[0] || String(row[0]).trim() === '') continue;

      if (String(row[8]).trim() !== 'Lunas') {
        totalInstallmentOutstanding += parseFloat(row[5]) || 0;
      }

      const item = {};
      for (let j = 0; j < instHeaders.length; j++) item[instHeaders[j]] = serializeValue(row[j]);
      item['total_amount'] = parseFloat(row[3]) || 0;
      item['paid_amount']  = parseFloat(row[4]) || 0;
      item['remaining']    = parseFloat(row[5]) || 0;

      const rawDate = row[6];
      item['start_date'] = rawDate instanceof Date ? rawDate.toISOString() : (rawDate ? String(rawDate).trim() : '');

      const rawDueDate = row[7];
      item['due_date'] = rawDueDate instanceof Date ? rawDueDate.toISOString() : (rawDueDate ? String(rawDueDate).trim() : '');

      installments.push(item);
    }

    const balance = totalIncome - (totalExpense + totalReceivable + totalInstallmentPaid);

    return createResponse(true, {
      dashboard: {
        income:                 totalIncome,
        expense:                totalExpense + totalInstallmentPaid,
        debt:                   0,
        receivable:             totalReceivable,
        balance:                balance,
        installmentPaid:        totalInstallmentPaid,
        installmentOutstanding: totalInstallmentOutstanding
      },
      transactions: transactions,
      installments: installments
    }, "Data aplikasi berhasil dimuat.");
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

/* ================================================================
   DASHBOARD
   ================================================================ */

function getDashboardData(spreadsheetId) {
  try {
    const ss = openSS(spreadsheetId);
    
    // Auto-setup jika sheet TRANSACTIONS tidak ditemukan
    let sheet = ss.getSheetByName("TRANSACTIONS");
    if (!sheet) {
      setupDatabase(spreadsheetId);
      sheet = ss.getSheetByName("TRANSACTIONS");
    }

    const data = sheet.getDataRange().getValues();
    let totalIncome = 0, totalExpense = 0, totalReceivable = 0, totalInstallmentPaid = 0;

    for (let i = 1; i < data.length; i++) {
      const type   = String(data[i][2]).trim();
      const amount = parseFloat(data[i][4]) || 0;
      if      (type === "Pendapatan")  totalIncome   += amount;
      else if (type === "Pengeluaran") totalExpense  += amount;
      else if (type === "Piutang")     totalReceivable += amount;
      else if (type === "Cicilan")     totalInstallmentPaid += amount;
    }

    // Ringkasan cicilan dari sheet INSTALLMENTS
    const installSheet = ss.getSheetByName("INSTALLMENTS");
    let totalInstallmentOutstanding = 0;
    if (installSheet) {
      const iData = installSheet.getDataRange().getValues();
      for (let i = 1; i < iData.length; i++) {
        if (String(iData[i][8]).trim() !== 'Lunas') { // kolom status bergeser dari index 7 ke index 8
          totalInstallmentOutstanding += parseFloat(iData[i][5]) || 0; // kolom sisa (tetap index 5)
        }
      }
    }

    const balance = totalIncome - (totalExpense + totalReceivable + totalInstallmentPaid);

    return createResponse(true, {
      income:                   totalIncome,
      expense:                  totalExpense + totalInstallmentPaid,
      debt:                     0,
      receivable:               totalReceivable,
      balance:                  balance,
      installmentPaid:          totalInstallmentPaid,
      installmentOutstanding:   totalInstallmentOutstanding
    }, "Data dashboard berhasil dimuat.");
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

/* ================================================================
   TRANSACTIONS — CRUD
   ================================================================ */

function getTransactions(spreadsheetId) {
  try {
    const ss = openSS(spreadsheetId);
    
    let sheet = ss.getSheetByName("TRANSACTIONS");
    if (!sheet) {
      setupDatabase(spreadsheetId);
      sheet = ss.getSheetByName("TRANSACTIONS");
    }

    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    const transactions = [];

    for (let i = data.length - 1; i > 0; i--) {
      const row = data[i];
      if (!row[0] || String(row[0]).trim() === '') continue;

      const trx = {};
      for (let j = 0; j < headers.length; j++) {
        trx[headers[j]] = serializeValue(row[j]);
      }
      trx['amount'] = parseFloat(row[4]) || 0;

      const rawDate = row[1];
      if (rawDate instanceof Date) {
        trx['date'] = rawDate.toISOString();
      } else if (rawDate && String(rawDate).trim() !== '') {
        const parsed = new Date(String(rawDate).trim());
        trx['date'] = isNaN(parsed.getTime()) ? String(rawDate).trim() : parsed.toISOString();
      } else {
        trx['date'] = '';
      }
      transactions.push(trx);
    }

    return createResponse(true, transactions, "Data transaksi berhasil dimuat.");
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

function addTransaction(trxData, spreadsheetId) {
  try {
    const ss = openSS(spreadsheetId);
    
    let sheet = ss.getSheetByName("TRANSACTIONS");
    if (!sheet) {
      setupDatabase(spreadsheetId);
      sheet = ss.getSheetByName("TRANSACTIONS");
    }

    const newId = "TRX-" + new Date().getTime();
    sheet.appendRow([
      newId,
      trxData.date,
      trxData.type,
      trxData.category,
      parseFloat(trxData.amount) || 0,
      trxData.description || '',
      trxData.status || 'Selesai',
      new Date()
    ]);

    // Jika tipe Cicilan dan ada installment_id, catat setoran ke sheet INSTALLMENTS
    if (trxData.type === "Cicilan" && trxData.installment_id) {
      recordInstallmentPayment(trxData.installment_id, parseFloat(trxData.amount) || 0, spreadsheetId);
    }

    return createResponse(true, { id: newId }, "Transaksi berhasil disimpan.");
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

function deleteTransaction(id, spreadsheetId) {
  try {
    const ss = openSS(spreadsheetId);
    
    let sheet = ss.getSheetByName("TRANSACTIONS");
    if (!sheet) {
      setupDatabase(spreadsheetId);
      sheet = ss.getSheetByName("TRANSACTIONS");
    }

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(id).trim()) {
        sheet.deleteRow(i + 1);
        return createResponse(true, null, "Transaksi berhasil dihapus.");
      }
    }
    return createResponse(false, null, "ID Transaksi tidak ditemukan.");
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

function updateTransaction(trxData, spreadsheetId) {
  try {
    const ss = openSS(spreadsheetId);
    
    let sheet = ss.getSheetByName("TRANSACTIONS");
    if (!sheet) {
      setupDatabase(spreadsheetId);
      sheet = ss.getSheetByName("TRANSACTIONS");
    }

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(trxData.id).trim()) {
        const rowNum = i + 1;
        // Update: date, type, category, amount, description, status — digabung jadi 1x panggilan setValues()
        sheet.getRange(rowNum, 2, 1, 6).setValues([[
          trxData.date,
          trxData.type,
          trxData.category,
          parseFloat(trxData.amount) || 0,
          trxData.description || '',
          trxData.status || 'Selesai'
        ]]);
        // timestamp tidak diubah (biarkan original)
        
        return createResponse(true, { id: trxData.id }, "Transaksi berhasil diperbarui.");
      }
    }
    return createResponse(false, null, "ID Transaksi tidak ditemukan.");
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

/* ================================================================
   INSTALLMENTS — Cicilan (Total / Setor / Sisa)
   Sheet kolom: id | name | creditor | total_amount | paid_amount | remaining | start_date | due_date | status | description | timestamp
   ================================================================ */

function getInstallments(spreadsheetId) {
  try {
    const ss = openSS(spreadsheetId);
    
    let sheet = ss.getSheetByName("INSTALLMENTS");
    if (!sheet) {
      setupDatabase(spreadsheetId);
      sheet = ss.getSheetByName("INSTALLMENTS");
    }

    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    const list    = [];

    for (let i = data.length - 1; i > 0; i--) {
      const row = data[i];
      if (!row[0] || String(row[0]).trim() === '') continue;

      const item = {};
      for (let j = 0; j < headers.length; j++) {
        item[headers[j]] = serializeValue(row[j]);
      }
      item['total_amount'] = parseFloat(row[3]) || 0;
      item['paid_amount']  = parseFloat(row[4]) || 0;
      item['remaining']    = parseFloat(row[5]) || 0;

      const rawDate = row[6];
      if (rawDate instanceof Date) item['start_date'] = rawDate.toISOString();
      else item['start_date'] = rawDate ? String(rawDate).trim() : '';

      const rawDueDate = row[7];
      if (rawDueDate instanceof Date) item['due_date'] = rawDueDate.toISOString();
      else item['due_date'] = rawDueDate ? String(rawDueDate).trim() : '';

      list.push(item);
    }

    return createResponse(true, list, "Data cicilan berhasil dimuat.");
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

function addInstallment(trxData, spreadsheetId) {
  try {
    const ss = openSS(spreadsheetId);
    
    let sheet = ss.getSheetByName("INSTALLMENTS");
    if (!sheet) {
      setupDatabase(spreadsheetId);
      sheet = ss.getSheetByName("INSTALLMENTS");
    }

    const newId      = "INST-" + new Date().getTime();
    const total      = parseFloat(trxData.total_amount) || 0;
    const paid       = 0;
    const remaining  = total;

    sheet.appendRow([
      newId,
      trxData.name        || '',
      trxData.creditor    || '',
      total,
      paid,
      remaining,
      trxData.start_date  || new Date(),
      trxData.due_date    || '', // due_date
      'Berjalan',          // status
      trxData.description || '',
      new Date()           // timestamp
    ]);

    return createResponse(true, { id: newId }, "Cicilan baru berhasil ditambahkan.");
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

function recordInstallmentPayment(installmentId, payAmount, spreadsheetId) {
  try {
    const ss = openSS(spreadsheetId);
    
    let sheet = ss.getSheetByName("INSTALLMENTS");
    if (!sheet) {
      setupDatabase(spreadsheetId);
      sheet = ss.getSheetByName("INSTALLMENTS");
    }

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(installmentId).trim()) {
        const rowNum     = i + 1;
        const total      = parseFloat(data[i][3]) || 0;
        const oldPaid    = parseFloat(data[i][4]) || 0;
        const newPaid    = oldPaid + (parseFloat(payAmount) || 0);
        const remaining  = Math.max(0, total - newPaid);
        const status     = remaining <= 0 ? 'Lunas' : 'Berjalan';

        sheet.getRange(rowNum, 5, 1, 2).setValues([[newPaid, remaining]]); // paid_amount + remaining
        sheet.getRange(rowNum, 9).setValue(status);     // status bergeser dari kolom 8 ke 9

        return createResponse(true, { paid: newPaid, remaining: remaining, status: status },
          "Pembayaran cicilan berhasil dicatat.");
      }
    }
    return createResponse(false, null, "ID Cicilan tidak ditemukan.");
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

function deleteInstallment(id, spreadsheetId) {
  try {
    const ss = openSS(spreadsheetId);
    
    let sheet = ss.getSheetByName("INSTALLMENTS");
    if (!sheet) {
      setupDatabase(spreadsheetId);
      sheet = ss.getSheetByName("INSTALLMENTS");
    }

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(id).trim()) {
        sheet.deleteRow(i + 1);
        return createResponse(true, null, "Cicilan berhasil dihapus.");
      }
    }
    return createResponse(false, null, "ID Cicilan tidak ditemukan.");
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

function updateInstallment(trxData, spreadsheetId) {
  try {
    const ss = openSS(spreadsheetId);
    
    let sheet = ss.getSheetByName("INSTALLMENTS");
    if (!sheet) {
      setupDatabase(spreadsheetId);
      sheet = ss.getSheetByName("INSTALLMENTS");
    }

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(trxData.id).trim()) {
        const rowNum = i + 1;
        const total  = parseFloat(trxData.total_amount) || 0;
        const paid   = parseFloat(trxData.paid_amount) || 0;
        const remaining = Math.max(0, total - paid);
        const status = remaining <= 0 ? 'Lunas' : 'Berjalan';
        
        // Update: name, creditor, total_amount, paid_amount, remaining, start_date, due_date, status, description — digabung jadi 1x panggilan setValues()
        sheet.getRange(rowNum, 2, 1, 9).setValues([[
          trxData.name || '',
          trxData.creditor || '',
          total,
          paid,
          remaining,
          trxData.start_date || new Date(),
          trxData.due_date || '',
          status,
          trxData.description || ''
        ]]);
        
        return createResponse(true, { id: trxData.id, status: status }, "Cicilan berhasil diperbarui.");
      }
    }
    return createResponse(false, null, "ID Cicilan tidak ditemukan.");
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}
