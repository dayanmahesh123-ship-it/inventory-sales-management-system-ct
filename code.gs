// ============================================================
// MAHESH WEB APP SOLUTION — Management System Backend
// Demo Business: Liyanage Electronics
// Version: 6.0 — Complete with Add/Edit/Delete Product
// ============================================================
// 
// 🇱🇰 සිංහල උපදෙස්:
// ─────────────────────────────────────────────────────────
// 1. Google Sheets open කරන්න
// 2. Extensions > Apps Script click කරන්න
// 3. මේ code එක paste කරන්න
// 4. Save කරන්න (Ctrl+S)
// 5. ⚙️ Mahesh App > 📋 Initialize All Sheets — run කරන්න
// 6. Deploy > New Deployment > Web App select කරන්න
//    - Execute as: Me
//    - Who has access: Anyone
// 7. Deploy click කරලා URL එක copy කරන්න
// 8. index.html එකේ API_URL variable එකට ඒ URL paste කරන්න
// ─────────────────────────────────────────────────────────
//
// 📌 Actions supported by doPost():
//   • addProduct     — නව භාණ්ඩයක් එකතු කිරීම
//   • updateProduct  — දැනට ඇති භාණ්ඩයක් සංස්කරණය
//   • deleteProduct  — භාණ්ඩයක් මකා දැමීම
//   • addSale        — නව අලෙවියක් සැකසීම
//   • processReturn  — ආපසු ලැබීමක් සැකසීම
//
// 📌 Actions supported by doGet():
//   • (default)      — සියලුම data ලබා ගැනීම
//   • getProducts    — භාණ්ඩ පමණක්
//   • getSales       — අලෙවි පමණක්
//   • getReturns     — ආපසු ලැබීම් පමණක්
//   • getRestockLog  — Restock log පමණක්
// ============================================================

const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();

// ⚠️ මෙය ඔබගේ සැබෑ email එකට වෙනස් කරන්න
// Low stock alerts මෙම email එකට එවනු ලැබේ
const ADMIN_EMAIL = 'admin@liyanageelectronics.lk';


// ============================================================
// 📧 EMAIL ALERT HELPER — අඩු තොග email ඇඟවීම
// ============================================================
function sendLowStockAlert(productId, productName, currentQty, minStock) {
  try {
    const subject = '⚠️ Low Stock Alert — ' + productName + ' [' + productId + ']';
    const htmlBody = 
      '<div style="font-family:Arial,sans-serif;max-width:500px;margin:0 auto;border:1px solid #e2e8f0;border-radius:12px;overflow:hidden">' +
        '<div style="background:linear-gradient(135deg,#f59e0b,#ef4444);padding:20px;text-align:center">' +
          '<h2 style="color:#fff;margin:0;font-size:18px">⚠️ Low Stock Alert</h2>' +
          '<p style="color:rgba(255,255,255,0.85);margin:4px 0 0;font-size:12px">Liyanage Electronics — Inventory System</p>' +
        '</div>' +
        '<div style="padding:24px">' +
          '<table style="width:100%;border-collapse:collapse;font-size:14px">' +
            '<tr><td style="padding:8px 0;color:#64748b">Product ID</td><td style="padding:8px 0;font-weight:700;text-align:right">' + productId + '</td></tr>' +
            '<tr><td style="padding:8px 0;color:#64748b">Product Name</td><td style="padding:8px 0;font-weight:700;text-align:right">' + productName + '</td></tr>' +
            '<tr style="border-top:1px solid #f1f5f9"><td style="padding:8px 0;color:#64748b">Current Stock</td><td style="padding:8px 0;font-weight:700;color:#ef4444;text-align:right">' + currentQty + ' units</td></tr>' +
            '<tr><td style="padding:8px 0;color:#64748b">Min Stock Level</td><td style="padding:8px 0;font-weight:700;text-align:right">' + minStock + ' units</td></tr>' +
          '</table>' +
          '<div style="margin-top:20px;padding:12px;background:#fef2f2;border-radius:8px;border-left:4px solid #ef4444">' +
            '<p style="margin:0;font-size:13px;color:#991b1b"><strong>Action Required:</strong> Please reorder this item immediately.</p>' +
          '</div>' +
        '</div>' +
        '<div style="padding:12px 24px;background:#f8fafc;text-align:center;font-size:11px;color:#94a3b8">' +
          'Powered by Mahesh Web App Solution | Automated Alert' +
        '</div>' +
      '</div>';

    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: subject,
      htmlBody: htmlBody,
      name: 'Liyanage Electronics System'
    });

    Logger.log('✅ Low stock alert sent for: ' + productName + ' (Qty: ' + currentQty + ')');
    return true;
  } catch (err) {
    Logger.log('❌ Email alert failed: ' + err.toString());
    return false;
  }
}

// Stock එක අඩුනම් email alert එවන්න
function checkAndAlertLowStock(productId, productName, newQty, minStock) {
  if (newQty <= minStock && newQty >= 0) {
    sendLowStockAlert(productId, productName, newQty, minStock);
  }
}


// ============================================================
// 📋 SHEET SETUP — Sheets සියල්ල initialize කිරීම
// ============================================================
function setupSheets() {
  const sheetsConfig = {
    // ────── Products Sheet (භාණ්ඩ) ──────
    Products: {
      headers: [
        'ProductID',    // භාණ්ඩ ID (MOB001, LAP001, etc.)
        'ProductName',  // භාණ්ඩයේ නම
        'Category',     // ප්‍රවර්ගය (Mobile Phones, Laptops, Appliances, Accessories)
        'Brand',        // වෙළඳ නාමය (Samsung, Apple, etc.)
        'Model',        // මාදිලිය
        'UnitPrice',    // විකුණුම් මිල (Rs.)
        'CostPrice',    // මිලදී ගත් මිල (Rs.)
        'StockQty',     // තොග ප්‍රමාණය
        'MinStockLevel',// අවම තොග මට්ටම
        'WarrantyMonths',// වගකීම් මාස ගණන
        'Supplier',     // සැපයුම්කරු
        'DateAdded',    // එකතු කළ දිනය
        'Status'        // තත්වය (Active / Low Stock)
      ],
      data: [
        ['MOB001','Samsung Galaxy S24 Ultra','Mobile Phones','Samsung','S24 Ultra',189900,152000,15,5,12,'Samsung Sri Lanka',new Date('2024-12-01'),'Active'],
        ['MOB002','iPhone 15 Pro Max','Mobile Phones','Apple','15 Pro Max',249900,210000,8,3,12,'Apple Authorized',new Date('2024-12-05'),'Active'],
        ['MOB003','Samsung Galaxy A15','Mobile Phones','Samsung','A15',42900,33000,25,10,12,'Samsung Sri Lanka',new Date('2024-12-10'),'Active'],
        ['MOB004','Xiaomi Redmi Note 13','Mobile Phones','Xiaomi','Redmi Note 13',52900,40000,20,8,12,'Xiaomi Distributors',new Date('2025-01-02'),'Active'],
        ['APP001','LG 55" OLED Smart TV','Appliances','LG','OLED55C3',329900,275000,4,2,24,'LG Electronics Lanka',new Date('2024-11-15'),'Active'],
        ['APP002','Samsung 65" Crystal UHD TV','Appliances','Samsung','CU7000',219900,178000,6,2,24,'Samsung Sri Lanka',new Date('2024-11-20'),'Active'],
        ['APP003','LG 10kg Front Load Washer','Appliances','LG','FV1410S5W',159900,128000,5,2,36,'LG Electronics Lanka',new Date('2025-01-05'),'Active'],
        ['LAP001','Dell Inspiron 15','Laptops','Dell','Inspiron 3520',174900,142000,10,3,12,'Dell Technologies',new Date('2024-12-20'),'Active'],
        ['LAP002','HP Pavilion x360','Laptops','HP','Pavilion x360 14',189900,155000,7,3,12,'HP Lanka',new Date('2025-01-10'),'Active'],
        ['LAP003','Lenovo IdeaPad Slim 3','Laptops','Lenovo','IdeaPad Slim 3',134900,108000,12,4,12,'Lenovo Distributors',new Date('2025-01-12'),'Active'],
        ['ACC001','Samsung Galaxy Buds FE','Accessories','Samsung','Buds FE',18900,14000,30,10,6,'Samsung Sri Lanka',new Date('2025-01-15'),'Active'],
        ['ACC002','Anker PowerCore 20000mAh','Accessories','Anker','PowerCore 20K',8900,5800,40,15,18,'Anker Distributors',new Date('2025-01-18'),'Active'],
        ['ACC003','Logitech MX Master 3S','Accessories','Logitech','MX Master 3S',21900,16500,3,5,24,'Logitech Lanka',new Date('2025-01-20'),'Low Stock'],
        ['APP004','Philips Air Fryer XXL','Appliances','Philips','HD9270',49900,38000,2,5,24,'Philips Lanka',new Date('2025-02-01'),'Low Stock']
      ]
    },

    // ────── Sales Sheet (අලෙවි) ──────
    Sales: {
      headers: [
        'SaleID',         // අලෙවි ID (S0001, S0002, etc.)
        'Date',           // දිනය සහ වේලාව
        'ProductID',      // භාණ්ඩ ID
        'ProductName',    // භාණ්ඩයේ නම
        'Category',       // ප්‍රවර්ගය
        'Qty',            // ප්‍රමාණය
        'UnitPrice',      // ඒකක මිල
        'DiscountPct',    // වට්ටම් ප්‍රතිශතය
        'TotalAmount',    // මුළු මුදල
        'CustomerName',   // පාරිභෝගිකයාගේ නම
        'CustomerPhone',  // දුරකථන අංකය
        'PaymentMethod',  // ගෙවීම් ක්‍රමය (Cash/Card/Bank Transfer)
        'SoldBy',         // විකුණුම්කරු
        'ReturnStatus'    // ආපසු තත්වය ('' / Partial Return / Returned)
      ],
      data: [
        ['S0001',new Date('2025-05-20 09:30'),'MOB001','Samsung Galaxy S24 Ultra','Mobile Phones',1,189900,0,189900,'Kamal Perera','0771234567','Card','Nimal',''],
        ['S0001',new Date('2025-05-20 09:30'),'ACC002','Anker PowerCore 20000mAh','Accessories',2,8900,0,17800,'Kamal Perera','0771234567','Card','Nimal',''],
        ['S0002',new Date('2025-05-20 11:15'),'APP001','LG 55" OLED Smart TV','Appliances',1,329900,5,313405,'Saman Silva','0712345678','Card','Sunil',''],
        ['S0003',new Date('2025-05-21 10:00'),'MOB003','Samsung Galaxy A15','Mobile Phones',2,42900,0,85800,'Ruwan Fernando','0761122334','Cash','Nimal','Partial Return'],
        ['S0004',new Date('2025-05-21 14:30'),'LAP001','Dell Inspiron 15','Laptops',1,174900,0,174900,'Dilshan Jayawardena','0779988776','Card','Sunil',''],
        ['S0005',new Date('2025-05-22 09:00'),'MOB002','iPhone 15 Pro Max','Mobile Phones',1,249900,3,242403,'Nadeesha Kumari','0714455667','Card','Nimal',''],
        ['S0005',new Date('2025-05-22 09:00'),'ACC001','Samsung Galaxy Buds FE','Accessories',1,18900,3,18333,'Nadeesha Kumari','0714455667','Card','Nimal',''],
        ['S0006',new Date('2025-05-22 15:45'),'APP002','Samsung 65" Crystal UHD TV','Appliances',1,219900,0,219900,'Chaminda Bandara','0723344556','Cash','Sunil',''],
        ['S0007',new Date('2025-05-23 10:20'),'LAP003','Lenovo IdeaPad Slim 3','Laptops',1,134900,0,134900,'Amaya Ratnayake','0775566778','Cash','Nimal',''],
        ['S0008',new Date('2025-05-23 13:00'),'MOB004','Xiaomi Redmi Note 13','Mobile Phones',1,52900,0,52900,'Pradeep Wijesinghe','0769900112','Cash','Sunil','Returned'],
        ['S0009',new Date('2025-05-24 11:30'),'LAP002','HP Pavilion x360','Laptops',1,189900,0,189900,'Sachini De Silva','0711223344','Card','Nimal',''],
        ['S0010',new Date('2025-05-25 12:30'),'ACC002','Anker PowerCore 20000mAh','Accessories',3,8900,0,26700,'Thilina Gamage','0776677889','Cash','Sunil','']
      ]
    },

    // ────── Returns Sheet (ආපසු ලැබීම්) ──────
    Returns: {
      headers: [
        'ReturnID',      // ආපසු ID (R0001, R0002, etc.)
        'Date',          // දිනය
        'SaleID',        // මුල් අලෙවි ID
        'ProductID',     // භාණ්ඩ ID
        'ProductName',   // භාණ්ඩයේ නම
        'Qty',           // ආපසු ප්‍රමාණය
        'Reason',        // හේතුව
        'RefundAmount',  // ආපසු මුදල
        'ProcessedBy',   // සැකසූ පුද්ගලයා
        'Status'         // තත්වය (Completed / Pending)
      ],
      data: [
        ['R0001',new Date('2025-05-22'),'S0003','MOB003','Samsung Galaxy A15',1,'Defective unit — screen flickering',42900,'Nimal','Completed'],
        ['R0002',new Date('2025-05-24'),'S0008','MOB004','Xiaomi Redmi Note 13',1,'Customer changed mind within 7 days',52900,'Sunil','Completed']
      ]
    },

    // ────── RestockLog Sheet (තොග පිරවීම් ලොගය) ──────
    RestockLog: {
      headers: [
        'RestockID',   // Restock ID
        'Date',        // දිනය
        'ProductID',   // භාණ්ඩ ID
        'ProductName', // භාණ්ඩයේ නම
        'Qty',         // ප්‍රමාණය
        'Supplier',    // සැපයුම්කරු
        'UnitCost',    // ඒකක මිල
        'TotalCost',   // මුළු මිල
        'ReceivedBy',  // ලැබුවේ
        'Notes'        // සටහන්
      ],
      data: [
        ['RS001',new Date('2025-05-15'),'MOB001','Samsung Galaxy S24 Ultra',10,'Samsung Sri Lanka',152000,1520000,'Nimal','Monthly restock'],
        ['RS002',new Date('2025-05-15'),'ACC002','Anker PowerCore 20000mAh',20,'Anker Distributors',5800,116000,'Sunil','High demand item'],
        ['RS003',new Date('2025-05-18'),'LAP001','Dell Inspiron 15',5,'Dell Technologies',142000,710000,'Nimal','New batch arrival']
      ]
    }
  };

  // Sheets හදන්න / data දාන්න
  for (const [sheetName, config] of Object.entries(sheetsConfig)) {
    let sheet = SPREADSHEET.getSheetByName(sheetName);
    if (sheet) {
      sheet.clear();
    } else {
      sheet = SPREADSHEET.insertSheet(sheetName);
    }

    // Headers set කරන්න
    sheet.getRange(1, 1, 1, config.headers.length).setValues([config.headers]);

    // Data set කරන්න
    if (config.data.length > 0) {
      sheet.getRange(2, 1, config.data.length, config.headers.length).setValues(config.data);
    }

    // Header styling
    sheet.getRange(1, 1, 1, config.headers.length)
      .setFontWeight('bold')
      .setBackground('#1e293b')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center')
      .setFontSize(10);

    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, config.headers.length);
  }

  // Success message
  SpreadsheetApp.getUi().alert(
    '✅ සියලුම Sheets සාර්ථකව සාදන ලදී!\n\n' +
    '• Products (භාණ්ඩ): ' + sheetsConfig.Products.data.length + ' items\n' +
    '• Sales (අලෙවි): ' + sheetsConfig.Sales.data.length + ' rows\n' +
    '• Returns (ආපසු): ' + sheetsConfig.Returns.data.length + ' records\n' +
    '• RestockLog (තොග): ' + sheetsConfig.RestockLog.data.length + ' records\n\n' +
    '📧 Low stock alerts → ' + ADMIN_EMAIL + '\n\n' +
    '🔜 ඊළඟ පියවර: Deploy > New Deployment > Web App'
  );
}


// ============================================================
// 📖 DATA RETRIEVAL — දත්ත ලබා ගැනීම
// ============================================================

// Sheet එකක data objects array එකක් ලෙස ලබා ගැනීම
function getSheetData(sheetName) {
  const sheet = SPREADSHEET.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  return rows.map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) {
      // Date objects → ISO string
      obj[h] = (row[i] instanceof Date) ? row[i].toISOString() : row[i];
    });
    return obj;
  });
}

// සියලුම data එකවර ලබා ගැනීම
function getAllData() {
  return {
    products: getSheetData('Products'),
    sales: getSheetData('Sales'),
    returns: getSheetData('Returns'),
    restockLog: getSheetData('RestockLog'),
    _timestamp: new Date().toISOString()
  };
}


// ============================================================
// 🌐 GET ENDPOINT — Frontend data ඉල්ලන විට
// ============================================================
function doGet(e) {
  var action = e && e.parameter && e.parameter.action;
  var result;

  switch (action) {
    case 'getProducts':
      result = { success: true, data: getSheetData('Products') };
      break;
    case 'getSales':
      result = { success: true, data: getSheetData('Sales') };
      break;
    case 'getReturns':
      result = { success: true, data: getSheetData('Returns') };
      break;
    case 'getRestockLog':
      result = { success: true, data: getSheetData('RestockLog') };
      break;
    default:
      // සියලුම data ලබා දෙන්න
      result = { success: true, data: getAllData() };
      break;
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}


// ============================================================
// 📮 POST ENDPOINT — Frontend data එවන විට
// ============================================================
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action = payload.action;

    // ══════════════════════════════════════════════
    // 📦 ADD PRODUCT — නව භාණ්ඩයක් එකතු කිරීම
    // ══════════════════════════════════════════════
    if (action === 'addProduct') {
      var p = payload.data;
      var sheet = SPREADSHEET.getSheetByName('Products');

      // Validation — අත්‍යවශ්‍ය fields check කරන්න
      if (!p.productId || !p.productName || !p.category || !p.brand) {
        return _jsonResponse({ success: false, error: 'අත්‍යවශ්‍ය fields නැත. ProductID, Name, Category, Brand අවශ්‍යයි.' });
      }

      // Duplicate check — එම ID දැනටමත් තිබේදැයි බලන්න
      var existing = getSheetData('Products');
      if (existing.find(function(x) { return x.ProductID === p.productId; })) {
        return _jsonResponse({ success: false, error: 'Product ID "' + p.productId + '" දැනටමත් පවතී.' });
      }

      var unitPrice = Number(p.unitPrice) || 0;
      var costPrice = Number(p.costPrice) || 0;
      var stockQty = Number(p.stockQty) || 0;
      var minStock = Number(p.minStockLevel) || 5;
      var warrantyMonths = Number(p.warrantyMonths) || 0;
      var status = stockQty <= minStock ? 'Low Stock' : 'Active';

      // Sheet එකට row එකතු කරන්න
      sheet.appendRow([
        p.productId,
        p.productName,
        p.category,
        p.brand,
        p.model || '',
        unitPrice,
        costPrice,
        stockQty,
        minStock,
        warrantyMonths,
        p.supplier || '',
        new Date(),
        status
      ]);

      sheet.autoResizeColumns(1, sheet.getLastColumn());

      // Low stock නම් email alert එවන්න
      if (status === 'Low Stock') {
        checkAndAlertLowStock(p.productId, p.productName, stockQty, minStock);
      }

      Logger.log('✅ Product added: ' + p.productId + ' — ' + p.productName);

      return _jsonResponse({
        success: true,
        message: p.productName + ' සාර්ථකව එකතු කරන ලදී!',
        productId: p.productId,
        status: status
      });
    }


    // ══════════════════════════════════════════════
    // ✏️ UPDATE PRODUCT — භාණ්ඩයක් සංස්කරණය කිරීම
    // ══════════════════════════════════════════════
    if (action === 'updateProduct') {
      var p = payload.data;
      var sheet = SPREADSHEET.getSheetByName('Products');

      if (!p.productId) {
        return _jsonResponse({ success: false, error: 'Product ID අවශ්‍යයි.' });
      }

      // Product එක සොයන්න
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
      var idColIdx = headers.indexOf('ProductID');
      var rowIndex = -1;

      for (var i = 0; i < allData.length; i++) {
        if (String(allData[i][idColIdx]).trim() === String(p.productId).trim()) {
          rowIndex = i;
          break;
        }
      }

      if (rowIndex === -1) {
        return _jsonResponse({ success: false, error: 'Product "' + p.productId + '" සොයාගත නොහැක.' });
      }

      var actualRow = rowIndex + 2; // +2 because header row = 1, index starts at 0

      // Update fields — ඇති fields පමණක් update කරන්න
      var colMap = {};
      headers.forEach(function(h, idx) { colMap[h] = idx + 1; }); // 1-based column index

      if (p.productName !== undefined) sheet.getRange(actualRow, colMap['ProductName']).setValue(p.productName);
      if (p.category !== undefined)    sheet.getRange(actualRow, colMap['Category']).setValue(p.category);
      if (p.brand !== undefined)       sheet.getRange(actualRow, colMap['Brand']).setValue(p.brand);
      if (p.model !== undefined)       sheet.getRange(actualRow, colMap['Model']).setValue(p.model);
      if (p.unitPrice !== undefined)   sheet.getRange(actualRow, colMap['UnitPrice']).setValue(Number(p.unitPrice) || 0);
      if (p.costPrice !== undefined)   sheet.getRange(actualRow, colMap['CostPrice']).setValue(Number(p.costPrice) || 0);
      if (p.supplier !== undefined)    sheet.getRange(actualRow, colMap['Supplier']).setValue(p.supplier);
      if (p.warrantyMonths !== undefined) sheet.getRange(actualRow, colMap['WarrantyMonths']).setValue(Number(p.warrantyMonths) || 0);

      // Stock update කරනවා නම් status එකත් update කරන්න
      if (p.stockQty !== undefined) {
        var newQty = Number(p.stockQty) || 0;
        var minS = (p.minStockLevel !== undefined) ? Number(p.minStockLevel) : Number(allData[rowIndex][colMap['MinStockLevel'] - 1]) || 5;
        sheet.getRange(actualRow, colMap['StockQty']).setValue(newQty);
        sheet.getRange(actualRow, colMap['MinStockLevel']).setValue(minS);

        var newStatus = newQty <= minS ? 'Low Stock' : 'Active';
        sheet.getRange(actualRow, colMap['Status']).setValue(newStatus);

        if (newStatus === 'Low Stock') {
          checkAndAlertLowStock(p.productId, p.productName || allData[rowIndex][colMap['ProductName'] - 1], newQty, minS);
        }
      } else if (p.minStockLevel !== undefined) {
        var minS = Number(p.minStockLevel) || 5;
        sheet.getRange(actualRow, colMap['MinStockLevel']).setValue(minS);
        var curQty = Number(allData[rowIndex][colMap['StockQty'] - 1]) || 0;
        sheet.getRange(actualRow, colMap['Status']).setValue(curQty <= minS ? 'Low Stock' : 'Active');
      }

      Logger.log('✅ Product updated: ' + p.productId);

      return _jsonResponse({
        success: true,
        message: p.productId + ' සාර්ථකව සංස්කරණය කරන ලදී!',
        productId: p.productId
      });
    }


    // ══════════════════════════════════════════════
    // 🗑️ DELETE PRODUCT — භාණ්ඩයක් මකා දැමීම
    // ══════════════════════════════════════════════
    if (action === 'deleteProduct') {
      var productId = String(payload.productId || '').trim();
      var sheet = SPREADSHEET.getSheetByName('Products');

      if (!productId) {
        return _jsonResponse({ success: false, error: 'Product ID අවශ්‍යයි.' });
      }

      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
      var idColIdx = headers.indexOf('ProductID');
      var nameColIdx = headers.indexOf('ProductName');
      var rowIndex = -1;
      var productName = '';

      for (var i = 0; i < allData.length; i++) {
        if (String(allData[i][idColIdx]).trim() === productId) {
          rowIndex = i;
          productName = allData[i][nameColIdx];
          break;
        }
      }

      if (rowIndex === -1) {
        return _jsonResponse({ success: false, error: 'Product "' + productId + '" සොයාගත නොහැක.' });
      }

      // Row delete කරන්න (rowIndex + 2 because header = row 1)
      sheet.deleteRow(rowIndex + 2);

      Logger.log('🗑️ Product deleted: ' + productId + ' — ' + productName);

      return _jsonResponse({
        success: true,
        message: productName + ' සාර්ථකව මකා දමන ලදී!',
        productId: productId,
        productName: productName
      });
    }


    // ══════════════════════════════════════════════
    // 🛒 ADD SALE — නව අලෙවියක් සැකසීම
    // ══════════════════════════════════════════════
    if (action === 'addSale') {
      // Concurrent sales වැළැක්වීමට lock භාවිතා කරනවා
      var lock = LockService.getScriptLock();
      try {
        lock.waitLock(15000);
      } catch (lockErr) {
        return _jsonResponse({ success: false, error: 'System busy. පසුව නැවත උත්සාහ කරන්න.' });
      }

      try {
        var order = payload.data;
        var items = order.items;
        var discPct = Number(order.discountPct) || 0;
        var custName = order.customerName || 'Walk-in';
        var custPhone = order.customerPhone || '';
        var payMethod = order.paymentMethod || 'Cash';
        var soldBy = order.soldBy || 'System';

        if (!items || items.length === 0) {
          return _jsonResponse({ success: false, error: 'භාණ්ඩ නැත. කරුණාකර items එකතු කරන්න.' });
        }

        var salesSheet = SPREADSHEET.getSheetByName('Sales');

        // ──── Next Sale ID generate කරන්න ────
        var maxNum = 0;
        if (salesSheet.getLastRow() >= 2) {
          var saleIds = salesSheet.getRange(2, 1, salesSheet.getLastRow() - 1, 1).getValues();
          saleIds.forEach(function(r) {
            var n = parseInt(String(r[0]).replace(/\D/g, '')) || 0;
            if (n > maxNum) maxNum = n;
          });
        }
        var saleId = 'S' + String(maxNum + 1).padStart(4, '0');

        // ──── Products validate + stock check ────
        var prodSheet = SPREADSHEET.getSheetByName('Products');
        var prodHeaders = prodSheet.getRange(1, 1, 1, prodSheet.getLastColumn()).getValues()[0];
        var prodData = prodSheet.getRange(2, 1, prodSheet.getLastRow() - 1, prodSheet.getLastColumn()).getValues();

        var validated = [];
        for (var idx = 0; idx < items.length; idx++) {
          var item = items[idx];
          var pi = -1;
          for (var j = 0; j < prodData.length; j++) {
            if (String(prodData[j][0]).trim() === String(item.productId).trim()) {
              pi = j;
              break;
            }
          }

          if (pi === -1) {
            return _jsonResponse({ success: false, error: 'භාණ්ඩය "' + item.productId + '" සොයාගත නොහැක.' });
          }

          var currentStock = Number(prodData[pi][7]) || 0;
          if (currentStock < item.qty) {
            return _jsonResponse({
              success: false,
              error: prodData[pi][1] + ': තොගයේ ඇත්තේ ' + currentStock + ' ක් පමණි. ' + item.qty + ' ක් ඉල්ලා ඇත.'
            });
          }

          var uPrice = Number(prodData[pi][5]) || 0;
          var lineTotal = Math.round(uPrice * item.qty * (1 - discPct / 100));

          validated.push({
            prodIdx: pi,
            productId: prodData[pi][0],
            productName: prodData[pi][1],
            category: prodData[pi][2],
            qty: item.qty,
            unitPrice: uPrice,
            lineTotal: lineTotal
          });
        }

        // ──── Sales rows write කරන්න ────
        var now = new Date();
        var saleRows = validated.map(function(v) {
          return [
            saleId, now, v.productId, v.productName, v.category,
            v.qty, v.unitPrice, discPct, v.lineTotal,
            custName, custPhone, payMethod, soldBy, ''
          ];
        });

        salesSheet.getRange(
          salesSheet.getLastRow() + 1, 1,
          saleRows.length, saleRows[0].length
        ).setValues(saleRows);

        // ──── Stock update + Low stock check ────
        var lowStockAlerts = [];
        for (var vi = 0; vi < validated.length; vi++) {
          var v = validated[vi];
          var row = v.prodIdx + 2;
          var newQty = Number(prodData[v.prodIdx][7]) - v.qty;
          var minStock = Number(prodData[v.prodIdx][8]) || 0;

          // Stock qty update
          prodSheet.getRange(row, 8).setValue(newQty);

          // Status update
          if (newQty <= minStock) {
            prodSheet.getRange(row, 13).setValue('Low Stock');
            lowStockAlerts.push({
              id: v.productId,
              name: v.productName,
              qty: newQty,
              min: minStock
            });
          }

          // In-memory data update for subsequent iterations
          prodData[v.prodIdx][7] = newQty;
        }

        // ──── Low stock email alerts ────
        lowStockAlerts.forEach(function(alert) {
          checkAndAlertLowStock(alert.id, alert.name, alert.qty, alert.min);
        });

        // ──── Response ────
        var grandTotal = validated.reduce(function(s, v) { return s + v.lineTotal; }, 0);
        var subtotal = validated.reduce(function(s, v) { return s + (v.unitPrice * v.qty); }, 0);

        Logger.log('✅ Sale completed: ' + saleId + ' — Rs. ' + grandTotal);

        return _jsonResponse({
          success: true,
          saleId: saleId,
          date: now.toISOString(),
          items: validated.map(function(v) {
            return {
              productId: v.productId,
              productName: v.productName,
              qty: v.qty,
              unitPrice: v.unitPrice,
              lineTotal: v.lineTotal
            };
          }),
          subtotal: subtotal,
          discountPct: discPct,
          discountAmount: subtotal - grandTotal,
          grandTotal: grandTotal,
          customerName: custName,
          customerPhone: custPhone,
          paymentMethod: payMethod,
          soldBy: soldBy,
          lowStockAlerts: lowStockAlerts.length > 0 ? lowStockAlerts : undefined
        });

      } finally {
        lock.releaseLock();
      }
    }


    // ══════════════════════════════════════════════
    // 🔄 PROCESS RETURN — ආපසු ලැබීමක් සැකසීම
    // ══════════════════════════════════════════════
    if (action === 'processReturn') {
      var lock = LockService.getScriptLock();
      try {
        lock.waitLock(15000);
      } catch (lockErr) {
        return _jsonResponse({ success: false, error: 'System busy. පසුව නැවත උත්සාහ කරන්න.' });
      }

      try {
        var ret = payload.data;
        var saleId = String(ret.saleId || '').trim().toUpperCase();
        var productId = String(ret.productId || '').trim();
        var returnQty = Number(ret.qty) || 0;
        var reason = ret.reason || 'හේතුවක් සඳහන් කර නැත';
        var processedBy = ret.processedBy || 'Admin';

        // Validation
        if (!saleId || !productId || returnQty <= 0) {
          return _jsonResponse({
            success: false,
            error: 'Sale ID, Product ID, සහ valid ප්‍රමාණය අවශ්‍යයි.'
          });
        }

        // ──── Sale record සොයන්න ────
        var salesSheet = SPREADSHEET.getSheetByName('Sales');
        var salesHeaders = salesSheet.getRange(1, 1, 1, salesSheet.getLastColumn()).getValues()[0];
        var salesData = salesSheet.getRange(2, 1, salesSheet.getLastRow() - 1, salesSheet.getLastColumn()).getValues();

        var saleRowIndex = -1;
        var saleLine = null;

        for (var i = 0; i < salesData.length; i++) {
          if (String(salesData[i][0]).trim().toUpperCase() === saleId &&
              String(salesData[i][2]).trim() === productId) {
            saleRowIndex = i;
            saleLine = {};
            salesHeaders.forEach(function(h, ci) {
              saleLine[h] = salesData[i][ci] instanceof Date ? salesData[i][ci].toISOString() : salesData[i][ci];
            });
            break;
          }
        }

        if (!saleLine) {
          return _jsonResponse({
            success: false,
            error: saleId + ' / ' + productId + ' — ගැලපෙන අලෙවි වාර්තාවක් සොයාගත නොහැක.'
          });
        }

        // ──── දැනටමත් return කළ ප්‍රමාණය check කරන්න ────
        var returnsData = getSheetData('Returns');
        var alreadyReturned = returnsData
          .filter(function(r) {
            return String(r.SaleID).toUpperCase() === saleId &&
                   r.ProductID === productId &&
                   r.Status === 'Completed';
          })
          .reduce(function(sum, r) { return sum + (Number(r.Qty) || 0); }, 0);

        var soldQty = Number(saleLine.Qty) || 0;
        var maxReturnable = soldQty - alreadyReturned;

        if (maxReturnable <= 0) {
          return _jsonResponse({
            success: false,
            error: 'මෙම අලෙවිය සඳහා සියලුම units දැනටමත් ආපසු ලබා ඇත.'
          });
        }

        if (returnQty > maxReturnable) {
          return _jsonResponse({
            success: false,
            error: 'උපරිම ආපසු ලබා ගත හැකි ප්‍රමාණය: ' + maxReturnable
          });
        }

        // ──── Refund ගණනය ────
        var unitPrice = Number(saleLine.UnitPrice) || 0;
        var discPct = Number(saleLine.DiscountPct) || 0;
        var refundPerUnit = Math.round(unitPrice * (1 - discPct / 100));
        var refundAmount = refundPerUnit * returnQty;

        // ──── Return ID generate කරන්න ────
        var retSheet = SPREADSHEET.getSheetByName('Returns');
        var maxRetNum = 0;
        if (retSheet.getLastRow() >= 2) {
          var retIds = retSheet.getRange(2, 1, retSheet.getLastRow() - 1, 1).getValues();
          retIds.forEach(function(r) {
            var n = parseInt(String(r[0]).replace(/\D/g, '')) || 0;
            if (n > maxRetNum) maxRetNum = n;
          });
        }
        var returnId = 'R' + String(maxRetNum + 1).padStart(4, '0');

        // ──── Returns sheet එකට write කරන්න ────
        retSheet.appendRow([
          returnId,
          new Date(),
          saleId,
          productId,
          saleLine.ProductName || '',
          returnQty,
          reason,
          refundAmount,
          processedBy,
          'Completed'
        ]);
        retSheet.autoResizeColumns(1, retSheet.getLastColumn());

        // ──── Sales sheet එකේ ReturnStatus update කරන්න ────
        var rsColIdx = salesHeaders.indexOf('ReturnStatus');
        if (rsColIdx !== -1) {
          var newReturnTotal = alreadyReturned + returnQty;
          var newReturnStatus = newReturnTotal >= soldQty ? 'Returned' : 'Partial Return';
          salesSheet.getRange(saleRowIndex + 2, rsColIdx + 1).setValue(newReturnStatus);
        }

        // ──── Products sheet එකේ stock restore කරන්න ────
        var prodSheet = SPREADSHEET.getSheetByName('Products');
        var prodData = prodSheet.getRange(2, 1, prodSheet.getLastRow() - 1, prodSheet.getLastColumn()).getValues();

        for (var pi = 0; pi < prodData.length; pi++) {
          if (String(prodData[pi][0]).trim() === productId) {
            var newQty = Number(prodData[pi][7]) + returnQty;
            var minStock = Number(prodData[pi][8]) || 0;

            // Stock qty restore
            prodSheet.getRange(pi + 2, 8).setValue(newQty);

            // Status update — Low Stock → Active (if stock restored above minimum)
            if (newQty > minStock && String(prodData[pi][12]) === 'Low Stock') {
              prodSheet.getRange(pi + 2, 13).setValue('Active');
            }
            break;
          }
        }

        Logger.log('✅ Return processed: ' + returnId + ' — Refund Rs. ' + refundAmount);

        return _jsonResponse({
          success: true,
          returnId: returnId,
          refundAmount: refundAmount,
          productName: saleLine.ProductName,
          productId: productId,
          qty: returnQty,
          customerName: saleLine.CustomerName || '',
          saleId: saleId,
          message: returnId + ' සාර්ථකව සකසන ලදී. ආපසු මුදල: Rs. ' + refundAmount.toLocaleString()
        });

      } finally {
        lock.releaseLock();
      }
    }


    // ══════════════════════════════════════════════
    // ❌ UNKNOWN ACTION
    // ══════════════════════════════════════════════
    return _jsonResponse({
      success: false,
      error: 'නොදන්නා action: ' + action + '. addProduct, updateProduct, deleteProduct, addSale, processReturn භාවිතා කරන්න.'
    });

  } catch (err) {
    Logger.log('❌ doPost Error: ' + err.toString());
    return _jsonResponse({ success: false, error: 'Server error: ' + err.toString() });
  }
}


// ============================================================
// 🔧 HELPER — JSON Response wrapper
// ============================================================
function _jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


// ============================================================
// 📋 CUSTOM MENU — Apps Script editor menu
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Mahesh App')
    .addItem('📋 Initialize All Sheets (සියලුම Sheets සාදන්න)', 'setupSheets')
    .addSeparator()
    .addItem('📧 Test Low Stock Email (Email test)', 'testLowStockEmail')
    .addItem('📊 View Data Summary (දත්ත සාරාංශය)', 'showDataSummary')
    .addSeparator()
    .addItem('🌐 Deploy Instructions (Deploy උපදෙස්)', 'showDeployInfo')
    .addToUi();
}


// ============================================================
// 🧪 TEST FUNCTIONS — පරීක්ෂණ functions
// ============================================================

// Low stock email test
function testLowStockEmail() {
  var sent = sendLowStockAlert('TEST001', 'Test Product — පරීක්ෂණ භාණ්ඩය', 2, 5);
  if (sent) {
    SpreadsheetApp.getUi().alert(
      '📧 Test email සාර්ථකව එවන ලදී!\n\n' +
      'To: ' + ADMIN_EMAIL + '\n\n' +
      '⚠️ ඔබගේ inbox එක check කරන්න.\n' +
      'ADMIN_EMAIL variable එක ඔබගේ සැබෑ email එකට වෙනස් කරන්න.'
    );
  } else {
    SpreadsheetApp.getUi().alert(
      '❌ Email එවීම අසාර්ථකයි.\n\n' +
      'MailApp permissions check කරන්න.\n' +
      'ADMIN_EMAIL: ' + ADMIN_EMAIL
    );
  }
}

// Data summary view
function showDataSummary() {
  var data = getAllData();
  var prodCount = data.products.length;
  var salesCount = data.sales.length;
  var returnsCount = data.returns.length;
  var restockCount = data.restockLog.length;

  var uniqueSales = [];
  data.sales.forEach(function(s) {
    if (uniqueSales.indexOf(s.SaleID) === -1) uniqueSales.push(s.SaleID);
  });

  var totalRevenue = data.sales.reduce(function(sum, s) {
    return sum + (Number(s.TotalAmount) || 0);
  }, 0);

  var totalRefunds = data.returns.reduce(function(sum, r) {
    return sum + (Number(r.RefundAmount) || 0);
  }, 0);

  var lowStockItems = data.products.filter(function(p) {
    return p.Status === 'Low Stock' || Number(p.StockQty) <= Number(p.MinStockLevel);
  });

  var msg =
    '📊 දත්ත සාරාංශය — Data Summary\n' +
    '═══════════════════════════════\n\n' +
    '📦 Products (භාණ්ඩ): ' + prodCount + '\n' +
    '🛒 Sales Records (අලෙවි): ' + salesCount + ' rows (' + uniqueSales.length + ' orders)\n' +
    '🔄 Returns (ආපසු): ' + returnsCount + '\n' +
    '📥 Restock Log (තොග): ' + restockCount + '\n\n' +
    '💰 Total Revenue (මුළු ආදායම): Rs. ' + totalRevenue.toLocaleString() + '\n' +
    '💸 Total Refunds (ආපසු මුදල): Rs. ' + totalRefunds.toLocaleString() + '\n' +
    '📈 Net Revenue (ශුද්ධ ආදායම): Rs. ' + (totalRevenue - totalRefunds).toLocaleString() + '\n\n';

  if (lowStockItems.length > 0) {
    msg += '⚠️ Low Stock Items (අඩු තොග):\n';
    lowStockItems.forEach(function(p) {
      msg += '   • ' + p.ProductName + ' — ' + p.StockQty + ' units (min: ' + p.MinStockLevel + ')\n';
    });
  } else {
    msg += '✅ සියලුම භාණ්ඩ ප්‍රමාණවත් තොගයක් ඇත.';
  }

  SpreadsheetApp.getUi().alert(msg);
}

// Deploy instructions
function showDeployInfo() {
  SpreadsheetApp.getUi().alert(
    '🌐 Web App Deploy කිරීමේ උපදෙස්\n' +
    '══════════════════════════════════\n\n' +
    '1️⃣  Deploy > New Deployment click කරන්න\n\n' +
    '2️⃣  Type: "Web app" select කරන්න\n\n' +
    '3️⃣  Settings:\n' +
    '    • Description: Liyanage Electronics API\n' +
    '    • Execute as: Me (මා ලෙස)\n' +
    '    • Who has access: Anyone (ඕනෑම කෙනෙක්)\n\n' +
    '4️⃣  Deploy button click කරන්න\n\n' +
    '5️⃣  Web app URL copy කරන්න\n\n' +
    '6️⃣  index.html file එකේ:\n' +
    '    API_URL = "ඔබගේ_URL_මෙතන_paste_කරන්න"\n\n' +
    '═══════════════════════════════════\n' +
    '⚡ Code එක වෙනස් කළ පසු Deploy > Manage Deployments\n' +
    '   > Edit (pencil icon) > Version: New > Deploy\n\n' +
    '📧 Low stock alerts: ' + ADMIN_EMAIL + '\n' +
    '   (ADMIN_EMAIL variable එක වෙනස් කරන්න)'
  );
}
