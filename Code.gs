// ============================================================================
// Fashion Shop POS & Inventory Management - Google Apps Script (Code.gs)
// With Secure User Authentication, Admin ID Control, & Google Sheets Sync
// ============================================================================

/**
 * Serves the POS web application when the deployed Web App URL is visited
 */
function doGet(e) {
  try {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle("Fashion Shop POS & Inventory Management")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    return ContentService.createTextOutput("POS API is running").setMimeType(ContentService.MimeType.TEXT);
  }
}

/**
 * Adds a custom menu directly inside Google Sheets to open POS & manage accounts
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏪 POS System')
    .addItem('সফটওয়্যার ড্যাশবোর্ড খুলুন (Open POS)', 'openPOSDialog')
    .addItem('পুরাতন ইনভয়েস সার্চ ও এডিট (Old Invoice Search & Edit)', 'openPOSDialog')
    .addItem('নতুন প্রোডাক্ট এন্ট্রি করুন (New Product Entry)', 'promptNewProductEntry')
    .addItem('শিট ট্যাব ও ইউজার লিস্ট প্রস্তুত করুন (Setup Sheets & Users)', 'setupShopSheets')
    .addItem('পূর্বের ডেমো ডাটা মুছে শিট ফ্রেশ করুন (Clear Demo Entries)', 'manualCleanDemoData')
    .addItem('প্রোডাক্ট শিট সাজান ও ফরম্যাট করুন (Format Product Sheet)', 'formatAndSyncProductSheet')
    .addItem('স্টক শিট সাজান ও ফরম্যাট করুন (Format Stock Sheet)', 'formatAndSyncStockSheet')
    .addItem('অ্যাডমিন পাসওয়ার্ড রিসেট করুন (Reset Admin Password)', 'resetAdminPasswordPrompt')
    .addToUi();
}

/**
 * Opens the POS inside a Google Sheets popup dialog
 */
function openPOSDialog() {
  var html = HtmlService.createHtmlOutputFromFile('Index')
    .setWidth(1250)
    .setHeight(880);
  SpreadsheetApp.getUi().showModalDialog(html, 'Fashion Shop POS & Inventory Management');
}

/**
 * Ensures all 7 sheets exist with formatted colored headers
 * Including the Users sheet for authentication & role control.
 * DOES NOT insert any fake or demo products/stock/sales.
 */
function setupShopSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheets = [
    {
      name: "Users",
      headers: ["Username", "Password", "Full Name", "Role", "Designation", "Created At"],
      color: "#0f172a",
      fontColor: "#ffffff"
    },
    {
      name: "Products",
      headers: ["Date", "ID", "Purchase No", "Supplier", "Product", "Qty", "Purchase Price", "Sale Price", "Total"],
      color: "#e2e8f0",
      fontColor: "#1e293b",
      isLightHeader: true
    },
    {
      name: "Stock",
      headers: ["Date", "Id", "Purchase No", "Product Name", "Category", "Purchased Qty", "Sold Qty", "Current Stock", "Sale Price", "Purchase Price"],
      color: "#e2e8f0",
      fontColor: "#1e293b",
      isLightHeader: true
    },
    {
      name: "Sales",
      headers: ["Date", "Invoice No", "Customer Name", "Phone", "Product Name", "Qty", "Sale Price", "Discount", "Purchase Price", "Payment", "Remarks", "Payment Method", "Profit"],
      color: "#1e293b",
      fontColor: "#ffffff"
    },
    {
      name: "Customers",
      headers: ["Customer ID", "Customer Name", "Phone", "Total Purchases", "Total Paid", "Current Due", "Status"],
      color: "#dc2626",
      fontColor: "#ffffff"
    },
    {
      name: "Due_Collections",
      headers: ["Collection Date", "Collection ID", "Customer Name", "Phone", "Amount Collected", "Payment Method", "Remaining Due", "Collected By", "Remarks"],
      color: "#d97706",
      fontColor: "#ffffff"
    },
    {
      name: "Expenses",
      headers: ["Date", "Expense ID", "Category", "Amount", "Note", "Payment Method", "Logged By"],
      color: "#0d9488",
      fontColor: "#ffffff"
    },
    {
      name: "Settings",
      headers: ["Setting Key", "Setting Value", "Updated At"],
      color: "#334155",
      fontColor: "#ffffff"
    }
  ];

  sheets.forEach(function(sh) {
    var sheet = ss.getSheetByName(sh.name) || ss.insertSheet(sh.name);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(sh.headers);
      sheet.getRange(1, 1, 1, sh.headers.length)
        .setBackground(sh.color)
        .setFontColor(sh.fontColor || "#ffffff")
        .setFontWeight("bold");
    }
  });

  // Specifically format the Products sheet (headers only, no demo rows)
  var prodSheet = ss.getSheetByName("Products");
  if (prodSheet) {
    if (prodSheet.getLastRow() === 0) {
      prodSheet.appendRow(["Date", "ID", "Purchase No", "Supplier", "Product", "Qty", "Purchase Price", "Sale Price", "Total"]);
    }
    formatProductSheet(prodSheet);
  }

  // Specifically format the Stock sheet (headers only, no demo rows)
  var stockSheet = ss.getSheetByName("Stock");
  if (stockSheet) {
    if (stockSheet.getLastRow() === 0) {
      stockSheet.appendRow(["Date", "Id", "Purchase No", "Product Name", "Category", "Purchased Qty", "Sold Qty", "Current Stock", "Sale Price", "Purchase Price"]);
    }
    formatStockSheet(stockSheet);
  }

  // Automatically clean up any pre-seeded demo rows from previous versions
  // (disabled: data must never be deleted automatically) removeDemoEntriesFromAllSheets();

  // Seed default 1 Admin + 5 Staff Accounts if Users sheet is empty
  var userSheet = ss.getSheetByName("Users");
  if (userSheet.getLastRow() <= 1) {
    var defaultAccounts = [
      ["admin", "admin1234", "সুপার অ্যাডমিন (মালিক)", "admin", "দোকানের মালিক ও চিফ অ্যাডমিন", new Date().toISOString().slice(0,10)],
      ["cashier1", "user1234", "ক্যাশিয়ার ১", "cashier", "কাউন্টার ১ বিক্রয়কর্মী", new Date().toISOString().slice(0,10)],
      ["cashier2", "user1234", "ক্যাশিয়ার ২", "cashier", "কাউন্টার ২ বিক্রয়কর্মী", new Date().toISOString().slice(0,10)],
      ["salesman", "user1234", "বিক্রয় কর্মী", "cashier", "সেলস ও ফ্লোর এক্সিকিউটিভ", new Date().toISOString().slice(0,10)],
      ["manager", "user1234", "স্টোর ম্যানেজার", "manager", "স্টোর ও স্টক সুপারভাইজার", new Date().toISOString().slice(0,10)],
      ["inventory", "user1234", "স্টক ইনচার্জ", "manager", "ইনভেন্টরি ও গোডাউন ইনচার্জ", new Date().toISOString().slice(0,10)]
    ];

    defaultAccounts.forEach(function(row) {
      userSheet.appendRow(row);
    });

    userSheet.getRange(2, 1, defaultAccounts.length, 6).setFontSize(10);
  }

  // Seed default Company Info in Settings sheet if empty
  var settingsSheet = ss.getSheetByName("Settings");
  if (settingsSheet && settingsSheet.getLastRow() <= 1) {
    var defaultSettings = [
      ["name", "Men's & women Fashion", new Date().toISOString()],
      ["address", "ঢাকা, বাংলাদেশ", new Date().toISOString()],
      ["phone", "01322-388151", new Date().toISOString()],
      ["tagline", "স্মার্ট শপ ইনভেন্টরি ও ক্যাশ মেমো", new Date().toISOString()],
      ["email", "fashionpos@gmail.com", new Date().toISOString()]
    ];

    defaultSettings.forEach(function(row) {
      settingsSheet.appendRow(row);
    });

    settingsSheet.getRange(2, 1, defaultSettings.length, 3).setFontSize(10);
  }

  SpreadsheetApp.getActive().toast("সকল ৮টি শিট ট্যাব সফলভাবে প্রস্তুত হয়েছে (কোনো ডেমো এন্ট্রি ছাড়া)!", "POS System", 5);
}

/**
 * Removes pre-seeded demo entries from Products, Stock, Sales, Customers, Due, Expenses.
 * Real user entries are strictly preserved and untouched.
 */
function removeDemoEntriesFromAllSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Clean demo rows in Products sheet
  var prodSheet = ss.getSheetByName("Products");
  if (prodSheet && prodSheet.getLastRow() > 1) {
    var pData = prodSheet.getDataRange().getValues();
    var demoPurNos = ["PUR00006", "PUR13178", "PUR00009", "PUR0003"];
    var demoProdIds = ["P0001", "P0002", "P0003", "P0004", "P0005"];
    for (var r = pData.length - 1; r >= 1; r--) {
      var row = pData[r];
      var pId = String(row[1] || "").trim();
      var purNo = String(row[2] || "").trim();
      var supp = String(row[3] || "").trim().toLowerCase();
      var prod = String(row[4] || "").trim().toLowerCase();
      
      var isDemo = (supp === "abcd" && (demoPurNos.indexOf(purNo) !== -1 || purNo === "N/A")) ||
                   (demoProdIds.indexOf(pId) !== -1 && supp === "abcd");
      if (isDemo) {
        prodSheet.deleteRow(r + 1);
      }
    }
  }

  // 2. Clean demo rows in Stock sheet
  var stockSheet = ss.getSheetByName("Stock");
  if (stockSheet && stockSheet.getLastRow() > 1) {
    var sData = stockSheet.getDataRange().getValues();
    var demoStockPurs = ["PUR13181", "PUR13175", "PUR13170", "PUR13165", "PUR13160", "PUR00006", "PUR13178", "PUR00009"];
    for (var r2 = sData.length - 1; r2 >= 1; r2--) {
      var row2 = sData[r2];
      var sPurNo = String(row2[2] || "").trim();
      var sProd = String(row2[3] || "").trim();
      var isStockDemo = demoStockPurs.indexOf(sPurNo) !== -1 || 
                        sProd === "পাঞ্জাবি প্রিমিয়াম" || sProd === "জর্জেট শাড়ি" || 
                        sProd === "ফরমাল শার্ট" || sProd === "সিল্ক থ্রি-পিস";
      if (isStockDemo) {
        stockSheet.deleteRow(r2 + 1);
      }
    }
  }

  // 3. Clean demo rows in Sales sheet
  var salesSheet = ss.getSheetByName("Sales");
  if (salesSheet && salesSheet.getLastRow() > 1) {
    var slData = salesSheet.getDataRange().getValues();
    for (var r3 = slData.length - 1; r3 >= 1; r3--) {
      var row3 = slData[r3];
      var inv = String(row3[1] || "").trim();
      var cName = String(row3[2] || "").trim();
      var phone = String(row3[3] || "").trim();
      var isSaleDemo = (inv === "INV-1001" || inv === "INV-1002") && 
                       (cName === "আব্দুল করিম" || cName === "ফারহানা আক্তার" || phone === "01711000000" || phone === "01811000000");
      if (isSaleDemo) {
        salesSheet.deleteRow(r3 + 1);
      }
    }
  }

  // 4. Clean demo customers
  var custSheet = ss.getSheetByName("Customers");
  if (custSheet && custSheet.getLastRow() > 1) {
    var cData = custSheet.getDataRange().getValues();
    for (var r4 = cData.length - 1; r4 >= 1; r4--) {
      var cRow = cData[r4];
      var cId = String(cRow[0] || "").trim();
      var custName = String(cRow[1] || "").trim();
      if ((cId === "C001" || cId === "C002") && (custName === "আব্দুল করিম" || custName === "ফারহানা আক্তার")) {
        custSheet.deleteRow(r4 + 1);
      }
    }
  }

  // 5. Clean demo due collections
  var dueSheet = ss.getSheetByName("Due_Collections");
  if (dueSheet && dueSheet.getLastRow() > 1) {
    var dData = dueSheet.getDataRange().getValues();
    for (var r5 = dData.length - 1; r5 >= 1; r5--) {
      var dRow = dData[r5];
      var dId = String(dRow[1] || "").trim();
      var dCust = String(dRow[2] || "").trim();
      if (dId === "DUE-01" && dCust === "ফারহানা আক্তার") {
        dueSheet.deleteRow(r5 + 1);
      }
    }
  }

  // 6. Clean demo expenses
  var expSheet = ss.getSheetByName("Expenses");
  if (expSheet && expSheet.getLastRow() > 1) {
    var eData = expSheet.getDataRange().getValues();
    for (var r6 = eData.length - 1; r6 >= 1; r6--) {
      var eRow = eData[r6];
      var eId = String(eRow[1] || "").trim();
      var eCat = String(eRow[2] || "").trim();
      if (eId === "EXP-01" && (eCat === "চা-নাস্তা / আপ্যায়ন" || eCat.indexOf("চা-নাস্তা") !== -1)) {
        expSheet.deleteRow(r6 + 1);
      }
    }
  }
}

/**
 * Manual menu trigger to clean all demo data
 */
function manualCleanDemoData() {
  removeDemoEntriesFromAllSheets();
  SpreadsheetApp.getActive().toast("শিট থেকে সকল পূর্বের ডেমো এন্ট্রি মুছে সম্পূর্ণ পরিষ্কার করা হয়েছে!", "POS System", 5);
}

/**
 * Secure Login Authentication function called from frontend
 */
function authenticateUser(username, password) {
  setupShopSheets();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var userSheet = ss.getSheetByName("Users");

  var cleanUsername = String(username || "").trim().toLowerCase();
  var cleanPassword = String(password || "").trim();

  if (!cleanUsername || !cleanPassword) {
    return { success: false, message: "দয়া করে ইউজার আইডি এবং পাসওয়ার্ড দিন।" };
  }

  var data = userSheet.getDataRange().getValues();
  // Header row is index 0: ["Username", "Password", "Full Name", "Role", "Designation", "Created At"]
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var u = String(row[0] || "").trim().toLowerCase();
    var p = String(row[1] || "").trim();

    if (u === cleanUsername && p === cleanPassword) {
      return {
        success: true,
        message: "লগইন সফল হয়েছে!",
        user: {
          username: String(row[0]),
          name: String(row[2] || row[0]),
          role: String(row[3] || "cashier"),
          designation: String(row[4] || "স্টাফ")
        }
      };
    }
  }

  return { success: false, message: "ভুল ইউজার আইডি অথবা পাসওয়ার্ড! আবার চেষ্টা করুন।" };
}

/**
 * Gets all user accounts (Admin only view)
 */
function getUsersList(adminUsername, adminPassword) {
  var auth = authenticateUser(adminUsername, adminPassword);
  if (!auth.success || auth.user.role !== "admin") {
    return { success: false, message: "শুধুমাত্র অ্যাডমিন ইউজার তালিকা দেখতে পারেন।" };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var userSheet = ss.getSheetByName("Users");
  var data = userSheet.getDataRange().getValues();
  var users = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    users.push({
      username: String(row[0]),
      password: String(row[1]),
      name: String(row[2]),
      role: String(row[3]),
      designation: String(row[4]),
      createdAt: String(row[5] || "")
    });
  }

  return { success: true, users: users };
}

/**
 * Create or update a user account (Admin only)
 */
function saveOrUpdateUser(adminUsername, adminPassword, userData) {
  var auth = authenticateUser(adminUsername, adminPassword);
  if (!auth.success || auth.user.role !== "admin") {
    return { success: false, message: "শুধুমাত্র অ্যাডমিন নতুন ইউজার তৈরি বা পরিবর্তন করতে পারবেন।" };
  }

  var targetUsername = String(userData.username || "").trim().toLowerCase();
  var targetPassword = String(userData.password || "").trim();
  var targetName = String(userData.name || targetUsername).trim();
  var targetRole = String(userData.role || "cashier").trim();
  var targetDesignation = String(userData.designation || "").trim();

  if (!targetUsername || !targetPassword) {
    return { success: false, message: "ইউজার আইডি ও পাসওয়ার্ড আবশ্যক।" };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var userSheet = ss.getSheetByName("Users");
  var data = userSheet.getDataRange().getValues();
  var foundRow = -1;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === targetUsername) {
      foundRow = i + 1; // 1-indexed for Sheets
      break;
    }
  }

  if (foundRow > 0) {
    // Update existing user
    userSheet.getRange(foundRow, 2).setValue(targetPassword);
    userSheet.getRange(foundRow, 3).setValue(targetName);
    userSheet.getRange(foundRow, 4).setValue(targetRole);
    userSheet.getRange(foundRow, 5).setValue(targetDesignation);
    return { success: true, message: "ইউজার '" + targetUsername + "' সফলভাবে আপডেট হয়েছে!" };
  } else {
    // Add new user
    userSheet.appendRow([
      targetUsername,
      targetPassword,
      targetName,
      targetRole,
      targetDesignation,
      new Date().toISOString().slice(0,10)
    ]);
    return { success: true, message: "নতুন ইউজার '" + targetUsername + "' সফলভাবে তৈরি হয়েছে!" };
  }
}

/**
 * Delete a user (Admin only, cannot delete self)
 */
function deleteUser(adminUsername, adminPassword, targetUsername) {
  var auth = authenticateUser(adminUsername, adminPassword);
  if (!auth.success || auth.user.role !== "admin") {
    return { success: false, message: "অনুমতি নেই।" };
  }

  var cleanTarget = String(targetUsername || "").trim().toLowerCase();
  if (cleanTarget === "admin" || cleanTarget === adminUsername.toLowerCase()) {
    return { success: false, message: "মেইন অ্যাডমিন একাউন্ট মুছে ফেলা যাবে না।" };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var userSheet = ss.getSheetByName("Users");
  var data = userSheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === cleanTarget) {
      userSheet.deleteRow(i + 1);
      return { success: true, message: "ইউজার '" + targetUsername + "' সফলভাবে মুছে ফেলা হয়েছে।" };
    }
  }

  return { success: false, message: "ইউজার পাওয়া যায়নি।" };
}

/**
 * Helper prompt from Sheets menu to quickly reset Admin password if forgotten
 */
function resetAdminPasswordPrompt() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt("Admin Password Reset", "নতুন অ্যাডমিন পাসওয়ার্ড লিখুন:", ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    var newPass = response.getResponseText().trim();
    if (newPass) {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var userSheet = ss.getSheetByName("Users");
      if (userSheet) {
        var data = userSheet.getDataRange().getValues();
        for (var i = 1; i < data.length; i++) {
          if (String(data[i][0]).trim().toLowerCase() === "admin") {
            userSheet.getRange(i + 1, 2).setValue(newPass);
            ui.alert("সফল!", "অ্যাডমিন পাসওয়ার্ড সফলভাবে '" + newPass + "' এ পরিবর্তিত হয়েছে।", ui.ButtonSet.OK);
            return;
          }
        }
      }
    }
  }
}

/**
 * Retrieves the Company Information (Name, Address, Phone, Tagline, Email) from Settings sheet or Properties
 */
function getCompanyInfo() {
  var defaultInfo = {
    name: "Men's & women Fashion",
    address: "ঢাকা, বাংলাদেশ",
    phone: "01322-388151",
    tagline: "স্মার্ট শপ ইনভেন্টরি ও ক্যাশ মেমো",
    email: "fashionpos@gmail.com"
  };

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Settings");
    if (!sheet) {
      sheet = ss.insertSheet("Settings");
      sheet.appendRow(["Setting Key", "Setting Value", "Updated At"]);
      sheet.getRange(1, 1, 1, 3).setBackground("#334155").setFontColor("#ffffff").setFontWeight("bold");
    }

    if (sheet.getLastRow() > 1) {
      var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
      var info = {};
      values.forEach(function(row) {
        var key = String(row[0] || "").trim();
        var val = String(row[1] || "").trim();
        if (key) {
          info[key] = val;
        }
      });
      return {
        name: info.name || defaultInfo.name,
        address: info.address || defaultInfo.address,
        phone: info.phone || defaultInfo.phone,
        tagline: info.tagline || defaultInfo.tagline,
        email: info.email || defaultInfo.email
      };
    }
  } catch (e) {
    console.error("Error reading Settings sheet:", e);
  }

  // Fallback to ScriptProperties
  try {
    var prop = PropertiesService.getScriptProperties().getProperty("COMPANY_INFO");
    if (prop) {
      var parsed = JSON.parse(prop);
      return Object.assign({}, defaultInfo, parsed);
    }
  } catch (e) {}

  return defaultInfo;
}

/**
 * Persists company info to Settings sheet and Script Properties
 */
function saveCompanyInfoDirect(companyData) {
  if (!companyData || typeof companyData !== "object") return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Settings") || ss.insertSheet("Settings");

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Setting Key", "Setting Value", "Updated At"]);
    sheet.getRange(1, 1, 1, 3).setBackground("#334155").setFontColor("#ffffff").setFontWeight("bold");
  }

  var rowIndices = {};
  if (sheet.getLastRow() > 1) {
    var vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    vals.forEach(function(r, idx) {
      var k = String(r[0] || "").trim();
      if (k) rowIndices[k] = idx + 2;
    });
  }

  var now = new Date().toISOString();
  var keys = ["name", "address", "phone", "tagline", "email"];
  keys.forEach(function(k) {
    var v = companyData[k] !== undefined ? String(companyData[k]) : "";
    if (rowIndices[k]) {
      sheet.getRange(rowIndices[k], 2, 1, 2).setValues([[v, now]]);
    } else {
      sheet.appendRow([k, v, now]);
    }
  });

  try {
    PropertiesService.getScriptProperties().setProperty("COMPANY_INFO", JSON.stringify(companyData));
  } catch (e) {}
}

/**
 * Saves Company Info from frontend (Admin permission check)
 */
function saveCompanyInfo(adminUsername, adminPassword, companyData) {
  var auth = authenticateUser(adminUsername, adminPassword);
  if (!auth.success || auth.user.role !== "admin") {
    return { success: false, message: "শুধুমাত্র অনুমোদিত অ্যাডমিন কোম্পানির তথ্য পরিবর্তন করতে পারবেন।" };
  }

  saveCompanyInfoDirect(companyData);
  return {
    success: true,
    message: "কোম্পানির তথ্য সফলভাবে গুগল শিটে সংরক্ষিত হয়েছে!",
    companyInfo: getCompanyInfo()
  };
}

/**
 * Loads all data from Google Sheets into the POS frontend
 */
function getShopData() {
  setupShopSheets();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = {
    products: [],
    stock: [],
    sales: [],
    customers: [],
    duePayments: [],
    expenses: []
  };

  // Products (Supports both 9-col screenshot format and 7-col legacy)
  var prodSheet = ss.getSheetByName("Products");
  if (prodSheet && prodSheet.getLastRow() > 1) {
    var lastCol = Math.max(7, prodSheet.getLastColumn());
    var prodValues = prodSheet.getRange(2, 1, prodSheet.getLastRow() - 1, lastCol).getValues();
    result.products = prodValues
      .filter(function(r) { return r && (String(r[0] || "").trim() !== "" || String(r[1] || "").trim() !== "" || String(r[4] || "").trim() !== ""); })
      .map(function(r) {
      if (lastCol >= 9 || (String(r[1]).indexOf("P") === 0 && isNaN(Number(r[1])))) {
        var qty = Number(r[5]) || 0;
        var buyPrice = Number(r[6]) || 0;
        var sPrice = Number(r[7]) || 0;
        var total = Number(r[8]) || (qty * buyPrice);
        return {
          date: String(r[0] || ""),
          id: String(r[1] || ""),
          productId: String(r[1] || ""),
          purchaseNo: String(r[2] || ""),
          supplier: String(r[3] || ""),
          product: String(r[4] || ""),
          name: String(r[4] || ""),
          qty: qty,
          currentStock: qty,
          purchasePrice: buyPrice,
          salePrice: sPrice,
          total: total,
          category: String(r[3] || "General"),
          status: qty < 10 ? "Low Stock" : "In Stock"
        };
      } else {
        return {
          id: String(r[0]),
          name: String(r[1]),
          category: String(r[2]),
          purchasePrice: Number(r[3]) || 0,
          salePrice: Number(r[4]) || 0,
          currentStock: Number(r[5]) || 0,
          qty: Number(r[5]) || 0,
          status: String(r[6] || "In Stock")
        };
      }
    });
  }

  // Stock
  var stockSheet = ss.getSheetByName("Stock");
  if (stockSheet && stockSheet.getLastRow() > 1) {
    var stockValues = stockSheet.getRange(2, 1, stockSheet.getLastRow() - 1, 10).getValues();
    result.stock = stockValues
      .filter(function(r) { return r && (String(r[1] || "").trim() !== "" || String(r[3] || "").trim() !== ""); })
      .map(function(r) {
      return {
        date: String(r[0] || ""),
        id: String(r[1] || ""),
        productId: String(r[1] || ""),
        purchaseNo: String(r[2] || ""),
        productName: String(r[3] || ""),
        category: String(r[4] || ""),
        purchasedQty: Number(r[5]) || 0,
        soldQty: Number(r[6]) || 0,
        currentStock: Number(r[7]) || 0,
        salePrice: Number(r[8]) || 0,
        purchasePrice: Number(r[9]) || 0
      };
    });
  }

  // Sales
  var salesSheet = ss.getSheetByName("Sales");
  if (salesSheet && salesSheet.getLastRow() > 1) {
    var salesValues = salesSheet.getRange(2, 1, salesSheet.getLastRow() - 1, 13).getValues();
    result.sales = salesValues
      .filter(function(r) { return r && (String(r[1] || "").trim() !== "" || String(r[2] || "").trim() !== ""); })
      .map(function(r) {
      return {
        date: String(r[0]),
        invoiceNo: String(r[1]),
        customerName: String(r[2]),
        phone: String(r[3]),
        productName: String(r[4]),
        qty: Number(r[5]) || 1,
        salePrice: Number(r[6]) || 0,
        discount: Number(r[7]) || 0,
        purchasePrice: Number(r[8]) || 0,
        payment: Number(r[9]) || 0,
        remarks: String(r[10] || "-"),
        paymentMethod: String(r[11] || "Cash"),
        profit: Number(r[12]) || 0,
        subtotal: Number(r[6]) || 0,
        netTotal: (Number(r[6]) || 0) - (Number(r[7]) || 0),
        due: Math.max(0, ((Number(r[6]) || 0) - (Number(r[7]) || 0)) - (Number(r[9]) || 0)),
        itemsSummary: String(r[4]) + " (" + (Number(r[5]) || 1) + "x" + (Number(r[6]) || 0) + ")"
      };
    });
  }

  // Customers
  var custSheet = ss.getSheetByName("Customers");
  if (custSheet && custSheet.getLastRow() > 1) {
    var custValues = custSheet.getRange(2, 1, custSheet.getLastRow() - 1, 7).getValues();
    result.customers = custValues
      .filter(function(r) { return r && (String(r[0] || "").trim() !== "" || String(r[1] || "").trim() !== ""); })
      .map(function(r) {
      return {
        id: String(r[0]),
        name: String(r[1]),
        phone: String(r[2]),
        totalPurchases: Number(r[3]) || 0,
        totalPaid: Number(r[4]) || 0,
        currentDue: Number(r[5]) || 0,
        status: String(r[6] || "Regular")
      };
    });
  }

  // Due Payments
  var dueSheet = ss.getSheetByName("Due_Collections");
  if (dueSheet && dueSheet.getLastRow() > 1) {
    var dueValues = dueSheet.getRange(2, 1, dueSheet.getLastRow() - 1, 9).getValues();
    result.duePayments = dueValues
      .filter(function(r) { return r && (String(r[1] || "").trim() !== "" || String(r[2] || "").trim() !== ""); })
      .map(function(r) {
      return {
        date: String(r[0]),
        id: String(r[1]),
        customerName: String(r[2]),
        phone: String(r[3]),
        amount: Number(r[4]) || 0,
        paymentMethod: String(r[5] || "Cash"),
        remainingDue: Number(r[6]) || 0,
        collectedBy: String(r[7] || ""),
        remarks: String(r[8] || "")
      };
    });
  }

  // Expenses
  var expSheet = ss.getSheetByName("Expenses");
  if (expSheet && expSheet.getLastRow() > 1) {
    var expValues = expSheet.getRange(2, 1, expSheet.getLastRow() - 1, 7).getValues();
    result.expenses = expValues
      .filter(function(r) { return r && (String(r[1] || "").trim() !== "" || String(r[2] || "").trim() !== ""); })
      .map(function(r) {
      return {
        date: String(r[0]),
        id: String(r[1]),
        category: String(r[2]),
        amount: Number(r[3]) || 0,
        note: String(r[4] || ""),
        paymentMethod: String(r[5] || "Cash"),
        loggedBy: String(r[6] || "")
      };
    });
  }

  result.companyInfo = getCompanyInfo();
  return JSON.stringify(result);
}

/**
 * Saves all records from the frontend into the Google Sheets
 */
function saveAllShopData_LEGACY_UNUSED(jsonData) {
  var data = typeof jsonData === "string" ? JSON.parse(jsonData) : jsonData;
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Products (Formatted identical to POS Products Grid Screenshot)
  if (data.products !== undefined) {
    var prodSheet = ss.getSheetByName("Products") || ss.insertSheet("Products");
    prodSheet.clear();
    var pHeaders = ["Date", "ID", "Purchase No", "Supplier", "Product", "Qty", "Purchase Price", "Sale Price", "Total"];
    var pRows = [pHeaders];
    (data.products || []).forEach(function(p) {
      var qty = Number(p.qty !== undefined ? p.qty : (p.currentStock !== undefined ? p.currentStock : 0));
      var buyPrice = Number(p.purchasePrice || 0);
      var sPrice = Number(p.salePrice || 0);
      var total = Number(p.total !== undefined ? p.total : (qty * buyPrice));
      var dateStr = p.date || new Date().toLocaleString();
      var pNo = p.purchaseNo || "N/A";
      var supp = p.supplier || p.category || "abcd";
      var prodName = p.product || p.name || "";
      var idStr = p.id || p.productId || "";
      pRows.push([dateStr, idStr, pNo, supp, prodName, qty, buyPrice, sPrice, total]);
    });
    prodSheet.getRange(1, 1, pRows.length, pHeaders.length).setValues(pRows);
    formatProductSheet(prodSheet);
  }

  // 2. Stock (Formatted identical to POS Stock Grid Screenshot)
  if (data.stock !== undefined) {
    var stockSheet = ss.getSheetByName("Stock") || ss.insertSheet("Stock");
    stockSheet.clear();
    var sHeaders = ["Date", "Id", "Purchase No", "Product Name", "Category", "Purchased Qty", "Sold Qty", "Current Stock", "Sale Price", "Purchase Price"];
    var sRows = [sHeaders];
    (data.stock || []).forEach(function(s) {
      sRows.push([
        s.date || new Date().toLocaleString(),
        s.id || s.productId || "",
        s.purchaseNo || "",
        s.productName || "",
        s.category || "",
        Number(s.purchasedQty !== undefined ? s.purchasedQty : (s.purchaseQty || 0)),
        Number(s.soldQty || 0),
        Number(s.currentStock || 0),
        Number(s.salePrice || 0),
        Number(s.purchasePrice || 0)
      ]);
    });
    stockSheet.getRange(1, 1, sRows.length, sHeaders.length).setValues(sRows);
    formatStockSheet(stockSheet);
  }

  // 3. Sales
  if (data.sales !== undefined) {
    var salesSheet = ss.getSheetByName("Sales") || ss.insertSheet("Sales");
    salesSheet.clear();
    var slHeaders = ["Date", "Invoice No", "Customer Name", "Phone", "Product Name", "Qty", "Sale Price", "Discount", "Purchase Price", "Payment", "Remarks", "Payment Method", "Profit"];
    var slRows = [slHeaders];
    (data.sales || []).forEach(function(s) {
      var prodName = s.productName || (s.itemsSummary || "").split("(")[0].trim() || (s.items && s.items[0] ? s.items[0].productName : "Item");
      var qty = s.qty || (s.items && s.items.reduce(function(a, it) { return a + it.qty; }, 0)) || 1;
      var saleP = s.salePrice || s.subtotal || s.netTotal || 0;
      var disc = s.discount || 0;
      var purchaseP = s.purchasePrice || 0;
      var pay = s.payment || s.netTotal || 0;
      var rem = s.remarks || "-";
      var method = s.paymentMethod || "Cash";
      var prof = s.profit !== undefined ? s.profit : Math.max(0, (s.netTotal || saleP) - purchaseP);

      slRows.push([
        s.date || "",
        s.invoiceNo || "",
        s.customerName || "",
        s.phone || "N/A",
        prodName,
        qty,
        saleP,
        disc,
        purchaseP,
        pay,
        rem,
        method,
        prof
      ]);
    });
    salesSheet.getRange(1, 1, slRows.length, slHeaders.length).setValues(slRows);
    salesSheet.getRange(1, 1, 1, slHeaders.length).setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold");
  }

  // 4. Customers
  if (data.customers !== undefined) {
    var custSheet = ss.getSheetByName("Customers") || ss.insertSheet("Customers");
    custSheet.clear();
    var cHeaders = ["Customer ID", "Customer Name", "Phone", "Total Purchases", "Total Paid", "Current Due", "Status"];
    var cRows = [cHeaders];
    (data.customers || []).forEach(function(c) {
      cRows.push([c.id || "", c.name || "", c.phone || "", c.totalPurchases || 0, c.totalPaid || 0, c.currentDue || 0, c.status || "Regular"]);
    });
    custSheet.getRange(1, 1, cRows.length, cHeaders.length).setValues(cRows);
    custSheet.getRange(1, 1, 1, cHeaders.length).setBackground("#dc2626").setFontColor("#ffffff").setFontWeight("bold");
  }

  // 5. Due Collections
  if (data.duePayments !== undefined) {
    var dueSheet = ss.getSheetByName("Due_Collections") || ss.insertSheet("Due_Collections");
    dueSheet.clear();
    var dHeaders = ["Collection Date", "Collection ID", "Customer Name", "Phone", "Amount Collected", "Payment Method", "Remaining Due", "Collected By", "Remarks"];
    var dRows = [dHeaders];
    (data.duePayments || []).forEach(function(d) {
      dRows.push([
        d.date || "",
        d.id || "",
        d.customerName || "",
        d.phone || "",
        d.amount || 0,
        d.paymentMethod || "Cash",
        d.remainingDue || 0,
        d.collectedBy || "",
        d.remarks || ""
      ]);
    });
    dueSheet.getRange(1, 1, dRows.length, dHeaders.length).setValues(dRows);
    dueSheet.getRange(1, 1, 1, dHeaders.length).setBackground("#d97706").setFontColor("#ffffff").setFontWeight("bold");
  }

  // 6. Expenses
  if (data.expenses !== undefined) {
    var expSheet = ss.getSheetByName("Expenses") || ss.insertSheet("Expenses");
    expSheet.clear();
    var eHeaders = ["Date", "Expense ID", "Category", "Amount", "Note", "Payment Method", "Logged By"];
    var eRows = [eHeaders];
    (data.expenses || []).forEach(function(e) {
      eRows.push([
        e.date || "",
        e.id || "",
        e.category || "",
        e.amount || 0,
        e.note || "",
        e.paymentMethod || "Cash",
        e.loggedBy || ""
      ]);
    });
    expSheet.getRange(1, 1, eRows.length, eHeaders.length).setValues(eRows);
    expSheet.getRange(1, 1, 1, eHeaders.length).setBackground("#0d9488").setFontColor("#ffffff").setFontWeight("bold");
  }

  // 7. Settings / Company Info
  if (data.companyInfo !== undefined && typeof data.companyInfo === "object") {
    saveCompanyInfoDirect(data.companyInfo);
  }

  return { success: true, message: "গুগল শিটের সব ট্যাব সফলভাবে আপডেট হয়েছে!" };
}

/**
 * Deletes a row directly from a sheet by ID/Invoice No and handles cross-sheet stock/customer recalculation
 */
function deleteSheetEntry(sheetName, idValue) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false, message: "শীট পাওয়া যায়নি: " + sheetName };

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, message: "শীটে কোনো ডাটা নেই!" };

  var targetId = String(idValue).trim().toLowerCase();
  var lastCol = sheet.getLastColumn();
  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  var idColIdx = 1; // 0-indexed column 1 (Column B)
  if (sheetName === "Customers") idColIdx = 0; // Column A: Customer ID

  var foundRowIdx = -1;
  var rowData = null;
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][idColIdx] || "").trim().toLowerCase() === targetId) {
      foundRowIdx = r + 1; // 1-indexed sheet row
      rowData = values[r];
      break;
    }
  }

  if (foundRowIdx === -1) {
    return { success: false, message: "আইডি পাওয়া যায়নি: " + idValue };
  }

  // If deleting from Sales: restore Stock and adjust Customer Due
  if (false && sheetName === "Sales" && rowData) { // adjustment is done by the app state, not twice
    var custName = String(rowData[2] || "").trim().toLowerCase();
    var prodName = String(rowData[4] || "").trim().toLowerCase();
    var qty = Number(rowData[5]) || 1;
    var netTotal = (Number(rowData[6]) || 0) - (Number(rowData[7]) || 0);
    var payment = Number(rowData[9]) || 0;
    var due = Math.max(0, netTotal - payment);

    // Restore Stock in Stock Sheet
    var stockSheet = ss.getSheetByName("Stock");
    if (stockSheet && stockSheet.getLastRow() > 1) {
      var sValues = stockSheet.getRange(2, 1, stockSheet.getLastRow() - 1, 10).getValues();
      for (var s = 0; s < sValues.length; s++) {
        var sName = String(sValues[s][3] || "").trim().toLowerCase();
        if (sName === prodName || sName.indexOf(prodName) !== -1 || prodName.indexOf(sName) !== -1) {
          var curStock = Number(sValues[s][7]) || 0;
          var curSold = Number(sValues[s][6]) || 0;
          stockSheet.getRange(s + 2, 8).setValue(curStock + qty); // Current Stock
          stockSheet.getRange(s + 2, 7).setValue(Math.max(0, curSold - qty)); // Sold Qty
          break;
        }
      }
    }

    // Adjust Customer
    if (custName) {
      var custSheet = ss.getSheetByName("Customers");
      if (custSheet && custSheet.getLastRow() > 1) {
        var cValues = custSheet.getRange(2, 1, custSheet.getLastRow() - 1, 7).getValues();
        for (var c = 0; c < cValues.length; c++) {
          var cName = String(cValues[c][1] || "").trim().toLowerCase();
          if (cName === custName) {
            var curPurchases = Number(cValues[c][3]) || 0;
            var curPaid = Number(cValues[c][4]) || 0;
            var curDue = Number(cValues[c][5]) || 0;
            var newDue = Math.max(0, curDue - due);
            custSheet.getRange(c + 2, 4).setValue(Math.max(0, curPurchases - netTotal));
            custSheet.getRange(c + 2, 5).setValue(Math.max(0, curPaid - payment));
            custSheet.getRange(c + 2, 6).setValue(newDue);
            custSheet.getRange(c + 2, 7).setValue(newDue > 0 ? "Due" : "Regular");
            break;
          }
        }
      }
    }
  }

  // Delete the row
  sheet.deleteRow(foundRowIdx);

  return { success: true, message: sheetName + " থেকে সফলভাবে ডিলিট করা হয়েছে!" };
}

/**
 * Updates an existing invoice directly in the Sales sheet and updates Customers sheet
 */
function updateInvoiceInSheet(jsonData) {
  try {
    var data = typeof jsonData === "string" ? JSON.parse(jsonData) : jsonData;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var salesSheet = ss.getSheetByName("Sales");
    if (!salesSheet) return { success: false, message: "Sales শিট পাওয়া যায়নি!" };

    var lastRow = salesSheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: "Sales শিটে কোনো ডাটা নেই!" };

    var invNo = String(data.invoiceNo || "").trim().toLowerCase();
    var sValues = salesSheet.getRange(1, 1, lastRow, 13).getValues();

    var targetRow = -1;
    var oldRow = null;
    for (var i = 1; i < sValues.length; i++) {
      if (String(sValues[i][1] || "").trim().toLowerCase() === invNo) {
        targetRow = i + 1; // 1-indexed sheet row
        oldRow = sValues[i];
        break;
      }
    }

    if (targetRow === -1) {
      return { success: false, message: "ইনভয়েস পাওয়া যায়নি: " + data.invoiceNo };
    }

    // Old values
    var oldCustName = String(oldRow[2] || "").trim();
    var oldSaleP = Number(oldRow[6]) || 0;
    var oldDisc = Number(oldRow[7]) || 0;
    var oldNet = oldSaleP - oldDisc;
    var oldPaid = Number(oldRow[9]) || 0;
    var oldDue = Math.max(0, oldNet - oldPaid);
    var purchaseP = Number(oldRow[8]) || 0;

    // New values
    var newCustName = String(data.customerName || oldCustName).trim();
    var newPhone = String(data.phone || oldRow[3] || "N/A").trim();
    var newSaleP = data.subtotal !== undefined ? Number(data.subtotal) : (data.salePrice !== undefined ? Number(data.salePrice) : oldSaleP);
    var newDisc = data.discount !== undefined ? Number(data.discount) : oldDisc;
    var newNet = data.netTotal !== undefined ? Number(data.netTotal) : (newSaleP - newDisc);
    var newPaid = data.payment !== undefined ? Number(data.payment) : oldPaid;
    var newRemarks = data.remarks !== undefined ? String(data.remarks) : String(oldRow[10] || "-");
    var newMethod = data.paymentMethod !== undefined ? String(data.paymentMethod) : String(oldRow[11] || "Cash");
    var newProfit = data.profit !== undefined ? Number(data.profit) : Math.max(0, newNet - purchaseP);

    // Update row in Sales sheet
    // Headers: Date (1), Invoice No (2), Customer Name (3), Phone (4), Product Name (5), Qty (6), Sale Price (7), Discount (8), Purchase Price (9), Payment (10), Remarks (11), Payment Method (12), Profit (13)
    salesSheet.getRange(targetRow, 3).setValue(newCustName);
    salesSheet.getRange(targetRow, 4).setValue(newPhone);
    salesSheet.getRange(targetRow, 7).setValue(newSaleP);
    salesSheet.getRange(targetRow, 8).setValue(newDisc);
    salesSheet.getRange(targetRow, 10).setValue(newPaid);
    salesSheet.getRange(targetRow, 11).setValue(newRemarks);
    salesSheet.getRange(targetRow, 12).setValue(newMethod);
    salesSheet.getRange(targetRow, 13).setValue(newProfit);

    // If date is provided and valid
    if (data.date) {
      salesSheet.getRange(targetRow, 1).setValue(data.date);
    }

    // Update Customers Sheet balance
    var custSheet = ss.getSheetByName("Customers");
    if (custSheet && custSheet.getLastRow() > 1) {
      var cValues = custSheet.getRange(2, 1, custSheet.getLastRow() - 1, 7).getValues();
      var netDiff = newNet - oldNet;
      var paidDiff = newPaid - oldPaid;
      var dueDiff = (newNet - newPaid) - oldDue;

      for (var c = 0; c < cValues.length; c++) {
        var cName = String(cValues[c][1] || "").trim().toLowerCase();
        if (cName === oldCustName.toLowerCase() || cName === newCustName.toLowerCase()) {
          var curPurchases = Number(cValues[c][3]) || 0;
          var curPaid = Number(cValues[c][4]) || 0;
          var curDue = Number(cValues[c][5]) || 0;

          var updatedPurchases = Math.max(0, curPurchases + netDiff);
          var updatedPaid = Math.max(0, curPaid + paidDiff);
          var updatedDue = Math.max(0, curDue + dueDiff);

          custSheet.getRange(c + 2, 2).setValue(newCustName);
          custSheet.getRange(c + 2, 3).setValue(newPhone);
          custSheet.getRange(c + 2, 4).setValue(updatedPurchases);
          custSheet.getRange(c + 2, 5).setValue(updatedPaid);
          custSheet.getRange(c + 2, 6).setValue(updatedDue);
          custSheet.getRange(c + 2, 7).setValue(updatedDue > 0 ? "Due" : "Regular");
          break;
        }
      }
    }

    return { success: true, message: "ইনভয়েস " + data.invoiceNo + " সফলভাবে শিটে আপডেট হয়েছে!" };
  } catch (err) {
    return { success: false, message: "আপডেট করতে ত্রুটি: " + err.message };
  }
}

/**
 * Formats the Stock sheet in Google Sheets to match the POS UI table screenshot EXACTLY:
 * Headers: Date | Id | Purchase No | Product Name | Category | Purchased Qty | Sold Qty | Current Stock | Sale Price | Purchase Price
 *
 * Visual Crafting:
 * - Header Row: Light subtle slate (#e2e8f0), dark text (#1e293b), bold, row height 32, frozen 1st row
 * - Col 1 (Date): text/date format, gray-600 color (#475569)
 * - Col 2 (Id): Bold font, Emerald Green color (#047857), text format (@)
 * - Col 3 (Purchase No): gray-600 color (#475569), text format (@)
 * - Col 4 (Product Name): Bold text, dark color (#0f172a)
 * - Col 5 (Category): gray-600 color (#475569)
 * - Col 6 (Purchased Qty): Center aligned, number format #,##0
 * - Col 7 (Sold Qty): Center aligned, Blue bold color (#1d4ed8), number format #,##0
 * - Col 8 (Current Stock): Center aligned, Bold font, dark color (#0f172a), number format #,##0
 * - Col 9 (Sale Price): Right aligned, Bold font, Green color (#065f46), number format #,##0
 * - Col 10 (Purchase Price): Right aligned, text color (#334155), number format #,##0
 * - Subtle table borders (#cbd5e1)
 * - Proportional column widths
 */
function formatStockSheet(stockSheet) {
  if (!stockSheet) return;
  var lastRow = stockSheet.getLastRow();
  var numCols = 10;
  if (lastRow === 0) return;

  // 1. Header row formatting
  var headerRange = stockSheet.getRange(1, 1, 1, numCols);
  headerRange
    .setBackground("#e2e8f0")
    .setFontColor("#1e293b")
    .setFontWeight("bold")
    .setFontSize(10)
    .setVerticalAlignment("middle");
  stockSheet.setRowHeight(1, 32);
  stockSheet.setFrozenRows(1);

  // Column alignments for header
  stockSheet.getRange(1, 1, 1, 5).setHorizontalAlignment("left");
  stockSheet.getRange(1, 6, 1, 3).setHorizontalAlignment("center");
  stockSheet.getRange(1, 9, 1, 2).setHorizontalAlignment("right");

  // 2. Data rows formatting
  if (lastRow > 1) {
    var dataRowsCount = lastRow - 1;

    // Reset base font
    var dataRange = stockSheet.getRange(2, 1, dataRowsCount, numCols);
    dataRange
      .setFontFamily("Roboto")
      .setFontSize(10)
      .setVerticalAlignment("middle");

    // Row heights for data
    for (var r = 2; r <= lastRow; r++) {
      stockSheet.setRowHeight(r, 26);
    }

    // Col 1: Date (Gray-600)
    stockSheet.getRange(2, 1, dataRowsCount, 1)
      .setHorizontalAlignment("left")
      .setFontColor("#475569");

    // Col 2: Id (Emerald Green bold - matches UI P0008)
    stockSheet.getRange(2, 2, dataRowsCount, 1)
      .setHorizontalAlignment("left")
      .setFontWeight("bold")
      .setFontColor("#047857")
      .setNumberFormat("@");

    // Col 3: Purchase No (Gray-600 text - PUR13181)
    stockSheet.getRange(2, 3, dataRowsCount, 1)
      .setHorizontalAlignment("left")
      .setFontColor("#475569")
      .setNumberFormat("@");

    // Col 4: Product Name (Bold dark text - pant)
    stockSheet.getRange(2, 4, dataRowsCount, 1)
      .setHorizontalAlignment("left")
      .setFontWeight("bold")
      .setFontColor("#0f172a");

    // Col 5: Category (Gray-600 - clothing)
    stockSheet.getRange(2, 5, dataRowsCount, 1)
      .setHorizontalAlignment("left")
      .setFontColor("#475569");

    // Col 6: Purchased Qty (Center number)
    stockSheet.getRange(2, 6, dataRowsCount, 1)
      .setHorizontalAlignment("center")
      .setFontColor("#1e293b")
      .setNumberFormat("#,##0");

    // Col 7: Sold Qty (Blue color - 0)
    stockSheet.getRange(2, 7, dataRowsCount, 1)
      .setHorizontalAlignment("center")
      .setFontColor("#1d4ed8")
      .setFontWeight("bold")
      .setNumberFormat("#,##0");

    // Col 8: Current Stock (Center bold - 10)
    stockSheet.getRange(2, 8, dataRowsCount, 1)
      .setHorizontalAlignment("center")
      .setFontWeight("bold")
      .setFontColor("#0f172a")
      .setNumberFormat("#,##0");

    // Col 9: Sale Price (Right bold emerald - 600)
    stockSheet.getRange(2, 9, dataRowsCount, 1)
      .setHorizontalAlignment("right")
      .setFontWeight("bold")
      .setFontColor("#065f46")
      .setNumberFormat("#,##0");

    // Col 10: Purchase Price (Right dark slate - 400)
    stockSheet.getRange(2, 10, dataRowsCount, 1)
      .setHorizontalAlignment("right")
      .setFontColor("#334155")
      .setNumberFormat("#,##0");

    // Subtle cell borders
    stockSheet.getRange(1, 1, lastRow, numCols)
      .setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
  }

  // Set proportional column widths matching the UI screenshot
  var colWidths = [155, 85, 110, 160, 120, 110, 90, 110, 100, 110];
  for (var i = 0; i < colWidths.length; i++) {
    stockSheet.setColumnWidth(i + 1, colWidths[i]);
  }
}

/**
 * Menu action / helper to reformat the Stock sheet anytime directly in Google Sheets
 */
function formatAndSyncStockSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var stockSheet = ss.getSheetByName("Stock");
  if (!stockSheet) {
    setupShopSheets();
    stockSheet = ss.getSheetByName("Stock");
  }

  // Ensure 10 exact headers
  var expectedHeaders = ["Date", "Id", "Purchase No", "Product Name", "Category", "Purchased Qty", "Sold Qty", "Current Stock", "Sale Price", "Purchase Price"];
  if (stockSheet.getLastRow() === 0) {
    stockSheet.appendRow(expectedHeaders);
  } else {
    stockSheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
  }

  formatStockSheet(stockSheet);
  SpreadsheetApp.getActive().toast("স্টক শিট সফলভাবে সফটওয়্যারের মতো সেইম টু সেইম ফরম্যাট করা হয়েছে!", "POS System", 5);
}

/**
 * Formats the Products sheet in Google Sheets to match the POS UI table screenshot EXACTLY:
 * Headers: Date | ID | Purchase No | Supplier | Product | Qty | Purchase Price | Sale Price | Total
 *
 * Visual Crafting:
 * - Header Row: Light subtle slate (#e2e8f0), dark text (#1e293b), bold, row height 32, frozen 1st row
 * - Col 1 (Date): text/date format, gray-600 color (#475569)
 * - Col 2 (ID): Bold font, Emerald Green color (#047857), text format (@)
 * - Col 3 (Purchase No): gray-600 color (#475569), text format (@)
 * - Col 4 (Supplier): text color (#334155), text format (@)
 * - Col 5 (Product): text color (#0f172a), font size 10
 * - Col 6 (Qty): Center aligned, number format #,##0, text color (#1e293b)
 * - Col 7 (Purchase Price): Right aligned, number format #,##0, text color (#334155)
 * - Col 8 (Sale Price): Right aligned, number format #,##0, text color (#334155)
 * - Col 9 (Total): Right aligned, number format #,##0, bold font, text color (#0f172a)
 * - Subtle cell borders (#cbd5e1)
 * - Proportional column widths
 */
function formatProductSheet(prodSheet) {
  if (!prodSheet) return;
  var lastRow = prodSheet.getLastRow();
  var numCols = 9;
  if (lastRow === 0) return;

  // 1. Header row formatting
  var headerRange = prodSheet.getRange(1, 1, 1, numCols);
  headerRange
    .setBackground("#e2e8f0")
    .setFontColor("#1e293b")
    .setFontWeight("bold")
    .setFontSize(10)
    .setVerticalAlignment("middle");
  prodSheet.setRowHeight(1, 32);
  prodSheet.setFrozenRows(1);

  // Column alignments for header
  prodSheet.getRange(1, 1, 1, 5).setHorizontalAlignment("left");
  prodSheet.getRange(1, 6, 1, 1).setHorizontalAlignment("center");
  prodSheet.getRange(1, 7, 1, 3).setHorizontalAlignment("right");

  // 2. Data rows formatting
  if (lastRow > 1) {
    var dataRowsCount = lastRow - 1;

    // Reset base font
    var dataRange = prodSheet.getRange(2, 1, dataRowsCount, numCols);
    dataRange
      .setFontFamily("Roboto")
      .setFontSize(10)
      .setVerticalAlignment("middle");

    // Row heights for data
    for (var r = 2; r <= lastRow; r++) {
      prodSheet.setRowHeight(r, 26);
    }

    // Col 1: Date (Gray-600)
    prodSheet.getRange(2, 1, dataRowsCount, 1)
      .setHorizontalAlignment("left")
      .setFontColor("#475569");

    // Col 2: ID (Emerald Green bold - matches UI P0001, P0002)
    prodSheet.getRange(2, 2, dataRowsCount, 1)
      .setHorizontalAlignment("left")
      .setFontWeight("bold")
      .setFontColor("#047857")
      .setNumberFormat("@");

    // Col 3: Purchase No (Gray-600 text - PUR00006)
    prodSheet.getRange(2, 3, dataRowsCount, 1)
      .setHorizontalAlignment("left")
      .setFontColor("#475569")
      .setNumberFormat("@");

    // Col 4: Supplier (Slate-700 - abcd)
    prodSheet.getRange(2, 4, dataRowsCount, 1)
      .setHorizontalAlignment("left")
      .setFontColor("#334155");

    // Col 5: Product (Dark text - SHIRT, pant, jacket, t-shirt)
    prodSheet.getRange(2, 5, dataRowsCount, 1)
      .setHorizontalAlignment("left")
      .setFontWeight("normal")
      .setFontColor("#0f172a");

    // Col 6: Qty (Center number)
    prodSheet.getRange(2, 6, dataRowsCount, 1)
      .setHorizontalAlignment("center")
      .setFontColor("#1e293b")
      .setFontWeight("bold")
      .setNumberFormat("#,##0");

    // Col 7: Purchase Price (Right aligned)
    prodSheet.getRange(2, 7, dataRowsCount, 1)
      .setHorizontalAlignment("right")
      .setFontColor("#334155")
      .setNumberFormat("#,##0");

    // Col 8: Sale Price (Right aligned)
    prodSheet.getRange(2, 8, dataRowsCount, 1)
      .setHorizontalAlignment("right")
      .setFontColor("#334155")
      .setNumberFormat("#,##0");

    // Col 9: Total (Right bold dark text - 20000, 6000)
    prodSheet.getRange(2, 9, dataRowsCount, 1)
      .setHorizontalAlignment("right")
      .setFontWeight("bold")
      .setFontColor("#0f172a")
      .setNumberFormat("#,##0");

    // Subtle cell borders
    prodSheet.getRange(1, 1, lastRow, numCols)
      .setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
  }

  // Set proportional column widths matching the screenshot
  var colWidths = [155, 80, 110, 100, 140, 75, 110, 100, 100];
  for (var i = 0; i < colWidths.length; i++) {
    prodSheet.setColumnWidth(i + 1, colWidths[i]);
  }
}

/**
 * Menu action / helper to reformat the Products sheet anytime directly in Google Sheets
 */
function formatAndSyncProductSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var prodSheet = ss.getSheetByName("Products");
  if (!prodSheet) {
    setupShopSheets();
    prodSheet = ss.getSheetByName("Products");
  }

  // Ensure 9 exact headers
  var expectedHeaders = ["Date", "ID", "Purchase No", "Supplier", "Product", "Qty", "Purchase Price", "Sale Price", "Total"];
  if (prodSheet.getLastRow() === 0) {
    prodSheet.appendRow(expectedHeaders);
  } else {
    prodSheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
  }

  formatProductSheet(prodSheet);
  SpreadsheetApp.getActive().toast("প্রোডাক্ট শিট সফলভাবে সফটওয়্যারের মতো সেইম টু সেইম ফরম্যাট করা হয়েছে!", "POS System", 5);
}

/**
 * Adds a new product entry directly into Google Sheets following the strict rule:
 * 1. Stock Sheet: If the SAME Product Name and SAME Price (Sale Price and Purchase Price) already exists,
 *    adds the new entry's Quantity to the existing Stock row's Purchased Qty and Current Stock.
 *    No new row is created in Stock Sheet.
 * 2. Product Sheet: EVERY new entry is ALWAYS added as a new separate row, even if Product Name and Price are identical.
 */
function addNewProductEntry(productData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  setupShopSheets();

  var data = typeof productData === "string" ? JSON.parse(productData) : (productData || {});
  var prodName = String(data.productName || data.name || data.product || "").trim();
  var qty = Number(data.qty !== undefined ? data.qty : (data.currentStock || 1));
  var buyPrice = Number(data.purchasePrice || 0);
  var salePrice = Number(data.salePrice || 0);
  var category = String(data.category || "clothing").trim();
  var supplier = String(data.supplier || "abcd").trim();
  var dateStr = data.date || new Date().toLocaleString();

  if (!prodName) {
    return { success: false, message: "প্রোডাক্টের নাম পাওয়া যায়নি।" };
  }

  var prodSheet = ss.getSheetByName("Products") || ss.insertSheet("Products");
  var stockSheet = ss.getSheetByName("Stock") || ss.insertSheet("Stock");

  // Determine next Product ID and Purchase No if not supplied
  var pId = data.id || data.productId;
  var pNo = data.purchaseNo;

  if (!pId) {
    var maxId = 0;
    if (prodSheet.getLastRow() > 1) {
      var ids = prodSheet.getRange(2, 2, prodSheet.getLastRow() - 1, 1).getValues();
      ids.forEach(function(r) {
        var m = String(r[0]).match(/^P(\d+)$/i);
        if (m) {
          var n = parseInt(m[1], 10);
          if (n > maxId) maxId = n;
        }
      });
    }
    pId = "P" + String(maxId + 1).padStart(4, '0');
  }

  if (!pNo) {
    var maxPur = 0;
    if (prodSheet.getLastRow() > 1) {
      var pnos = prodSheet.getRange(2, 3, prodSheet.getLastRow() - 1, 1).getValues();
      pnos.forEach(function(r) {
        var m = String(r[0]).match(/^PUR(\d+)$/i);
        if (m) {
          var n = parseInt(m[1], 10);
          if (n > maxPur) maxPur = n;
        }
      });
    }
    pNo = "PUR" + String(maxPur + 1).padStart(5, '0');
  }

  var total = qty * buyPrice;

  // 1. PRODUCT SHEET: সর্বদা নতুন রো হিসেবে যোগ হবে
  // কলাম: Date | ID | Purchase No | Supplier | Product | Qty | Purchase Price | Sale Price | Total
  prodSheet.appendRow([dateStr, pId, pNo, supplier, prodName, qty, buyPrice, salePrice, total]);
  formatProductSheet(prodSheet);

  // 2. STOCK SHEET: Same Product Name AND Same Price (Sale & Purchase) হলে আগের রো-এর সঙ্গে কোয়ান্টিটি যোগ হবে
  // কলাম: Date | Id | Purchase No | Product Name | Category | Purchased Qty | Sold Qty | Current Stock | Sale Price | Purchase Price
  var stockMerged = false;
  var matchedStockRow = -1;
  var matchedCurrentStock = 0;

  if (stockSheet.getLastRow() > 1) {
    var stockValues = stockSheet.getRange(2, 1, stockSheet.getLastRow() - 1, 10).getValues();
    for (var i = 0; i < stockValues.length; i++) {
      var row = stockValues[i];
      var existingName = String(row[3] || "").trim().toLowerCase();
      var existingSalePrice = Number(row[8] || 0);
      var existingBuyPrice = Number(row[9] || 0);

      // Same Product + Same Price
      if (existingName === prodName.toLowerCase() && existingSalePrice === salePrice && existingBuyPrice === buyPrice) {
        matchedStockRow = i + 2; // Row index in sheet (1-based, skipping header)
        var prevPurchased = Number(row[5] || 0);
        var prevStock = Number(row[7] || 0);
        var newPurchased = prevPurchased + qty;
        var newStock = prevStock + qty;
        matchedCurrentStock = newStock;

        // আপডেট করুন (নতুন রো তৈরি হবে না)
        stockSheet.getRange(matchedStockRow, 1).setValue(dateStr); // Date
        stockSheet.getRange(matchedStockRow, 6).setValue(newPurchased); // Purchased Qty
        stockSheet.getRange(matchedStockRow, 8).setValue(newStock); // Current Stock
        stockMerged = true;
        break;
      }
    }
  }

  // যদি একই প্রোডাক্ট ও একই দামের রেকর্ড স্টকে না থাকে, তবে নতুন রো হিসেবে যুক্ত হবে
  if (!stockMerged) {
    stockSheet.appendRow([dateStr, pId, pNo, prodName, category, qty, 0, qty, salePrice, buyPrice]);
    matchedCurrentStock = qty;
  }
  formatStockSheet(stockSheet);

  var msg = stockMerged
    ? "✅ '" + prodName + "' স্টকে পূর্বে একই মূল্যে থাকায় স্টক শিটে আগের স্টকের সাথে +" + qty + " যোগ হয়েছে (মোট স্টক: " + matchedCurrentStock + ")। এবং প্রোডাক্ট শিটে আলাদা নতুন রো যোগ হয়েছে।"
    : "✅ '" + prodName + "' প্রোডাক্ট শিট এবং স্টক শিট উভয় জায়গায় সফলভাবে নতুন রেকর্ড হিসেবে তৈরি হয়েছে।";

  SpreadsheetApp.getActive().toast(msg, "POS System", 6);

  return {
    success: true,
    message: msg,
    stockMerged: stockMerged,
    productId: pId,
    purchaseNo: pNo,
    currentStock: matchedCurrentStock
  };
}

/**
 * Google Sheets UI Prompt to quickly enter a new product from the spreadsheet menu
 */
function promptNewProductEntry() {
  var ui = SpreadsheetApp.getUi();
  var namePrompt = ui.prompt("New Product Entry", "প্রোডাক্টের নাম লিখুন (যেমন: SHIRT, pant, jacket):", ui.ButtonSet.OK_CANCEL);
  if (namePrompt.getSelectedButton() != ui.Button.OK) return;
  var prodName = namePrompt.getResponseText().trim();
  if (!prodName) {
    ui.alert("ত্রুটি", "প্রোডাক্টের নাম আবশ্যক!", ui.ButtonSet.OK);
    return;
  }

  var qtyPrompt = ui.prompt("New Product Entry", "'" + prodName + "' এর পরিমাণ (Qty) লিখুন (যেমন: 10):", ui.ButtonSet.OK_CANCEL);
  if (qtyPrompt.getSelectedButton() != ui.Button.OK) return;
  var qty = parseInt(qtyPrompt.getResponseText().trim(), 10) || 1;

  var buyPrompt = ui.prompt("New Product Entry", "ক্রয় মূল্য (Purchase Price) লিখুন (যেমন: 300):", ui.ButtonSet.OK_CANCEL);
  if (buyPrompt.getSelectedButton() != ui.Button.OK) return;
  var buyPrice = parseFloat(buyPrompt.getResponseText().trim()) || 0;

  var salePrompt = ui.prompt("New Product Entry", "বিক্রয় মূল্য (Sale Price) লিখুন (যেমন: 500):", ui.ButtonSet.OK_CANCEL);
  if (salePrompt.getSelectedButton() != ui.Button.OK) return;
  var salePrice = parseFloat(salePrompt.getResponseText().trim()) || 0;

  var res = addNewProductEntry({
    productName: prodName,
    category: "clothing",
    supplier: "abcd",
    qty: qty,
    purchasePrice: buyPrice,
    salePrice: salePrice
  });

  ui.alert("ফলাফল", res.message, ui.ButtonSet.OK);
}

/**
 * Searches product details by Product ID from Stock and Products sheets.
 * Used for instant Auto-Fill in POS Sale Popup.
 * Returns product name, current stock, sale price, purchase price, category, etc.
 */
function getProductDetailsById(productId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetId = String(productId || "").trim().toLowerCase();

  if (!targetId) {
    return { success: false, message: "প্রোডাক্ট আইডি খালি রাখা যাবে না।" };
  }

  var targetNum = targetId.replace(/\D/g, "");
  var targetNumInt = targetNum ? parseInt(targetNum, 10) : null;

  // 1. First priority: Check Stock Sheet for live current stock
  var stockSheet = ss.getSheetByName("Stock");
  if (stockSheet && stockSheet.getLastRow() > 1) {
    var stockValues = stockSheet.getRange(2, 1, stockSheet.getLastRow() - 1, 10).getValues();
    for (var i = 0; i < stockValues.length; i++) {
      var row = stockValues[i];
      var sId = String(row[1] || "").trim().toLowerCase();
      var sNo = String(row[2] || "").trim().toLowerCase();
      var sNum = sId.replace(/\D/g, "");
      var numMatch = targetNumInt !== null && sNum && parseInt(sNum, 10) === targetNumInt;

      if (sId === targetId || sNo === targetId || numMatch) {
        return {
          success: true,
          product: {
            id: String(row[1] || productId),
            productId: String(row[1] || productId),
            purchaseNo: String(row[2] || ""),
            name: String(row[3] || ""),
            productName: String(row[3] || ""),
            category: String(row[4] || ""),
            purchasedQty: Number(row[5]) || 0,
            soldQty: Number(row[6]) || 0,
            currentStock: Number(row[7]) || 0,
            qty: Number(row[7]) || 0,
            salePrice: Number(row[8]) || 0,
            purchasePrice: Number(row[9]) || 0
          }
        };
      }
    }
  }

  // 2. Secondary check: Products Sheet
  var prodSheet = ss.getSheetByName("Products");
  if (prodSheet && prodSheet.getLastRow() > 1) {
    var prodValues = prodSheet.getRange(2, 1, prodSheet.getLastRow() - 1, 9).getValues();
    for (var j = 0; j < prodValues.length; j++) {
      var pRow = prodValues[j];
      var pId = String(pRow[1] || "").trim().toLowerCase();
      var pPur = String(pRow[2] || "").trim().toLowerCase();
      var pNum = pId.replace(/\D/g, "");
      var pNumMatch = targetNumInt !== null && pNum && parseInt(pNum, 10) === targetNumInt;

      if (pId === targetId || pPur === targetId || pNumMatch) {
        return {
          success: true,
          product: {
            id: String(pRow[1] || productId),
            productId: String(pRow[1] || productId),
            purchaseNo: String(pRow[2] || ""),
            supplier: String(pRow[3] || ""),
            name: String(pRow[4] || ""),
            productName: String(pRow[4] || ""),
            currentStock: Number(pRow[5]) || 0,
            qty: Number(pRow[5]) || 0,
            purchasePrice: Number(pRow[6]) || 0,
            salePrice: Number(pRow[7]) || 0,
            total: Number(pRow[8]) || 0
          }
        };
      }
    }
  }

  return { success: false, message: "প্রোডাক্ট আইডি '" + productId + "' পাওয়া যায়নি।" };
}

/**
 * Searches products by query (ID, English Name, or Bengali Alias)
 */
function searchProductInSheet(searchQuery) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var q = String(searchQuery || "").trim().toLowerCase();
  if (!q) return { success: true, results: [] };

  var qNum = q.replace(/\D/g, "");
  var qNumInt = qNum ? parseInt(qNum, 10) : null;

  var bengaliMap = {
    "shirt": ["শার্ট", "শাট", "shirt"],
    "pant": ["প্যান্ট", "পেন্ট", "pant"],
    "panjabi": ["পাঞ্জাবি", "পাঞ্জাবী", "panjabi"],
    "t-shirt": ["টি শার্ট", "টি-শার্ট", "t-shirt", "tshirt", "টিশার্ট"],
    "polo": ["পোলো", "polo"],
    "jeans": ["জিন্স", "jeans"],
    "sharee": ["শাড়ি", "শাড়ী", "sharee", "saree"],
    "kurti": ["কুর্তি", "kurti"],
    "three_piece": ["থ্রি পিস", "থ্রিপিস", "three piece"],
    "shoes": ["জুতো", "জুতা", "shoes", "shoe"]
  };

  var results = [];
  var stockSheet = ss.getSheetByName("Stock");
  if (stockSheet && stockSheet.getLastRow() > 1) {
    var data = stockSheet.getRange(2, 1, stockSheet.getLastRow() - 1, 10).getValues();
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var id = String(row[1] || "").trim();
      var name = String(row[3] || "").trim();
      var idLower = id.toLowerCase();
      var nameLower = name.toLowerCase();
      var idNum = idLower.replace(/\D/g, "");
      var idNumMatch = qNumInt !== null && idNum && parseInt(idNum, 10) === qNumInt;

      var isMatch = idLower.indexOf(q) > -1 || nameLower.indexOf(q) > -1 || idNumMatch;

      if (!isMatch) {
        for (var key in bengaliMap) {
          if (nameLower.indexOf(key) > -1) {
            var aliases = bengaliMap[key];
            for (var a = 0; a < aliases.length; a++) {
              if (aliases[a].indexOf(q) > -1 || q.indexOf(aliases[a]) > -1) {
                isMatch = true;
                break;
              }
            }
          }
          if (isMatch) break;
        }
      }

      if (isMatch) {
        results.push({
          id: id,
          productId: id,
          purchaseNo: String(row[2] || ""),
          name: name,
          productName: name,
          category: String(row[4] || ""),
          currentStock: Number(row[7]) || 0,
          salePrice: Number(row[8]) || 0,
          purchasePrice: Number(row[9]) || 0
        });
      }
    }
  }

  return { success: true, results: results };
}

/**
 * Searches due customers by Name, Phone, or Invoice Number
 */
function searchDueCustomersInSheet(searchQuery, searchType) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var q = String(searchQuery || "").trim().toLowerCase();
  var type = String(searchType || "all").toLowerCase();

  var result = {
    success: true,
    customers: [],
    matchingInvoices: []
  };

  var rawData = getShopData();
  var shopData = typeof rawData === "string" ? JSON.parse(rawData) : (rawData || {});
  var customers = shopData.customers || [];
  var sales = shopData.sales || [];

  if (!q) {
    result.customers = customers.filter(function(c) { return c.currentDue > 0; });
    return result;
  }

  // Find sales that match invoice search
  var matchedSales = sales.filter(function(s) {
    return (s.invoiceNo || "").toLowerCase().indexOf(q) > -1;
  });
  result.matchingInvoices = matchedSales;

  result.customers = customers.filter(function(c) {
    var nameMatch = (c.name || "").toLowerCase().indexOf(q) > -1;
    var phoneMatch = (c.phone || "").toLowerCase().indexOf(q) > -1;
    var invMatch = matchedSales.some(function(s) {
      return (s.customerName || "").toLowerCase() === (c.name || "").toLowerCase() ||
             (c.phone && s.phone && s.phone === c.phone);
    });

    if (type === "name") return nameMatch;
    if (type === "phone") return phoneMatch;
    if (type === "invoice") return invMatch;
    return nameMatch || phoneMatch || invMatch;
  });

  return result;
}



// ============================================================================
// VERCEL API LAYER (added) — safe real-time sync, role control, no auto-delete
// ============================================================================
// Secret is built in below (no Script properties needed).
// Deploy: Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone
// Put the deployment URL in Vercel env GAS_URL and the same secret in GAS_SECRET.

var API_ALL = ['admin', 'manager', 'cashier', 'stock'];
var API_FN = {
  authenticateUser:     { auth: false },
  getShopData:          { roles: API_ALL },
  getVersion:           { roles: API_ALL },
  getCompanyInfo:       { roles: API_ALL },
  saveAllShopData:      { roles: API_ALL, write: true },
  deleteSheetEntry:     { roles: ['admin', 'manager'], write: true },
  updateInvoiceInSheet: { roles: ['admin', 'manager'], write: true },
  getUsersList:         { roles: ['admin'] },
  saveOrUpdateUser:     { roles: ['admin'], write: true },
  deleteUser:           { roles: ['admin'], write: true },
  saveCompanyInfo:      { roles: ['admin'], write: true },
  uploadFile:           { roles: ['admin', 'manager', 'stock'], write: true }
};

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getVersion() {
  return Number(PropertiesService.getScriptProperties().getProperty('DATA_VERSION') || 0);
}

function bumpVersion_() {
  var p = PropertiesService.getScriptProperties();
  p.setProperty('DATA_VERSION', String(Number(p.getProperty('DATA_VERSION') || 0) + 1));
}

function lookupUser_(auth) {
  if (!auth || !auth.u || !auth.p) return null;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  if (!sh || sh.getLastRow() < 2) return null;
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
  var u = String(auth.u).trim().toLowerCase();
  var p = String(auth.p).trim();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() === u && String(rows[i][1]).trim() === p) {
      return { username: String(rows[i][0]), role: String(rows[i][3] || "cashier").trim().toLowerCase() };
    }
  }
  return null;
}

function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents);
    var secret = 'E-o8ieNwKFusT9USq7oweKi2A6Gyb9bAf_hoQObj'; // must be identical to FALLBACK_SECRET in api/gas.js
    if (!secret || req.secret !== secret) return json_({ ok: false, error: 'Unauthorized' });
    var spec = API_FN[req.fn];
    if (!spec) return json_({ ok: false, error: 'Unknown function' });

    var role = null;
    if (spec.auth !== false) {
      var user = lookupUser_(req.auth);
      if (!user) return json_({ ok: false, error: 'লগইন সেশন অবৈধ, আবার লগইন করুন' });
      role = user.role;
      if (spec.roles.indexOf(role) === -1) return json_({ ok: false, error: 'এই কাজের অনুমতি আপনার নেই' });
    }

    var lock = LockService.getScriptLock();
    if (spec.write) lock.waitLock(30000);
    var result;
    try {
      result = dispatch_(req.fn, req.args || [], role);
      if (spec.write) bumpVersion_();
    } finally {
      if (spec.write) lock.releaseLock();
    }
    return json_({ ok: true, result: result, v: getVersion() });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function dispatch_(fn, a, role) {
  switch (fn) {
    case 'authenticateUser':     return authenticateUser(a[0], a[1]);
    case 'getShopData':          return filterShopData_(getShopData(), role);
    case 'getVersion':           return getVersion();
    case 'getCompanyInfo':       return getCompanyInfo();
    case 'saveAllShopData':      return saveAllShopDataSafe_(a[0], role);
    case 'deleteSheetEntry':     return deleteSheetEntry(a[0], a[1]);
    case 'updateInvoiceInSheet': return updateInvoiceInSheet(a[0]);
    case 'getUsersList':         return getUsersList(a[0], a[1]);
    case 'saveOrUpdateUser':     return saveOrUpdateUser(a[0], a[1], a[2]);
    case 'deleteUser':           return deleteUser(a[0], a[1], a[2]);
    case 'saveCompanyInfo':      return saveCompanyInfo(a[0], a[1], a[2]);
    case 'uploadFile':           return uploadFile_(a[0], a[1], a[2]);
  }
  throw new Error('Unknown function');
}

/** Role based view of data: salesman has no expenses; stock user sees only products/stock */
function filterShopData_(data, role) {
  var d = typeof data === "string" ? JSON.parse(data) : data;
  if (role === 'cashier') {
    d.expenses = [];
  } else if (role === 'stock') {
    d.sales = []; d.customers = []; d.duePayments = []; d.expenses = [];
  }
  return d;
}

/**
 * SAFE SAVE: never clears a sheet and never deletes a row.
 * Rows are matched by ID: existing rows updated, new rows appended.
 * Rows missing from the incoming data are left untouched.
 * Only deleteSheetEntry (user pressed delete) removes a row.
 */
function mergeRows_(sheet, headers, newRows, keyFn, protectCols) {
  var width = headers.length;
  if (sheet.getLastRow() === 0) sheet.appendRow(headers);
  var last = sheet.getLastRow();
  var existing = last > 1 ? sheet.getRange(2, 1, last - 1, width).getValues() : [];
  var index = {};
  existing.forEach(function (r, i) { index[keyFn(r)] = i; });
  var toAppend = [];
  var changed = false;
  newRows.forEach(function (nr) {
    var k = keyFn(nr);
    if (!k || k === '|') return;
    if (index.hasOwnProperty(k)) {
      var i = index[k];
      var old = existing[i];
      var merged = nr.slice();
      protectCols.forEach(function (c) { merged[c] = old[c]; });
      for (var c = 0; c < width; c++) {
        if (String(old[c]) !== String(merged[c])) { existing[i] = merged; changed = true; break; }
      }
    } else {
      index[k] = existing.length + toAppend.length;
      toAppend.push(nr);
    }
  });
  if (changed && existing.length) sheet.getRange(2, 1, existing.length, width).setValues(existing);
  if (toAppend.length) sheet.getRange(sheet.getLastRow() + 1, 1, toAppend.length, width).setValues(toAppend);
}

function saveAllShopDataSafe_(jsonData, role) {
  var data = typeof jsonData === "string" ? JSON.parse(jsonData) : jsonData;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var canAll = (role === 'admin' || role === 'manager');
  var allow = {
    products:    canAll || role === 'stock',
    stock:       canAll || role === 'stock' || role === 'cashier',
    sales:       canAll || role === 'cashier',
    customers:   canAll || role === 'cashier',
    duePayments: canAll,
    expenses:    canAll
  };
  function sh(name) { return ss.getSheetByName(name) || ss.insertSheet(name); }
  var now = new Date().toLocaleString();

  if (allow.products && data.products) {
    var pH = ["Date", "ID", "Purchase No", "Supplier", "Product", "Qty", "Purchase Price", "Sale Price", "Total"];
    var pRows = data.products.map(function (p) {
      var qty = Number(p.qty !== undefined ? p.qty : (p.currentStock !== undefined ? p.currentStock : 0));
      var buy = Number(p.purchasePrice || 0);
      return [p.date || now, p.id || p.productId || "", p.purchaseNo || "N/A", p.supplier || p.category || "abcd",
              p.product || p.name || "", qty, buy, Number(p.salePrice || 0), Number(p.total !== undefined ? p.total : qty * buy)];
    });
    mergeRows_(sh("Products"), pH, pRows, function (r) { return String(r[1]) + '|' + String(r[2]); }, [0]);
  }

  if (allow.stock && data.stock) {
    var sH = ["Date", "Id", "Purchase No", "Product Name", "Category", "Purchased Qty", "Sold Qty", "Current Stock", "Sale Price", "Purchase Price"];
    var sRows = data.stock.map(function (s) {
      return [s.date || now, s.id || s.productId || "", s.purchaseNo || "", s.productName || "", s.category || "",
              Number(s.purchasedQty !== undefined ? s.purchasedQty : (s.purchaseQty || 0)), Number(s.soldQty || 0),
              Number(s.currentStock || 0), Number(s.salePrice || 0), Number(s.purchasePrice || 0)];
    });
    // salesman may only change Sold Qty / Current Stock of existing rows
    var sProtect = role === 'cashier' ? [0, 2, 3, 4, 5, 8, 9] : [0];
    mergeRows_(sh("Stock"), sH, sRows, function (r) { return String(r[1]) + '|' + String(r[2]); }, sProtect);
  }

  if (allow.sales && data.sales) {
    var slH = ["Date", "Invoice No", "Customer Name", "Phone", "Product Name", "Qty", "Sale Price", "Discount", "Purchase Price", "Payment", "Remarks", "Payment Method", "Profit"];
    var slRows = data.sales.map(function (s) {
      var prodName = s.productName || (s.itemsSummary || "").split("(")[0].trim() || (s.items && s.items[0] ? s.items[0].productName : "Item");
      var qty = s.qty || (s.items && s.items.reduce(function (a, it) { return a + it.qty; }, 0)) || 1;
      var saleP = s.salePrice || s.subtotal || s.netTotal || 0;
      var purchaseP = s.purchasePrice || 0;
      var prof = s.profit !== undefined ? s.profit : Math.max(0, (s.netTotal || saleP) - purchaseP);
      return [s.date || "", s.invoiceNo || "", s.customerName || "", s.phone || "N/A", prodName, qty, saleP, s.discount || 0,
              purchaseP, s.payment || s.netTotal || 0, s.remarks || "-", s.paymentMethod || "Cash", prof];
    });
    // existing invoices: salesman cannot alter anything except via admin/manager edit
    var slProtect = role === 'cashier' ? [0, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12] : [0];
    mergeRows_(sh("Sales"), slH, slRows, function (r) { return String(r[1]); }, slProtect);
  }

  if (allow.customers && data.customers) {
    var cH = ["Customer ID", "Customer Name", "Phone", "Total Purchases", "Total Paid", "Current Due", "Status"];
    var cRows = data.customers.map(function (c) {
      return [c.id || "", c.name || "", c.phone || "", c.totalPurchases || 0, c.totalPaid || 0, c.currentDue || 0, c.status || "Regular"];
    });
    mergeRows_(sh("Customers"), cH, cRows, function (r) { return String(r[0]); }, []);
  }

  if (allow.duePayments && data.duePayments) {
    var dH = ["Collection Date", "Collection ID", "Customer Name", "Phone", "Amount Collected", "Payment Method", "Remaining Due", "Collected By", "Remarks"];
    var dRows = data.duePayments.map(function (d) {
      return [d.date || "", d.id || "", d.customerName || "", d.phone || "", d.amount || 0, d.paymentMethod || "Cash", d.remainingDue || 0, d.collectedBy || "", d.remarks || ""];
    });
    mergeRows_(sh("Due_Collections"), dH, dRows, function (r) { return String(r[1]); }, [0]);
  }

  if (allow.expenses && data.expenses) {
    var eH = ["Date", "Expense ID", "Category", "Amount", "Note", "Payment Method", "Logged By"];
    var eRows = data.expenses.map(function (e) {
      return [e.date || "", e.id || "", e.category || "", e.amount || 0, e.note || "", e.paymentMethod || "Cash", e.loggedBy || ""];
    });
    mergeRows_(sh("Expenses"), eH, eRows, function (r) { return String(r[1]); }, [0]);
  }

  if (role === 'admin' && data.companyInfo && typeof data.companyInfo === "object") {
    saveCompanyInfoDirect(data.companyInfo);
  }
  return { success: true, message: "গুগল শিট আপডেট হয়েছে (কোনো ডাটা মোছা হয়নি)" };
}

/** Pictures/files: stored in Google Drive, only the link goes into the sheet (a cell holds max 50,000 chars) */
function uploadFile_(name, mimeType, base64) {
  var folders = DriveApp.getFoldersByName("POS_Uploads");
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder("POS_Uploads");
  var blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType || "application/octet-stream", name || "file");
  var f = folder.createFile(blob);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { success: true, id: f.getId(), url: "https://drive.google.com/uc?export=view&id=" + f.getId() };
}
