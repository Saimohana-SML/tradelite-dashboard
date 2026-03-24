// Configuration: Define the names of your Google Sheets tabs
const SHEET_PRODUCTS = "Products";
const SHEET_ORDERS = "Orders";
const SHEET_CATEGORIES = "Categories";
const SHEET_EMPLOYEES = "Employees";

/**
 * Main GET handler. 
 * Can serve HTML pages OR JSON data based on parameters.
 */
function doGet(e) {
  var action = e.parameter.action;
  var page = e.parameter.page || 'admin-dashboard'; 

  // If there's an action, it's a DATA request (read OR write via GET)
  if (action) {
    try {
      var result;
      switch (action) {
        // ── READ actions ──
        case 'getProducts':         result = getProducts(); break;
        case 'getOrders':           result = getOrders(); break;
        case 'getEmployees':        result = getEmployees(); break;
        case 'getCategories':       result = getCategories(); break;
        case 'generateCatalogJson': result = generateCatalogJson(); break;

        // ── WRITE actions via GET (avoids POST→redirect→GET body-drop bug) ──
        case 'addProduct': {
          var payload = JSON.parse(e.parameter.payload || '{}');
          result = addProduct(payload);
          break;
        }
        case 'editProduct': {
          var payload = JSON.parse(e.parameter.payload || '{}');
          result = editProduct(payload);
          break;
        }
        case 'addEmployee': {
          var payload = JSON.parse(e.parameter.payload || '{}');
          result = addEmployee(payload);
          break;
        }
        case 'addOrder': {
          var payload = JSON.parse(e.parameter.payload || '{}');
          result = addOrder(payload);
          break;
        }
        case 'updateOrderStatus': {
          var payload = JSON.parse(e.parameter.payload || '{}');
          result = updateOrderStatus(payload);
          break;
        }

        default: return createJsonResponse({ status: 'error', message: 'Unknown action: ' + action });
      }
      return createJsonResponse({ status: 'success', data: result });
    } catch (err) {
      return createJsonResponse({ status: 'error', message: err.toString() });
    }
  }

  // Otherwise, it's a PAGE request
  try {
    return HtmlService.createTemplateFromFile(page)
        .evaluate()
        .setTitle('TradeLite Business Suite')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput("<h1>Error 404</h1><p>Page '" + page + "' not found in Script project.</p><p>Technical Error: " + err.toString() + "</p>");
  }
}

/**
 * Helper to include other HTML files
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Get the main URL of the script
 */
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * Generate a URL for a specific page
 */
function getPageUrl(page) {
  return getScriptUrl() + "?page=" + (page || 'admin-dashboard');
}



/**
 * Handle POST Requests
 */
function doPost(e) {
  let requestBody;
  
  try {
    if (e.postData && e.postData.contents) {
      requestBody = JSON.parse(e.postData.contents);
    } else {
      throw new Error("No payload found");
    }
  } catch (error) {
    return createJsonResponse({ status: "error", message: "Invalid JSON body" }, 400);
  }

  const action = requestBody.action;
  const payload = requestBody.payload;

  try {
    if (action === "addProduct") {
      return createJsonResponse({ status: "success", data: addProduct(payload) });
    } else if (action === "editProduct") {
      return createJsonResponse({ status: "success", data: editProduct(payload) });
    } else if (action === "addOrder") {
      return createJsonResponse({ status: "success", data: addOrder(payload) });
    } else if (action === "updateOrderStatus") {
      return createJsonResponse({ status: "success", data: updateOrderStatus(payload) });
    } else if (action === "addEmployee") {
      return createJsonResponse({ status: "success", data: addEmployee(payload) });
    } else {
      return createJsonResponse({ status: "error", message: "Invalid action" }, 400);
    }
  } catch (error) {
    return createJsonResponse({ status: "error", message: error.toString() }, 500);
  }
}

// ==========================================
// ROBUST DATA HELPERS (NEW)
// ==========================================

/**
 * Normalizes a header string (lowercase, removes spaces/underscores) for matching.
 */
function normalizeStr(str) {
  return str.toString().toLowerCase().replace(/[^a-z0-9]/g, "");
}

/**
 * Finds the index (0-based) of a header in an array.
 */
function getHeaderIndex(headers, target) {
  const normalizedTarget = normalizeStr(target);
  return headers.findIndex(h => normalizeStr(h) === normalizedTarget);
}

/**
 * Converts a sheet's data into an array of objects based on headers.
 * Automatically filters out rows where the specified 'idColumn' is empty.
 */
function sheetToObjects(sheetName, idColumnName) {
  const sheet = getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const idIndex = getHeaderIndex(headers, idColumnName);
  const results = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // Always skip completely empty rows (all cells blank)
    const isRowEmpty = row.every(cell => cell === null || cell === undefined || cell.toString().trim() === "");
    if (isRowEmpty) continue;

    // If we found the ID column, skip rows where the ID cell is empty
    if (idIndex !== -1 && (!row[idIndex] || row[idIndex].toString().trim() === "")) continue;

    const obj = {};
    headers.forEach((header, index) => {
      const key = header.toString().replace(/ /g, "_");
      obj[key] = row[index];
    });
    results.push(obj);
  }
  return results;
}

// ==========================================
// BUSINESS LOGIC: PRODUCTS
// ==========================================

function getProducts() {
  const products = sheetToObjects(SHEET_PRODUCTS, "ProductID");
  // Ensure every product has the standardized keys the frontend expects
  return products.map(p => ({
    productId: p.ProductID || p.ProductID_ || "",
    name: p.ProductName || p.Name || p.Product_Name || "",
    price: parseFloat(p.Price || p.Selling_Price || 0),
    category: p.Category || "",
    subCategory: p.SubCategory || p.Sub_Category || "",
    description: p.Description || "",
    tags: p.Tags ? p.Tags.split(",").map(t => t.trim()) : [],
    images: p.ImageURL ? p.ImageURL.split(",").map(t => t.trim()) : [],
    availability: (parseInt(p.InStock || 0) > 0) ? "In Stock" : "Out of Stock",
    featured: p.Featured || p.Status || "Active",
    sortOrder: p.SortOrder || ""
  }));
}

function addProduct(payload) {
  const sheet = getSheetByName(SHEET_PRODUCTS);
  const productId = payload.productId || Utilities.getUuid();
  
  const newRow = [
    productId,
    payload.name || "",
    payload.price || 0,
    payload.category || "",
    payload.subCategory || "",
    payload.description || "",
    Array.isArray(payload.tags) ? payload.tags.join(", ") : (payload.tags || ""),
    Array.isArray(payload.images) ? payload.images.join(", ") : (payload.images || payload.ImageURL || ""),
    payload.inStock || payload.availability || payload.stock || 0,
    payload.featured || payload.status || "Active",
    payload.sortOrder || ""
  ];
  
  sheet.appendRow(newRow);
  return { id: productId, message: "Product added successfully" };
}

function editProduct(payload) {
  const sheet = getSheetByName(SHEET_PRODUCTS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIndex = getHeaderIndex(headers, "ProductID");
  const productId = payload.productId;
  
  const rowIndex = data.findIndex((row, idx) => idx > 0 && row[idIndex].toString() === productId.toString()) + 1;
  if (rowIndex <= 0) throw new Error("Product ID not found.");

  const updatedRow = [
    productId,
    payload.name || "",
    payload.price || 0,
    payload.category || "",
    payload.subCategory || "",
    payload.description || "",
    Array.isArray(payload.tags) ? payload.tags.join(", ") : (payload.tags || ""),
    Array.isArray(payload.images) ? payload.images.join(", ") : (payload.images || payload.ImageURL || ""),
    payload.inStock || payload.availability || payload.stock || 0,
    payload.featured || payload.status || "Active",
    payload.sortOrder || ""
  ];
  
  sheet.getRange(rowIndex, 1, 1, updatedRow.length).setValues([updatedRow]);
  return { id: productId, message: "Product updated successfully" };
}

// ==========================================
// BUSINESS LOGIC: CATEGORIES & CATALOG
// ==========================================

function getCategories() {
  const sheet = getSheetByName(SHEET_CATEGORIES);
  const data = sheet.getDataRange().getValues();
  const catMap = {};
  
  data.forEach(row => {
    if (row[0]) {
      const subs = row[1] ? row[1].toString().split(",").map(s => s.trim()).filter(s => s) : [];
      catMap[row[0].toString().trim()] = subs;
    }
  });
  return catMap;
}

function generateCatalogJson() {
  const allProducts = getProducts();
  const activeProducts = allProducts.filter(p => ["active", "published"].includes(p.featured.toLowerCase()));
  
  return { 
    lastGenerated: new Date().toISOString(), 
    count: activeProducts.length, 
    catalog: activeProducts 
  };
}

// ==========================================
// BUSINESS LOGIC: ORDERS 
// ==========================================

function getOrders() {
  const rawOrders = sheetToObjects(SHEET_ORDERS, "Order_ID");
  return rawOrders.map(o => {
    let items = [];
    if (o.Items_JSON) {
      try { items = JSON.parse(o.Items_JSON); } catch(e) {}
    }
    return {
      ...o,
      Items: items
    };
  });
}

function addOrder(payload) {
  const sheet = getSheetByName(SHEET_ORDERS);
  const orderId = Utilities.getUuid();
  const newRow = [
    orderId,
    new Date(),
    payload.customerName || "Unknown",
    payload.status || "New",
    JSON.stringify(payload.items || []),
    payload.totalAmount || 0
  ];
  sheet.appendRow(newRow);
  return { id: orderId, message: "Order processed successfully" };
}

function updateOrderStatus(payload) {
  const sheet = getSheetByName(SHEET_ORDERS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIndex = getHeaderIndex(headers, "Order_ID");
  const statusIndex = getHeaderIndex(headers, "Status");
  
  const rowIndex = data.findIndex((row, idx) => idx > 0 && row[idIndex].toString() === payload.orderId.toString()) + 1;
  if (rowIndex <= 0) throw new Error("Order ID not found.");

  sheet.getRange(rowIndex, statusIndex + 1).setValue(payload.status);
  return { id: payload.orderId, status: payload.status, message: "Order status updated." };
}

// ==========================================
// EMPLOYEES CORE LOGIC
// ==========================================
function getEmployees() {
  return sheetToObjects(SHEET_EMPLOYEES, "Employee_ID");
}

function addEmployee(payload) {
  const sheet = getSheetByName(SHEET_EMPLOYEES);
  sheet.appendRow([
    Utilities.getUuid(),
    payload.name || "Unknown",
    payload.mobile || "",
    payload.email || "",
    payload.status || "Active",
    new Date()
  ]);
  return { message: "Employee registered successfully" };
}

// ==========================================
// GLOBALS & UTILITIES
// ==========================================

function getSheetByName(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    if (sheetName === SHEET_PRODUCTS) {
      sheet.appendRow(["ProductID", "ProductName", "Price", "Category", "SubCategory", "Description", "Tags", "ImageURL", "InStock", "Featured", "SortOrder"]);
    } else if (sheetName === SHEET_ORDERS) {
      sheet.appendRow(["Order_ID", "Timestamp", "Customer_Name", "Status", "Items_JSON", "Total_Amount"]);
    } else if (sheetName === SHEET_EMPLOYEES) {
      sheet.appendRow(["Employee_ID", "Name", "Mobile", "Email", "Status", "JoinedDate"]);
    }
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function doOptions(e) {
  return createJsonResponse({ status: "ok" });
}
