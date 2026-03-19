// Configuration: Define the names of your Google Sheets tabs
const SHEET_PRODUCTS = "Products";
const SHEET_ORDERS = "Orders";
const SHEET_CATEGORIES = "Categories";
const SHEET_EMPLOYEES = "Employees";

/**
 * Handle GET Requests
 * Used for fetching data from Google Sheets (e.g. loading the product list).
 * Example URL: https://script.google.com/.../exec?action=getProducts
 */
function doGet(e) {
  // Always include CORS headers for frontend integration
  const action = e.parameter.action;
  
  try {
    if (action === "getProducts") {
      return createJsonResponse(getProducts());
    } else if (action === "getOrders") {
      return createJsonResponse(getOrders()); 
    } else if (action === "getCategories") {
      return createJsonResponse(getCategories());
    } else if (action === "generateCatalogJson") {
      return createJsonResponse(generateCatalogJson());
    } else if (action === "getEmployees") {
      return createJsonResponse(getEmployees());
    } else {
      return createJsonResponse({ status: "error", message: "Invalid action parameter" }, 400);
    }
  } catch (error) {
    return createJsonResponse({ status: "error", message: error.toString() }, 500);
  }
}

/**
 * Handle POST Requests
 * Used for receiving form data from the Vanilla JS frontend (e.g. creating a new product).
 */
function doPost(e) {
  let requestBody;
  
  try {
    // Parse the incoming JSON body from the frontend fetch() request
    if (e.postData && e.postData.contents) {
      requestBody = JSON.parse(e.postData.contents);
    } else {
      throw new Error("No payload found");
    }
  } catch (error) {
    return createJsonResponse({ status: "error", message: "Invalid JSON body or empty payload" }, 400);
  }

  const action = requestBody.action;
  const payload = requestBody.payload;

  try {
    if (action === "addProduct") {
      const result = addProduct(payload);
      return createJsonResponse({ status: "success", data: result });
    } else if (action === "editProduct") {
      const result = editProduct(payload);
      return createJsonResponse({ status: "success", data: result });
    } else if (action === "addOrder") {
      const result = addOrder(payload);
      return createJsonResponse({ status: "success", data: result });
    } else if (action === "updateOrderStatus") {
      const result = updateOrderStatus(payload);
      return createJsonResponse({ status: "success", data: result });
    } else if (action === "addEmployee") {
      const result = addEmployee(payload);
      return createJsonResponse({ status: "success", data: result });
    } else {
      return createJsonResponse({ status: "error", message: "Invalid action" }, 400);
    }
  } catch (error) {
    return createJsonResponse({ status: "error", message: error.toString() }, 500);
  }
}

// ==========================================
// BUSINESS LOGIC: PRODUCTS
// ==========================================

function addProduct(payload) {
  const sheet = getSheetByName(SHEET_PRODUCTS);
  
  // Handle IDs
  const productId = payload.productId || Utilities.getUuid();
  
  // Map fields directly to the user's 11 exact columns:
  // ProductID, ProductName, Price, Category, SubCategory, Description, Tags, ImageURL, InStock, Featured, SortOrder
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
  
  // Append to the next available row in the Google Sheet
  sheet.appendRow(newRow);
  
  return { id: productId, message: "Product added successfully" };
}

function editProduct(payload) {
  const sheet = getSheetByName(SHEET_PRODUCTS);
  const data = sheet.getDataRange().getValues();
  const productId = payload.productId;
  
  if (!productId) throw new Error("Product ID is required for editing.");
  
  // Find the row with matching ID (assuming ID is in Column A, index 0)
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) { // Skip headers
    if (data[i][0].toString() === productId.toString()) {
      rowIndex = i + 1; // Apps Script ranges are 1-indexed
      break;
    }
  }
  
  if (rowIndex === -1) throw new Error("Product ID not found.");
  
  // Create updated row exactly matching the current structure
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
  
  // Write the updated row back to the exact range
  sheet.getRange(rowIndex, 1, 1, 11).setValues([updatedRow]);
  
  return { id: productId, message: "Product updated successfully" };
}

function getProducts() {
  const sheet = getSheetByName(SHEET_PRODUCTS);
  const data = sheet.getDataRange().getValues();
  
  // If only headers exist (or sheet is completely empty)
  if (data.length <= 1) return []; 
  
  const headers = data[0];
  const products = [];
  
  // Convert 2D array from sheet into JSON objects matching the headers
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const product = {};
    for (let j = 0; j < headers.length; j++) {
      // Clean header name e.g., 'Created At' to 'Created_At' or just use as is
      product[headers[j]] = row[j];
    }
    products.push(product);
  }
  return products;
}

// ==========================================
// BUSINESS LOGIC: CATEGORIES & CATALOG
// ==========================================

function getCategories() {
  const sheet = getSheetByName(SHEET_CATEGORIES);
  const data = sheet.getDataRange().getValues();
  const catMap = {};
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue; // Skip empty rows
    
    // SubCategories are expected to be comma separated in column 2
    const subs = row[1] ? row[1].toString().split(",").map(s => s.trim()).filter(s => s) : [];
    catMap[row[0].trim()] = subs;
  }
  
  return catMap;
}

function generateCatalogJson() {
  const sheet = getSheetByName(SHEET_PRODUCTS);
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) return { lastGenerated: new Date().toISOString(), count: 0, catalog: [] };
  
  const products = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    // Filter active products
    const status = row[9] ? row[9].toString().toLowerCase() : "";
    if (status !== "active" && status !== "published") continue; 
    
    const qty = parseInt(row[8]) || 0;
    const availabilityStatus = qty > 0 ? "In Stock" : "Out of Stock";
    
    products.push({
      productId: row[0].toString(),
      name: row[1].toString(),
      price: parseFloat(row[2]) || 0,
      category: row[3].toString(),
      subCategory: row[4].toString(),
      description: row[5].toString(),
      tags: row[6].toString() ? row[6].toString().split(',').map(tag => tag.trim()) : [],
      images: row[7].toString() ? row[7].toString().split(',').map(img => img.trim()) : [],
      availability: availabilityStatus,
      featured: status,
      sortOrder: row[10].toString()
    });
  }
  
  return { 
    lastGenerated: new Date().toISOString(), 
    count: products.length, 
    catalog: products 
  };
}

// ==========================================
// BUSINESS LOGIC: ORDERS 
// ==========================================

function getOrders() {
  const sheet = getSheetByName(SHEET_ORDERS);
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) return []; 
  
  const headers = data[0];
  const orders = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const order = {};
    for (let j = 0; j < headers.length; j++) {
      order[headers[j]] = row[j];
    }
    
    // Automatically parse the JSON items if they exist so frontend receives a proper array
    if (order["Items_JSON"]) {
      try {
        order["Items"] = JSON.parse(order["Items_JSON"]);
      } catch(e) {
        order["Items"] = [];
      }
    }
    
    orders.push(order);
  }
  return orders;
}

function addOrder(payload) {
  const sheet = getSheetByName(SHEET_ORDERS);
  const orderId = Utilities.getUuid();
  
  const newRow = [
    orderId,
    new Date(),
    payload.customerName || "Unknown",
    payload.status || "New",
    JSON.stringify(payload.items || []), // Store items array as JSON string
    payload.totalAmount || 0
  ];
  
  sheet.appendRow(newRow);
  
  // ----------------------------------------
  // INTEGRATION: EMAIL (GmailApp)
  // ----------------------------------------
  // Uncomment the lines below to auto-email the admin when a new order arrives!
  /*
  const adminEmail = Session.getActiveUser().getEmail(); 
  const subject = `New Order Received: ${orderId}`;
  const body = `A new order has been placed by ${payload.customerName}.\nTotal Amount: ₹${payload.totalAmount}`;
  GmailApp.sendEmail(adminEmail, subject, body);
  */
  
  // ----------------------------------------
  // INTEGRATION: DOCUMENTS (Google Docs)
  // ----------------------------------------
  // You can automatically generate an invoice document using DocumentApp here.
  
  return { id: orderId, message: "Order processed successfully" };
}

function updateOrderStatus(payload) {
  const sheet = getSheetByName(SHEET_ORDERS);
  const data = sheet.getDataRange().getValues();
  const orderId = payload.orderId;
  const newStatus = payload.status;
  
  if (!orderId || !newStatus) throw new Error("Order ID and Status are required.");
  
  // Find the row with matching ID (assuming Order_ID is in Column A, index 0)
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) { // Skip headers
    if (data[i][0].toString() === orderId.toString()) {
      rowIndex = i + 1;
      break;
    }
  }
  
  if (rowIndex === -1) throw new Error("Order ID not found.");
  
  // Update the Status column (Assuming Status is Column D, index 3 meaning column index 4 for getRange)
  // Let's verify headers: ["Order_ID", "Timestamp", "Customer_Name", "Status", "Items_JSON", "Total_Amount"] => index 3 => Column D (row: rowIndex, col: 4)
  sheet.getRange(rowIndex, 4).setValue(newStatus);
  
  return { id: orderId, status: newStatus, message: "Order status updated." };
}

// ==========================================
// EMPLOYEES CORE LOGIC
// ==========================================
function getEmployees() {
  const sheet = getSheetByName(SHEET_EMPLOYEES);
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) return []; // Only headers or empty
  
  const headers = data[0];
  const employees = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    let employeeData = {};
    for (let j = 0; j < headers.length; j++) {
      employeeData[headers[j]] = row[j];
    }
    employees.push(employeeData);
  }
  return employees;
}

function addEmployee(payload) {
  const sheet = getSheetByName(SHEET_EMPLOYEES);
  const employeeId = Utilities.getUuid();
  
  const newRow = [
    employeeId,
    payload.name || "Unknown",
    payload.mobile || "",
    payload.email || "",
    payload.role || "Sales Representative",
    payload.status || "Active",
    new Date()
  ];
  
  sheet.appendRow(newRow);
  
  return { id: employeeId, message: "Employee registered successfully" };
}

// ==========================================
// GLOBALS & UTILITIES
// ==========================================

/**
 * Helper to grab a sheet by name. If it doesn't exist, it auto-creates it
 * and applies the necessary header row so it's ready to use instantly.
 */
function getSheetByName(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    
    // Auto-setup headers dynamically based on which sheet is requested
    if (sheetName === SHEET_PRODUCTS) {
      sheet.appendRow(["ProductID", "ProductName", "Price", "Category", "SubCategory", "Description", "Tags", "ImageURL", "InStock", "Featured", "SortOrder"]);
      sheet.getRange(1, 1, 1, 11).setFontWeight("bold");
    } else if (sheetName === SHEET_ORDERS) {
      sheet.appendRow(["Order_ID", "Timestamp", "Customer_Name", "Status", "Items_JSON", "Total_Amount"]);
      sheet.getRange(1, 1, 1, 6).setFontWeight("bold");
    } else if (sheetName === SHEET_CATEGORIES) {
      sheet.appendRow(["Women", "Kurtis, Sarees, Dupattas, Suits, Bottoms"]);
      sheet.appendRow(["Men", "Shirts, Ethnic Wear, Bottomwear, Casuals"]);
      sheet.appendRow(["Kids", "Ethnic Wear, Frocks, Boys Wear, Accessories"]);
      sheet.appendRow(["Accessories", "Jewellery, Bags, Footwear"]);
    } else if (sheetName === SHEET_EMPLOYEES) {
      sheet.appendRow(["Employee_ID", "Name", "Mobile", "Email", "Role", "Status", "JoinedDate"]);
      sheet.getRange(1, 1, 1, 7).setFontWeight("bold");
    }
    
    // Freeze the top row so headers stay visible when scrolling
    sheet.setFrozenRows(1);
  }
  
  return sheet;
}

/**
 * Formats the response into a proper JSON Web App output
 */
function createJsonResponse(data, statusCode = 200) {
  // Note: Google Apps Script 'ContentService' currently lacks a direct way to set HTTP Status codes dynamically 
  // without returning custom error pages, so we handle success/error strictly via the JSON payload body.
  const stringified = JSON.stringify(data);
  return ContentService.createTextOutput(stringified).setMimeType(ContentService.MimeType.JSON);
}

// -------------------------------------------------------------
// REQUIRED: Avoid CORS Preflight Issues with Vanilla JS Setup 
// -------------------------------------------------------------
function doOptions(e) {
  return createJsonResponse({ status: "ok" });
}
