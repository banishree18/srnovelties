// ============================================
// SR NOVELTIES AUTOMATION – Code.gs
// ============================================
const SS = SpreadsheetApp.getActiveSpreadsheet();
const ORDER_SHEET = SS.getSheetByName("Order") || SS.insertSheet("Order");
const INVENTORY_SHEET = SS.getSheetByName("Inventory") || SS.insertSheet("Inventory");
const CUSTOMER_SHEET = SS.getSheetByName("Customer") || SS.insertSheet("Customer");

// Owner's email — stock alerts and "product not found" alerts go here.
const OWNER_EMAIL = "monali.kothari2911@gmail.com";

// ---------- Initialize headers if sheets are empty ----------
function initSheets() {
  if (ORDER_SHEET.getLastRow() === 0) {
    ORDER_SHEET.appendRow(["Order_Date", "Order_ID", "Customer_ID", "Customer_Name",
      "product_Name", "Phone", "Category", "Qty", "Unit_price",
      "Total_amount", "Address", "Payment_Status", "Notes"]);
  }
  if (CUSTOMER_SHEET.getLastRow() === 0) {
    CUSTOMER_SHEET.appendRow(["Customer_ID", "Customer_Name", "Phone",
      "Total_Orders", "Address", "Last_Order_Date"]);
  }
  if (INVENTORY_SHEET.getLastRow() === 0) {
    INVENTORY_SHEET.appendRow(["Product_ID", "Product_Name", "Category", "Supplier_Name",
      "Unit_price", "Units_sold", "Total_Amount", "Initial_Stock",
      "Stock_Quantity", "Warning", "Reorder_level"]);
  }
}

// ---------- Generate unique Order ID ----------
function generateOrderID() {
  const orders = ORDER_SHEET.getDataRange().getValues();
  let maxId = 0;
  orders.forEach(row => {
    const id = row[1]; // Order_ID column
    if (typeof id === 'string' && id.startsWith('ORD-')) {
      const num = parseInt(id.replace('ORD-', '')) || 0;
      if (num > maxId) maxId = num;
    }
  });
  return `ORD-${(maxId + 1).toString().padStart(3, '0')}`;
}

// ---------- Serve product list to the landing page ----------
// Whatever rows/categories exist in the Inventory sheet, this returns them as JSON.
// This is what makes new products AND new categories show up on the website automatically —
// the frontend does not hardcode any category name, it just groups by whatever comes back here.
function doGet(e) {
  const inventory = INVENTORY_SHEET.getDataRange().getValues();
  const headers = inventory[0].map(h => h.toString().trim().toLowerCase());

  const idCol = headers.indexOf("product_id");
  const nameCol = headers.indexOf("product_name");
  const catCol = headers.indexOf("category");
  const priceCol = headers.indexOf("unit_price");
  const stockCol = headers.indexOf("stock_quantity");
  const subcatCol = headers.indexOf("subcategory");
  const imageCol = headers.indexOf("image_url");
  const descCol = headers.indexOf("description");

  const products = [];
  for (let i = 1; i < inventory.length; i++) {
    const row = inventory[i];
    const name = (row[nameCol] || "").toString().trim();
    if (!name) continue;
    products.push({
      id: idCol >= 0 ? row[idCol] : "",
      name: name,
      category: (row[catCol] || "Other").toString().trim() || "Other",
      price: Number(row[priceCol]) || 0,
      stock: Number(row[stockCol]) || 0,
      subcategory: subcatCol >= 0 ? (row[subcatCol] || "").toString().trim() : "",
      image: imageCol >= 0 ? (row[imageCol] || "").toString().trim() : "",
      description: descCol >= 0 ? (row[descCol] || "").toString().trim() : ""
    });
  }

  const responseObj = { status: "success", products: products };

  // JSONP support — this is the fix. If the request includes a "callback" param,
  // wrap the response as JavaScript instead of plain JSON.
  const callback = e && e.parameter && e.parameter.callback;
  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + JSON.stringify(responseObj) + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService.createTextOutput(JSON.stringify(responseObj))
    .setMimeType(ContentService.MimeType.JSON);
}
// ---------- Main Webhook ----------
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const orderDate = new Date();
    const orderId = generateOrderID();
    const qty = parseInt(data.qty) || 1;

    // 1. Check inventory BEFORE saving anything
    const stockResult = updateInventory(data.product, qty, Number(data.unit_price) || 0);

    if (!stockResult.success) {
      sendOwnerWarning(data, stockResult);
      return ContentService.createTextOutput(JSON.stringify({
        status: "failed",
        message: stockResult.matched
          ? `Sorry, only ${stockResult.available} unit(s) of "${data.product}" left in stock.`
          : `"${data.product}" could not be matched in our inventory. Our team has been notified.`,
        orderId: orderId
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 2. Find or create the customer – SAME id used in both sheets
    const customerId = getOrCreateCustomer(data.name, data.phone, data.address);

    // 3. Save to Order Sheet
    ORDER_SHEET.appendRow([
      orderDate,
      orderId,
      customerId,
      data.name,
      data.product,
      data.phone,
      data.category,
      qty,
      data.unit_price,
      data.total_amount,
      data.address || "",
      "Pending",
      data.note || ""
    ]);

    // 4. Confirmation back to the customer (shown as the "Order placed!" alert on the site)
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "Order placed successfully!",
      orderId: orderId
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    console.error(err);
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: "Server error: " + err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ---------- Inventory check + update ----------
// Returns { success, available, matched }
// matched=false means the product name wasn't found in the sheet at all (a data problem, not a stock problem)
function updateInventory(productName, qty, unitPrice) {
  const inventory = INVENTORY_SHEET.getDataRange().getValues();
  const headers = inventory[0].map(h => h.toString().toLowerCase().trim());

  const nameCol = headers.indexOf("product_name") + 1 || 2;
  const stockCol = headers.indexOf("stock_quantity") + 1 || 9;
  const soldCol = headers.indexOf("units_sold") + 1 || 0;
  const totalCol = headers.indexOf("total_amount") + 1 || 0;

  const target = (productName || "").toString().trim().toLowerCase();
  let rowIndex = -1;

  // Pass 1: exact match (preferred, avoids picking the wrong size/variant)
  for (let i = 1; i < inventory.length; i++) {
    const rowProduct = (inventory[i][nameCol - 1] || "").toString().trim().toLowerCase();
    if (rowProduct === target) { rowIndex = i; break; }
  }
  // Pass 2: partial match, only if no exact match was found
  if (rowIndex === -1) {
    for (let i = 1; i < inventory.length; i++) {
      const rowProduct = (inventory[i][nameCol - 1] || "").toString().trim().toLowerCase();
      if (rowProduct && rowProduct.includes(target)) { rowIndex = i; break; }
    }
  }

  if (rowIndex === -1) {
    // Product name from the website doesn't exist in the Inventory sheet at all
    return { success: false, available: 0, matched: false };
  }

  const currentStock = Number(inventory[rowIndex][stockCol - 1]) || 0;

  if (currentStock < qty) {
    return { success: false, available: currentStock, matched: true };
  }

  const newStock = currentStock - qty;
  INVENTORY_SHEET.getRange(rowIndex + 1, stockCol).setValue(newStock);

  if (soldCol) {
    const currentSold = Number(inventory[rowIndex][soldCol - 1]) || 0;
    INVENTORY_SHEET.getRange(rowIndex + 1, soldCol).setValue(currentSold + qty);
  }
  if (totalCol) {
    const currentTotal = Number(inventory[rowIndex][totalCol - 1]) || 0;
    INVENTORY_SHEET.getRange(rowIndex + 1, totalCol).setValue(currentTotal + qty * unitPrice);
  }

  return { success: true, available: newStock, matched: true };
}

// ---------- Owner warning email ----------
function sendOwnerWarning(orderData, stockResult) {
  const notFound = !stockResult.matched;
  const subject = notFound
    ? `⚠️ PRODUCT NOT FOUND IN INVENTORY - ${orderData.product}`
    : `⚠️ LOW STOCK ALERT - ${orderData.product}`;

  const body = notFound
    ? `A customer tried to order "${orderData.product}" but this exact name wasn't found in the Inventory sheet's Product_Name column.\n\n` +
      `Customer: ${orderData.name} (${orderData.phone})\nQty requested: ${orderData.qty}\n\n` +
      `Fix: make sure the Inventory sheet's Product_Name matches the website's product name (or add the missing row).`
    : `Product: ${orderData.product}\nRequested: ${orderData.qty}\nAvailable: ${stockResult.available}\n` +
      `Customer: ${orderData.name} (${orderData.phone})\n\nThis order was NOT saved – please restock and ask the customer to reorder.`;

  MailApp.sendEmail(OWNER_EMAIL, subject, body);
}

// ---------- Find or create a customer, return one consistent Customer_ID ----------
function getOrCreateCustomer(name, phone, address) {
  const customers = CUSTOMER_SHEET.getDataRange().getValues();
  const headers = customers[0].map(h => h.toString().trim().toLowerCase());
  const cleanPhone = (phone || "").toString().trim();

  const idCol      = headers.indexOf("customer_id");
  const nameCol    = headers.indexOf("customer_name");
  const phoneCol   = headers.indexOf("phone");
  const totalCol   = headers.indexOf("total_orders");
  const joinCol    = headers.indexOf("join date");
  const activityCol= headers.indexOf("activity_status");
  const typeCol    = headers.indexOf("customer_type");

  for (let i = 1; i < customers.length; i++) {
    const rowPhone = (customers[i][phoneCol] || "").toString().trim();
    if (rowPhone && rowPhone === cleanPhone) {
      const newTotal = (Number(customers[i][totalCol]) || 0) + 1;
      CUSTOMER_SHEET.getRange(i + 1, totalCol + 1).setValue(newTotal);
      if (activityCol >= 0) {
        CUSTOMER_SHEET.getRange(i + 1, activityCol + 1).setValue("Active - " + new Date().toLocaleDateString("en-IN"));
      }
      if (typeCol >= 0) {
        CUSTOMER_SHEET.getRange(i + 1, typeCol + 1).setValue(newTotal >= 5 ? "Premium" : "Regular");
      }
      return customers[i][idCol]; // existing Customer_ID
    }
  }

  // New customer
  const newId = `CUST-${Date.now().toString().slice(-6)}`;
  const row = [];
  row[idCol] = newId;
  row[nameCol] = name;
  row[phoneCol] = cleanPhone;
  row[totalCol] = 1;
  if (joinCol >= 0) row[joinCol] = new Date();
  if (activityCol >= 0) row[activityCol] = "Active - " + new Date().toLocaleDateString("en-IN");
  if (typeCol >= 0) row[typeCol] = "Regular";
  CUSTOMER_SHEET.appendRow(row);
  return newId;
}
function testMail() {
  MailApp.sendEmail(OWNER_EMAIL, "Test Email", "This is a permission test.");
}
function checkWhichSheet() {
  Logger.log("Spreadsheet Name: " + SS.getName());
  Logger.log("Spreadsheet URL: " + SS.getUrl());
  Logger.log("Order sheet last row: " + ORDER_SHEET.getLastRow());
  Logger.log("Customer sheet last row: " + CUSTOMER_SHEET.getLastRow());
}
