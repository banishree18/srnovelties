// ============================================
// SR NOVELTIES AUTOMATION – Code.gs
// Version 2 – adds order validation, server-side pricing and LockService
// ============================================
const SS = SpreadsheetApp.getActiveSpreadsheet();
const ORDER_SHEET = SS.getSheetByName("Order") || SS.insertSheet("Order");
const INVENTORY_SHEET = SS.getSheetByName("Inventory") || SS.insertSheet("Inventory");
const CUSTOMER_SHEET = SS.getSheetByName("Customer") || SS.insertSheet("Customer");

// Owner's email — stock alerts and "product not found" alerts go here.
const OWNER_EMAIL = "monali.kothari2911@gmail.com";

// Order rules (change these numbers if the shop needs different limits)
const MAX_QTY_PER_ORDER = 20;   // largest quantity allowed in one order
const MAX_NAME_LENGTH = 60;
const MAX_ADDRESS_LENGTH = 300;
const MAX_NOTE_LENGTH = 300;

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

  // JSONP support. If the request includes a "callback" param,
  // wrap the response as JavaScript instead of plain JSON.
  // Only simple names are accepted, so nobody can inject code through the callback.
  const callback = e && e.parameter && e.parameter.callback;
  if (callback && /^[A-Za-z_$][\w$.]{0,60}$/.test(callback)) {
    return ContentService
      .createTextOutput(callback + "(" + JSON.stringify(responseObj) + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService.createTextOutput(JSON.stringify(responseObj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- Small helper to send a JSON reply ----------
function jsonReply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- Check the order data BEFORE touching any sheet ----------
// Returns { ok: true, clean: {...} } or { ok: false, message: "..." }
function validateOrder(data) {
  if (!data || typeof data !== "object") {
    return { ok: false, message: "Invalid order. Please try again." };
  }

  // Hidden "honeypot" field: real customers never fill it, spam bots do.
  // (Only works once a hidden input named "website" is added to the form.)
  if (data.website) {
    return { ok: false, message: "Order could not be placed." };
  }

  const name = (data.name || "").toString().trim();
  const product = (data.product || "").toString().trim();
  const address = (data.address || "").toString().trim();
  const note = (data.note || "").toString().trim();

  // Phone: keep only digits, drop a leading 91 / 0, must be a 10-digit Indian mobile number
  let phone = (data.phone || "").toString().replace(/\D/g, "");
  if (phone.length === 12 && phone.startsWith("91")) phone = phone.slice(2);
  if (phone.length === 11 && phone.startsWith("0")) phone = phone.slice(1);

  const qty = Number(data.qty);

  if (name.length < 2) return { ok: false, message: "Please enter your name." };
  if (name.length > MAX_NAME_LENGTH) return { ok: false, message: "Name is too long." };
  if (!/^[6-9]\d{9}$/.test(phone)) return { ok: false, message: "Please enter a valid 10-digit mobile number." };
  if (!product) return { ok: false, message: "Please choose a product." };
  if (!Number.isInteger(qty) || qty < 1 || qty > MAX_QTY_PER_ORDER) {
    return { ok: false, message: `Quantity must be between 1 and ${MAX_QTY_PER_ORDER}.` };
  }
  if (address.length > MAX_ADDRESS_LENGTH) return { ok: false, message: "Address is too long." };
  if (note.length > MAX_NOTE_LENGTH) return { ok: false, message: "Note is too long." };

  return {
    ok: true,
    clean: {
      name: safeText(name),
      phone: phone,
      product: product,
      qty: qty,
      address: safeText(address),
      note: safeText(note)
    }
  };
}

// Stops text starting with = + - @ from being treated as a formula in Google Sheets
function safeText(text) {
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

// ---------- Main Webhook ----------
function doPost(e) {
  // LockService: only one order is processed at a time,
  // so two orders arriving together can't get the same Order ID or overwrite stock.
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000); // wait up to 20 seconds for the previous order to finish
  } catch (lockErr) {
    return jsonReply({ status: "error", message: "We're busy right now. Please try again in a moment." });
  }

  try {
    let data;
    try {
      data = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      return jsonReply({ status: "failed", message: "Invalid order data. Please try again." });
    }

    // 1. Validate the customer's details
    const check = validateOrder(data);
    if (!check.ok) {
      return jsonReply({ status: "failed", message: check.message });
    }
    const order = check.clean;

    // 2. Check inventory and take the REAL price from the Inventory sheet
    //    (the price sent by the browser is ignored, so nobody can change it)
    const stockResult = updateInventory(order.product, order.qty);

    if (!stockResult.success) {
      sendOwnerWarning(order, stockResult);
      return jsonReply({
        status: "failed",
        message: stockResult.matched
          ? `Sorry, only ${stockResult.available} unit(s) of "${order.product}" left in stock.`
          : `"${order.product}" could not be matched in our inventory. Our team has been notified.`
      });
    }

    const orderDate = new Date();
    const orderId = generateOrderID();
    const unitPrice = stockResult.unitPrice;
    const totalAmount = unitPrice * order.qty;

    // 3. Find or create the customer – SAME id used in both sheets
    const customerId = getOrCreateCustomer(order.name, order.phone, order.address);

    // 4. Save to Order Sheet (price, total and category come from the Inventory sheet)
    ORDER_SHEET.appendRow([
      orderDate,
      orderId,
      customerId,
      order.name,
      stockResult.productName,
      order.phone,
      stockResult.category,
      order.qty,
      unitPrice,
      totalAmount,
      order.address,
      "Pending",
      order.note
    ]);

    // 5. Confirmation back to the customer (shown as the "Order placed!" alert on the site)
    return jsonReply({
      status: "success",
      message: "Order placed successfully!",
      orderId: orderId,
      total: totalAmount
    });

  } catch (err) {
    console.error(err);
    return jsonReply({ status: "error", message: "Something went wrong. Please try again or order on WhatsApp." });
  } finally {
    lock.releaseLock();
  }
}

// ---------- Inventory check + update ----------
// Returns { success, available, matched, unitPrice, category, productName }
// matched=false means the product name wasn't found in the sheet at all (a data problem, not a stock problem)
function updateInventory(productName, qty) {
  const inventory = INVENTORY_SHEET.getDataRange().getValues();
  const headers = inventory[0].map(h => h.toString().toLowerCase().trim());

  const nameCol = headers.indexOf("product_name") + 1 || 2;
  const stockCol = headers.indexOf("stock_quantity") + 1 || 9;
  const priceCol = headers.indexOf("unit_price") + 1 || 5;
  const catCol = headers.indexOf("category") + 1 || 3;
  const soldCol = headers.indexOf("units_sold") + 1 || 0;
  const totalCol = headers.indexOf("total_amount") + 1 || 0;

  const target = (productName || "").toString().trim().toLowerCase();
  let rowIndex = -1;

  // Exact match only. The website sends the exact Product_Name it received from doGet,
  // so a partial match is never needed and could pick the wrong size/variant.
  for (let i = 1; i < inventory.length; i++) {
    const rowProduct = (inventory[i][nameCol - 1] || "").toString().trim().toLowerCase();
    if (rowProduct === target) { rowIndex = i; break; }
  }

  if (rowIndex === -1) {
    return { success: false, available: 0, matched: false };
  }

  const row = inventory[rowIndex];
  const currentStock = Number(row[stockCol - 1]) || 0;
  const unitPrice = Number(row[priceCol - 1]) || 0;

  if (currentStock < qty) {
    return { success: false, available: currentStock, matched: true };
  }
  if (unitPrice <= 0) {
    // No valid price in the sheet – treat as a data problem so the owner is told
    return { success: false, available: currentStock, matched: false };
  }

  const newStock = currentStock - qty;
  INVENTORY_SHEET.getRange(rowIndex + 1, stockCol).setValue(newStock);

  if (soldCol) {
    const currentSold = Number(row[soldCol - 1]) || 0;
    INVENTORY_SHEET.getRange(rowIndex + 1, soldCol).setValue(currentSold + qty);
  }
  if (totalCol) {
    const currentTotal = Number(row[totalCol - 1]) || 0;
    INVENTORY_SHEET.getRange(rowIndex + 1, totalCol).setValue(currentTotal + qty * unitPrice);
  }

  return {
    success: true,
    available: newStock,
    matched: true,
    unitPrice: unitPrice,
    category: (row[catCol - 1] || "").toString().trim(),
    productName: (row[nameCol - 1] || "").toString().trim()
  };
}

// ---------- Owner warning email ----------
function sendOwnerWarning(orderData, stockResult) {
  const notFound = !stockResult.matched;
  const subject = notFound
    ? `⚠️ PRODUCT NOT FOUND / NO PRICE IN INVENTORY - ${orderData.product}`
    : `⚠️ LOW STOCK ALERT - ${orderData.product}`;

  const body = notFound
    ? `A customer tried to order "${orderData.product}" but it wasn't found in the Inventory sheet's Product_Name column, or its Unit_price is empty.\n\n` +
      `Customer: ${orderData.name} (${orderData.phone})\nQty requested: ${orderData.qty}\n\n` +
      `Fix: make sure the Product_Name matches the website exactly and Unit_price is filled in.`
    : `Product: ${orderData.product}\nRequested: ${orderData.qty}\nAvailable: ${stockResult.available}\n` +
      `Customer: ${orderData.name} (${orderData.phone})\n\nThis order was NOT saved – please restock and ask the customer to reorder.`;

  try {
    MailApp.sendEmail(OWNER_EMAIL, subject, body);
  } catch (mailErr) {
    console.error("Email failed: " + mailErr); // don't break the order reply if email quota is used up
  }
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
    // Compare last 10 digits, so old rows saved as "+91 98765..." still match
    const rowPhone = (customers[i][phoneCol] || "").toString().replace(/\D/g, "").slice(-10);
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