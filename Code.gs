// Code.gs - Imprint Prints POS & ERP System Backend

const SHEETS = [
  { name: 'Users', headers: ['user_id', 'name', 'email', 'password', 'role', 'status', 'created_at'] },
  { name: 'Customers', headers: ['customer_id', 'name', 'phone', 'email', 'address', 'institution', 'customer_type', 'reward_points', 'total_orders', 'total_spent', 'created_at'] },
  { name: 'Suppliers', headers: ['supplier_id', 'company_name', 'contact_person', 'phone', 'email', 'address', 'previous_due', 'balance', 'status'] },
  { name: 'PaperStock', headers: ['id', 'item_code', 'barcode', 'name', 'category', 'subcategory', 'brand', 'gsm', 'size', 'unit', 'supplier', 'purchase_price', 'selling_price', 'current_stock', 'minimum_stock', 'reorder_level', 'rack_location', 'image_drive_id', 'status', 'created_at', 'updated_at'] },
  { name: 'Services', headers: ['service_id', 'service_name', 'price', 'cost', 'unit', 'description', 'status'] },
  { name: 'Purchases', headers: ['purchase_id', 'supplier_id', 'date', 'total_amount', 'discount', 'net_amount', 'paid', 'due', 'status'] },
  { name: 'PurchaseItems', headers: ['id', 'purchase_id', 'item_id', 'quantity', 'unit_price', 'total'] },
  { name: 'Sales', headers: ['sale_id', 'customer_id', 'date', 'total_amount', 'discount', 'net_amount', 'paid', 'due', 'payment_method', 'status', 'operator'] },
  { name: 'SaleItems', headers: ['id', 'sale_id', 'item_id', 'type', 'quantity', 'unit_price', 'total', 'details'] },
  { name: 'Orders', headers: ['order_id', 'customer_id', 'operator', 'status', 'timeline', 'notes', 'est_delivery', 'delivery_method', 'created_at'] },
  { name: 'OrderItems', headers: ['id', 'order_id', 'item_id', 'type', 'quantity', 'details', 'status'] },
  { name: 'Payments', headers: ['payment_id', 'ref_id', 'type', 'amount', 'method', 'date'] },
  { name: 'Expenses', headers: ['expense_id', 'category', 'amount', 'date', 'description', 'recorded_by'] },
  { name: 'Settings', headers: ['key', 'value'] },
  { name: 'ActivityLogs', headers: ['log_id', 'user_id', 'action', 'details', 'date'] },
  { name: 'ImportLogs', headers: ['id', 'module', 'status', 'date'] },
  { name: 'Categories', headers: ['id', 'name', 'type'] },
  { name: 'SubCategories', headers: ['id', 'category_id', 'name'] },
  { name: 'InventoryMovements', headers: ['id', 'item_id', 'type', 'quantity', 'date', 'ref_id'] }
];

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Imprint Prints ERP')
    .setFaviconUrl('https://img.icons8.com/color/48/000000/print.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ----- SYSTEM INIT -----

function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  SHEETS.forEach(s => {
    let sheet = ss.getSheetByName(s.name);
    if (!sheet) {
      sheet = ss.insertSheet(s.name);
      sheet.appendRow(s.headers);
      sheet.getRange(1, 1, 1, s.headers.length).setFontWeight('bold');
    }
  });
  return { success: true, message: 'Sheets initialized.' };
}

function setupDemoData() {
  initSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Clear existing
  SHEETS.forEach(s => {
    let sheet = ss.getSheetByName(s.name);
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, s.headers.length).clearContent();
    }
  });

  // Users
  insertRow('Users', ['U1001', 'Admin User', 'admin@imprint.com', 'admin123', 'Admin', 'Active', new Date().toISOString()]);
  insertRow('Users', ['U1002', 'Cashier One', 'cashier@imprint.com', 'cashier123', 'Cashier', 'Active', new Date().toISOString()]);
  
  // Settings
  const settings = [
    ['business_name', 'Imprint Prints'],
    ['phone', '+8801700000000'],
    ['email', 'contact@imprintprints.com'],
    ['address', 'Dhaka, Bangladesh'],
    ['currency', 'BDT'],
    ['invoice_prefix', 'INV-'],
    ['vat', '5']
  ];
  settings.forEach(s => insertRow('Settings', s));

  // Customers
  for(let i=1; i<=50; i++) {
    insertRow('Customers', [`C${2000+i}`, `Customer ${i}`, `01711000${i.toString().padStart(2,'0')}`, `cust${i}@test.com`, `Address ${i}`, `University ${i%5}`, 'Retail', 0, 0, 0, new Date().toISOString()]);
  }

  // Suppliers
  for(let i=1; i<=20; i++) {
    insertRow('Suppliers', [`S${3000+i}`, `Supplier Company ${i}`, `Contact ${i}`, `01811000${i.toString().padStart(2,'0')}`, `sup${i}@test.com`, `Supplier Addr ${i}`, 0, 0, 'Active']);
  }

  // Services
  const demoServices = ['Black Print', 'Color Print', 'Scanning', 'Photocopy', 'Spiral Binding', 'Hard Binding', 'Lamination', 'Poster Print', 'Banner Print', 'PVC ID Card'];
  demoServices.forEach((srv, idx) => {
    insertRow('Services', [`SRV${4000+idx}`, srv, (idx+1)*5, (idx+1)*2, 'pcs', srv + ' Description', 'Active']);
  });

  // PaperStock
  for(let i=1; i<=100; i++) {
    insertRow('PaperStock', [`P${5000+i}`, `ITM${i}`, `890100${i}`, `Paper Item ${i}`, 'Paper', 'A4', 'BrandX', '70', 'A4', 'ream', `S3001`, 200, 250, 100, 20, 30, `Rack-${i%5}`, '', 'Active', new Date().toISOString(), new Date().toISOString()]);
  }

  // Expenses
  const expCategories = ['Rent', 'Salary', 'Internet', 'Electricity', 'Maintenance'];
  for(let i=1; i<=20; i++) {
    insertRow('Expenses', [`EXP${i}`, expCategories[i%5], i*100, new Date().toISOString(), `Expense Note ${i}`, 'U1001']);
  }

  return { success: true, message: 'Demo data injected successfully.' };
}

// ----- DATABASE ABSTRACTION -----

function getSheetData(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  const rows = data.slice(1);
  return rows.map(row => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function insertRow(sheetName, values) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return null;
  sheet.appendRow(values);
  return values;
}

function updateRow(sheetName, idField, idValue, updateObj) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return false;
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIndex = headers.indexOf(idField);
  if (idIndex === -1) return false;

  for (let i = 1; i < data.length; i++) {
    if (data[i][idIndex] == idValue) {
      headers.forEach((h, colIdx) => {
        if (updateObj[h] !== undefined) {
          sheet.getRange(i + 1, colIdx + 1).setValue(updateObj[h]);
        }
      });
      return true;
    }
  }
  return false;
}

function deleteRow(sheetName, idField, idValue) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return false;
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIndex = headers.indexOf(idField);
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIndex] == idValue) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

// ----- API ENDPOINTS -----

function processRequest(action, payload) {
  try {
    switch (action) {
      case 'init': return initSheets();
      case 'setup_demo': return setupDemoData();
      case 'login': return doLogin(payload);
      
      // Generic CRUD Get
      case 'get_data': return { success: true, data: getSheetData(payload.sheet) };
      
      // Generic CRUD Insert
      case 'insert_data':
        const headers = SHEETS.find(s => s.name === payload.sheet).headers;
        const row = headers.map(h => payload.data[h] || '');
        if (payload.sheet === 'Sales') row[0] = 'INV' + new Date().getTime(); // auto ID
        insertRow(payload.sheet, row);
        return { success: true, message: 'Record added successfully' };
        
      // Generic CRUD Update
      case 'update_data':
        updateRow(payload.sheet, payload.idField, payload.idValue, payload.data);
        return { success: true, message: 'Record updated successfully' };
        
      // Generic CRUD Delete
      case 'delete_data':
        deleteRow(payload.sheet, payload.idField, payload.idValue);
        return { success: true, message: 'Record deleted successfully' };
        
      case 'save_sale': return processSale(payload);
      case 'get_dashboard': return getDashboardStats();
      default: return { success: false, message: 'Unknown action' };
    }
  } catch (e) {
    return { success: false, message: e.message, stack: e.stack };
  }
}

function doLogin(cred) {
  const users = getSheetData('Users');
  const user = users.find(u => (u.email === cred.email || u.name === cred.email) && u.password === cred.password && u.status === 'Active');
  if (user) {
    delete user.password; // secure
    return { success: true, user: user };
  }
  return { success: false, message: 'Invalid credentials or inactive account' };
}

function processSale(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const saleId = 'INV-' + new Date().getTime();
    const date = new Date().toISOString();
    
    // Save Sale
    insertRow('Sales', [
      saleId, payload.customer_id, date, payload.total, payload.discount, 
      payload.net_amount, payload.paid, payload.due, payload.payment_method, 'Completed', payload.operator
    ]);
    
    // Save Items & Update Stock
    payload.items.forEach(item => {
      insertRow('SaleItems', [
        'SI' + new Date().getTime() + Math.floor(Math.random()*1000),
        saleId, item.id, item.type, item.qty, item.price, item.qty * item.price, JSON.stringify(item.details || {})
      ]);
      
      if (item.type === 'product') {
        const stock = getSheetData('PaperStock').find(p => p.id == item.id);
        if (stock) {
          const newStock = parseFloat(stock.current_stock) - parseFloat(item.qty);
          updateRow('PaperStock', 'id', item.id, { current_stock: newStock, updated_at: date });
        }
      }
    });
    
    // Update Customer
    if (payload.customer_id && payload.customer_id !== 'Walk-in') {
      const cust = getSheetData('Customers').find(c => c.customer_id == payload.customer_id);
      if (cust) {
        updateRow('Customers', 'customer_id', payload.customer_id, {
          total_orders: (parseInt(cust.total_orders)||0) + 1,
          total_spent: (parseFloat(cust.total_spent)||0) + parseFloat(payload.net_amount)
        });
      }
    }
    
    return { success: true, saleId: saleId };
  } finally {
    lock.releaseLock();
  }
}

function getDashboardStats() {
  const sales = getSheetData('Sales');
  const expenses = getSheetData('Expenses');
  const orders = getSheetData('Orders');
  const stock = getSheetData('PaperStock');
  
  const today = new Date().toISOString().split('T')[0];
  const thisMonth = today.substring(0, 7);
  
  let todaySales = 0;
  let monthlySales = 0;
  
  sales.forEach(s => {
    if (s.date.startsWith(today)) todaySales += parseFloat(s.net_amount||0);
    if (s.date.startsWith(thisMonth)) monthlySales += parseFloat(s.net_amount||0);
  });
  
  let lowStockCount = stock.filter(s => parseFloat(s.current_stock) <= parseFloat(s.minimum_stock)).length;
  
  return {
    success: true,
    data: {
      todaySales: todaySales,
      monthlySales: monthlySales,
      pendingOrders: orders.filter(o => o.status !== 'Delivered' && o.status !== 'Cancelled').length,
      lowStock: lowStockCount,
      recentSales: sales.slice(-10).reverse()
    }
  };
}
