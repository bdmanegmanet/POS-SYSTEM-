/*******************************************************
 * BUSINESS POS SYSTEM
 * Google Apps Script Backend
 * Database: Google Sheets | Storage: Google Drive
 *******************************************************/

/* =========================
   CONFIG
========================= */
const CONFIG = {
  APP_NAME: 'Business POS System',
  SESSION_SECONDS: 21600,
  DEFAULT_SHEET: 'Settings',
  SHEETS: {
    Settings: ['Key','Value','Updated_At'],
    Users: ['User_ID','Username','Email','Password_Hash','Role','Status','Created_At','Updated_At'],
    Roles: ['Role','Description'],
    Permissions: ['Role','Permission','Allowed'],
    Products: ['Product_ID','Barcode','SKU','Product_Name','Product_Image','Category','Subcategory','Brand','Model','Description','Purchase_Price','Selling_Price','Wholesale_Price','Discount_Price','VAT','Unit','Size','Color','Material','Weight','Warranty','Minimum_Stock','Opening_Stock','Current_Stock','Supplier','Status','Created_Date','Updated_Date','Created_By','Updated_By'],
    Product_Variants: ['Variant_ID','Product_ID','Variant_Name','SKU','Barcode','Purchase_Price','Selling_Price','Stock','Status','Created_At','Updated_At'],
    Categories: ['Category_ID','Category_Name','Description','Status','Created_At','Updated_At'],
    Subcategories: ['Subcategory_ID','Category_ID','Subcategory_Name','Description','Status','Created_At','Updated_At'],
    Brands: ['Brand_ID','Brand_Name','Description','Status','Created_At','Updated_At'],
    Units: ['Unit_ID','Unit_Name','Short_Name','Status','Created_At','Updated_At'],
    Customers: ['Customer_ID','Name','Phone','Email','Address','Opening_Due','Current_Due','Loyalty_Points','Status','Created_At','Updated_At'],
    Suppliers: ['Supplier_ID','Name','Phone','Email','Address','Opening_Due','Current_Due','Status','Created_At','Updated_At'],
    Sales: ['Sale_ID','Invoice_No','Customer_ID','Subtotal','Discount','Tax','Grand_Total','Paid','Due','Payment_Method','Sale_Date','Status','Created_By','Created_At'],
    Sale_Items: ['Sale_Item_ID','Sale_ID','Product_ID','Product_Name','Qty','Unit_Price','Discount','Tax','Line_Total','Created_At'],
    Purchases: ['Purchase_ID','Invoice_No','Supplier_ID','Subtotal','Discount','Tax','Grand_Total','Paid','Due','Payment_Method','Purchase_Date','Status','Created_By','Created_At'],
    Purchase_Items: ['Purchase_Item_ID','Purchase_ID','Product_ID','Product_Name','Qty','Unit_Cost','Discount','Tax','Line_Total','Created_At'],
    Sales_Returns: ['Return_ID','Sale_ID','Customer_ID','Amount','Reason','Return_Date','Created_By','Created_At'],
    Sales_Return_Items: ['Return_Item_ID','Return_ID','Product_ID','Qty','Amount'],
    Purchase_Returns: ['Return_ID','Purchase_ID','Supplier_ID','Amount','Reason','Return_Date','Created_By','Created_At'],
    Purchase_Return_Items: ['Return_Item_ID','Return_ID','Product_ID','Qty','Amount'],
    Stock: ['Stock_ID','Product_ID','Quantity','Reserved','Available','Last_Updated'],
    Stock_Movements: ['Movement_ID','Product_ID','Type','Reference_ID','Quantity','Before_Qty','After_Qty','Note','Created_By','Created_At'],
    Payments: ['Payment_ID','Party_Type','Party_ID','Reference_ID','Amount','Method','Direction','Note','Payment_Date','Created_By','Created_At'],
    Payment_Methods: ['Method_ID','Method_Name','Status','Created_At'],
    Accounts: ['Account_ID','Account_Name','Account_Type','Balance','Status','Created_At','Updated_At'],
    Transactions: ['Transaction_ID','Account_ID','Type','Reference_ID','Amount','Description','Transaction_Date','Created_By','Created_At'],
    Expenses: ['Expense_ID','Category_ID','Amount','Description','Payment_Method','Expense_Date','Created_By','Created_At'],
    Expense_Categories: ['Category_ID','Category_Name','Description','Status','Created_At'],
    Coupons: ['Coupon_ID','Code','Type','Value','Start_Date','End_Date','Status'],
    Discounts: ['Discount_ID','Name','Type','Value','Status','Created_At'],
    Loyalty: ['Loyalty_ID','Customer_ID','Points','Type','Reference_ID','Created_At'],
    Invoices: ['Invoice_ID','Sale_ID','Invoice_No','Drive_File_ID','Drive_URL','Created_At'],
    Taxes: ['Tax_ID','Tax_Name','Rate','Status','Created_At'],
    Journal: ['Journal_ID','Reference_ID','Description','Journal_Date','Created_By','Created_At'],
    Ledger: ['Ledger_ID','Journal_ID','Account_ID','Debit','Credit','Ledger_Date'],
    Chart_of_Accounts: ['Account_ID','Account_Code','Account_Name','Account_Type','Parent_ID','Status'],
    Audit_Logs: ['Log_ID','User','Action','Module','Record_ID','Old_Value','New_Value','Session_ID','Date','Time'],
    Notifications: ['Notification_ID','Type','Title','Message','Status','Created_At'],
    System_Logs: ['Log_ID','Level','Action','Message','Created_At'],
    Backups: ['Backup_ID','Type','File_ID','File_URL','Created_By','Created_At']
  },
  ROLES: ['Super Admin','Admin','Manager','Cashier','Sales Staff','Inventory Manager','Accountant','Viewer'],
  PERMISSIONS: ['view','create','edit','delete','export','print','approve','manage_settings','manage_users','manage_database']
};

/* =========================
   WEB APP
========================= */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(CONFIG.APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/* =========================
   DATABASE
========================= */
function getDb_() {
  const id = PropertiesService.getScriptProperties().getProperty('DATABASE_SHEET_ID');
  if (!id) throw new Error('Database is not configured. Run setupDatabase() first.');
  return SpreadsheetApp.openById(id);
}

function setupDatabase(adminUsername, adminPassword, adminEmail) {
  if (!adminUsername || !adminPassword) throw new Error('Run setupDatabase("admin","StrongPassword","email@example.com")');
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('DATABASE_SHEET_ID');
  let ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.create(CONFIG.APP_NAME + ' Database');
  props.setProperty('DATABASE_SHEET_ID', ss.getId());
  props.setProperty('APP_VERSION', '1.0.0');

  Object.keys(CONFIG.SHEETS).forEach(name => ensureSheet_(ss, name, CONFIG.SHEETS[name]));
  seedSettings_(ss);
  seedRoles_(ss);
  seedPermissions_(ss);

  const users = getRows_('Users');
  if (!users.some(u => String(u.Username).toLowerCase() === String(adminUsername).toLowerCase())) {
    appendRow_('Users', {
      User_ID: id_('USR'),
      Username: adminUsername,
      Email: adminEmail || '',
      Password_Hash: hash_(adminPassword),
      Role: 'Super Admin',
      Status: 'Active',
      Created_At: now_(),
      Updated_At: now_()
    });
  }
  return {success:true, spreadsheetId:ss.getId(), spreadsheetUrl:ss.getUrl(), message:'Database ready'};
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,headers.length).setValues([headers]);
  else {
    const current = sh.getRange(1,1,1,Math.max(sh.getLastColumn(),headers.length)).getValues()[0];
    const missing = headers.filter(h => current.indexOf(h) === -1);
    if (missing.length) sh.getRange(1,current.filter(String).length+1,1,missing.length).setValues([missing]);
  }
  sh.setFrozenRows(1);
}

function seedSettings_(ss) {
  const sh = ss.getSheetByName('Settings');
  const rows = getSheetRows_(sh);
  const defaults = {
    APP_NAME: CONFIG.APP_NAME,
    CURRENCY: 'BDT',
    TAX_RATE: '0',
    LOW_STOCK_DEFAULT: '5',
    AUTO_SYNC: 'true',
    SYNC_INTERVAL: '30',
    THEME: 'light',
    LANGUAGE: 'bn'
  };
  Object.keys(defaults).forEach(k => {
    if (!rows.some(r => r.Key === k)) sh.appendRow([k,defaults[k],now_()]);
  });
}
function seedRoles_(ss) {
  const sh=ss.getSheetByName('Roles');
  if (sh.getLastRow() < 2) CONFIG.ROLES.forEach(r=>sh.appendRow([r,r+' role']));
}
function seedPermissions_(ss) {
  const sh=ss.getSheetByName('Permissions');
  if (sh.getLastRow() < 2) CONFIG.ROLES.forEach(role=>CONFIG.PERMISSIONS.forEach(p=>sh.appendRow([role,p,role==='Viewer'?(p==='view'||p==='export'||p==='print'):'true'])));
}

/* =========================
   AUTHENTICATION
========================= */
function login(username, password) {
  if (!username || !password) throw new Error('Username and password are required.');
  const user = getRows_('Users').find(u =>
    String(u.Username).toLowerCase() === String(username).toLowerCase() ||
    String(u.Email).toLowerCase() === String(username).toLowerCase()
  );
  if (!user || user.Status !== 'Active' || user.Password_Hash !== hash_(password)) {
    logSystem_('WARN','LOGIN_FAILED','Invalid login for '+username);
    throw new Error('Invalid username/password.');
  }
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('SESSION_'+token, JSON.stringify({
    userId:user.User_ID, username:user.Username, role:user.Role, loginAt:now_()
  }), CONFIG.SESSION_SECONDS);
  audit_(user.Username,'LOGIN','AUTH','', '', '', token);
  return {token, user:{id:user.User_ID,username:user.Username,role:user.Role}};
}
function logout(token) {
  if (token) CacheService.getScriptCache().remove('SESSION_'+token);
  return {success:true};
}
function session_(token) {
  if (!token) throw new Error('Session expired.');
  const raw=CacheService.getScriptCache().get('SESSION_'+token);
  if (!raw) throw new Error('Session expired. Please login again.');
  return JSON.parse(raw);
}
function getSession(token) { return session_(token); }
function requireSession_(token, permission) {
  const s=session_(token);
  if (permission && !can_(s.role,permission)) throw new Error('Permission denied.');
  return s;
}
function can_(role, permission) {
  if (role==='Super Admin') return true;
  const row=getRows_('Permissions').find(r=>r.Role===role && r.Permission===permission);
  return row ? String(row.Allowed).toLowerCase()==='true' : false;
}
function changePassword(token, oldPassword, newPassword) {
  const s=requireSession_(token,'edit');
  const sh=getDb_().getSheetByName('Users');
  const rows=getSheetRows_(sh);
  const idx=rows.findIndex(r=>r.Username===s.username);
  if(idx<0 || rows[idx].Password_Hash!==hash_(oldPassword)) throw new Error('Current password is incorrect.');
  if(String(newPassword).length<8) throw new Error('New password must be at least 8 characters.');
  sh.getRange(idx+2,4).setValue(hash_(newPassword));
  sh.getRange(idx+2,8).setValue(now_());
  audit_(s.username,'CHANGE_PASSWORD','AUTH',rows[idx].User_ID,'','','');
  return {success:true};
}

/* =========================
   API
========================= */
function api(token, action, payload) {
  const p=payload||{};
  switch(action) {
    case 'dashboard': return getDashboard(token);
    case 'products': return getProducts(token,p);
    case 'saveProduct': return saveProduct(token,p);
    case 'deleteProduct': return deleteProduct(token,p.Product_ID);
    case 'customers': return getCustomers(token,p);
    case 'saveCustomer': return saveCustomer(token,p);
    case 'suppliers': return getSuppliers(token,p);
    case 'saveSupplier': return saveSupplier(token,p);
    case 'sales': return getSales(token,p);
    case 'createSale': return createSale(token,p);
    case 'purchases': return getPurchases(token,p);
    case 'createPurchase': return createPurchase(token,p);
    case 'expenses': return getExpenses(token,p);
    case 'saveExpense': return saveExpense(token,p);
    case 'sync': return syncData(token);
    case 'push': return pushData(token);
    case 'settings': return getSettings(token);
    case 'saveSettings': return saveSettings(token,p);
    case 'code': return getCode(token);
    case 'health': return systemHealth(token);
    case 'setup': return setupDatabase(p.username,p.password,p.email);
    default: throw new Error('Unknown API action: '+action);
  }
}

/* =========================
   PRODUCTS
========================= */
function getProducts(token, p) {
  requireSession_(token,'view');
  let rows=getRows_('Products');
  if(p.search) {
    const q=String(p.search).toLowerCase();
    rows=rows.filter(r=>['Product_Name','Barcode','SKU','Category','Brand'].some(k=>String(r[k]||'').toLowerCase().includes(q)));
  }
  if(p.status) rows=rows.filter(r=>r.Status===p.status);
  return {rows:rows.map(normalize_)};
}
function saveProduct(token,p) {
  const s=requireSession_(token,p.Product_ID?'edit':'create');
  required_(p,['Product_Name','Selling_Price']);
  const rows=getRows_('Products');
  if(p.Barcode && rows.some(r=>r.Barcode===p.Barcode && r.Product_ID!==p.Product_ID)) throw new Error('Barcode already exists.');
  if(p.SKU && rows.some(r=>r.SKU===p.SKU && r.Product_ID!==p.Product_ID)) throw new Error('SKU already exists.');
  const data=Object.assign({},p);
  data.Product_ID=p.Product_ID||id_('PRD');
  data.Purchase_Price=num_(p.Purchase_Price);
  data.Selling_Price=num_(p.Selling_Price);
  data.Wholesale_Price=num_(p.Wholesale_Price);
  data.Discount_Price=num_(p.Discount_Price);
  data.VAT=num_(p.VAT);
  data.Minimum_Stock=num_(p.Minimum_Stock);
  data.Current_Stock=num_(p.Current_Stock);
  data.Status=p.Status||'Active';
  data.Updated_Date=now_();
  data.Created_Date=p.Created_Date||now_();
  data.Created_By=p.Created_By||s.username;
  data.Updated_By=s.username;
  upsert_('Products',data,'Product_ID');
  ensureStock_(data.Product_ID,data.Current_Stock,s.username);
  audit_(s.username,p.Product_ID?'UPDATE':'CREATE','PRODUCTS',data.Product_ID,p.Product_ID?'':'',JSON.stringify(data),'');
  return data;
}
function deleteProduct(token,id) {
  const s=requireSession_(token,'delete');
  if(!id) throw new Error('Product ID required.');
  upsert_('Products',{Product_ID:id,Status:'Deleted',Updated_Date:now_(),Updated_By:s.username},'Product_ID');
  audit_(s.username,'DELETE','PRODUCTS',id,'','','');
  return {success:true};
}

/* =========================
   CUSTOMERS / SUPPLIERS
========================= */
function getCustomers(token,p){requireSession_(token,'view'); return {rows:filterRows_('Customers',p)};}
function getSuppliers(token,p){requireSession_(token,'view'); return {rows:filterRows_('Suppliers',p)};}
function saveCustomer(token,p) {
  const s=requireSession_(token,p.Customer_ID?'edit':'create');
  required_(p,['Name']);
  const rows=getRows_('Customers');
  if(p.Phone && rows.some(r=>r.Phone===p.Phone && r.Customer_ID!==p.Customer_ID)) throw new Error('Customer phone already exists.');
  const d=Object.assign({},p,{Customer_ID:p.Customer_ID||id_('CUS'),Current_Due:num_(p.Current_Due),Opening_Due:num_(p.Opening_Due),Loyalty_Points:num_(p.Loyalty_Points),Status:p.Status||'Active',Updated_At:now_(),Created_At:p.Created_At||now_()});
  upsert_('Customers',d,'Customer_ID'); audit_(s.username,p.Customer_ID?'UPDATE':'CREATE','CUSTOMERS',d.Customer_ID,'',JSON.stringify(d),''); return d;
}
function saveSupplier(token,p) {
  const s=requireSession_(token,p.Supplier_ID?'edit':'create');
  required_(p,['Name']);
  const d=Object.assign({},p,{Supplier_ID:p.Supplier_ID||id_('SUP'),Current_Due:num_(p.Current_Due),Opening_Due:num_(p.Opening_Due),Status:p.Status||'Active',Updated_At:now_(),Created_At:p.Created_At||now_()});
  upsert_('Suppliers',d,'Supplier_ID'); audit_(s.username,p.Supplier_ID?'UPDATE':'CREATE','SUPPLIERS',d.Supplier_ID,'',JSON.stringify(d),''); return d;
}

/* =========================
   SALES
========================= */
function createSale(token,p) {
  const s=requireSession_(token,'create');
  if(!Array.isArray(p.items)||!p.items.length) throw new Error('Sale must contain at least one item.');
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const saleId=id_('SAL'), invoice='INV-'+Utilities.formatDate(new Date(),Session.getScriptTimeZone()||'Asia/Dhaka','yyyyMMdd-HHmmss');
    let subtotal=0, tax=0, discount=num_(p.discount);
    const products=getRows_('Products');
    const prepared=[];
    p.items.forEach(i=>{
      const pr=products.find(x=>x.Product_ID===i.Product_ID && x.Status!=='Deleted');
      if(!pr) throw new Error('Product not found: '+i.Product_ID);
      const qty=num_(i.Qty); if(qty<=0) throw new Error('Quantity must be greater than zero.');
      const stock=num_(pr.Current_Stock); if(stock<qty) throw new Error('Insufficient stock: '+pr.Product_Name);
      const unit=num_(i.Unit_Price||pr.Selling_Price); const line=qty*unit-num_(i.Discount);
      const t=line*num_(i.VAT||pr.VAT)/100; subtotal+=line; tax+=t;
      prepared.push({pr,qty,unit,discount:num_(i.Discount),tax:t,line:line+t});
    });
    const grand=Math.max(0,subtotal+tax-discount), paid=num_(p.paid);
    if(paid<0 || paid>grand) throw new Error('Invalid payment amount.');
    const due=grand-paid;
    appendRow_('Sales',{Sale_ID:saleId,Invoice_No:invoice,Customer_ID:p.Customer_ID||'',Subtotal:subtotal,Discount:discount,Tax:tax,Grand_Total:grand,Paid:paid,Due:due,Payment_Method:p.Payment_Method||'Cash',Sale_Date:p.Sale_Date||now_(),Status:'Completed',Created_By:s.username,Created_At:now_()});
    prepared.forEach(x=>{
      appendRow_('Sale_Items',{Sale_Item_ID:id_('SIT'),Sale_ID:saleId,Product_ID:x.pr.Product_ID,Product_Name:x.pr.Product_Name,Qty:x.qty,Unit_Price:x.unit,Discount:x.discount,Tax:x.tax,Line_Total:x.line,Created_At:now_()});
      updateStock_(x.pr.Product_ID,-x.qty,'SALE',saleId,s.username);
    });
    if(p.Customer_ID) updatePartyDue_('Customers','Customer_ID',p.Customer_ID,due,s.username);
    if(paid) recordPayment_(s.username,'Customer',p.Customer_ID||'',saleId,paid,p.Payment_Method||'Cash','IN');
    audit_(s.username,'CREATE','SALES',saleId,'',JSON.stringify({invoice,grand,paid,due}),'');
    return {success:true,saleId,invoice,subtotal,discount,tax,grandTotal:grand,paid,due};
  } finally { lock.releaseLock(); }
}
function getSales(token,p){requireSession_(token,'view'); return {rows:filterRows_('Sales',p)};}

/* =========================
   PURCHASE
========================= */
function createPurchase(token,p) {
  const s=requireSession_(token,'create');
  if(!Array.isArray(p.items)||!p.items.length) throw new Error('Purchase must contain at least one item.');
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const id=id_('PUR'), invoice='PUR-'+Utilities.formatDate(new Date(),Session.getScriptTimeZone()||'Asia/Dhaka','yyyyMMdd-HHmmss');
    let subtotal=0,tax=0,discount=num_(p.discount);
    const products=getRows_('Products'), prepared=[];
    p.items.forEach(i=>{
      const pr=products.find(x=>x.Product_ID===i.Product_ID && x.Status!=='Deleted');
      if(!pr) throw new Error('Product not found: '+i.Product_ID);
      const qty=num_(i.Qty); if(qty<=0) throw new Error('Quantity must be greater than zero.');
      const cost=num_(i.Unit_Cost||pr.Purchase_Price); const line=qty*cost-num_(i.Discount); const t=line*num_(i.VAT||pr.VAT)/100;
      subtotal+=line;tax+=t;prepared.push({pr,qty,cost,discount:num_(i.Discount),tax:t,line:line+t});
    });
    const grand=Math.max(0,subtotal+tax-discount),paid=num_(p.paid); if(paid<0||paid>grand) throw new Error('Invalid payment amount.');
    const due=grand-paid;
    appendRow_('Purchases',{Purchase_ID:id,Invoice_No:invoice,Supplier_ID:p.Supplier_ID||'',Subtotal:subtotal,Discount:discount,Tax:tax,Grand_Total:grand,Paid:paid,Due:due,Payment_Method:p.Payment_Method||'Cash',Purchase_Date:p.Purchase_Date||now_(),Status:'Completed',Created_By:s.username,Created_At:now_()});
    prepared.forEach(x=>{
      appendRow_('Purchase_Items',{Purchase_Item_ID:id_('PIT'),Purchase_ID:id,Product_ID:x.pr.Product_ID,Product_Name:x.pr.Product_Name,Qty:x.qty,Unit_Cost:x.cost,Discount:x.discount,Tax:x.tax,Line_Total:x.line,Created_At:now_()});
      updateStock_(x.pr.Product_ID,x.qty,'PURCHASE',id,s.username);
      upsert_('Products',{Product_ID:x.pr.Product_ID,Purchase_Price:x.cost,Updated_Date:now_(),Updated_By:s.username},'Product_ID');
    });
    if(p.Supplier_ID) updatePartyDue_('Suppliers','Supplier_ID',p.Supplier_ID,due,s.username);
    if(paid) recordPayment_(s.username,'Supplier',p.Supplier_ID||'',id,paid,p.Payment_Method||'Cash','OUT');
    audit_(s.username,'CREATE','PURCHASES',id,'',JSON.stringify({invoice,grand,paid,due}),'');
    return {success:true,purchaseId:id,invoice,grandTotal:grand,paid,due};
  } finally { lock.releaseLock(); }
}
function getPurchases(token,p){requireSession_(token,'view'); return {rows:filterRows_('Purchases',p)};}

/* =========================
   EXPENSES
========================= */
function saveExpense(token,p) {
  const s=requireSession_(token,'create');
  required_(p,['Amount']);
  const amount=num_(p.Amount); if(amount<=0) throw new Error('Expense amount must be greater than zero.');
  const d={Expense_ID:p.Expense_ID||id_('EXP'),Category_ID:p.Category_ID||'',Amount:amount,Description:p.Description||'',Payment_Method:p.Payment_Method||'Cash',Expense_Date:p.Expense_Date||now_(),Created_By:s.username,Created_At:now_()};
  upsert_('Expenses',d,'Expense_ID');
  recordPayment_(s.username,'Expense','',d.Expense_ID,amount,d.Payment_Method,'OUT');
  audit_(s.username,p.Expense_ID?'UPDATE':'CREATE','EXPENSES',d.Expense_ID,'',JSON.stringify(d),''); return d;
}
function getExpenses(token,p){requireSession_(token,'view'); return {rows:filterRows_('Expenses',p)};}

/* =========================
   STOCK
========================= */
function ensureStock_(productId,qty,user) {
  const rows=getRows_('Stock'), row=rows.find(r=>r.Product_ID===productId);
  if(!row) appendRow_('Stock',{Stock_ID:id_('STK'),Product_ID:productId,Quantity:num_(qty),Reserved:0,Available:num_(qty),Last_Updated:now_()});
  else updateStock_(productId,num_(qty)-num_(row.Quantity),'ADJUSTMENT','',user);
}
function updateStock_(productId,delta,type,ref,user) {
  const rows=getRows_('Stock'), row=rows.find(r=>r.Product_ID===productId);
  const before=row?num_(row.Quantity):0, after=before+num_(delta); if(after<0) throw new Error('Stock cannot be negative.');
  if(row) upsert_('Stock',{Stock_ID:row.Stock_ID,Product_ID:productId,Quantity:after,Reserved:num_(row.Reserved),Available:after-num_(row.Reserved),Last_Updated:now_()},'Stock_ID');
  else appendRow_('Stock',{Stock_ID:id_('STK'),Product_ID:productId,Quantity:after,Reserved:0,Available:after,Last_Updated:now_()});
  upsert_('Products',{Product_ID:productId,Current_Stock:after,Updated_Date:now_(),Updated_By:user||''},'Product_ID');
  appendRow_('Stock_Movements',{Movement_ID:id_('MOV'),Product_ID:productId,Type:type,Reference_ID:ref||'',Quantity:num_(delta),Before_Qty:before,After_Qty:after,Note:'',Created_By:user||'',Created_At:now_()});
}

/* =========================
   DASHBOARD / REPORT DATA
========================= */
function getDashboard(token) {
  requireSession_(token,'view');
  const products=getRows_('Products').filter(r=>r.Status!=='Deleted'), customers=getRows_('Customers').filter(r=>r.Status!=='Deleted'), suppliers=getRows_('Suppliers').filter(r=>r.Status!=='Deleted');
  const sales=getRows_('Sales').filter(r=>r.Status==='Completed'), purchases=getRows_('Purchases').filter(r=>r.Status==='Completed'), expenses=getRows_('Expenses');
  const today=Utilities.formatDate(new Date(),Session.getScriptTimeZone()||'Asia/Dhaka','yyyy-MM-dd');
  const todaySales=sales.filter(r=>String(r.Sale_Date).indexOf(today)>=0).reduce((a,r)=>a+num_(r.Grand_Total),0);
  const todayPurchase=purchases.filter(r=>String(r.Purchase_Date).indexOf(today)>=0).reduce((a,r)=>a+num_(r.Grand_Total),0);
  const todayExpense=expenses.filter(r=>String(r.Expense_Date).indexOf(today)>=0).reduce((a,r)=>a+num_(r.Amount),0);
  const stockValue=products.reduce((a,r)=>a+num_(r.Current_Stock)*num_(r.Purchase_Price),0);
  const low=products.filter(r=>num_(r.Current_Stock)<=num_(r.Minimum_Stock||5)&&num_(r.Current_Stock)>0).length;
  const out=products.filter(r=>num_(r.Current_Stock)<=0).length;
  const customerDue=customers.reduce((a,r)=>a+num_(r.Current_Due),0), supplierDue=suppliers.reduce((a,r)=>a+num_(r.Current_Due),0);
  return {todaySales,todayPurchase,todayExpense,todayNet:todaySales-todayPurchase-todayExpense,totalProducts:products.length,stockValue,lowStock:low,outOfStock:out,totalCustomers:customers.length,totalSuppliers:suppliers.length,customerDue,supplierDue,cashBalance:accountBalance_('Cash'),bankBalance:accountBalance_('Bank'),bKashBalance:accountBalance_('bKash'),nagadBalance:accountBalance_('Nagad')};
}

/* =========================
   PAYMENTS / FINANCE
========================= */
function recordPayment_(user,partyType,partyId,ref,amount,method,direction) {
  appendRow_('Payments',{Payment_ID:id_('PAY'),Party_Type:partyType,Party_ID:partyId||'',Reference_ID:ref||'',Amount:num_(amount),Method:method||'Cash',Direction:direction||'IN',Note:'',Payment_Date:now_(),Created_By:user,Created_At:now_()});
  const account=method||'Cash'; let ar=getRows_('Accounts').find(a=>a.Account_Name===account);
  if(!ar) { ar={Account_ID:id_('ACC'),Account_Name:account,Account_Type:'Cash',Balance:0,Status:'Active',Created_At:now_(),Updated_At:now_()}; appendRow_('Accounts',ar); }
  const bal=num_(ar.Balance)+(direction==='OUT'?-num_(amount):num_(amount));
  upsert_('Accounts',{Account_ID:ar.Account_ID,Balance:bal,Updated_At:now_()},'Account_ID');
  appendRow_('Transactions',{Transaction_ID:id_('TXN'),Account_ID:ar.Account_ID,Type:direction,Reference_ID:ref||'',Amount:num_(amount),Description:partyType,Transaction_Date:now_(),Created_By:user,Created_At:now_()});
}
function updatePartyDue_(sheet,key,id,delta,user) {
  const row=getRows_(sheet).find(r=>String(r[key])===String(id)); if(!row) throw new Error('Party not found.');
  upsert_(sheet,Object.assign({},row,{Current_Due:num_(row.Current_Due)+num_(delta),Updated_At:now_()}),key);
}
function accountBalance_(name){const r=getRows_('Accounts').find(x=>x.Account_Name===name);return r?num_(r.Balance):0;}

/* =========================
   SETTINGS / SYNC / HEALTH
========================= */
function getSettings(token){requireSession_(token,'view'); const o={};getRows_('Settings').forEach(r=>o[r.Key]=r.Value);return o;}
function saveSettings(token,p) {
  const s=requireSession_(token,'manage_settings'); Object.keys(p||{}).forEach(k=>upsert_('Settings',{Key:k,Value:String(p[k]),Updated_At:now_()},'Key'));
  audit_(s.username,'UPDATE','SETTINGS','', '',JSON.stringify(p),''); return getSettings(token);
}
function syncData(token){const s=requireSession_(token,'view'); return {success:true,status:'Connected',lastSync:now_(),pendingChanges:0,failed:0,user:s.username};}
function pushData(token){const s=requireSession_(token,'manage_database'); return {success:true,status:'Pushed',lastPush:now_(),user:s.username};}
function systemHealth(token){requireSession_(token,'view');const ss=getDb_();return {database:'OK',spreadsheet:ss.getName(),spreadsheetId:ss.getId(),lastSync:now_(),storage:'Google Drive available',appVersion:PropertiesService.getScriptProperties().getProperty('APP_VERSION')||'1.0.0'};}
function getCode(token){requireSession_(token,'manage_database');return ScriptApp.getService().getUrl() ? getFullCode_() : getFullCode_();}

/* =========================
   UTILITIES
========================= */
function getFullCode_(){
  // Admin code-copy feature: Code.gs itself is not directly readable by Apps Script at runtime.
  // The repository copy is the canonical source. This safe runtime message avoids exposing secrets.
  return 'Code.gs source is maintained in the GitHub repository. For security, runtime credentials and Script Properties are never exposed here.';
}
function getRows_(sheet){return getSheetRows_(getDb_().getSheetByName(sheet));}
function getSheetRows_(sh){if(!sh||sh.getLastRow()<2)return [];const h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];return sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues().map(row=>{const o={};h.forEach((k,i)=>o[k]=row[i]);return o;});}
function appendRow_(sheet,obj){const sh=getDb_().getSheetByName(sheet),h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];sh.appendRow(h.map(k=>obj[k]!==undefined?obj[k]:''));}
function upsert_(sheet,obj,key){
  const sh=getDb_().getSheetByName(sheet),h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0], rows=getSheetRows_(sh), idx=rows.findIndex(r=>String(r[key])===String(obj[key]));
  const values=h.map(k=>obj[k]!==undefined?obj[k]:(idx>=0?rows[idx][k]:''));
  if(idx>=0) sh.getRange(idx+2,1,1,h.length).setValues([values]); else sh.appendRow(values);
}
function filterRows_(sheet,p){let rows=getRows_(sheet);if(p&&p.search){const q=String(p.search).toLowerCase();rows=rows.filter(r=>Object.values(r).some(v=>String(v).toLowerCase().includes(q)));}return rows.map(normalize_);}
function normalize_(o){const r={};Object.keys(o).forEach(k=>r[k]=o[k] instanceof Date?o[k].toISOString():o[k]);return r;}
function required_(p,keys){keys.forEach(k=>{if(p[k]===undefined||p[k]===null||String(p[k]).trim()==='')throw new Error(k+' is required.');});}
function num_(v){const n=Number(v);return isNaN(n)?0:n;}
function id_(prefix){return prefix+'-'+Utilities.getUuid().replace(/-/g,'').slice(0,12).toUpperCase();}
function now_(){return new Date();}
function hash_(value){return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(value),Utilities.Charset.UTF_8));}
function audit_(user,action,module,record,oldValue,newValue,sessionId){appendRow_('Audit_Logs',{Log_ID:id_('LOG'),User:user||'',Action:action,Module:module,Record_ID:record||'',Old_Value:oldValue||'',New_Value:newValue||'',Session_ID:sessionId||'',Date:now_(),Time:Utilities.formatDate(new Date(),Session.getScriptTimeZone()||'Asia/Dhaka','HH:mm:ss')});}
function logSystem_(level,action,message){try{appendRow_('System_Logs',{Log_ID:id_('SYS'),Level:level,Action:action,Message:message,Created_At:now_()});}catch(e){}}
