/**
 * AIC Inventory Management System
 * Google Apps Script Backend - Template Handler
 * File: gstemplate.gs
 */

/**
 * Main function to serve the web app
 * This is the entry point for the Google Apps Script web app
 */
function doGet(e) {
  const page = e.parameter.page || 'dashboard';
  
  try {
    // Create template from base template file
    const template = HtmlService.createTemplateFromFile('template');
    
    // Map pages to their template files
    const pageTemplates = {
      'dashboard': 'index',
      'inventory': 'inventory',
      'suppliers': 'suppliers',
      'customers': 'customers',
      'purchases': 'purchases',
      'sales': 'sales',
      'receipts': 'receipts',
      'payments': 'payments',
      'reports': 'reports'
    };
    
    // Get the correct template name
    const contentTemplate = pageTemplates[page] || 'index';
    template.contentTemplate = contentTemplate;
    
    // Return the evaluated template
    return template.evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
      
  } catch (error) {
    Logger.log('Error in doGet: ' + error.toString());
    return HtmlService.createHtmlOutput('<h1>Error Loading App</h1><p>' + error.toString() + '</p>');
  }
}

/**
 * Helper function to include HTML files
 * This function reads the content of an HTML file and returns it as string
 * @param {string} filename - The name of the file to include
 * @return {string} The content of the HTML file
 */
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (error) {
    Logger.log('Error including file ' + filename + ': ' + error.toString());
    return '<div class="error">Error loading ' + filename + '</div>';
  }
}

/**
 * Helper function to get the script URL
 * This function returns the URL of the web app
 * @return {string} The URL of the Google Apps Script web app
 */
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * Get the active Google Sheet
 * @return {Sheet} The active sheet
 */
function getActiveSheet() {
  return SpreadsheetApp.getActiveSheet();
}

/**
 * Get a specific sheet by name
 * @param {string} sheetName - The name of the sheet
 * @return {Sheet} The requested sheet
 */
function getSheetByName(sheetName) {
  try {
    return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  } catch (error) {
    Logger.log('Sheet not found: ' + sheetName);
    return null;
  }
}

/**
 * Create a new sheet if it doesn't exist
 * @param {string} sheetName - The name of the new sheet
 * @return {Sheet} The created or existing sheet
 */
function createSheetIfNotExists(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  return sheet;
}

/**
 * Get all data from a sheet
 * @param {string} sheetName - The name of the sheet
 * @return {Array<Array>} All data in the sheet
 */
function getSheetData(sheetName) {
  const sheet = getSheetByName(sheetName);
  
  if (!sheet) {
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  return data;
}

/**
 * Get sheet data as objects with headers
 * @param {string} sheetName - The name of the sheet
 * @return {Array<Object>} Array of objects with sheet data
 */
function getSheetDataAsObjects(sheetName) {
  const data = getSheetData(sheetName);
  
  if (data.length === 0) {
    return [];
  }
  
  const headers = data[0];
  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j] || '';
    }
    result.push(obj);
  }
  
  return result;
}

/**
 * Add a new row to a sheet
 * @param {string} sheetName - The name of the sheet
 * @param {Array} rowData - The data to add
 * @return {boolean} True if successful
 */
function addRowToSheet(sheetName, rowData) {
  try {
    const sheet = getSheetByName(sheetName);
    
    if (!sheet) {
      return false;
    }
    
    sheet.appendRow(rowData);
    return true;
  } catch (error) {
    Logger.log('Error adding row to sheet: ' + error.toString());
    return false;
  }
}

/**
 * Update a row in a sheet
 * @param {string} sheetName - The name of the sheet
 * @param {number} rowIndex - The row index (1-based)
 * @param {Array} rowData - The new data
 * @return {boolean} True if successful
 */
function updateRowInSheet(sheetName, rowIndex, rowData) {
  try {
    const sheet = getSheetByName(sheetName);
    
    if (!sheet || rowIndex < 1) {
      return false;
    }
    
    const range = sheet.getRange(rowIndex, 1, 1, rowData.length);
    range.setValues([rowData]);
    return true;
  } catch (error) {
    Logger.log('Error updating row in sheet: ' + error.toString());
    return false;
  }
}

/**
 * Delete a row from a sheet
 * @param {string} sheetName - The name of the sheet
 * @param {number} rowIndex - The row index (1-based)
 * @return {boolean} True if successful
 */
function deleteRowFromSheet(sheetName, rowIndex) {
  try {
    const sheet = getSheetByName(sheetName);
    
    if (!sheet || rowIndex < 2) {
      return false; // Don't delete header row
    }
    
    sheet.deleteRow(rowIndex);
    return true;
  } catch (error) {
    Logger.log('Error deleting row from sheet: ' + error.toString());
    return false;
  }
}

/**
 * Clear all data from a sheet (keeps headers)
 * @param {string} sheetName - The name of the sheet
 * @return {boolean} True if successful
 */
function clearSheetData(sheetName) {
  try {
    const sheet = getSheetByName(sheetName);
    
    if (!sheet) {
      return false;
    }
    
    const range = sheet.getDataRange();
    const lastRow = range.getLastRow();
    
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
    
    return true;
  } catch (error) {
    Logger.log('Error clearing sheet: ' + error.toString());
    return false;
  }
}

/**
 * Initialize required sheets for the inventory system
 */
function initializeSheets() {
  const sheetNames = [
    'Inventory',
    'Suppliers',
    'Customers',
    'Purchase Orders',
    'Sales Orders',
    'Receipts',
    'Payments',
    'Settings'
  ];
  
  // Inventory sheet headers
  const inventoryHeaders = [
    'Item ID',
    'Item Name',
    'Category',
    'Quantity',
    'Unit Price',
    'Reorder Level',
    'Supplier ID',
    'Location',
    'Last Updated',
    'Status'
  ];
  
  // Suppliers sheet headers
  const suppliersHeaders = [
    'Supplier ID',
    'Supplier Name',
    'Contact Person',
    'Email',
    'Phone',
    'Address',
    'City',
    'Country',
    'Payment Terms',
    'Status'
  ];
  
  // Customers sheet headers
  const customersHeaders = [
    'Customer ID',
    'Customer Name',
    'Email',
    'Phone',
    'Address',
    'City',
    'Country',
    'Credit Limit',
    'Account Status',
    'Registration Date'
  ];
  
  // Purchase Orders sheet headers
  const purchaseOrdersHeaders = [
    'PO ID',
    'Supplier ID',
    'Order Date',
    'Delivery Date',
    'Total Amount',
    'Payment Status',
    'Order Status',
    'Notes',
    'Created Date',
    'Last Modified'
  ];
  
  // Sales Orders sheet headers
  const salesOrdersHeaders = [
    'SO ID',
    'Customer ID',
    'Order Date',
    'Delivery Date',
    'Total Amount',
    'Payment Status',
    'Order Status',
    'Notes',
    'Created Date',
    'Last Modified'
  ];
  
  // Receipts sheet headers
  const receiptsHeaders = [
    'Receipt ID',
    'PO ID',
    'Receipt Date',
    'Quantity Received',
    'Condition',
    'Notes',
    'Received By',
    'Created Date',
    'Last Modified'
  ];
  
  // Payments sheet headers
  const paymentsHeaders = [
    'Payment ID',
    'Reference ID',
    'Payment Type',
    'Amount',
    'Payment Date',
    'Payment Method',
    'Status',
    'Notes',
    'Created Date',
    'Last Modified'
  ];
  
  try {
    const inventorySheet = createSheetIfNotExists('Inventory');
    if (inventorySheet.getLastRow() === 0) {
      inventorySheet.appendRow(inventoryHeaders);
    }
    
    const suppliersSheet = createSheetIfNotExists('Suppliers');
    if (suppliersSheet.getLastRow() === 0) {
      suppliersSheet.appendRow(suppliersHeaders);
    }
    
    const customersSheet = createSheetIfNotExists('Customers');
    if (customersSheet.getLastRow() === 0) {
      customersSheet.appendRow(customersHeaders);
    }
    
    const poSheet = createSheetIfNotExists('Purchase Orders');
    if (poSheet.getLastRow() === 0) {
      poSheet.appendRow(purchaseOrdersHeaders);
    }
    
    const soSheet = createSheetIfNotExists('Sales Orders');
    if (soSheet.getLastRow() === 0) {
      soSheet.appendRow(salesOrdersHeaders);
    }
    
    const receiptsSheet = createSheetIfNotExists('Receipts');
    if (receiptsSheet.getLastRow() === 0) {
      receiptsSheet.appendRow(receiptsHeaders);
    }
    
    const paymentsSheet = createSheetIfNotExists('Payments');
    if (paymentsSheet.getLastRow() === 0) {
      paymentsSheet.appendRow(paymentsHeaders);
    }
    
    Logger.log('Sheets initialized successfully');
    return true;
  } catch (error) {
    Logger.log('Error initializing sheets: ' + error.toString());
    return false;
  }
}

/**
 * Generate a unique ID
 * @param {string} prefix - Prefix for the ID
 * @return {string} Unique ID
 */
function generateUniqueId(prefix = 'ID') {
  const timestamp = Date.now().toString(36).toUpperCase();
  const random = Math.random().toString(36).substring(2, 8).toUpperCase();
  return `${prefix}-${timestamp}-${random}`;
}

/**
 * Validate email address
 * @param {string} email - Email to validate
 * @return {boolean} True if valid email
 */
function isValidEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * Format currency
 * @param {number} value - The value to format
 * @return {string} Formatted currency string
 */
function formatCurrency(value) {
  return new Intl.NumberFormat('en-US', {
    style: 'currency',
    currency: 'USD'
  }).format(value);
}

/**
 * Format date
 * @param {Date} date - The date to format
 * @return {string} Formatted date string
 */
function formatDate(date) {
  if (!(date instanceof Date)) {
    date = new Date(date);
  }
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm:ss');
}

/**
 * Get current timestamp
 * @return {string} Current timestamp
 */
function getCurrentTimestamp() {
  return formatDate(new Date());
}

/**
 * Log action for audit trail
 * @param {string} action - The action performed
 * @param {string} details - Additional details
 * @param {string} user - User who performed the action
 */
function logAction(action, details = '', user = 'System') {
  try {
    const logsSheet = createSheetIfNotExists('Logs');
    const headers = ['Timestamp', 'User', 'Action', 'Details', 'Status'];
    
    if (logsSheet.getLastRow() === 0) {
      logsSheet.appendRow(headers);
    }
    
    logsSheet.appendRow([
      getCurrentTimestamp(),
      user,
      action,
      details,
      'Completed'
    ]);
  } catch (error) {
    Logger.log('Error logging action: ' + error.toString());
  }
}

/**
 * Export sheet data to PDF (placeholder for future implementation)
 * @param {string} sheetName - Name of the sheet to export
 * @param {string} fileName - Name for the PDF file
 * @return {string} File ID or error message
 */
function exportSheetToPdf(sheetName, fileName) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      return 'Sheet not found';
    }
    
    // This would require additional setup for PDF export
    Logger.log('PDF export scheduled for: ' + fileName);
    return 'PDF export queued';
  } catch (error) {
    Logger.log('Error exporting to PDF: ' + error.toString());
    return 'Error: ' + error.toString();
  }
}

/**
 * On open trigger - Run initialization
 */
function onOpen() {
  initializeSheets();
  Logger.log('Inventory App opened - Sheets initialized');
}

// Call initialization when script loads
initializeSheets();
