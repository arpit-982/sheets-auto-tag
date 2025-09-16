/**
 * @OnlyCurrentDoc
 */

/**
 * Main function to set up the spreadsheet for the auto-tagging workflow.
 * It creates a metadata section, data headers, supporting sheets, and applies filters.
 */
function initializePlugin() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = ss.getActiveSheet();
    mainSheet.setName("Processing"); // Standardize the main sheet name

    // Set up the main sheet layout (metadata + headers)
    setupMainSheetLayout(mainSheet);
    
    // Create and configure the 'Rules' and 'Accounts' sheets
    setupSupportingSheets(ss);
    
    // Add the 'Funding Account' dropdown using data from the 'Accounts' sheet
    createFundingAccountDropdown(mainSheet, ss);
    
    // Automatically apply filters to the data headers
    applyAutomaticFilters(mainSheet);

    SpreadsheetApp.getUi().alert('✅ Initialization Complete!\n\nThe sheet is now configured with metadata controls and data filters.');
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + error.message);
  }
}

/**
 * Sets up the top rows for metadata and the data headers on the main sheet.
 */
function setupMainSheetLayout(sheet) {
  sheet.clear(); // Start with a clean slate
  sheet.setFrozenRows(4); // Freeze the metadata and header rows for easy scrolling

  // --- Setup Metadata Section (Rows 1-3) ---
  const metadataSection = sheet.getRange("A1:B3");
  metadataSection.setFontWeight("bold");
  sheet.getRange("A1").setValue("Funding Account:");
  sheet.getRange("A2").setValue("Processing Status:");
  sheet.getRange("B2").setValue("Ready").setFontWeight("normal");
  sheet.getRange("A3").setValue("Last Run:").setValue("N/A").setFontWeight("normal");

  // --- Setup Data Headers (Row 4) ---
const headers = [
  "Sr No", "Transaction Date", "Narration", "Withdrawal", "Deposit", "Balance",
  "User Context", "Tags", "LLM Confidence", "Final Entry"
];
  sheet.getRange(4, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
}

/**
 * Creates and configures the 'Rules' and 'Accounts' sheets.
 */
function setupSupportingSheets(spreadsheet) {
  // --- Setup 'Accounts' Sheet ---
  const accountsSheetName = 'Accounts';
  let accountsSheet = spreadsheet.getSheetByName(accountsSheetName);
  if (!accountsSheet) {
    accountsSheet = spreadsheet.insertSheet(accountsSheetName);
    accountsSheet.clear();
    accountsSheet.getRange("A1:C1").setValues([['Account', 'Type', 'Usage Count']]).setFontWeight("bold");
    populateAccountsSheet(accountsSheet); // Only populate if new sheet
  } else {
    // Sheet exists, just ensure headers are correct
    const existingHeaders = accountsSheet.getRange("A1:C1").getValues()[0];
    if (existingHeaders[0] !== 'Account' || existingHeaders[1] !== 'Type' || existingHeaders[2] !== 'Usage Count') {
      accountsSheet.getRange("A1:C1").setValues([['Account', 'Type', 'Usage Count']]).setFontWeight("bold");
    }
  }

  // --- Setup 'Rules' Sheet with enhanced columns ---
  const rulesSheetName = 'Rules';
  let rulesSheet = spreadsheet.getSheetByName(rulesSheetName);
  if (!rulesSheet) {
    rulesSheet = spreadsheet.insertSheet(rulesSheetName);
    rulesSheet.clear();
    const rulesHeaders = ["ID", "Priority", "Active", "Condition", "Pattern / Value", "Action Type", "Action Value"];
    rulesSheet.getRange(1, 1, 1, rulesHeaders.length).setValues([rulesHeaders]).setFontWeight("bold");
    rulesSheet.getRange("C2:C").insertCheckboxes(); // Add checkboxes to the 'Active' column
  } else {
    // Sheet exists, just ensure headers are correct and checkboxes are in Active column
    const rulesHeaders = ["ID", "Priority", "Active", "Condition", "Pattern / Value", "Action Type", "Action Value"];
    rulesSheet.getRange(1, 1, 1, rulesHeaders.length).setValues([rulesHeaders]).setFontWeight("bold");
    // Only add checkboxes if Active column doesn't already have them
    try {
      rulesSheet.getRange("C2:C").insertCheckboxes();
    } catch (e) {
      // Checkboxes already exist, ignore error
    }
  }
}

/**
 * Applies filters automatically to the data header row.
 */
function applyAutomaticFilters(sheet) {
  const headerRow = 4;
  const range = sheet.getRange(headerRow, 1, sheet.getMaxRows() - headerRow, sheet.getLastColumn());
  
  // Always remove existing filters before creating a new one to avoid errors.
  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }
  range.createFilter();
}


/**
 * Populates the 'Accounts' sheet with the full chart of accounts.
 * This is your existing function, kept for its comprehensive account list.
 */
function populateAccountsSheet(sheet) {
  // This is the comprehensive list from your original 'Validation.js' file.
const accounts = [
    ['Expenses:Entertainment:Dining Out', 'Expense', 0],
    ['Expenses:Entertainment:Movies & Shows', 'Expense', 0],
    ['Expenses:Entertainment:Other Entertainment', 'Expense', 0],
    ['Expenses:Entertainment:Parties', 'Expense', 0],
    ['Expenses:Household:Alcohol', 'Expense', 0],
    ['Expenses:Household:Food', 'Expense', 0],
    ['Expenses:Household:Groceries', 'Expense', 0],
    ['Expenses:Household:Health and Wellness', 'Expense', 0],
    ['Expenses:Household:Help', 'Expense', 0],
    ['Expenses:Household:Medicines', 'Expense', 0],
    ['Expenses:Household:Other Household', 'Expense', 0],
    ['Expenses:Household:Biduiee', 'Expense', 0],
    ['Expenses:Rent:House', 'Expense', 0],
    ['Expenses:Rent:Internet and Phone', 'Expense', 0],
    ['Expenses:Rent:Washing Machine', 'Expense', 0],
    ['Expenses:Shopping:Clothes and Apparels', 'Expense', 0],
    ['Expenses:Shopping:Electronics and Accessories', 'Expense', 0],
    ['Expenses:Shopping:Gifts', 'Expense', 0],
    ['Expenses:Shopping:Other Shopping', 'Expense', 0],
    ['Expenses:Shopping:Subscriptions and Digital Purchases', 'Expense', 0],
    ['Expenses:Transport:Fuel', 'Expense', 0],
    ['Expenses:Transport:Taxis', 'Expense', 0],
    ['Expenses:Travel', 'Expense', 0],
    ['Expenses:Travel:General', 'Expense', 0],
    ['Expenses:Travel:Stay', 'Expense', 0],
    ['Expenses:Travel:Misc. Expenses', 'Expense', 0],
    ['Expenses:Travel:Trains and Flights', 'Expense', 0],
    ['Expenses:Travel:Buses and Cabs', 'Expense', 0],
    ['Expenses:Utilities:Electricity', 'Expense', 0],
    ['Expenses:Utilities:Gas', 'Expense', 0],
    ['Expenses:Utilities:Other Utilities', 'Expense', 0],
    ['Expenses:Utilities:Water', 'Expense', 0],
    ['Expenses:Others:To Family', 'Expense', 0],
    ['Expenses:Others:Other Charges', 'Expense', 0],
    ['Expenses:Others:Insurance Premium', 'Expense', 0],
    ['Expenses:Others:Taxes', 'Expense', 0],
    ['Income:Employer:Salary', 'Income', 0],
    ['Income:Employer:Bonus', 'Income', 0],
    ['Income:Others', 'Income', 0],
    ['Income:Reimbursements', 'Income', 0],
    ['Income:Refund:Credit Card:SBI', 'Income', 0],
    ['Assets:Checking', 'Asset', 0],
    ['Assets:Checking:Bank of Baroda', 'Asset', 0],
    ['Assets:Checking:Punjab National Bank', 'Asset', 0],
    ['Assets:Checking:Wallet', 'Asset', 0],
    ['Assets:Mutual Funds:Canara Robeco ELSS', 'Asset', 0],
    ['Assets:Mutual Funds:Liquid Fund', 'Asset', 0],
    ['Assets:Mutual Funds:Mirae ELSS', 'Asset', 0],
    ['Assets:Mutual Funds:Quant ELSS', 'Asset', 0],
    ['Assets:Mutual Funds:DSP ELSS', 'Asset', 0],
    ['Assets:Mutual Funds:Quant Active', 'Asset', 0],
    ['Assets:Other Investments:Employee Provident Fund', 'Asset', 0],
    ['Assets:Other Investments:NPS', 'Asset', 0],
    ['Assets:Other Investments:Public Provident Fund', 'Asset', 0],
    ['Assets:Receivables:Misc Receivables', 'Asset', 0],
    ['Assets:Receivables:Advance Payments', 'Asset', 0],
    ['Assets:Shares:AU Bank', 'Asset', 0],
    ['Assets:Shares:LIC', 'Asset', 0],
    ['Assets:Other Assets:Secutity Deposit', 'Asset', 0],
    ['Liabilities:Payables:Akshu', 'Liability', 0],
    ['Liabilities:Payables:Ananya', 'Liability', 0],
    ['Liabilities:Payables:Anna', 'Liability', 0],
    ['Liabilities:Payables:Misc Payables', 'Liability', 0],
    ['Liabilities:Payables:Mummy', 'Liability', 0],
    ['Liabilities:Other Liabilities:Akshara Investment', 'Liability', 0],
    ['Liabilities:Other Liabilities:Akshu Investment', 'Liability', 0],
    ['Liabilities:Credit Card:SBI', 'Liability', 0],
    ['Equity', 'Equity', 0]
  ];
  sheet.getRange(2, 1, accounts.length, 3).setValues(accounts);
}