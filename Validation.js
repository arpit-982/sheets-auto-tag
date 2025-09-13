function initializePlugin() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    
    // Step 1: Validate required columns
    validateRequiredColumns(sheet);
    
    // Step 2: Add dynamic columns if missing
    addDynamicColumns(sheet);
    
    // Step 3: Create supporting sheets
    createSupportingSheets();
    
    SpreadsheetApp.getUi().alert('✅ Plugin initialized successfully!\n\nCheck out the new "Rules" and "Accounts" sheets.');
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + error.message);
  }
}

function validateRequiredColumns(sheet) {
  const requiredColumns = ['Sr No', 'Transaction Date', 'Withdrawal', 'Deposit', 'Balance', 'Narration', 'User Context'];
  const headers = sheet.getRange(1, 1, 1, 7).getValues()[0];
  
  for (let i = 0; i < requiredColumns.length; i++) {
    if (headers[i] !== requiredColumns[i]) {
      throw new Error(`Column ${i + 1} should be "${requiredColumns[i]}" but found "${headers[i]}"`);
    }
  }
  
  console.log('✅ All required columns validated');
}

function addDynamicColumns(sheet) {
  const dynamicColumns = ['Tags', 'LLM Confidence', 'Final Entry'];
  const lastCol = sheet.getLastColumn();
  
  // Check if dynamic columns already exist
  if (lastCol >= 10) {
    const existingHeaders = sheet.getRange(1, 8, 1, 3).getValues()[0];
    if (existingHeaders[0] === 'Tags') {
      console.log('✅ Dynamic columns already exist');
      return;
    }
  }
  
  // Add dynamic column headers
  for (let i = 0; i < dynamicColumns.length; i++) {
    sheet.getRange(1, 8 + i).setValue(dynamicColumns[i]);
  }
  
  // Make headers bold
  sheet.getRange(1, 8, 1, 3).setFontWeight('bold');
  
  console.log('✅ Dynamic columns added');
}

function createSupportingSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create Rules sheet if it doesn't exist
  if (!spreadsheet.getSheetByName('Rules')) {
    const rulesSheet = spreadsheet.insertSheet('Rules');
    const rulesHeaders = ['ID', 'Priority', 'Active', 'Condition Type', 'Pattern', 'Target Account', 'Tags', 'Created Date'];
    rulesSheet.getRange(1, 1, 1, rulesHeaders.length).setValues([rulesHeaders]);
    rulesSheet.getRange(1, 1, 1, rulesHeaders.length).setFontWeight('bold');
    console.log('✅ Rules sheet created');
  }
  
  // Create Accounts sheet if it doesn't exist
  if (!spreadsheet.getSheetByName('Accounts')) {
    const accountsSheet = spreadsheet.insertSheet('Accounts');
    const accountsHeaders = ['Account', 'Type', 'Usage Count'];
    accountsSheet.getRange(1, 1, 1, accountsHeaders.length).setValues([accountsHeaders]);
    accountsSheet.getRange(1, 1, 1, accountsHeaders.length).setFontWeight('bold');
    
    // Pre-populate with your account hierarchy
    populateAccountsSheet(accountsSheet);
    console.log('✅ Accounts sheet created and populated');
  }
}

function populateAccountsSheet(sheet) {
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
  
  // Add all accounts starting from row 2
  sheet.getRange(2, 1, accounts.length, 3).setValues(accounts);
}