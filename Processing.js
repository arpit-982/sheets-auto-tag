function processAllRows() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 5) { // Changed from 2 to 5 since data starts on row 5
    SpreadsheetApp.getUi().alert('No data rows found to process.');
    return;
  }
  
  processRowRange(sheet, 5, lastRow); // Changed from 2 to 5
}

function processSelectedRows() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const selection = sheet.getActiveRange();
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();
  
  // Skip metadata and header rows - data starts at row 5
  const actualStartRow = Math.max(startRow, 5); // Changed from 2 to 5
  const endRow = Math.min(actualStartRow + numRows - 1, sheet.getLastRow());
  
  if (actualStartRow > sheet.getLastRow()) {
    SpreadsheetApp.getUi().alert('No data rows selected to process.');
    return;
  }
  
  processRowRange(sheet, actualStartRow, endRow);
}

function processCurrentRow() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const activeRange = sheet.getActiveRange();
  const currentRow = activeRange.getRow(); // Fixed: use activeRange.getRow() instead of activeCell.getRow()
  
  if (currentRow < 5) { // Changed from 2 to 5
    SpreadsheetApp.getUi().alert('Please select a transaction row (not header or metadata)');
    return;
  }
  
  processRowRange(sheet, currentRow, currentRow);
}

function processRowRange(sheet, startRow, endRow) {
  try {
    Logger.log('=== processRowRange called ===');
    Logger.log('Processing rows ' + startRow + ' to ' + endRow);
    
    // Read headers from row 4 instead of row 1
    const headers = sheet.getRange(4, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colIndices = getColumnIndices(headers);
    
    Logger.log('Column indices: ' + JSON.stringify(colIndices));
    
    // Read funding account from cell B1
    const fundingAccount = sheet.getRange('B1').getValue() || '';
    Logger.log('Funding account: ' + fundingAccount);
    
    // FIX: Stop processing if no funding account is selected
    if (!fundingAccount || fundingAccount.trim() === '') {
      SpreadsheetApp.getUi().alert('❌ Please select a Funding Account in cell B1 before processing transactions.');
      return;
    }
    
    let processedCount = 0;
    const totalRows = endRow - startRow + 1;
    
    // Show initial toast
    SpreadsheetApp.getActiveSpreadsheet().toast(`Starting processing... (0/${totalRows})`, 'Ledger Tools', 5);
    
    // Process each row in range
    for (let rowNum = startRow; rowNum <= endRow; rowNum++) {
      const row = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
      
      Logger.log('Processing row ' + rowNum + ': ' + JSON.stringify(row));
      
      // Skip empty rows (no Sr No)
      if (!row[colIndices.srNo]) {
        Logger.log('Skipping row ' + rowNum + ' - no Sr No');
        continue;
      }
      
      // Process the transaction with funding account
      Logger.log('Calling processTransaction for row ' + rowNum);
      const result = processTransaction(row, colIndices, fundingAccount);
      Logger.log('processTransaction returned: ' + JSON.stringify(result));
      
      // Update the row with results
      sheet.getRange(rowNum, colIndices.tags + 1).setValue(result.tags);
      sheet.getRange(rowNum, colIndices.confidence + 1).setValue(result.confidence);
      sheet.getRange(rowNum, colIndices.finalEntry + 1).setValue(result.finalEntry);
      
      processedCount++;
    }
    
    Logger.log('Completed processing ' + processedCount + ' rows');
    
    SpreadsheetApp.getActiveSpreadsheet().toast(`Completed! Processed ${processedCount}/${totalRows} rows`, 'Ledger Tools', 5);
    SpreadsheetApp.getUi().alert(`✅ Processed ${processedCount} transactions in rows ${startRow} to ${endRow}`);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error processing transactions: ' + error.message);
    console.log('Processing error:', error);
  }
}

function getColumnIndices(headers) {
  // Debug: Log the actual headers being read
  Logger.log('Headers array length: ' + headers.length);
  Logger.log('Headers content: ' + JSON.stringify(headers));

  // Check for exact matches first
  for (let i = 0; i < headers.length; i++) {
    Logger.log('Header[' + i + ']: "' + headers[i] + '" (type: ' + typeof headers[i] + ')');
  }

  // Helper function to find column index with flexible matching
  function findColumnIndex(searchTerms) {
    for (let term of searchTerms) {
      const index = headers.indexOf(term);
      if (index !== -1) return index;
    }
    return -1;
  }

  const indices = {
    srNo: findColumnIndex(['Sr No', 'Sr. No', 'SrNo']),
    date: findColumnIndex(['Transaction Date', 'Date', 'Txn Date']),
    withdrawal: findColumnIndex(['Withdrawal', 'Withdrawal Amount', 'Debit', 'Debit Amount']),
    deposit: findColumnIndex(['Deposit', 'Deposit Amount', 'Credit', 'Credit Amount']),
    balance: findColumnIndex(['Balance', 'Available Balance']),
    narration: findColumnIndex(['Narration', 'Description', 'Details']),
    userContext: findColumnIndex(['User Context', 'UserContext', 'Context']),
    tags: findColumnIndex(['Tags', 'Tag']),
    confidence: findColumnIndex(['LLM Confidence', 'Confidence']),
    finalEntry: findColumnIndex(['Final Entry', 'FinalEntry', 'Entry'])
  };

  Logger.log('Final column indices: ' + JSON.stringify(indices));
  return indices;
}

/**
 * Orchestrates the processing for a single transaction row.
 * It tries the rule engine first, then falls back to the LLM.
 * @returns {object} An object with {tags, confidence, finalEntry}.
 */
function processTransaction(row, colIndices, fundingAccount) {
  Logger.log('=== processTransaction called ===');
  
  const narration = String(row[colIndices.narration] || '');
  const userContext = String(row[colIndices.userContext] || '');
  const withdrawalRaw = row[colIndices.withdrawal] || 0;
  const depositRaw = row[colIndices.deposit] || 0;
  const dateValue = row[colIndices.date];

  // FIX: Parse amounts properly, handling commas and strings
  const withdrawal = parseFloat(String(withdrawalRaw).replace(/,/g, '')) || 0;
  const deposit = parseFloat(String(depositRaw).replace(/,/g, '')) || 0;

  Logger.log('Extracted values - Narration: ' + narration + ', Withdrawal: ' + withdrawal + ', Deposit: ' + deposit);

  if (!narration || (withdrawal === 0 && deposit === 0) || !dateValue) {
    Logger.log('Skipping row - invalid data');
    return { tags: "", confidence: "", finalEntry: "" };
  }
  
  const date = new Date(dateValue);
  const amount = deposit || withdrawal;
  const isCredit = deposit > 0;

  Logger.log('Calling applyRules with amount: ' + amount + ', isCredit: ' + isCredit);
  
  // 1. Try to apply deterministic rules first.
  let result = applyRules(narration, amount, date, fundingAccount, isCredit, userContext);
  Logger.log('applyRules returned: ' + (result ? 'SUCCESS' : 'NULL'));

  // 2. If no rule matched, fall back to your existing LLM suggestion logic.
if (!result) {
    const suggestion = createBasicSuggestion(narration, amount, userContext);
    
    // FIX: Use suggestion.payee instead of generic payee extraction
    const payee = suggestion.payee || userContext.trim() || 'Misc Expense';
    
    const finalEntry = formatLedgerCliEntry(date, payee, suggestion.account, amount, fundingAccount, isCredit, suggestion.tags, null, userContext);
    result = {
      finalEntry: finalEntry,
      tags: suggestion.tags,
      confidence: suggestion.confidence
    };
  }

  return result;
}

function createBasicSuggestion(narration, amount, userContext) {
  let settings;
  try {
    settings = getSettings();
  } catch (error) {
    console.log('Settings error:', error.message);
    return createFallbackSuggestion(narration, amount, userContext);
  }
  
  // Try LLM if configured
  if (settings.provider && settings.apiKey) {
    try {
      // Fixed: ensure narration is string before substring
      const narrationStr = String(narration || '');
      console.log('Attempting LLM call for:', narrationStr.substring(0, 50) + '...');
      return createLLMSuggestion(narration, amount, userContext);
    } catch (error) {
      console.log('LLM failed, using fallback:', error.message);
      // Don't show alert every time, just log and fallback
    }
  } else {
    console.log('LLM not configured, using fallback');
  }
  
  return createFallbackSuggestion(narration, amount, userContext);
}

function createLLMSuggestion(narration, amount, userContext) {
  const prompt = `Categorize this transaction and suggest a meaningful, properly formatted description.

Transaction: "${narration}"
Amount: ${amount}
User Context: "${userContext || 'none'}"

Generate a clean, properly capitalized description. Examples:
- "sbi card payment" → "SBI Card Payment"
- "bhidu cat boarding" → "Bhidu Cat Boarding"
- "payment at himachal stay" → "Himachal Stay"

Common accounts:
- Expenses:Household:Food (food, dining)
- Expenses:Transport:Taxis (auto, uber, transport)
- Expenses:Household:Other Household (small misc expenses)
- Expenses:Shopping:Subscriptions and Digital Purchases (apps, subscriptions)
- Liabilities:Payables:Ananya (splits, payments to people)
- Expenses:Others:Other Charges (default)

Respond with only this JSON format:
{"account":"Expenses:Household:Food","payee":"Properly Formatted Name","tags":"food","confidence":0.8}`;


  try {
    const response = callLLM(prompt, 0.3, 500);
    console.log('LLM Response:', response);
    
    if (!response || response.trim() === '') {
      throw new Error('Empty LLM response');
    }
    
    // Try to parse JSON
    const jsonMatch = response.match(/\{[^}]+\}/);
    if (!jsonMatch) {
      throw new Error('No JSON found in response');
    }
    
    const parsed = JSON.parse(jsonMatch[0]);
  return {
    account: parsed.account || 'Expenses:Others:Other Charges',
    payee: parsed.payee || userContext.trim() || 'Misc Expense', // ADD fallback
    tags: parsed.tags || 'ai',
    confidence: Math.min(parsed.confidence || 0.7, 0.95)
  };
    
  } catch (error) {
    console.log('LLM processing failed:', error.message);
    throw error;
  }
}

function createFallbackSuggestion(narration, amount, userContext) {
  // Use User Context as payee if available
  const payee = userContext && userContext.trim() ? userContext.trim() : 'Misc Expense';
  const narrationStr = String(narration || '').toLowerCase();
  const userContextStr = String(userContext || '').toLowerCase();
  
  let account = 'Expenses:Others:Other Charges';
  let tags = 'other';
  let confidence = 0.3;
  
  // Basic pattern matching with improved string handling
  if (userContextStr) {
    if (userContextStr.includes('food')) {
      account = 'Expenses:Household:Food';
      tags = 'food';
      confidence = 0.7;
    } else if (userContextStr.includes('transport') || userContextStr.includes('taxi') || userContextStr.includes('auto')) {
      account = 'Expenses:Transport:Taxis';
      tags = 'transport';
      confidence = 0.7;
    } else if (userContextStr.includes('thadi')) {
      account = 'Expenses:Household:Other Household';
      tags = 'thadi';
      confidence = 0.7;
    }
  }
  
  // Fallback to narration patterns if no user context
  if (confidence === 0.3 && narrationStr) {
    if (narrationStr.includes('upi') && Math.abs(amount) < 100) {
      account = 'Expenses:Household:Other Household';
      tags = 'thadi';
      confidence = 0.6;
    } else if (narrationStr.includes('salary')) {
      account = 'Income:Employer:Salary';
      tags = 'salary';
      confidence = 0.8;
    }
  }
  
  return { account, tags, confidence, payee };
}

function getAccountList() {
  try {
    const accountsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Accounts');
    if (!accountsSheet) {
      console.log('Accounts sheet not found');
      return ['Expenses:Others:Other Charges', 'Expenses:Household:Food', 'Expenses:Transport:Taxis'];
    }
    
    const data = accountsSheet.getRange('A2:A50').getValues();
    const accounts = data.map(row => row[0]).filter(account => account && account.trim());
    console.log('Found accounts:', accounts.length);
    return accounts;
  } catch (error) {
    console.log('Error getting accounts:', error.message);
    return ['Expenses:Others:Other Charges', 'Expenses:Household:Food', 'Expenses:Transport:Taxis'];
  }
}

function getSettings() {
  const properties = PropertiesService.getScriptProperties();
  return {
    provider: properties.getProperty('LLM_PROVIDER'),
    model: properties.getProperty('LLM_MODEL'), 
    apiKey: properties.getProperty('LLM_API_KEY'),
    customBaseUrl: properties.getProperty('LLM_BASE_URL')
  };
}

function formatFinalEntry(date, narration, account, amount, tags, userContext, fundingAccount) {
  const formattedDate = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "yyyy/MM/dd");
  const sourceAccount = fundingAccount || 'Assets:Checking:Bank of Baroda'; // Use funding account from B1
  
  // Use user context as description if provided, otherwise use narration
  let description = userContext && userContext.trim() ? userContext.trim() : String(narration || '');
  
  // Create human-readable format
  if (amount > 0) {
    // Income/Deposit
    return `${formattedDate} ${description} → ${sourceAccount} (${amount}) ← ${account} (-${amount}) #${tags}`;
  } else {
    // Expense/Withdrawal  
    return `${formattedDate} ${description} → ${account} (${Math.abs(amount)}) ← ${sourceAccount} (${amount}) #${tags}`;
  }
}

function testLLMDirectly() {
  try {
    const response = callLLM('Say hello', 0.5, 100);
    console.log('LLM Response:', response);
    SpreadsheetApp.getUi().alert('Success! LLM Response: "' + response + '"');
    return response;
  } catch (error) {
    console.log('Error:', error.message);
    SpreadsheetApp.getUi().alert('LLM Test Failed: ' + error.message);
    return null;
  }
}

function applyRules(narration, amount, date, fundingAccount, isCredit, userContext = '') {
  Logger.log('=== applyRules called ===');
  Logger.log('Narration: ' + narration);
  Logger.log('Amount: ' + amount);
  Logger.log('IsCredit: ' + isCredit);
  
  const rulesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Rules');
  if (!rulesSheet) {
    Logger.log('ERROR: Rules sheet not found');
    return null;
  }

  const lastRow = rulesSheet.getLastRow();
  Logger.log('Rules sheet last row: ' + lastRow);
  
  if (lastRow < 2) {
    Logger.log('ERROR: No data rows in Rules sheet');
    return null;
  }

  const rulesData = rulesSheet.getRange(2, 1, lastRow - 1, 8).getValues();
  Logger.log('Number of rules loaded: ' + rulesData.length);
  Logger.log('First rule data: ' + JSON.stringify(rulesData[0]));
  
  const lowerNarration = narration.toLowerCase();

  for (let ruleIndex = 0; ruleIndex < rulesData.length; ruleIndex++) {
    const rule = rulesData[ruleIndex];
    Logger.log('\n--- Processing Rule ' + (ruleIndex + 1) + ' ---');
    Logger.log('Rule array: ' + JSON.stringify(rule));
    
    const [id, priority, isActive, conditionStr, patternStr, actionType, actionValue] = rule;
  Logger.log('Parsed - ID: ' + id + ', Active: ' + isActive + ', Condition: ' + conditionStr);
  
  if (!isActive) {
    Logger.log('Rule inactive, skipping');
    continue;
  }
  
  if (!conditionStr || !patternStr || !actionType || !actionValue) {
    Logger.log('Rule incomplete, skipping');
    continue;
  }

  // FIX: Convert to string and handle null/undefined values
  const conditionString = String(conditionStr || '');
  const patternString = String(patternStr || '');
  
  const conditions = conditionString.includes(' AND ') 
    ? conditionString.split(' AND ').map(c => c.trim())
    : [conditionString.trim()];
  
  const patterns = patternString.includes(';') 
    ? patternString.split(';').map(p => p.trim())
    : [patternString.trim()];

  Logger.log('Conditions: ' + JSON.stringify(conditions));
  Logger.log('Patterns: ' + JSON.stringify(patterns));

    if (conditions.length !== patterns.length) {
      Logger.log('Condition/pattern count mismatch, skipping');
      continue;
    }

    let allConditionsMet = true;
    for (let i = 0; i < conditions.length; i++) {
      const condition = conditions[i];
      const currentPattern = patterns[i];
      
      Logger.log('Testing condition: ' + condition + ' with pattern: ' + currentPattern);
      
      let conditionMet = false;
      try {
        if (condition.startsWith('Narration REGEX')) {
          conditionMet = new RegExp(currentPattern, 'i').test(narration);
          Logger.log('REGEX test result: ' + conditionMet);
        } else if (condition.startsWith('Narration CONTAINS')) {
          conditionMet = lowerNarration.includes(currentPattern.toLowerCase());
          Logger.log('CONTAINS test result: ' + conditionMet);
        } else if (condition.startsWith('Amount ==')) {
          conditionMet = amount == parseFloat(currentPattern);
          Logger.log('Amount == test result: ' + conditionMet);
        } else if (condition.startsWith('Amount >')) {
          conditionMet = amount > parseFloat(currentPattern);
          Logger.log('Amount > test result: ' + conditionMet);
        } else if (condition.startsWith('Amount <')) {
          conditionMet = amount < parseFloat(currentPattern);
          Logger.log('Amount < test result: ' + conditionMet);
        }
      } catch (e) { 
        Logger.log('Rule Error ID ' + id + ': ' + e.message);
      }
      
      if (!conditionMet) { 
        allConditionsMet = false; 
        Logger.log('Condition failed, rule rejected');
        break; 
      }
    }

    if (allConditionsMet) {
      Logger.log('All conditions met! Processing action...');
      try {
        const cleanActionValue = actionValue.replace(/'/g, '"');
        Logger.log('Clean action value: ' + cleanActionValue);
        const params = JSON.parse(cleanActionValue);
        Logger.log('Parsed params: ' + JSON.stringify(params));
        
        let finalEntry;
        if (actionType === 'CREATE_ENTRY') {
          finalEntry = formatLedgerCliEntry(date, params.payee || narration, params.account, amount, fundingAccount, isCredit, params.tags, params, userContext);
        } else if (actionType === 'CREATE_TRANSFER') {
          finalEntry = formatLedgerCliEntry(date, params.payee || narration, params.to_account, amount, fundingAccount, false, params.tags, params, userContext);
        }

        Logger.log('Generated entry: ' + finalEntry);
        return { 
          finalEntry, 
          tags: params.tags || '', 
          confidence: 1.0 
        };
      } catch (e) { 
        Logger.log('Rule Action Error ID ' + id + ': ' + e.message);
        return null; 
      }
    }
  }
  
  Logger.log('No rules matched');
  return null;
}

/**
 * Formats a ledger entry in the ledger-cli style, supporting both simple and split transactions.
 * @returns {string} The formatted, multi-line ledger entry.
 */
function formatLedgerCliEntry(date, payee, targetAccount, amount, fundingAccount, isCredit, tags, actionData = null, userContext = null) {
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const dd = String(date.getDate()).padStart(2, '0');
  const formattedDate = `${yyyy}/${mm}/${dd}`;
  const totalAmount = Math.abs(amount);

  // Check if this is a split transaction
  if (actionData && actionData.split_type && actionData.split_type !== 'none') {
    return generateSplitLedgerEntry(formattedDate, payee, targetAccount, totalAmount, fundingAccount, isCredit, tags, actionData, userContext);
  }

  // Standard non-split entry
  const formattedAmount = `₹${totalAmount.toFixed(2)}`;

  // Build the entry with comments
  let entry = `${formattedDate} ${payee}`;

  // Add user context as first comment if enabled
  if (actionData && actionData.include_user_context && userContext && userContext.trim()) {
    entry += `\n    ;${userContext.trim()}`;
  }

  // Add tags as individual comments if provided
  if (tags && tags.trim()) {
    const tagArray = tags.split(',').map(tag => tag.trim());
    const tagComments = tagArray.map(tag => `;${tag}`).join(' ');
    entry += `\n    ${tagComments}`;
  }

  if (isCredit) {
    // Income: Money flows TO the funding account FROM the target account
    entry += `\n    ${fundingAccount}    ${formattedAmount}\n    ${targetAccount}`;
  } else {
    // Expense/Transfer: Money flows FROM the funding account TO the target account
    entry += `\n    ${targetAccount}    ${formattedAmount}\n    ${fundingAccount}`;
  }

  return entry;
}

function generateSplitLedgerEntry(formattedDate, payee, targetAccount, totalAmount, fundingAccount, isCredit, tags, actionData, userContext = null) {
  // Build the entry with comments
  let entry = `${formattedDate} ${payee}`;

  // Add user context as first comment if enabled
  if (actionData && actionData.include_user_context && userContext && userContext.trim()) {
    entry += `\n    ;${userContext.trim()}`;
  }

  // Add tags as individual comments if provided
  if (tags && tags.trim()) {
    const tagArray = tags.split(',').map(tag => tag.trim());
    const tagComments = tagArray.map(tag => `;${tag}`).join(' ');
    entry += `\n    ${tagComments}`;
  }

  if (isCredit) {
    // For credit transactions, splits don't make much sense in the expense sharing context
    // Just fall back to standard entry
    const formattedAmount = `₹${totalAmount.toFixed(2)}`;
    return entry + `\n    ${fundingAccount}    ${formattedAmount}\n    ${targetAccount}`;
  }

  // Expense split logic
  const splitType = actionData.split_type;
  const splitConfig = actionData.split_config;

  if (splitType === 'fifty_fifty') {
    const yourShare = Math.ceil(totalAmount / 2); // You get the extra rupee
    const theirShare = totalAmount - yourShare;

    entry += `\n    ${targetAccount}    ₹${yourShare.toFixed(2)}`;
    entry += `\n    ${splitConfig.split_account}    ₹${theirShare.toFixed(2)}`;
    entry += `\n    ${fundingAccount}`;

  } else if (splitType === 'three_way') {
    const yourShare = Math.ceil(totalAmount / 3); // You get the extra rupee(s)
    const remainingAmount = totalAmount - yourShare;
    const share1 = Math.floor(remainingAmount / 2);
    const share2 = remainingAmount - share1;

    entry += `\n    ${targetAccount}    ₹${yourShare.toFixed(2)}`;
    entry += `\n    ${splitConfig.split_accounts[0]}    ₹${share1.toFixed(2)}`;
    entry += `\n    ${splitConfig.split_accounts[1]}    ₹${share2.toFixed(2)}`;
    entry += `\n    ${fundingAccount}`;

  } else if (splitType === 'custom') {
    const yourSharePercent = splitConfig.your_share_percent;
    const yourShare = Math.floor((totalAmount * yourSharePercent) / 100);
    let remainingAmount = totalAmount - yourShare;

    entry += `\n    ${targetAccount}    ₹${yourShare.toFixed(2)}`;

    // Add each custom split
    splitConfig.custom_splits.forEach(function(split, index) {
      const isLast = index === splitConfig.custom_splits.length - 1;
      let splitAmount;

      if (isLast) {
        // Last entry gets any remaining amount to ensure total balance
        splitAmount = remainingAmount;
      } else {
        splitAmount = Math.floor((totalAmount * split.percent) / 100);
        remainingAmount -= splitAmount;
      }

      entry += `\n    ${split.account}    ₹${splitAmount.toFixed(2)}`;
    });

    entry += `\n    ${fundingAccount}`;
  }

  return entry;
}