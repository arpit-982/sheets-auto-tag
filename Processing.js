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
    // Read headers from row 4 instead of row 1
    const headers = sheet.getRange(4, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colIndices = getColumnIndices(headers);
    
    // Read funding account from cell B1
    const fundingAccount = sheet.getRange('B1').getValue() || 'Assets:Checking:Bank of Baroda';
    
    let processedCount = 0;
    const totalRows = endRow - startRow + 1;
    
    // Show initial toast
    SpreadsheetApp.getActiveSpreadsheet().toast(`Starting processing... (0/${totalRows})`, 'Ledger Tools', 5);
    
    // Process each row in range
    for (let rowNum = startRow; rowNum <= endRow; rowNum++) {
      const row = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
      
      // Skip empty rows (no Sr No)
      if (!row[colIndices.srNo]) continue;
      
      // Update progress
      processedCount++;
      if (processedCount % 2 === 0 || processedCount === totalRows) {
        SpreadsheetApp.getActiveSpreadsheet().toast(`Processing... (${processedCount}/${totalRows})`, 'Ledger Tools', 2);
      }
      
      // Process the transaction with funding account
      const result = processTransaction(row, colIndices, fundingAccount);
      
      // Update the row with results
      sheet.getRange(rowNum, colIndices.tags + 1).setValue(result.tags);
      sheet.getRange(rowNum, colIndices.confidence + 1).setValue(result.confidence);
      sheet.getRange(rowNum, colIndices.finalEntry + 1).setValue(result.finalEntry);
    }
    
    SpreadsheetApp.getActiveSpreadsheet().toast(`Completed! Processed ${processedCount}/${totalRows} rows`, 'Ledger Tools', 5);
    SpreadsheetApp.getUi().alert(`✅ Processed ${processedCount} transactions in rows ${startRow} to ${endRow}`);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error processing transactions: ' + error.message);
    console.log('Processing error:', error);
  }
}

function getColumnIndices(headers) {
  return {
    srNo: headers.indexOf('Sr No'),
    date: headers.indexOf('Transaction Date'),
    withdrawal: headers.indexOf('Withdrawal'),
    deposit: headers.indexOf('Deposit'),
    balance: headers.indexOf('Balance'),
    narration: headers.indexOf('Narration'),
    userContext: headers.indexOf('User Context'),
    tags: headers.indexOf('Tags'),
    confidence: headers.indexOf('LLM Confidence'),
    finalEntry: headers.indexOf('Final Entry')
  };
}

/**
 * Orchestrates the processing for a single transaction row.
 * It tries the rule engine first, then falls back to the LLM.
 * @returns {object} An object with {tags, confidence, finalEntry}.
 */
function processTransaction(row, colIndices, fundingAccount) {
  const narration = String(row[colIndices.narration] || '');
  const userContext = String(row[colIndices.userContext] || '');
  const withdrawal = row[colIndices.withdrawal] || 0;
  const deposit = row[colIndices.deposit] || 0;
  const dateValue = row[colIndices.date];

  if (!narration || (withdrawal === 0 && deposit === 0) || !dateValue) {
    return { tags: "", confidence: "", finalEntry: "" }; // Skip empty or invalid rows
  }
  const date = new Date(dateValue);

  const amount = deposit || withdrawal;
  const isCredit = deposit > 0;

  // 1. Try to apply deterministic rules first.
  let result = applyRules(narration, amount, date, fundingAccount, isCredit);

  // 2. If no rule matched, fall back to your existing LLM suggestion logic.
  if (!result) {
    const suggestion = createBasicSuggestion(narration, amount, userContext);
    const payee = narration.split('/')[1] || narration; // Simple payee extraction
    const finalEntry = formatLedgerCliEntry(date, payee, suggestion.account, amount, fundingAccount, isCredit);
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
  const prompt = `Categorize this transaction into the most appropriate account and suggest tags.

Transaction: "${narration}"
Amount: ${amount}
Context: "${userContext || 'none'}"

Common accounts:
- Expenses:Household:Food (food, dining)
- Expenses:Transport:Taxis (auto, uber, transport)
- Expenses:Household:Other Household (small misc expenses)
- Expenses:Shopping:Subscriptions and Digital Purchases (apps, subscriptions)
- Liabilities:Payables:Ananya (splits, payments to people)
- Expenses:Others:Other Charges (default)

Respond with only this JSON format:
{"account":"Expenses:Household:Food","tags":"food","confidence":0.8}`;

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
      tags: parsed.tags || 'ai',
      confidence: Math.min(parsed.confidence || 0.7, 0.95)
    };
    
  } catch (error) {
    console.log('LLM processing failed:', error.message);
    throw error;
  }
}

function createFallbackSuggestion(narration, amount, userContext) {
  // Fixed: Convert to string and handle null/undefined values
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
  
  return { account, tags, confidence };
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

/**
 * Tries to find a matching rule and generate a complete ledger entry.
 * @returns {object|null} A result object with the final entry, tags, and confidence, or null if no match.
 */
function applyRules(narration, amount, date, fundingAccount, isCredit) {
  const rulesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Rules');
  if (!rulesSheet) return null;

  const rulesData = rulesSheet.getRange(2, 1, rulesSheet.getLastRow() - 1, 7).getValues();
  const lowerNarration = narration.toLowerCase();

  for (const rule of rulesData) {
    const isActive = rule[2]; // Active column
    if (!isActive) continue;

    const conditionStr = rule[3], patternStr = String(rule[4]), actionType = rule[5], actionValue = rule[6];
    if (!conditionStr || !patternStr || !actionType || !actionValue) continue;

    const conditions = conditionStr.split(' AND ').map(c => c.trim());
    const patterns = patternStr.split(';').map(p => p.trim());
    if (conditions.length !== patterns.length) continue;

    let allConditionsMet = true;
    for (let i = 0; i < conditions.length; i++) {
      const [field, operator] = conditions[i].split(' ');
      const currentPattern = patterns[i];
      let conditionMet = false;
      try {
        if (field === 'Narration') {
          if (operator === 'CONTAINS' && lowerNarration.includes(currentPattern.toLowerCase())) conditionMet = true;
          if (operator === 'REGEX' && new RegExp(currentPattern, 'i').test(narration)) conditionMet = true;
        } else if (field === 'Amount') {
          const numPattern = parseFloat(currentPattern);
          if (operator === '>' && amount > numPattern) conditionMet = true;
          if (operator === '<' && amount < numPattern) conditionMet = true;
          if (operator === '==' && amount == numPattern) conditionMet = true;
        }
      } catch (e) { console.error(`Rule Error ID ${rule[0]}: ${e}`); }
      if (!conditionMet) { allConditionsMet = false; break; }
    }

    if (allConditionsMet) {
      try {
        const params = JSON.parse(actionValue);
        let finalEntry;
        if (actionType === 'CREATE_ENTRY') {
          finalEntry = formatLedgerCliEntry(date, params.payee || narration, params.account, amount, fundingAccount, isCredit);
        } else if (actionType === 'CREATE_TRANSFER') {
          finalEntry = formatLedgerCliEntry(date, params.payee || narration, params.to_account, amount, fundingAccount, false);
        } else { continue; }

        return { finalEntry, tags: params.tags || '', confidence: 1.0 };
      } catch (e) { console.error(`Rule Action Error ID ${rule[0]}: ${e}`); return null; }
    }
  }
  return null; // No rule matched
}

/**
 * Formats a standard, two-posting ledger entry in the ledger-cli style.
 * @returns {string} The formatted, multi-line ledger entry.
 */
function formatLedgerCliEntry(date, payee, targetAccount, amount, fundingAccount, isCredit) {
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const dd = String(date.getDate()).padStart(2, '0');
  const formattedDate = `${yyyy}/${mm}/${dd}`;
  const formattedAmount = `₹${amount.toFixed(2)}`;

  if (isCredit) {
    // Income: Money flows TO the funding account FROM the target account
    return `${formattedDate} ${payee}\n    ${fundingAccount}    ${formattedAmount}\n    ${targetAccount}`;
  } else {
    // Expense/Transfer: Money flows FROM the funding account TO the target account
    return `${formattedDate} ${payee}\n    ${targetAccount}    ${formattedAmount}\n    ${fundingAccount}`;
  }
}