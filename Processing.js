function processAllRows() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No data rows found to process.');
    return;
  }
  
  processRowRange(sheet, 2, lastRow);
}

function processSelectedRows() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const selection = sheet.getActiveRange();
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();
  
  // Skip header row
  const actualStartRow = Math.max(startRow, 2);
  const endRow = Math.min(actualStartRow + numRows - 1, sheet.getLastRow());
  
  if (actualStartRow > sheet.getLastRow()) {
    SpreadsheetApp.getUi().alert('No data rows selected to process.');
    return;
  }
  
  processRowRange(sheet, actualStartRow, endRow);
}

function processCurrentRow() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const currentRow = sheet.getActiveCell().getRow();
  
  if (currentRow < 2) {
    SpreadsheetApp.getUi().alert('Please select a transaction row (not header)');
    return;
  }
  
  processRowRange(sheet, currentRow, currentRow);
}

function processRowRange(sheet, startRow, endRow) {
  try {
    // Get column indices
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colIndices = getColumnIndices(headers);
    
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
      
      // Process the transaction
      const result = processTransaction(row, colIndices);
      
      // Update the row with results
      sheet.getRange(rowNum, colIndices.tags + 1).setValue(result.tags);
      sheet.getRange(rowNum, colIndices.confidence + 1).setValue(result.confidence);
      sheet.getRange(rowNum, colIndices.finalEntry + 1).setValue(result.finalEntry);
    }
    
    SpreadsheetApp.getActiveSpreadsheet().toast(`Completed! Processed ${processedCount}/${totalRows} rows`, 'Ledger Tools', 5);
    SpreadsheetApp.getUi().alert(`✅ Processed ${processedCount} transactions in rows ${startRow} to ${endRow}`);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error processing transactions: ' + error.message);
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

function processTransaction(row, colIndices) {
  // Extract transaction data
  const narration = row[colIndices.narration] || '';
  const userContext = row[colIndices.userContext] || '';
  const withdrawal = row[colIndices.withdrawal] || 0;
  const deposit = row[colIndices.deposit] || 0;
  const date = row[colIndices.date];
  
  // Calculate amount (negative for withdrawals, positive for deposits)
  const amount = deposit ? deposit : -withdrawal;
  
  // For now, let's create basic rule-based suggestions
  // Later we'll add LLM integration
  const suggestion = createBasicSuggestion(narration, amount, userContext);
  
  // Format the final entry
  const finalEntry = formatFinalEntry(date, narration, suggestion.account, amount, suggestion.tags, userContext);
  
  return {
    tags: suggestion.tags,
    confidence: suggestion.confidence,
    finalEntry: finalEntry
  };
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
      console.log('Attempting LLM call for:', narration.substring(0, 50) + '...');
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
    const response = callLLM(prompt, 0.3, 500); // Adequate token limit
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
  // Your existing basic pattern matching logic
  const lowerNarration = narration.toLowerCase();
  const lowerContext = userContext.toLowerCase();
  
  let account = 'Expenses:Others:Other Charges';
  let tags = 'other';
  let confidence = 0.3;
  
  // Basic pattern matching (keep your existing logic here)
  if (userContext) {
    if (lowerContext.includes('food')) {
      account = 'Expenses:Household:Food';
      tags = 'food';
      confidence = 0.7;
    } // ... etc
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
    
    const data = accountsSheet.getRange('A2:A50').getValues(); // Get first 50 accounts
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

function formatFinalEntry(date, narration, account, amount, tags, userContext) {
  const formattedDate = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "yyyy/MM/dd");
  const sourceAccount = 'Assets:Checking:Bank of Baroda'; // Default source account
  
  // Use user context as description if provided, otherwise use narration
  let description = userContext && userContext.trim() ? userContext.trim() : narration;
  
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
    return response; // Explicit return
  } catch (error) {
    console.log('Error:', error.message);
    SpreadsheetApp.getUi().alert('LLM Test Failed: ' + error.message);
    return null; // Explicit return
  }
}