/**
 * RuleBuilder.js - Unified Rule Management System
 */

function createRuleFromSelection() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const activeRange = sheet.getActiveRange();
  const currentRow = activeRange.getRow();

  Logger.log("Selected row: " + currentRow);

  if (currentRow < 5) {
    SpreadsheetApp.getUi().alert(
      "Please select a transaction row (not header or metadata)"
    );
    return;
  }

  const headers = sheet.getRange(4, 1, 1, sheet.getLastColumn()).getValues()[0];
  Logger.log("Headers: " + JSON.stringify(headers));

  const colIndices = getColumnIndices(headers);
  Logger.log("Column indices: " + JSON.stringify(colIndices));

  const row = sheet
    .getRange(currentRow, 1, 1, sheet.getLastColumn())
    .getValues()[0];
  Logger.log("Row data: " + JSON.stringify(row));

  const transactionData = {
    narration: String(row[colIndices.narration] || ""),
    userContext: String(row[colIndices.userContext] || ""),
    withdrawal:
      parseFloat(String(row[colIndices.withdrawal] || "0").replace(/,/g, "")) ||
      0,
    deposit:
      parseFloat(String(row[colIndices.deposit] || "0").replace(/,/g, "")) || 0,
  };

  Logger.log("Final transaction data: " + JSON.stringify(transactionData));

  openRuleBuilder("create_from_transaction", transactionData);
}

function createNewRule() {
  openRuleBuilder("create_new");
}

function editExistingRule() {
  openRuleBuilder("edit");
}

function openRuleBuilder(mode, transactionData = null) {
  Logger.log(
    "openRuleBuilder: mode=" +
      mode +
      ", data=" +
      JSON.stringify(transactionData)
  );
  const htmlTemplate = HtmlService.createTemplate(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: Google Sans, Arial, sans-serif; padding: 20px; margin: 0; }
        .container { max-width: 600px; margin: 0 auto; }
        .section { margin-bottom: 25px; padding: 15px; border: 1px solid #e0e0e0; border-radius: 8px; }
        .section-title { font-weight: bold; margin-bottom: 15px; color: #1a73e8; font-size: 16px; }
        
        .form-row { margin-bottom: 15px; }
        .form-row label { display: block; font-weight: 500; margin-bottom: 5px; }
        .form-row input, .form-row select, .form-row textarea { 
          width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px; font-size: 14px;
          box-sizing: border-box;
        }
        
        .form-row-inline { display: flex; gap: 10px; align-items: end; }
        .form-row-inline > div { flex: 1; }
        
        .button { 
          background: #1a73e8; color: white; padding: 10px 20px; border: none; 
          border-radius: 4px; cursor: pointer; margin: 5px; font-size: 14px;
        }
        .button:hover { background: #1557b0; }
        .button-secondary { background: #f8f9fa; color: #3c4043; border: 1px solid #dadce0; }
        .button-secondary:hover { background: #e8eaed; }
        .button-danger { background: #ea4335; }
        .button-danger:hover { background: #d33b2c; }
        
        .preview-box { 
          background: #f8f9fa; padding: 15px; border-radius: 4px; 
          font-family: monospace; white-space: pre-line; margin-top: 10px;
          border: 1px solid #e0e0e0;
        }
        
        .status { 
  padding: 15px; 
  margin: 15px 0; 
  border-radius: 8px; 
  font-weight: 500;
  border-left: 4px solid;
}
.success { 
  background: #d4edda; 
  color: #155724; 
  border-color: #28a745;
}
.error { 
  background: #f8d7da; 
  color: #721c24; 
  border-color: #dc3545;
}
.info { 
  background: #d1ecf1; 
  color: #0c5460; 
  border-color: #17a2b8;
}
        
        .condition-row { 
          display: flex; gap: 10px; align-items: end; margin-bottom: 10px; 
          padding: 10px; background: #f8f9fa; border-radius: 4px;
        }
        .condition-row > div { flex: 1; }
        .remove-condition { 
          background: #ea4335; color: white; border: none; padding: 8px 12px; 
          border-radius: 4px; cursor: pointer; flex: 0 0 auto;
        }
        
        #ruleSelector { margin-bottom: 20px; }
        
        .hidden { display: none; }
      </style>
    </head>
    <body>
      <div class="container">
        <h2 id="formTitle">Rule Builder</h2>
        
         <div id="status"></div>

        <!-- Rule Selector (for edit mode) -->
        <div id="ruleSelector" class="section hidden">
          <div class="section-title">Select Rule to Edit</div>
          <div class="form-row-inline">
            <div>
              <select id="existingRuleSelect">
                <option value="">Loading rules...</option>
              </select>
            </div>
            <div style="flex: 0 0 auto;">
              <button type="button" class="button" onclick="loadSelectedRule()">Load Rule</button>
            </div>
          </div>
        </div>
        
        <!-- Rule Metadata -->
        <div class="section">
          <div class="section-title">Rule Information</div>
          <div class="form-row-inline">
            <div>
              <label for="ruleId">Rule ID</label>
              <input type="text" id="ruleId" placeholder="Auto-generated">
            </div>
            <div>
              <label for="priority">Priority</label>
              <input type="number" id="priority" value="1" min="1">
            </div>
            <div style="flex: 0 0 100px;">
              <label>&nbsp;</label>
              <label style="display: flex; align-items: center; gap: 5px;">
                <input type="checkbox" id="isActive" checked> Active
              </label>
            </div>
          </div>
        </div>
        
        <!-- Conditions -->
<div class="section">
  <div class="section-title">Conditions</div>
  <div id="conditionsContainer">
    <!-- Conditions will be added dynamically -->
  </div>
  <div style="display: flex; gap: 10px; margin-top: 10px;">
    <button type="button" class="button-secondary button" onclick="addCondition()">Add Condition</button>
   <button type="button" class="button-secondary button" onclick="generateRegexPattern()">🤖 Generate Regex</button>
   <button type="button" class="button-secondary" onclick="checkTransactionData()">Debug Data</button>
  </div>
</div>
        <!-- Action -->
        <div class="section">
          <div class="section-title">Action</div>
          <div class="form-row">
            <label for="actionType">Action Type</label>
            <select id="actionType" onchange="updateActionFields()">
              <option value="CREATE_ENTRY">Create Entry</option>
              <option value="CREATE_TRANSFER">Create Transfer</option>
            </select>
          </div>
          
          <div class="form-row">
            <label for="account">Account</label>
            <select id="account">
              <option value="">Loading accounts...</option>
            </select>
          </div>
          
          <div class="form-row">
            <label for="payee">Payee</label>
            <input type="text" id="payee" placeholder="Transaction description">
          </div>
          
          <div class="form-row">
            <label for="tags">Tags (comma-separated)</label>
            <input type="text" id="tags" placeholder="tag1,tag2,tag3">
          </div>
        </div>
        
        <!-- Preview -->
        <div class="section">
          <div class="section-title">Preview</div>
          <button type="button" class="button-secondary" onclick="generatePreview()">Generate Preview</button>
          <div id="previewContainer" class="preview-box hidden"></div>
        </div>
        
        <!-- Actions -->
        <div style="text-align: center; margin-top: 30px;">
          <button type="button" class="button" onclick="saveRule()" id="saveButton">Save Rule</button>
          <button type="button" class="button-secondary" onclick="testRule()">Test Rule</button>
          <button type="button" class="button-danger hidden" onclick="deleteRule()" id="deleteButton">Delete Rule</button>
          <button type="button" class="button-secondary" onclick="google.script.host.close()">Cancel</button>
        </div>
        
       
      </div>
      
      <script>
        let currentMode = '<?= mode ?>';
let transactionData;
try {
  transactionData = <?= JSON.stringify(transactionData || null) ?>;
} catch(e) {
  transactionData = null;
}

// Debug: Check what we actually got
if (transactionData) {
  console.log('transactionData assigned successfully:', transactionData);
} else {
  console.log('transactionData is null/undefined');
}
alert('Raw template data: ' + '<?= JSON.stringify(transactionData || {}) ?>');
        let conditionCounter = 0;
        let allAccounts = [];
        let existingRules = [];
        
        // Initialize the form based on mode
        window.onload = function() {
          initializeForm();
          loadAccounts();
          if (currentMode === 'edit') {
            loadExistingRules();
          }
        };
        
        function initializeForm() {
          const title = document.getElementById('formTitle');
          const ruleSelector = document.getElementById('ruleSelector');
          const deleteButton = document.getElementById('deleteButton');
          const saveButton = document.getElementById('saveButton');
          
          switch(currentMode) {
            case 'create_from_transaction':
              title.textContent = 'Create Rule from Transaction';
              prefillFromTransaction();
              break;
            case 'create_new':
              title.textContent = 'Create New Rule';
              addCondition(); // Start with one condition
              break;
            case 'edit':
              title.textContent = 'Edit Rule';
              ruleSelector.classList.remove('hidden');
              deleteButton.classList.remove('hidden');
              saveButton.textContent = 'Update Rule';
              break;
          }
        }

function generateRegexPattern() {
  showStatus('Generating regex pattern...', 'info');
  
  // Try to get the raw template data directly
  var rawData = '<?= JSON.stringify(transactionData || null) ?>';
  var data = null;
  
  try {
    data = JSON.parse(rawData);
  } catch(e) {
    showStatus('Failed to parse transaction data', 'error');
    return;
  }
  
  if (!data || !data.narration) {
    showStatus('No transaction narration available', 'error');
    return;
  }
  
var prompt = 'Extract the merchant identifier from this transaction: "' + data.narration + '"\\n\\n' +
  'Rules:\\n' +
  '- For UPI transactions, extract the merchant name before the @ symbol\\n' +
  '- For NEFT transactions, extract the company name after the last slash\\n' +
  '- Create a simple regex pattern to match that merchant\\n\\n' +
  'Return only the regex pattern, nothing else.';

  google.script.run
    .withSuccessHandler(function(pattern) {
      var cleanPattern = pattern.trim().replace(/^["']|["']$/g, '');
      var lastConditionNum = conditionCounter || 1;
      var conditionSelect = document.getElementById('condition_' + lastConditionNum);
      var patternInput = document.getElementById('pattern_' + lastConditionNum);
      
if (conditionSelect && patternInput) {
  conditionSelect.value = 'Narration REGEX';
  patternInput.value = cleanPattern;
  
  // Show preview of what this pattern will match
  var previewDiv = document.getElementById('pattern_preview_1');
  if (previewDiv) {
    var testResult = data.narration.match(new RegExp(cleanPattern, 'i'));
    if (testResult) {
      previewDiv.textContent = 'Matches: "' + testResult[0] + '"';
      previewDiv.style.color = '#28a745';
    } else {
      previewDiv.textContent = 'No match found in current transaction';
      previewDiv.style.color = '#dc3545';
    }
  }
  
  showStatus('Regex pattern generated: ' + cleanPattern, 'success');
}
  else {
        showStatus('Please add a condition first', 'error');
      }
    })
    .withFailureHandler(function(error) {
      showStatus('Failed to generate regex: ' + error, 'error');
    })
    .callLLMForRegex(prompt);
}
        
        
        function prefillFromTransaction() {
  // Add initial condition based on transaction
  addCondition();
  
  console.log('Transaction data:', transactionData); // Debug line
  
  // Try to extract meaningful pattern from narration
  if (transactionData.narration) {
    console.log('Narration:', transactionData.narration); // Debug line
    
    // Extract UPI ID or merchant info
    const upiMatch = transactionData.narration.match(/([^\/]+@[^\/]+)/);
    console.log('UPI Match:', upiMatch); // Debug line
    
    if (upiMatch) {
      setTimeout(() => { // Add delay to ensure DOM is ready
        document.getElementById('condition_1').value = 'Narration CONTAINS';
        document.getElementById('pattern_1').value = upiMatch[1];
        console.log('Prefilled pattern:', upiMatch[1]); // Debug line
      }, 100);
    }
  }
  
  // Prefill payee from user context
  if (transactionData.userContext) {
    setTimeout(() => {
      document.getElementById('payee').value = transactionData.userContext;
    }, 100);
  }
  
  // Suggest account based on user context
  setTimeout(() => {
    suggestAccount();
  }, 100);
}

function addCondition() {
  conditionCounter++;
  const container = document.getElementById('conditionsContainer');
  
  const conditionDiv = document.createElement('div');
  conditionDiv.className = 'condition-row';
  conditionDiv.id = 'condition_row_' + conditionCounter;
  
  conditionDiv.innerHTML = 
  '<div>' +
    '<label>Condition</label>' +
    '<select id="condition_' + conditionCounter + '">' +
      '<option value="Narration CONTAINS">Narration Contains</option>' +
      '<option value="Narration REGEX">Narration Regex</option>' +
      '<option value="Amount ==">Amount Equals</option>' +
      '<option value="Amount >">Amount Greater Than</option>' +
      '<option value="Amount <">Amount Less Than</option>' +
      '<option value="User_Context CONTAINS">User Context Contains</option>' +
    '</select>' +
  '</div>' +
  '<div>' +
    '<label>Pattern</label>' +
    '<input type="text" id="pattern_' + conditionCounter + '" placeholder="Pattern or value">' +
    '<div id="pattern_preview_' + conditionCounter + '" style="font-size: 11px; color: #666; margin-top: 5px; font-family: monospace;"></div>' +
  '</div>' +
  (conditionCounter > 1 ? '<button type="button" class="remove-condition" onclick="removeCondition(' + conditionCounter + ')">×</button>' : '<div></div>');
  
  container.appendChild(conditionDiv);
}

function removeCondition(id) {
  const element = document.getElementById('condition_row_' + id);
  if (element) {
    element.remove();
  }
}

function updateActionFields() {
  const actionType = document.getElementById('actionType').value;
  const accountLabel = document.querySelector('label[for="account"]');
  
  if (actionType === 'CREATE_TRANSFER') {
    accountLabel.textContent = 'To Account';
  } else {
    accountLabel.textContent = 'Account';
  }
}

function loadAccounts() {
  google.script.run
    .withSuccessHandler(function(accounts) {
      allAccounts = accounts;
      const accountSelect = document.getElementById('account');
      accountSelect.innerHTML = '<option value="">Select account...</option>';
      
      accounts.forEach(account => {
        const option = document.createElement('option');
        option.value = account;
        option.textContent = account;
        accountSelect.appendChild(option);
      });
    })
    .withFailureHandler(function(error) {
      showStatus('Failed to load accounts: ' + error, 'error');
    })
    .getAccountList();
}

function loadExistingRules() {
  google.script.run
    .withSuccessHandler(function(rules) {
      existingRules = rules;
      const select = document.getElementById('existingRuleSelect');
      select.innerHTML = '<option value="">Select a rule to edit...</option>';
      
      rules.forEach((rule, index) => {
        const option = document.createElement('option');
        option.value = index;
  option.textContent = rule.id + ' - ' + rule.condition + ' (Priority: ' + rule.priority + ')';
        select.appendChild(option);
      });
    })
    .withFailureHandler(function(error) {
      showStatus('Failed to load rules: ' + error, 'error');
    })
    .getExistingRules();
}

function loadSelectedRule() {
  const selectIndex = document.getElementById('existingRuleSelect').value;
  if (!selectIndex) return;
  
  const rule = existingRules[parseInt(selectIndex)];
  if (!rule) return;
  
  // Populate form with rule data
  document.getElementById('ruleId').value = rule.id;
  document.getElementById('ruleId').readOnly = true;
  document.getElementById('priority').value = rule.priority;
  document.getElementById('isActive').checked = rule.active;
  
  // Clear existing conditions and add rule conditions
  document.getElementById('conditionsContainer').innerHTML = '';
  conditionCounter = 0;
  
  const conditions = rule.condition.includes(' AND ') ? rule.condition.split(' AND ') : [rule.condition];
  const patterns = rule.pattern.includes(';') ? rule.pattern.split(';') : [rule.pattern];
  
  for (let i = 0; i < conditions.length; i++) {
    addCondition();
  document.getElementById('condition_' + conditionCounter).value = conditions[i].trim();
  document.getElementById('pattern_' + conditionCounter).value = patterns[i] ? patterns[i].trim() : '';
  }
  
  // Parse action value JSON
  try {
    const actionData = JSON.parse(rule.actionValue.replace(/'/g, '"'));
    document.getElementById('actionType').value = rule.actionType;
    document.getElementById('account').value = actionData.account || actionData.to_account || '';
    document.getElementById('payee').value = actionData.payee || '';
    document.getElementById('tags').value = actionData.tags || '';
  } catch (e) {
    showStatus('Error parsing rule action: ' + e.message, 'error');
  }
  
  updateActionFields();
}

function suggestAccount() {
  const userContext = transactionData.userContext ? transactionData.userContext.toLowerCase() : '';
  
  const suggestions = {
    'food': 'Expenses:Household:Food',
    'dining': 'Expenses:Entertainment:Dining Out',
    'transport': 'Expenses:Transport:Taxis',
    'subscription': 'Expenses:Shopping:Subscriptions and Digital Purchases',
    'card payment': 'Liabilities:Credit Card:SBI',
    'payment': 'Liabilities:Payables:Ananya'
  };
  
  for (const [keyword, account] of Object.entries(suggestions)) {
    if (userContext.includes(keyword)) {
      document.getElementById('account').value = account;
      break;
    }
  }
}

function generatePreview() {
  const ruleData = collectFormData();
  if (!ruleData) return;
  
  // Generate preview of what this rule would create
  google.script.run
    .withSuccessHandler(function(preview) {
      const container = document.getElementById('previewContainer');
      container.textContent = preview;
      container.classList.remove('hidden');
    })
    .withFailureHandler(function(error) {
      showStatus('Preview generation failed: ' + error, 'error');
    })
    .generateRulePreview(ruleData, transactionData);
}

function collectFormData() {
  const conditions = [];
  const patterns = [];
  
  // Collect all conditions
  for (let i = 1; i <= conditionCounter; i++) {
  const conditionEl = document.getElementById('condition_' + i);
  const patternEl = document.getElementById('pattern_' + i);
    
    if (conditionEl && patternEl && conditionEl.value && patternEl.value) {
      conditions.push(conditionEl.value);
      patterns.push(patternEl.value);
    }
  }
  
  if (conditions.length === 0) {
    showStatus('Please add at least one condition', 'error');
    return null;
  }
  
  const actionType = document.getElementById('actionType').value;
  const account = document.getElementById('account').value;
  const payee = document.getElementById('payee').value;
  const tags = document.getElementById('tags').value;
  
  if (!account) {
    showStatus('Please select an account', 'error');
    return null;
  }
  
  const actionValue = {
    payee: payee,
    tags: tags
  };
  
  if (actionType === 'CREATE_TRANSFER') {
    actionValue.to_account = account;
  } else {
    actionValue.account = account;
  }
  
  return {
    id: document.getElementById('ruleId').value || generateRuleId(),
    priority: parseInt(document.getElementById('priority').value) || 1,
    active: document.getElementById('isActive').checked,
    condition: conditions.join(' AND '),
    pattern: patterns.join(';'),
    actionType: actionType,
    actionValue: JSON.stringify(actionValue)
  };
}

function generateRuleId() {
  return 'R' + String(Date.now()).slice(-6);
}

function saveRule() {
  const ruleData = collectFormData();
  if (!ruleData) return;
  
  showStatus('Saving rule...', 'info');
  
  google.script.run
    .withSuccessHandler(function(result) {
      showStatus('Rule saved successfully!', 'success');
      if (currentMode !== 'edit') {
        // Clear form for new rule
        setTimeout(() => google.script.host.close(), 1500);
      }
    })
    .withFailureHandler(function(error) {
      showStatus('Failed to save rule: ' + error, 'error');
    })
    .saveRuleToSheet(ruleData, currentMode === 'edit');
}

function testRule() {
  const ruleData = collectFormData();
  if (!ruleData || !transactionData.narration) return;
  
  google.script.run
    .withSuccessHandler(function(result) {
      if (result.matches) {
        showStatus('✅ Rule matches this transaction!', 'success');
      } else {
        showStatus('❌ Rule does not match this transaction', 'error');
      }
    })
    .withFailureHandler(function(error) {
      showStatus('Test failed: ' + error, 'error');
    })
    .testRuleMatch(ruleData, transactionData);
}

function deleteRule() {
  const ruleId = document.getElementById('ruleId').value;
  if (!ruleId) return;
  
    if (!confirm('Are you sure you want to delete rule ' + ruleId + '?')) return;
  
  google.script.run
    .withSuccessHandler(function(result) {
      showStatus('Rule deleted successfully!', 'success');
      setTimeout(() => google.script.host.close(), 1500);
    })
    .withFailureHandler(function(error) {
      showStatus('Failed to delete rule: ' + error, 'error');
    })
    .deleteRuleFromSheet(ruleId);
}

function showStatus(message, type) {
  const statusDiv = document.getElementById('status');
    statusDiv.innerHTML = '<div class="status ' + type + '">' + message + '</div>';
  setTimeout(() => statusDiv.innerHTML = '', 5000);
}

function checkTransactionData() {
  if (transactionData) {
    showStatus('Transaction data found: ' + JSON.stringify(transactionData), 'success');
  } else {
    showStatus('Transaction data is null or undefined', 'error');
  }
}
      </script>
    </body>
    </html>
  `);

  htmlTemplate.mode = mode;
  htmlTemplate.transactionData = transactionData;

  const html = htmlTemplate.evaluate().setWidth(700).setHeight(800);

  SpreadsheetApp.getUi().showModalDialog(html, "Rule Builder");
}

// Server-side functions for Rule Builder

function getAccountList() {
  try {
    const accountsSheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Accounts");
    if (!accountsSheet) {
      return [
        "Expenses:Others:Other Charges",
        "Expenses:Household:Food",
        "Expenses:Transport:Taxis",
      ];
    }

    // Get all data instead of just first 50 rows
    const lastRow = accountsSheet.getLastRow();
    const data = accountsSheet.getRange("A2:A" + lastRow).getValues();
    const accounts = data
      .map((row) => row[0])
      .filter((account) => account && account.trim());
    return accounts;
  } catch (error) {
    throw new Error("Failed to load accounts: " + error.message);
  }
}

function getExistingRules() {
  try {
    const rulesSheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rules");
    if (!rulesSheet || rulesSheet.getLastRow() < 2) {
      return [];
    }

    const rulesData = rulesSheet
      .getRange(2, 1, rulesSheet.getLastRow() - 1, 8)
      .getValues();
    return rulesData.map((rule) => ({
      id: rule[0],
      priority: rule[1],
      active: rule[2],
      condition: rule[3],
      pattern: rule[4],
      actionType: rule[5],
      actionValue: rule[6],
    }));
  } catch (error) {
    throw new Error("Failed to load rules: " + error.message);
  }
}

function saveRuleToSheet(ruleData, isEdit = false) {
  try {
    const rulesSheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rules");
    if (!rulesSheet) {
      throw new Error("Rules sheet not found");
    }

    const ruleArray = [
      ruleData.id,
      ruleData.priority,
      ruleData.active,
      ruleData.condition,
      ruleData.pattern,
      ruleData.actionType,
      ruleData.actionValue,
    ];

    if (isEdit) {
      // Find existing rule and update it
      const data = rulesSheet
        .getRange(2, 1, rulesSheet.getLastRow() - 1, 8)
        .getValues();
      for (let i = 0; i < data.length; i++) {
        if (data[i][0] === ruleData.id) {
          rulesSheet.getRange(i + 2, 1, 1, 7).setValues([ruleArray]);
          return { success: true };
        }
      }
      throw new Error("Rule not found for editing");
    } else {
      // Add new rule
      rulesSheet.appendRow(ruleArray);
      return { success: true };
    }
  } catch (error) {
    throw new Error("Failed to save rule: " + error.message);
  }
}

function deleteRuleFromSheet(ruleId) {
  try {
    const rulesSheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rules");
    if (!rulesSheet) {
      throw new Error("Rules sheet not found");
    }

    const data = rulesSheet
      .getRange(2, 1, rulesSheet.getLastRow() - 1, 8)
      .getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === ruleId) {
        rulesSheet.deleteRow(i + 2);
        return { success: true };
      }
    }
    throw new Error("Rule not found");
  } catch (error) {
    throw new Error("Failed to delete rule: " + error.message);
  }
}

function generateRulePreview(ruleData, transactionData) {
  try {
    // Simulate what the rule would generate
    const date = new Date();
    const amount = transactionData.deposit || transactionData.withdrawal || 100;
    const isCredit = transactionData.deposit > 0;

    const actionData = JSON.parse(ruleData.actionValue);
    const payee = actionData.payee || "Sample Transaction";
    const account = actionData.account || actionData.to_account;
    const tags = actionData.tags || "";

    return formatLedgerCliEntry(
      date,
      payee,
      account,
      amount,
      "Assets:Checking:Punjab National Bank",
      isCredit,
      tags
    );
  } catch (error) {
    throw new Error("Preview generation failed: " + error.message);
  }
}

function testRuleMatch(ruleData, transactionData) {
  try {
    // Test if the rule would match the transaction
    const narration = transactionData.narration;
    const amount = transactionData.deposit || transactionData.withdrawal;

    const conditions = ruleData.condition.includes(" AND ")
      ? ruleData.condition.split(" AND ")
      : [ruleData.condition];
    const patterns = ruleData.pattern.includes(";")
      ? ruleData.pattern.split(";")
      : [ruleData.pattern];

    for (let i = 0; i < conditions.length; i++) {
      const condition = conditions[i].trim();
      const pattern = patterns[i].trim();

      if (condition === "Narration CONTAINS") {
        if (!narration.toLowerCase().includes(pattern.toLowerCase())) {
          return { matches: false };
        }
      } else if (condition === "Amount ==") {
        if (amount != parseFloat(pattern)) {
          return { matches: false };
        }
      }
      // Add more condition types as needed
    }

    return { matches: true };
  } catch (error) {
    throw new Error("Test failed: " + error.message);
  }
}

function callLLMForRegex(prompt) {
  try {
    Logger.log(
      "Calling LLM for regex with prompt: " + prompt.substring(0, 100) + "..."
    );
    const result = callLLM(prompt, 0.3, 500);
    Logger.log("LLM regex result: " + result);
    return result;
  } catch (error) {
    Logger.log("LLM regex error: " + error.message);
    throw new Error("LLM call failed: " + error.message);
  }
}
