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
      <title>Rule Builder</title>
      <style>
        body {
          font-family: 'Google Sans', 'Roboto', Arial, sans-serif;
          padding: 0; margin: 0;
          background-color: #ffffff;
          font-size: 14px;
          color: #3c4043;
        }
        .container { max-width: 100%; margin: 0; padding: 16px; }
        .section {
          margin-bottom: 24px;
          padding: 0;
          border: none;
          border-radius: 0;
        }
        .section-title {
          font-weight: 500;
          margin-bottom: 16px;
          color: #1f1f1f;
          font-size: 16px;
          line-height: 24px;
        }

        .form-row { margin-bottom: 16px; }
        .form-row label {
          display: block;
          font-weight: 400;
          margin-bottom: 6px;
          font-size: 14px;
          color: #5f6368;
          line-height: 20px;
        }
        .form-row input, .form-row select, .form-row textarea {
          width: 100%;
          padding: 8px 12px;
          border: 1px solid #dadce0;
          border-radius: 4px;
          font-size: 14px;
          box-sizing: border-box;
          background-color: #ffffff;
          color: #3c4043;
          line-height: 20px;
          font-family: 'Google Sans', 'Roboto', Arial, sans-serif;
        }
        .form-row input:focus, .form-row select:focus, .form-row textarea:focus {
          outline: none;
          border-color: #1a73e8;
          box-shadow: inset 0 0 0 1px #1a73e8;
        }
        .form-row input[readonly] {
          background-color: #f8f9fa;
          color: #5f6368;
        }

        .form-row-inline { display: flex; gap: 12px; align-items: end; }
        .form-row-inline > div { flex: 1; }
        
        .button {
          background: #1a73e8;
          color: #ffffff;
          padding: 8px 16px;
          border: none;
          border-radius: 4px;
          cursor: pointer;
          margin: 4px;
          font-size: 14px;
          font-weight: 500;
          font-family: 'Google Sans', 'Roboto', Arial, sans-serif;
          line-height: 20px;
          min-height: 36px;
          transition: background-color 0.1s ease;
        }
        .button:hover { background: #1765cc; }
        .button:active { background: #1557b0; }

        .button-secondary {
          background: #ffffff;
          color: #1a73e8;
          border: 1px solid #dadce0;
          font-weight: 500;
        }
        .button-secondary:hover {
          background: #f8f9fa;
          border-color: #c2c7d0;
        }
        .button-secondary:active {
          background: #f1f3f4;
        }

        .button-danger {
          background: #d93025;
          color: #ffffff;
        }
        .button-danger:hover { background: #c5221f; }
        .button-danger:active { background: #b52d20; }
        
        .preview-box {
          background: #f8f9fa;
          padding: 16px;
          border-radius: 4px;
          font-family: 'Roboto Mono', monospace;
          white-space: pre-line;
          margin-top: 12px;
          border: 1px solid #dadce0;
          font-size: 13px;
          color: #3c4043;
        }

        .status {
          padding: 12px 16px;
          margin: 16px 0;
          border-radius: 4px;
          font-weight: 400;
          font-size: 14px;
          line-height: 20px;
        }
        .success {
          background: #e8f5e8;
          color: #137333;
          border: 1px solid #34a853;
        }
        .error {
          background: #fce8e6;
          color: #d93025;
          border: 1px solid #ea4335;
        }
        .info {
          background: #e3f2fd;
          color: #1565c0;
          border: 1px solid #4285f4;
        }
        
        .condition-row {
          display: grid;
          grid-template-columns: 1fr 1fr auto;
          grid-template-rows: auto auto;
          gap: 12px;
          margin-bottom: 12px;
          padding: 12px;
          background: #f8f9fa;
          border-radius: 4px;
          border: 1px solid #e8eaed;
        }
        .remove-condition {
          background: #d93025;
          color: #ffffff;
          border: none;
          padding: 6px 12px;
          border-radius: 4px;
          cursor: pointer;
          flex: 0 0 auto;
          font-size: 13px;
          font-weight: 500;
          font-family: 'Google Sans', 'Roboto', Arial, sans-serif;
          min-height: 32px;
          transition: background-color 0.1s ease;
        }
        .remove-condition:hover { background: #c5221f; }

        .pattern-preview {
          font-size: 12px;
          margin-top: 0;
          font-family: 'Roboto Mono', monospace;
          padding: 6px 8px;
          border-radius: 4px;
          min-height: 18px;
          line-height: 18px;
          border: 1px solid transparent;
          grid-column: 1 / -1;
          width: 100%;
          box-sizing: border-box;
        }

        .pattern-preview.match {
          color: #137333;
          background: #e8f5e8;
          border-color: #34a853;
        }

        .pattern-preview.no-match {
          color: #d93025;
          background: #fce8e6;
          border-color: #ea4335;
        }

        .pattern-preview.error {
          color: #f57c00;
          background: #fff8e1;
          border-color: #ffb300;
        }
        
        #formTitle {
          font-size: 22px;
          font-weight: 400;
          color: #3c4043;
          margin: 0 0 24px 0;
          padding: 0;
          line-height: 28px;
        }

        #ruleSelector { margin-bottom: 24px; }

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
          <div class="form-row">
            <select id="existingRuleSelect" style="margin-bottom: 8px;">
              <option value="">Loading rules...</option>
            </select>
            <button type="button" class="button" onclick="loadSelectedRule()" style="width: 100%;">Load Rule</button>
          </div>
        </div>
        
        <!-- Rule Metadata -->
        <div class="section">
          <div class="section-title">Rule Information</div>
          <div class="form-row">
            <label for="ruleId">Rule ID</label>
            <input type="text" id="ruleId" readonly>
          </div>
          <div class="form-row-inline">
            <div style="flex: 2;">
              <label for="priority">Priority</label>
              <input type="number" id="priority" value="1" min="1">
            </div>
            <div style="flex: 1; display: flex; align-items: end; padding-bottom: 8px;">
              <label style="display: flex; align-items: center; gap: 8px; font-size: 14px; color: #5f6368; cursor: pointer;">
                <input type="checkbox" id="isActive" checked style="margin: 0; width: 16px; height: 16px;"> Active
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
          <div style="display: flex; gap: 8px; margin-top: 16px; flex-wrap: wrap;">
            <button type="button" class="button-secondary" onclick="addCondition()">Add Condition</button>
            <button type="button" class="button-secondary" onclick="generateRegexPattern()">🤖 Generate Regex</button>
            <button type="button" class="button-secondary" onclick="debugTransactionData()">Debug Data</button>
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
        <div style="display: flex; flex-wrap: wrap; gap: 8px; margin-top: 24px; justify-content: flex-start;">
          <button type="button" class="button" onclick="saveRule()" id="saveButton">Save Rule</button>
          <button type="button" class="button-secondary" onclick="testRule()">Test Rule</button>
          <button type="button" class="button-danger hidden" onclick="deleteRule()" id="deleteButton">Delete Rule</button>
          <button type="button" class="button-secondary" onclick="google.script.host.close()">Cancel</button>
        </div>
      </div>
      
      <script>
        let currentMode = '<?= mode ?>';
        let transactionData = null;
        let conditionCounter = 0;
        let allAccounts = [];
        let existingRules = [];
        
        // Initialize transaction data safely
        try {
          const rawData = '<?= JSON.stringify(transactionData || null) ?>';
          if (rawData && rawData !== 'null') {
            transactionData = JSON.parse(rawData);
          }
        } catch(e) {
          console.warn('Failed to parse transaction data:', e);
        }
        
        // Initialize the form based on mode
        window.onload = function() {
          debugLog('Window onload - initializing form');
          initializeForm();
          loadAccounts();
          if (currentMode === 'edit') {
            loadExistingRules();
          } else {
            // Load the next Rule ID for new rules
            debugLog('Loading next rule ID for new rule');
            loadNextRuleId();
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

        function loadNextRuleId() {
          debugLog('loadNextRuleId called');
          google.script.run
            .withSuccessHandler(function(nextRuleId) {
              debugLog('Received next rule ID: ' + nextRuleId);
              const ruleIdField = document.getElementById('ruleId');
              if (ruleIdField) {
                ruleIdField.value = nextRuleId;
                debugLog('Set rule ID field to: ' + nextRuleId);
              } else {
                debugLog('Rule ID field not found');
              }
            })
            .withFailureHandler(function(error) {
              debugLog('Failed to load next rule ID: ' + error);
              showStatus('Failed to load next rule ID: ' + error, 'error');
              const ruleIdField = document.getElementById('ruleId');
              if (ruleIdField) {
                ruleIdField.value = 'R001'; // Fallback
              }
            })
            .getNextRuleId();
        }
        
        function buildConditionHTML(id) {
          const removeButton = id > 1
            ? '<button type="button" class="remove-condition" onclick="removeCondition(' + id + ')">×</button>'
            : '<div style="flex: 0 0 auto; width: 40px;"></div>';

          return [
            '<div style="flex: 1;">',
              '<label>Condition</label>',
              '<select id="condition_' + id + '" data-condition-id="' + id + '" onchange="validatePatternRealtime(' + id + ')">',
                '<option value="Narration CONTAINS">Narration Contains</option>',
                '<option value="Narration REGEX">Narration Regex</option>',
                '<option value="Amount ==">Amount Equals</option>',
                '<option value="Amount >">Amount Greater Than</option>',
                '<option value="Amount <">Amount Less Than</option>',
                '<option value="User_Context CONTAINS">User Context Contains</option>',
              '</select>',
            '</div>',
            '<div style="flex: 1;">',
              '<label>Pattern</label>',
              '<input type="text" id="pattern_' + id + '" data-condition-id="' + id + '" placeholder="Pattern or value" oninput="validatePatternRealtime(' + id + ')">',
            '</div>',
            removeButton,
            '<div id="pattern_preview_' + id + '" class="pattern-preview" style="grid-column: 1 / -1;"></div>'
          ].join('');
        }
        
        function addCondition() {
          conditionCounter++;
          const container = document.getElementById('conditionsContainer');
          
          const conditionDiv = document.createElement('div');
          conditionDiv.className = 'condition-row';
          conditionDiv.id = 'condition_row_' + conditionCounter;
          conditionDiv.dataset.conditionId = conditionCounter;
          
          conditionDiv.innerHTML = buildConditionHTML(conditionCounter);
          container.appendChild(conditionDiv);
        }
        
        function removeCondition(id) {
          const element = document.getElementById('condition_row_' + id);
          if (element) {
            element.remove();
          }
        }
        
        function validatePatternRealtime(conditionId) {
          if (!transactionData || !transactionData.narration) return;
          
          const conditionSelect = document.getElementById('condition_' + conditionId);
          const patternInput = document.getElementById('pattern_' + conditionId);
          const previewDiv = document.getElementById('pattern_preview_' + conditionId);
          
          if (!conditionSelect || !patternInput || !previewDiv) return;
          
          const condition = conditionSelect.value;
          const pattern = patternInput.value.trim();
          
          if (!pattern) {
            previewDiv.textContent = '';
            previewDiv.className = 'pattern-preview';
            return;
          }
          
          try {
            let matches = false;
            let result = '';
            
            if (condition === 'Narration CONTAINS') {
              matches = transactionData.narration.toLowerCase().includes(pattern.toLowerCase());
              result = matches ? '✓ Contains "' + pattern + '"' : '✗ Does not contain "' + pattern + '"';
            } else if (condition === 'Narration REGEX') {
              const regex = new RegExp(pattern, 'i');
              const match = transactionData.narration.match(regex);
              matches = !!match;
              result = matches ? '✓ Matches: "' + match[0] + '"' : '✗ No match';
            } else if (condition.startsWith('Amount')) {
              const amount = transactionData.deposit || transactionData.withdrawal || 0;
              const targetAmount = parseFloat(pattern);
              if (condition === 'Amount ==') {
                matches = Math.abs(amount - targetAmount) < 0.01;
              } else if (condition === 'Amount >') {
                matches = amount > targetAmount;
              } else if (condition === 'Amount <') {
                matches = amount < targetAmount;
              }
              result = matches ? '✓ Amount condition met' : '✗ Amount condition not met';
            } else if (condition === 'User_Context CONTAINS') {
              const userContext = transactionData.userContext || '';
              matches = userContext.toLowerCase().includes(pattern.toLowerCase());
              result = matches ? '✓ User context contains "' + pattern + '"' : '✗ User context does not contain "' + pattern + '"';
            }
            
            previewDiv.textContent = result;
            previewDiv.className = 'pattern-preview ' + (matches ? 'match' : 'no-match');
            
          } catch (error) {
            previewDiv.textContent = '⚠ Invalid pattern: ' + error.message;
            previewDiv.className = 'pattern-preview error';
          }
        }
        
        function generateRegexPattern() {
          if (!transactionData || !transactionData.narration) {
            showStatus('No transaction narration available for regex generation', 'error');
            return;
          }
          
          // Find the last added condition
          const activeConditions = document.querySelectorAll('[id^="condition_row_"]');
          if (activeConditions.length === 0) {
            showStatus('Please add a condition first', 'error');
            return;
          }
          
          const lastConditionId = activeConditions[activeConditions.length - 1].dataset.conditionId;
          
          showStatus('Generating regex pattern...', 'info');
          
          const prompt = 'Extract the merchant identifier from this transaction: "' + transactionData.narration + '"\\n\\n' +
            'Rules:\\n' +
            '- For UPI transactions, extract the merchant name before the @ symbol\\n' +
            '- For NEFT transactions, extract the company name after the last slash\\n' +
            '- Create a simple regex pattern to match that merchant\\n\\n' +
            'Return only the regex pattern, nothing else.';

          google.script.run
            .withSuccessHandler(function(pattern) {
              const cleanPattern = pattern.trim().replace(/^["']|["']$/g, '');
              
              const conditionSelect = document.getElementById('condition_' + lastConditionId);
              const patternInput = document.getElementById('pattern_' + lastConditionId);
              
              if (conditionSelect && patternInput) {
                conditionSelect.value = 'Narration REGEX';
                patternInput.value = cleanPattern;
                
                // Trigger real-time validation
                validatePatternRealtime(lastConditionId);
                
                showStatus('Regex pattern generated: ' + cleanPattern, 'success');
              } else {
                showStatus('Could not find condition fields to update', 'error');
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
          
          // Try to extract meaningful pattern from narration
          if (transactionData && transactionData.narration) {
            // Extract UPI ID or merchant info
            const upiMatch = transactionData.narration.match(/([^\/]+@[^\/]+)/);
            
            if (upiMatch) {
              setTimeout(function() {
                document.getElementById('condition_1').value = 'Narration CONTAINS';
                document.getElementById('pattern_1').value = upiMatch[1];
                validatePatternRealtime(1);
              }, 100);
            }
          }
          
          // Prefill payee from user context
          if (transactionData && transactionData.userContext) {
            setTimeout(function() {
              document.getElementById('payee').value = transactionData.userContext;
            }, 100);
          }
          
          // Suggest account based on user context
          setTimeout(function() {
            suggestAccount();
          }, 100);
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
              
              accounts.forEach(function(account) {
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
              
              rules.forEach(function(rule, index) {
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
            validatePatternRealtime(conditionCounter);
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
          if (!transactionData || !transactionData.userContext) return;
          
          const userContext = transactionData.userContext.toLowerCase();
          
          const suggestions = {
            'food': 'Expenses:Household:Food',
            'dining': 'Expenses:Entertainment:Dining Out',
            'transport': 'Expenses:Transport:Taxis',
            'subscription': 'Expenses:Shopping:Subscriptions and Digital Purchases',
            'card payment': 'Liabilities:Credit Card:SBI',
            'payment': 'Liabilities:Payables:Ananya'
          };
          
          for (const keyword in suggestions) {
            if (userContext.includes(keyword)) {
              setTimeout(function() {
                document.getElementById('account').value = suggestions[keyword];
              }, 200);
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

          debugLog('Collected conditions: ' + JSON.stringify(conditions));
          debugLog('Collected patterns: ' + JSON.stringify(patterns));
          
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

          debugLog('Action value object: ' + JSON.stringify(actionValue));
          debugLog('Action value JSON: ' + JSON.stringify(actionValue));
          
          const ruleId = document.getElementById('ruleId').value;
          if (!ruleId) {
            showStatus('Rule ID is missing. Please refresh and try again.', 'error');
            return null;
          }

          return {
            id: ruleId,
            priority: parseInt(document.getElementById('priority').value) || 1,
            active: document.getElementById('isActive').checked,
            condition: conditions.join(' AND '),
            pattern: patterns.join(';'),
            actionType: actionType,
            actionValue: JSON.stringify(actionValue)
          };
        }
        
        
        function saveRule() {
          const ruleData = collectFormData();
          if (!ruleData) return;

          debugLog('About to save rule data: ' + JSON.stringify(ruleData));

          showStatus('Saving rule...', 'info');

          google.script.run
            .withSuccessHandler(function(result) {
              debugLog('Save success result: ' + JSON.stringify(result));
              showStatus('Rule saved successfully!', 'success');
              if (currentMode !== 'edit') {
                setTimeout(function() {
                  google.script.host.close();
                }, 1500);
              }
            })
            .withFailureHandler(function(error) {
              debugLog('Save error: ' + error);
              showStatus('Failed to save rule: ' + error, 'error');
            })
            .saveRuleToSheet(ruleData, currentMode === 'edit');
        }
        
        function testRule() {
          const ruleData = collectFormData();
          if (!ruleData) return;
          
          if (!transactionData || !transactionData.narration) {
            showStatus('No transaction data available for testing', 'error');
            return;
          }
          
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
              setTimeout(function() {
                google.script.host.close();
              }, 1500);
            })
            .withFailureHandler(function(error) {
              showStatus('Failed to delete rule: ' + error, 'error');
            })
            .deleteRuleFromSheet(ruleId);
        }
        
        function showStatus(message, type) {
          const statusDiv = document.getElementById('status');
          statusDiv.innerHTML = '<div class="status ' + type + '">' + message + '</div>';
          setTimeout(function() {
            statusDiv.innerHTML = '';
          }, 5000);
        }

        function debugLog(message) {
          // For debugging, we can show debug messages in the UI temporarily
          // and also log to browser console if available
          try {
            console.log('[DEBUG] ' + message);
          } catch(e) {
            // Browser console not available, ignore
          }

          // Also send to Apps Script logger
          google.script.run
            .withFailureHandler(function() {
              // Ignore logging failures
            })
            .logDebugMessage(message);
        }
        
        function debugTransactionData() {
          if (transactionData) {
            showStatus('Transaction data: ' + JSON.stringify(transactionData, null, 2), 'info');
          } else {
            showStatus('No transaction data available', 'error');
          }
        }

        function testSheetAccess() {
          showStatus('Testing Apps Script connection...', 'info');

          // First test basic Apps Script functionality
          google.script.run
            .withSuccessHandler(function(result) {
              showStatus('Apps Script working: ' + result, 'success');

              // Now test sheet access
              setTimeout(function() {
                showStatus('Testing Rules sheet access...', 'info');
                google.script.run
                  .withSuccessHandler(function(result) {
                    if (result.success) {
                      showStatus('Rules sheet accessible! Last row: ' + result.lastRow, 'success');
                    } else {
                      showStatus('Rules sheet access failed: ' + result.error, 'error');
                    }
                  })
                  .withFailureHandler(function(error) {
                    showStatus('Sheet access test failed: ' + error, 'error');
                  })
                  .testRulesSheetAccess();
              }, 1000);
            })
            .withFailureHandler(function(error) {
              showStatus('Apps Script connection failed: ' + error, 'error');
            })
            .simpleTest();
        }

        function testRuleId() {
          showStatus('Testing Rule ID generation...', 'info');
          google.script.run
            .withSuccessHandler(function(result) {
              showStatus('Next Rule ID: ' + result, 'success');
              // Also set it in the field
              const ruleIdField = document.getElementById('ruleId');
              if (ruleIdField) {
                ruleIdField.value = result;
              }
            })
            .withFailureHandler(function(error) {
              showStatus('Rule ID test failed: ' + error, 'error');
            })
            .getNextRuleId();
        }

        function testSave() {
          // Create a simple test rule to verify save functionality
          const testRuleData = {
            id: 'TEST001',
            priority: 1,
            active: true,
            condition: 'Narration CONTAINS',
            pattern: 'test',
            actionType: 'CREATE_ENTRY',
            actionValue: JSON.stringify({
              account: 'Expenses:Others:Other Charges',
              payee: 'Test Transaction',
              tags: 'test'
            })
          };

          showStatus('Testing save functionality...', 'info');

          google.script.run
            .withSuccessHandler(function(result) {
              showStatus('Save test successful: ' + JSON.stringify(result), 'success');
            })
            .withFailureHandler(function(error) {
              showStatus('Save test failed: ' + error, 'error');
            })
            .saveRuleToSheet(testRuleData, false);
        }
      </script>
    </body>
    </html>
  `);

  htmlTemplate.mode = mode;
  htmlTemplate.transactionData = transactionData;

  const html = htmlTemplate.evaluate().setWidth(400).setTitle('Rule Builder');
  SpreadsheetApp.getUi().showSidebar(html);
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

function getNextRuleId() {
  try {
    Logger.log("getNextRuleId function called");

    const rulesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rules");
    if (!rulesSheet || rulesSheet.getLastRow() < 2) {
      Logger.log("No Rules sheet or no data, returning R001");
      return "R001";
    }

    const lastRow = rulesSheet.getLastRow();
    Logger.log("Rules sheet last row: " + lastRow);

    const existingIds = rulesSheet
      .getRange(2, 1, lastRow - 1, 1)
      .getValues()
      .map(row => row[0])
      .filter(id => id && String(id).startsWith('R'));

    Logger.log("Existing rule IDs: " + JSON.stringify(existingIds));

    if (existingIds.length === 0) {
      Logger.log("No existing rule IDs found, returning R001");
      return "R001";
    }

    // Extract numeric parts and find the maximum
    const numbers = existingIds.map(id => {
      const match = String(id).match(/^R(\d+)$/);
      return match ? parseInt(match[1], 10) : 0;
    });

    Logger.log("Extracted numbers: " + JSON.stringify(numbers));

    const maxNumber = Math.max(...numbers);
    const nextNumber = maxNumber + 1;
    const nextRuleId = "R" + String(nextNumber).padStart(3, "0");

    Logger.log("Max number: " + maxNumber + ", Next number: " + nextNumber + ", Next rule ID: " + nextRuleId);

    return nextRuleId;
  } catch (error) {
    Logger.log("Error generating next rule ID: " + error.message);
    Logger.log("Error stack: " + error.stack);
    return "R001";
  }
}

function logDebugMessage(message) {
  Logger.log('[CLIENT DEBUG] ' + message);
}

function simpleTest() {
  Logger.log('simpleTest function called');
  return 'Apps Script is working!';
}

function testRulesSheetAccess() {
  try {
    const rulesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rules");
    if (!rulesSheet) {
      return { success: false, error: "Rules sheet not found" };
    }

    const lastRow = rulesSheet.getLastRow();
    Logger.log("Rules sheet accessible, last row: " + lastRow);

    // Try to read the first few rows to verify structure
    if (lastRow > 1) {
      const sampleData = rulesSheet.getRange(1, 1, Math.min(3, lastRow), 7).getValues();
      Logger.log("Sample data from Rules sheet: " + JSON.stringify(sampleData));
    }

    return { success: true, lastRow: lastRow };
  } catch (error) {
    Logger.log("Error accessing Rules sheet: " + error.message);
    return { success: false, error: error.message };
  }
}

function saveRuleToSheet(ruleData, isEdit = false) {
  try {
    Logger.log('Saving rule data: ' + JSON.stringify(ruleData));

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

    Logger.log('Rule array to save: ' + JSON.stringify(ruleArray));

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
      // Add new rule to the next available row (not at the very end)
      Logger.log('About to add row to Rules sheet');

      // Find the actual last row with data in column A (ID column)
      const values = rulesSheet.getRange('A:A').getValues();
      let lastDataRow = 1; // Start from row 1 (header)

      for (let i = values.length - 1; i >= 0; i--) {
        if (values[i][0] && values[i][0] !== '') {
          lastDataRow = i + 1;
          break;
        }
      }

      Logger.log('Last data row found: ' + lastDataRow);
      const nextRow = lastDataRow + 1;
      Logger.log('Inserting rule at row: ' + nextRow);

      // Insert the rule at the next row after the last data row
      rulesSheet.getRange(nextRow, 1, 1, ruleArray.length).setValues([ruleArray]);
      Logger.log('Successfully added rule at row: ' + nextRow);

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
    const amount =
      (transactionData &&
        (transactionData.deposit || transactionData.withdrawal)) ||
      100;
    const isCredit = transactionData && transactionData.deposit > 0;

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
    const narration = transactionData.narration || "";
    const amount = transactionData.deposit || transactionData.withdrawal || 0;
    const userContext = transactionData.userContext || "";

    const conditions = ruleData.condition.includes(" AND ")
      ? ruleData.condition.split(" AND ")
      : [ruleData.condition];
    const patterns = ruleData.pattern.includes(";")
      ? ruleData.pattern.split(";")
      : [ruleData.pattern];

    for (let i = 0; i < conditions.length; i++) {
      const condition = conditions[i].trim();
      const pattern = patterns[i] ? patterns[i].trim() : "";

      if (!pattern) continue;

      let conditionMet = false;

      if (condition === "Narration CONTAINS") {
        conditionMet = narration.toLowerCase().includes(pattern.toLowerCase());
      } else if (condition === "Narration REGEX") {
        const regex = new RegExp(pattern, "i");
        conditionMet = regex.test(narration);
      } else if (condition === "Amount ==") {
        const targetAmount = parseFloat(pattern);
        conditionMet = Math.abs(amount - targetAmount) < 0.01;
      } else if (condition === "Amount >") {
        const targetAmount = parseFloat(pattern);
        conditionMet = amount > targetAmount;
      } else if (condition === "Amount <") {
        const targetAmount = parseFloat(pattern);
        conditionMet = amount < targetAmount;
      } else if (condition === "User_Context CONTAINS") {
        conditionMet = userContext
          .toLowerCase()
          .includes(pattern.toLowerCase());
      }

      if (!conditionMet) {
        return {
          matches: false,
          failedCondition: condition,
          failedPattern: pattern,
        };
      }
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

function formatLedgerCliEntry(
  date,
  payee,
  targetAccount,
  amount,
  fundingAccount,
  isCredit,
  tags
) {
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  const formattedDate = `${yyyy}/${mm}/${dd}`;
  const formattedAmount = `₹${amount.toFixed(2)}`;

  const tagLine = tags ? `    ;${tags}` : "";

  if (isCredit) {
    // Income: Money flows TO the funding account FROM the target account
    return `${formattedDate} ${payee}${
      tagLine ? "\n" + tagLine : ""
    }\n    ${fundingAccount}    ${formattedAmount}\n    ${targetAccount}`;
  } else {
    // Expense/Transfer: Money flows FROM the funding account TO the target account
    return `${formattedDate} ${payee}${
      tagLine ? "\n" + tagLine : ""
    }\n    ${targetAccount}    ${formattedAmount}\n    ${fundingAccount}`;
  }
}

// Helper function to get column indices (used by createRuleFromSelection)
function getColumnIndices(headers) {
  // Helper function to find column index with flexible matching
  function findColumnIndex(searchTerms) {
    for (let term of searchTerms) {
      const index = headers.indexOf(term);
      if (index !== -1) return index;
    }
    return -1;
  }

  return {
    srNo: findColumnIndex(["Sr No", "Sr. No", "SrNo"]),
    date: findColumnIndex(["Transaction Date", "Date", "Txn Date"]),
    narration: findColumnIndex(["Narration", "Description", "Details"]),
    withdrawal: findColumnIndex(["Withdrawal", "Withdrawal Amount", "Debit", "Debit Amount"]),
    deposit: findColumnIndex(["Deposit", "Deposit Amount", "Credit", "Credit Amount"]),
    balance: findColumnIndex(["Balance", "Available Balance"]),
    userContext: findColumnIndex(["User Context", "UserContext", "Context"]),
    tags: findColumnIndex(["Tags", "Tag"]),
    confidence: findColumnIndex(["LLM Confidence", "Confidence"]),
    finalEntry: findColumnIndex(["Final Entry", "FinalEntry", "Entry"]),
  };
}
