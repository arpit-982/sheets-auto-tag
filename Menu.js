function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Ledger Tools')
    .addItem('🔧 Initialize Plugin', 'initializePlugin')
    .addSeparator()
    .addItem('⚡ Process All Rows', 'processAllRows')
    .addItem('📍 Process Selected Rows', 'processSelectedRows') 
    .addItem('🎯 Process Current Row', 'processCurrentRow')
    .addSeparator()
    .addItem('➕ Create Rule from Selection', 'createRuleFromSelection')  // NEW
    .addItem('📝 Create New Rule', 'createNewRule')                      // NEW
    .addItem('✏️ Edit Rule', 'editExistingRule')                         // NEW
    .addSeparator()
    .addItem('📋 Export All to Clipboard', 'exportAllToClipboard')
    .addItem('📋 Export Selected to Clipboard', 'exportSelectedToClipboard')
    .addSeparator()
    .addItem('⚙️ Settings', 'openSettings')
    .addToUi();
}

function openSettings() {
  SpreadsheetApp.getUi().alert('Settings panel coming soon!\n\nFor now, use the other menu options to process transactions.');
}