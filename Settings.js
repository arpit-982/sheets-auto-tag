function openSettings() {
  const html = HtmlService.createHtmlOutput(
    `
    <style>
      body { font-family: Google Sans, Arial, sans-serif; padding: 20px; }
      .provider-section { margin-bottom: 20px; padding: 15px; border: 1px solid #e0e0e0; border-radius: 8px; }
      .provider-header { font-weight: bold; margin-bottom: 10px; color: #1a73e8; }
      input[type="text"], input[type="password"], select { 
        width: 100%; padding: 8px; margin: 5px 0; border: 1px solid #ccc; border-radius: 4px; 
      }
      button { 
        background: #1a73e8; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; margin: 5px; 
      }
      button:hover { background: #1557b0; }
      .status { padding: 10px; margin: 10px 0; border-radius: 4px; }
      .success { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
      .error { background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
      .info { background: #d1ecf1; color: #0c5460; border: 1px solid #bee5eb; }
      .loading { opacity: 0.6; }
    </style>
    
    <h2>🔧 Ledger Tools Settings</h2>
    
    <div class="provider-section">
      <div class="provider-header">🤖 LLM Provider</div>
      <select id="provider" onchange="updateProvider()">
        <option value="">Select Provider</option>
        <option value="requesty">Requesty AI</option>
        <option value="gemini">Google Gemini API</option>
      </select>
      
      <div id="providerConfig" style="margin-top: 15px; display: none;">
        <div style="margin-top: 10px;">
          <label for="apiKey">API Key:</label>
          <input type="password" id="apiKey" placeholder="Enter your API key" onchange="onApiKeyChange()">
        </div>
        
        <div id="modelSection" style="display: none; margin-top: 10px;">
          <label for="model">Model:</label>
          <select id="model">
            <option value="">Loading models...</option>
          </select>
        </div>
        
        <div id="customUrlSection" style="margin-top: 10px; display: none;">
          <label for="customBaseUrl">Custom Base URL (optional):</label>
          <input type="text" id="customBaseUrl" placeholder="Leave empty for default">
        </div>
      </div>
    </div>
    
    <div style="margin-top: 20px;">
      <button onclick="saveSettings()">💾 Save Settings</button>
      <button onclick="testConnection()">🧪 Test Connection</button>
    </div>
    
    <div id="status"></div>
    
    <script>
    let isLoadingModels = false;
      const providers = {
        requesty: {
          name: "Requesty AI",
          models: [], // Will be fetched dynamically
          baseUrl: "https://router.requesty.ai/v1"
        },
        gemini: {
          name: "Google Gemini", 
          models: ["gemini-pro", "gemini-pro-vision", "gemini-1.5-flash", "gemini-1.5-pro"],
          baseUrl: "https://generativelanguage.googleapis.com/v1beta"
        }
      };
      
      function updateProvider() {
  const provider = document.getElementById('provider').value;
  const configDiv = document.getElementById('providerConfig');
  const modelSection = document.getElementById('modelSection');
  const customUrlSection = document.getElementById('customUrlSection');
  
  if (provider) {
    configDiv.style.display = 'block';
    customUrlSection.style.display = 'block';
    
    if (provider === 'gemini') {
      // Show model section but don't populate until API key is provided
      modelSection.style.display = 'block';
      const modelSelect = document.getElementById('model');
      modelSelect.innerHTML = '<option value="">Enter API key to load models...</option>';
    } else if (provider === 'requesty') {
      modelSection.style.display = 'block';
      const modelSelect = document.getElementById('model');
      modelSelect.innerHTML = '<option value="">Enter API key to load models...</option>';
    }
  } else {
    configDiv.style.display = 'none';
  }
}

function onApiKeyChange() {
  const provider = document.getElementById('provider').value;
  const apiKey = document.getElementById('apiKey').value;
  
  if (apiKey.trim() && document.hasFocus()) {
    if (provider === 'requesty') {
      fetchModelsForRequesty(apiKey.trim(), null);
    } else if (provider === 'gemini') {
      fetchModelsForGemini(apiKey.trim(), null);
    }
  }
}

function fetchModelsForGemini(apiKey, selectedModel) {
  if (isLoadingModels) return;
  
  isLoadingModels = true;
  showStatus('Loading Gemini models...', 'info');
  
  google.script.run
    .withSuccessHandler(function(models) {
      isLoadingModels = false;
      populateModels(models);
      document.getElementById('modelSection').style.display = 'block';
      
      if (selectedModel) {
        document.getElementById('model').value = selectedModel;
      }
      showStatus('Gemini models loaded successfully!', 'success');
    })
    .withFailureHandler(function(error) {
      isLoadingModels = false;
      showStatus('Failed to load Gemini models: ' + error, 'error');
    })
    .fetchGeminiModels(apiKey);
}
    
      
      function onModelsFetched(models) {
        if (models && models.length > 0) {
          populateModels(models);
          document.getElementById('modelSection').style.display = 'block';
          showStatus('✅ Models loaded successfully!', 'success');
        } else {
          showStatus('⚠️ No models found. Please check your API key.', 'error');
        }
      }
      
      function onModelsFetchError(error) {
        showStatus('❌ Failed to fetch models: ' + error, 'error');
        document.getElementById('modelSection').style.display = 'none';
      }
      
      function populateModels(models) {
        const modelSelect = document.getElementById('model');
        modelSelect.innerHTML = '';
        
        models.forEach(model => {
          const option = document.createElement('option');
          option.value = model;
          option.textContent = model;
          modelSelect.appendChild(option);
        });
      }
      
function loadSavedSettings(settings) {
  if (settings && settings.provider) {
    document.getElementById('provider').value = settings.provider;
    document.getElementById('apiKey').value = settings.apiKey || '';
    document.getElementById('customBaseUrl').value = settings.customBaseUrl || '';
    
    updateProvider();
    
    setTimeout(() => {
      if (settings.provider === 'gemini' && settings.apiKey) {
        fetchModelsForGemini(settings.apiKey, settings.model);
      } else if (settings.provider === 'requesty' && settings.apiKey) {
        fetchModelsForRequesty(settings.apiKey, settings.model);
      }
    }, 100);
  } else {
    document.getElementById('provider').value = 'gemini';
    updateProvider();
  }
}

function fetchModelsForRequesty(apiKey, selectedModel) {
  if (isLoadingModels) {
    return; // Already loading, don't start another fetch
  }
  
  isLoadingModels = true;
  showStatus('Loading models...', 'info');
  
  google.script.run
    .withSuccessHandler(function(models) {
      isLoadingModels = false;
      populateModels(models);
      document.getElementById('modelSection').style.display = 'block';
      
      if (selectedModel) {
        document.getElementById('model').value = selectedModel;
      }
      showStatus('Models loaded successfully!', 'success');
    })
    .withFailureHandler(function(error) {
      isLoadingModels = false;
      showStatus('Failed to load models: ' + error, 'error');
      const modelSelect = document.getElementById('model');
      modelSelect.innerHTML = '<option value="' + (selectedModel || '') + '">' + (selectedModel || 'Failed to load models') + '</option>';
      document.getElementById('modelSection').style.display = 'block';
    })
    .fetchRequestyModels(apiKey);
}
      
      function saveSettings() {
        const settings = {
          provider: document.getElementById('provider').value,
          model: document.getElementById('model').value,
          apiKey: document.getElementById('apiKey').value,
          customBaseUrl: document.getElementById('customBaseUrl').value
        };
        
        if (!settings.provider || !settings.model || !settings.apiKey) {
          showStatus('Please fill in all required fields', 'error');
          return;
        }
        
        google.script.run
          .withSuccessHandler(() => showStatus('Settings saved successfully!', 'success'))
          .withFailureHandler(error => showStatus('Error saving settings: ' + error, 'error'))
          .saveSettings(settings);
      }
      
      function testConnection() {
        showStatus('Testing connection...', 'info');
        google.script.run
          .withSuccessHandler(result => showStatus('✅ Connection successful!', 'success'))
          .withFailureHandler(error => showStatus('❌ Connection failed: ' + error, 'error'))
          .testLLMConnection();
      }
      
      function showStatus(message, type) {
        const statusDiv = document.getElementById('status');
        statusDiv.innerHTML = '<div class="status ' + type + '">' + message + '</div>';
      }
      
      window.onload = function() {
  google.script.run.withSuccessHandler(loadSavedSettings).getSettings();
};
    </script>
  `
  )
    .setWidth(500)
    .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(html, "⚙️ Settings");
}

function fetchRequestyModels(apiKey) {
  const baseUrl = "https://router.requesty.ai/v1";

  const options = {
    method: "GET",
    headers: {
      Authorization: "Bearer " + apiKey,
      "Content-Type": "application/json",
    },
  };

  try {
    const response = UrlFetchApp.fetch(baseUrl + "/models", options);
    const data = JSON.parse(response.getContentText());

    if (data.error) {
      throw new Error(data.error.message || "API Error");
    }

    // OpenAI-compatible format: {"data": [{"id": "model-name"}, ...]}
    if (data.data && Array.isArray(data.data)) {
      return data.data.map((model) => model.id).sort();
    }

    return [];
  } catch (error) {
    throw new Error("Failed to fetch models: " + error.message);
  }
}

function getSettings() {
  const properties = PropertiesService.getScriptProperties();
  return {
    provider: properties.getProperty("LLM_PROVIDER"),
    model: properties.getProperty("LLM_MODEL"),
    apiKey: properties.getProperty("LLM_API_KEY"),
    customBaseUrl: properties.getProperty("LLM_BASE_URL"),
  };
}

function saveSettings(settings) {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperties({
    LLM_PROVIDER: settings.provider,
    LLM_MODEL: settings.model,
    LLM_API_KEY: settings.apiKey,
    LLM_BASE_URL: settings.customBaseUrl || "",
  });
}

function testLLMConnection() {
  const settings = getSettings();
  if (!settings.provider || !settings.apiKey) {
    throw new Error("Please configure LLM settings first");
  }

  // Simple test call
  const result = callLLM("Test message", 0.1, 10);
  return result;
}

function fetchGeminiModels(apiKey) {
  const baseUrl = "https://generativelanguage.googleapis.com/v1beta";

  const options = {
    method: "GET",
    headers: {
      "Content-Type": "application/json",
    },
  };

  try {
    const response = UrlFetchApp.fetch(
      `${baseUrl}/models?key=${apiKey}`,
      options
    );
    const data = JSON.parse(response.getContentText());

    if (data.error) {
      throw new Error(data.error.message || "API Error");
    }

    // Filter models that support generateContent
    if (data.models && Array.isArray(data.models)) {
      const generateContentModels = data.models
        .filter(
          (model) =>
            model.supportedGenerationMethods &&
            model.supportedGenerationMethods.includes("generateContent")
        )
        .map((model) => model.name.replace("models/", "")) // Remove 'models/' prefix
        .sort();

      return generateContentModels;
    }

    return [];
  } catch (error) {
    throw new Error("Failed to fetch Gemini models: " + error.message);
  }
}
