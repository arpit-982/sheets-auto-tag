function callLLM(prompt, temperature = 0.3, maxTokens = 150) {
  const settings = getSettings();
  
  if (!settings.provider || !settings.apiKey) {
    throw new Error('LLM provider not configured. Please check Settings.');
  }
  
  switch (settings.provider) {
    case 'requesty':
      return callRequestyAPI(prompt, settings, temperature, maxTokens);
    case 'gemini':
      return callGeminiAPI(prompt, settings, temperature, maxTokens);
    default:
      throw new Error('Unsupported LLM provider: ' + settings.provider);
  }
}

function callRequestyAPI(prompt, settings, temperature, maxTokens) {
  const baseUrl = settings.customBaseUrl || 'https://router.requesty.ai/v1';
  
  const payload = {
    model: settings.model,
    messages: [
      {
        role: "user",
        content: prompt
      }
    ],
    temperature: temperature,
    max_tokens: maxTokens
  };
  
  const options = {
    method: 'POST',
    headers: {
      'Authorization': 'Bearer ' + settings.apiKey,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload)
  };
  
  try {
    const response = UrlFetchApp.fetch(baseUrl + '/chat/completions', options);
    const responseText = response.getContentText();
    
    console.log('Raw API Response Text:', responseText);
    
    const data = JSON.parse(responseText);
    console.log('Parsed API Data:', JSON.stringify(data, null, 2));
    
    if (data.error) {
      throw new Error(data.error.message || 'API Error');
    }
    
    // Check if choices exist
    if (!data.choices || !Array.isArray(data.choices) || data.choices.length === 0) {
      console.log('No choices in response');
      throw new Error('No choices returned by API');
    }
    
    const choice = data.choices[0];
    console.log('First choice:', JSON.stringify(choice, null, 2));
    
    if (!choice.message) {
      throw new Error('No message in choice');
    }
    
    const content = choice.message.content;
    console.log('Message content:', JSON.stringify(content));
    
    if (content === null || content === undefined) {
      throw new Error('Content is null/undefined');
    }
    
    if (content === '') {
      throw new Error('Content is empty string');
    }
    
    return content.trim();
    
  } catch (error) {
    console.log('Requesty API Error:', error.message);
    throw new Error('Requesty API call failed: ' + error.message);
  }
}

function callGeminiAPI(prompt, settings, temperature, maxTokens) {
  const baseUrl = settings.customBaseUrl || 'https://generativelanguage.googleapis.com/v1beta';
  
  const payload = {
    contents: [
      {
        parts: [
          {
            text: prompt
          }
        ]
      }
    ],
    generationConfig: {
      temperature: temperature,
      maxOutputTokens: maxTokens
    }
  };
  
  const options = {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload)
  };
  
  try {
    const response = UrlFetchApp.fetch(`${baseUrl}/models/${settings.model}:generateContent?key=${settings.apiKey}`, options);
    const data = JSON.parse(response.getContentText());
    
    if (data.error) {
      throw new Error(data.error.message || 'API Error');
    }
    
    return data.candidates[0].content.parts[0].text.trim();
  } catch (error) {
    throw new Error('Gemini API call failed: ' + error.message);
  }
}