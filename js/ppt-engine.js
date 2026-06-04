const PPTEngine = (() => {
  'use strict';

  const runtimeConfig = window.OFFICE_ASSISTANT_CONFIG || {};
  const isLocalPage = ['localhost', '127.0.0.1'].includes(window.location.hostname);
  const DEFAULT_SERVER_URL = runtimeConfig.defaultServerUrl || (isLocalPage ? 'http://localhost:3456' : '');
  const SYSTEM_PROMPT = 'You are a PowerPoint content assistant. Return clear, presentation-ready content.';
  const PROVIDERS = {
    gemini: { name: 'Gemini', type: 'gemini', endpoint: 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent', model: 'gemini-2.5-flash' },
    deepseek: { name: 'DeepSeek', type: 'openai', endpoint: 'https://api.deepseek.com/v1/chat/completions', model: 'deepseek-chat' },
    groq: { name: 'Groq', type: 'openai', endpoint: 'https://api.groq.com/openai/v1/chat/completions', model: 'llama-3.3-70b-versatile' },
    openai: { name: 'OpenAI', type: 'openai', endpoint: 'https://api.openai.com/v1/chat/completions', model: 'gpt-4o' },
    claude: { name: 'Claude', type: 'claude', endpoint: 'https://api.anthropic.com/v1/messages', model: 'claude-sonnet-4-20250514' },
    ollama: { name: 'Ollama', type: 'ollama', endpoint: 'http://localhost:11434', model: 'qwen2.5:7b' },
    custom: { name: 'Custom', type: 'openai', endpoint: '', model: '' },
  };
  const STYLES = {
    business: 'Business professional',
    academic: 'Academic rigorous',
    creative: 'Creative lively',
    minimal: 'Minimal modern',
    tech: 'Technology forward',
  };

  const store = {
    get(key, fallback = '') {
      try {
        const value = localStorage.getItem(key);
        return value == null ? fallback : value;
      } catch (_error) {
        return fallback;
      }
    },
    set(key, value) {
      try {
        localStorage.setItem(key, value || '');
      } catch (_error) {}
    },
    remove(key) {
      try {
        localStorage.removeItem(key);
      } catch (_error) {}
    },
  };

  let settings = { provider: 'gemini', endpoint: '', model: '', apikey: '', style: 'business' };
  let officeReady = false;
  let host = 'Browser';
  let chatHistory = [];
  let serverUrl = DEFAULT_SERVER_URL;
  let authToken = '';
  let mode = 'direct';

  function normalizeSettings(nextSettings = {}) {
    return {
      provider: nextSettings.provider || settings.provider || 'gemini',
      endpoint: nextSettings.endpoint || '',
      model: nextSettings.model || '',
      apikey: nextSettings.apikey || nextSettings.apiKey || '',
      style: nextSettings.style || settings.style || 'business',
    };
  }

  async function init(options = {}) {
    settings = normalizeSettings({
      provider: store.get('ppt_provider', options.provider || 'gemini'),
      endpoint: store.get('ppt_endpoint', options.endpoint || ''),
      model: store.get('ppt_model', options.model || ''),
      apikey: store.get('ppt_apikey', options.apikey || ''),
      style: store.get('ppt_style', options.style || 'business'),
    });
    serverUrl = store.get('ppt_server_url', options.serverUrl || DEFAULT_SERVER_URL).replace(/\/+$/, '');
    authToken = store.get('ppt_auth_token', options.token || '');
    mode = store.get('ppt_mode', options.mode || 'direct');

    try {
      chatHistory = JSON.parse(store.get('ppt_chat', '[]'));
    } catch (_error) {
      chatHistory = [];
    }

    const payload = await new Promise((resolve) => {
      const done = (nextHost) => {
        host = nextHost || 'Browser';
        officeReady = host !== 'Browser';
        resolve({ host, settings: getSettings(), mode: getMode() });
      };

      if (typeof Office !== 'undefined' && Office.onReady) {
        Office.onReady((info) => done(info.host ? String(info.host) : 'Browser'));
      } else {
        done('Browser');
      }
    });

    if (typeof options.onReady === 'function') {
      options.onReady(payload);
    }

    return payload;
  }

  function getSettings() {
    return { ...settings };
  }

  function saveSettings(nextSettings = {}) {
    settings = normalizeSettings({ ...settings, ...nextSettings });
    store.set('ppt_provider', settings.provider);
    store.set('ppt_endpoint', settings.endpoint);
    store.set('ppt_model', settings.model);
    store.set('ppt_apikey', settings.apikey);
    store.set('ppt_style', settings.style);
    return getSettings();
  }

  function getProviders() {
    return Object.entries(PROVIDERS).map(([id, provider]) => ({ id, ...provider }));
  }

  function getStyles() {
    return Object.entries(STYLES).map(([id, label]) => ({ id, label }));
  }

  function useServer(nextServerUrl, token) {
    serverUrl = (nextServerUrl || DEFAULT_SERVER_URL).replace(/\/+$/, '');
    authToken = token || authToken || store.get('ppt_auth_token', '');
    mode = 'proxy';
    store.set('ppt_server_url', serverUrl);
    store.set('ppt_mode', mode);
    if (authToken) {
      store.set('ppt_auth_token', authToken);
    }
    return getMode();
  }

  function useDirect() {
    mode = 'direct';
    store.set('ppt_mode', mode);
    return getMode();
  }

  function getMode() {
    return { mode, serverUrl, authenticated: Boolean(authToken) };
  }

  function logout() {
    authToken = '';
    store.remove('ppt_auth_token');
    store.remove('ppt_username');
    return getMode();
  }

  async function requestJson(url, options = {}) {
    const response = await fetch(url, options);
    const raw = await response.text();
    let data = null;

    if (raw) {
      try {
        data = JSON.parse(raw);
      } catch (_error) {
        data = { raw };
      }
    }

    if (!response.ok) {
      const message = data?.error?.message || data?.message || data?.raw || `Request failed (${response.status})`;
      throw new Error(message);
    }

    return data;
  }

  async function callServer(endpoint, body, method = 'POST') {
    if (!serverUrl) {
      throw new Error('Server mode needs a public HTTPS server URL.');
    }

    const headers = { 'Content-Type': 'application/json' };
    if (authToken) {
      headers.Authorization = `Bearer ${authToken}`;
    }

    const response = await requestJson(`${serverUrl}${endpoint}`, {
      method,
      headers,
      ...(body !== undefined ? { body: JSON.stringify(body) } : {}),
    });

    if (response?.success === false) {
      throw new Error(response?.error?.message || 'Server request failed');
    }

    return response?.data ?? response;
  }

  async function callServerGet(endpoint) {
    if (!serverUrl) {
      throw new Error('Server mode needs a public HTTPS server URL.');
    }

    const headers = authToken ? { Authorization: `Bearer ${authToken}` } : {};
    const response = await requestJson(`${serverUrl}${endpoint}`, { headers });
    if (response?.success === false) {
      throw new Error(response?.error?.message || 'Server request failed');
    }
    return response?.data ?? response;
  }

  async function register(username, password, displayName) {
    const data = await callServer('/api/auth/register', { username, password, displayName });
    authToken = data.token || '';
    if (authToken) {
      store.set('ppt_auth_token', authToken);
    }
    store.set('ppt_username', username);
    useServer(serverUrl, authToken);
    return data;
  }

  async function login(username, password) {
    const data = await callServer('/api/auth/login', { username, password });
    authToken = data.token || '';
    if (authToken) {
      store.set('ppt_auth_token', authToken);
    }
    store.set('ppt_username', username);
    useServer(serverUrl, authToken);
    return data;
  }

  function getUserInfo() {
    if (!authToken) {
      throw new Error('Please log in first');
    }
    return callServerGet('/api/auth/me');
  }

  async function loadServerSettings() {
    if (!authToken) {
      throw new Error('Please log in first');
    }
    const data = await callServerGet('/api/settings');
    saveSettings(data || {});
    return getSettings();
  }

  async function saveServerSettings(nextSettings) {
    if (!authToken) {
      throw new Error('Please log in first');
    }
    const payload = {
      ...nextSettings,
      apiKey: nextSettings.apikey || nextSettings.apiKey || '',
    };
    const result = await callServer('/api/settings', payload, 'PUT');
    saveSettings(nextSettings);
    return result;
  }

  function isOfficeReady() {
    return officeReady;
  }

  function getHost() {
    return host;
  }

  function getSelection() {
    return new Promise((resolve) => {
      if (!officeReady || typeof Office === 'undefined') {
        resolve('');
        return;
      }

      Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
        resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : '');
      });
    });
  }

  function writeToOffice(text) {
    return new Promise((resolve, reject) => {
      if (!officeReady || typeof Office === 'undefined') {
        reject(new Error('Office host is not available'));
        return;
      }

      Office.context.document.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(new Error(result.error?.message || 'Failed to write back to Office'));
        }
      });
    });
  }

  function persistChatHistory() {
    store.set('ppt_chat', JSON.stringify(chatHistory));
  }

  function getChatHistory() {
    return [...chatHistory];
  }

  function addChatMessage(role, content) {
    chatHistory.push({ id: `${Date.now()}-${Math.random()}`, role, content, createdAt: new Date().toISOString() });
    persistChatHistory();
    return getChatHistory();
  }

  function clearChatHistory() {
    chatHistory = [];
    persistChatHistory();
  }

  function exportChatHistory() {
    return chatHistory.map((item) => `${item.role}: ${item.content}`).join('\n\n');
  }

  async function loadChatHistory() {
    if (mode === 'proxy' && authToken) {
      const data = await callServerGet('/api/chat');
      chatHistory = data.messages || [];
      persistChatHistory();
      return getChatHistory();
    }
    return getChatHistory();
  }

  async function clearServerChat() {
    if (!authToken) {
      throw new Error('Please log in first');
    }
    await callServer('/api/chat', undefined, 'DELETE');
    chatHistory = [];
    persistChatHistory();
  }

  async function exportServerChat(format = 'markdown') {
    if (!authToken) {
      throw new Error('Please log in first');
    }
    return callServerGet(`/api/chat/export?format=${encodeURIComponent(format)}`);
  }

  function getDocuments() {
    if (!authToken) {
      throw new Error('Please log in first');
    }
    return callServerGet('/api/documents');
  }

  function createDocument(title, type, content, summary, tags = []) {
    if (!authToken) {
      throw new Error('Please log in first');
    }
    return callServer('/api/documents', { title, type, content, summary, tags, source: 'assistant' });
  }

  function getDocument(id) {
    if (!authToken) {
      throw new Error('Please log in first');
    }
    return callServerGet(`/api/documents/${id}`);
  }

  function updateDocument(id, payload) {
    if (!authToken) {
      throw new Error('Please log in first');
    }
    return callServer(`/api/documents/${id}`, payload, 'PUT');
  }

  function deleteDocument(id) {
    if (!authToken) {
      throw new Error('Please log in first');
    }
    return callServer(`/api/documents/${id}`, undefined, 'DELETE');
  }

  function buildSystemPrompt() {
    const styleLabel = STYLES[settings.style] || STYLES.business;
    return `${SYSTEM_PROMPT} Preferred style: ${styleLabel}.`;
  }

  async function callOpenAICompatible(provider, prompt) {
    const endpoint = settings.endpoint || provider.endpoint;
    const response = await requestJson(endpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${settings.apikey}`,
      },
      body: JSON.stringify({
        model: settings.model || provider.model,
        messages: [
          { role: 'system', content: buildSystemPrompt() },
          { role: 'user', content: prompt },
        ],
        temperature: 0.7,
      }),
    });

    return response.choices?.[0]?.message?.content || '';
  }

  async function callGemini(provider, prompt) {
    const endpoint = (settings.endpoint || provider.endpoint).replace(/\/+$/, '');
    const response = await requestJson(`${endpoint}?key=${encodeURIComponent(settings.apikey)}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        systemInstruction: { parts: [{ text: buildSystemPrompt() }] },
      }),
    });

    return response.candidates?.[0]?.content?.parts?.[0]?.text || '';
  }

  async function callClaude(provider, prompt) {
    const endpoint = settings.endpoint || provider.endpoint;
    const response = await requestJson(endpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': settings.apikey,
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify({
        model: settings.model || provider.model,
        system: buildSystemPrompt(),
        max_tokens: 4096,
        messages: [{ role: 'user', content: prompt }],
      }),
    });

    return response.content?.[0]?.text || '';
  }

  async function callOllama(provider, prompt) {
    const baseUrl = (settings.endpoint || provider.endpoint).replace(/\/+$/, '');
    const response = await requestJson(`${baseUrl}/api/generate`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        model: settings.model || provider.model,
        prompt: `${buildSystemPrompt()}\n\n${prompt}`,
        stream: false,
      }),
    });

    return response.response || '';
  }

  async function callDirectAI(prompt) {
    const provider = PROVIDERS[settings.provider] || PROVIDERS.gemini;

    if (provider.type !== 'ollama' && !settings.apikey) {
      throw new Error(`Please provide an API key for ${provider.name}`);
    }

    switch (provider.type) {
      case 'gemini':
        return callGemini(provider, prompt);
      case 'claude':
        return callClaude(provider, prompt);
      case 'ollama':
        return callOllama(provider, prompt);
      default:
        return callOpenAICompatible(provider, prompt);
    }
  }

  async function runAiAction({ endpoint, prompt, body }) {
    const result = mode === 'proxy'
      ? await callServer(endpoint, {
          ...body,
          provider: settings.provider,
          model: settings.model,
          apiKey: settings.apikey,
          style: settings.style,
        })
      : { text: await callDirectAI(prompt) };

    const text = result.text || '';
    if (text) {
      addChatMessage('assistant', text);
    }
    return text;
  }

  function checkConnection() {
    if (mode === 'proxy') {
      return callServer('/api/ai/check-connection', { provider: settings.provider }).then((data) => Boolean(data.connected));
    }

    if (settings.provider === 'ollama') {
      const baseUrl = (settings.endpoint || PROVIDERS.ollama.endpoint).replace(/\/+$/, '');
      return fetch(`${baseUrl}/api/tags`, { signal: AbortSignal.timeout(5000) }).then((response) => response.ok).catch(() => false);
    }

    return callDirectAI('Reply with OK only.').then(() => true).catch(() => false);
  }

  function generateOutline(topic) {
    if (!topic) throw new Error('Please enter a PPT topic');
    addChatMessage('user', topic);
    return runAiAction({ endpoint: '/api/ai/generate-outline', prompt: `Create a PPT outline for: ${topic}`, body: { topic } });
  }

  function generateFullContent(topic) {
    if (!topic) throw new Error('Please enter a PPT topic');
    addChatMessage('user', topic);
    return runAiAction({ endpoint: '/api/ai/generate-content', prompt: `Create full PPT content for: ${topic}`, body: { topic } });
  }

  function generateSlideContent(title) {
    if (!title) throw new Error('Please enter a slide topic');
    addChatMessage('user', title);
    return runAiAction({ endpoint: '/api/ai/generate-slide', prompt: `Create one slide for: ${title}`, body: { title } });
  }

  function polishText(text) {
    if (!text) throw new Error('Please enter text to polish');
    addChatMessage('user', text);
    return runAiAction({ endpoint: '/api/ai/polish', prompt: `Polish this PPT text:\n${text}`, body: { text } });
  }

  function expandContent(text) {
    if (!text) throw new Error('Please enter text to expand');
    addChatMessage('user', text);
    return runAiAction({ endpoint: '/api/ai/expand', prompt: `Expand this PPT content:\n${text}`, body: { text } });
  }

  function condenseContent(text) {
    if (!text) throw new Error('Please enter text to condense');
    addChatMessage('user', text);
    return runAiAction({ endpoint: '/api/ai/condense', prompt: `Condense this PPT content:\n${text}`, body: { text } });
  }

  function translate(text) {
    if (!text) throw new Error('Please enter text to translate');
    addChatMessage('user', text);
    return runAiAction({ endpoint: '/api/ai/translate', prompt: `Translate this text between Chinese and English:\n${text}`, body: { text } });
  }

  function generateSpeakerNotes(text) {
    if (!text) throw new Error('Please enter slide content first');
    addChatMessage('user', text);
    return runAiAction({ endpoint: '/api/ai/speaker-notes', prompt: `Write speaker notes for:\n${text}`, body: { text } });
  }

  function getDesignTips(text) {
    if (!text) throw new Error('Please enter slide content first');
    addChatMessage('user', text);
    return runAiAction({ endpoint: '/api/ai/design-tips', prompt: `Give design tips for:\n${text}`, body: { text } });
  }

  function chat(message, context = '') {
    if (!message) throw new Error('Please enter a message');
    addChatMessage('user', message);
    return runAiAction({ endpoint: '/api/ai/chat', prompt: context ? `Context:\n${context}\n\n${message}` : message, body: { message, context } });
  }

  return {
    init,
    useServer,
    useDirect,
    getMode,
    register,
    login,
    logout,
    getUserInfo,
    getSettings,
    getProviders,
    getStyles,
    saveSettings,
    checkConnection,
    loadServerSettings,
    saveServerSettings,
    isOfficeReady,
    getHost,
    getSelection,
    writeToOffice,
    getChatHistory,
    addChatMessage,
    clearChatHistory,
    exportChatHistory,
    loadChatHistory,
    clearServerChat,
    exportServerChat,
    getDocuments,
    createDocument,
    getDocument,
    updateDocument,
    deleteDocument,
    generateOutline,
    generateFullContent,
    generateSlideContent,
    polishText,
    expandContent,
    condenseContent,
    translate,
    generateSpeakerNotes,
    getDesignTips,
    chat,
  };
})();

if (typeof module !== 'undefined' && module.exports) {
  module.exports = PPTEngine;
}
