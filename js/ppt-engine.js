const PPTEngine = (() => {
  'use strict';

  const browserWindow = typeof window !== 'undefined' ? window : { location: { hostname: '' } };
  const runtimeConfig = browserWindow.OFFICE_ASSISTANT_CONFIG || {};
  const isLocalPage = ['localhost', '127.0.0.1'].includes(browserWindow.location.hostname);
  const DEFAULT_SERVER_URL = runtimeConfig.defaultServerUrl || (isLocalPage ? 'http://localhost:3456' : '');
  const SYSTEM_PROMPT = 'You are a PowerPoint content assistant. Return clear, presentation-ready content.';
  const PROVIDERS = {
    local: { name: '免配置模板', type: 'local', endpoint: '', model: 'local-template' },
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

  let settings = { provider: 'local', endpoint: '', model: '', apikey: '', style: 'business' };
  let officeReady = false;
  let host = 'Browser';
  let chatHistory = [];
  let serverUrl = DEFAULT_SERVER_URL;
  let authToken = '';
  let mode = 'direct';

  function normalizeSettings(nextSettings = {}) {
    return {
      provider: nextSettings.provider || settings.provider || 'local',
      endpoint: nextSettings.endpoint || '',
      model: nextSettings.model || '',
      apikey: nextSettings.apikey || nextSettings.apiKey || '',
      style: nextSettings.style || settings.style || 'business',
    };
  }

  async function init(options = {}) {
    settings = normalizeSettings({
      provider: store.get('ppt_provider', options.provider || 'local'),
      endpoint: store.get('ppt_endpoint', options.endpoint || ''),
      model: store.get('ppt_model', options.model || ''),
      apikey: store.get('ppt_apikey', options.apikey || ''),
      style: store.get('ppt_style', options.style || 'business'),
    });
    if (settings.provider === 'gemini' && !settings.apikey && !settings.endpoint) {
      settings.provider = 'local';
    }

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

  function canUseLocalProxy(provider) {
    return Boolean(serverUrl) && provider.type !== 'local' && provider.type !== 'ollama';
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
      throw new Error('服务端模式需要先填写服务端地址，例如 http://localhost:3456。');
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
      throw new Error('服务端模式需要先填写服务端地址，例如 http://localhost:3456。');
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

  function cleanInput(value, fallback = '这份演示') {
    const text = String(value || '').trim();
    return text || fallback;
  }

  function topicFromBody(body = {}) {
    return cleanInput(body.topic || body.title || body.message || body.text);
  }

  function styleHeading() {
    const label = STYLES[settings.style] || STYLES.business;
    return `风格：${label}`;
  }

  function makeBullets(topic, verbs) {
    return verbs.map((verb) => `- ${verb}${topic}的关键要点，给观众一个清晰、可执行的判断。`).join('\n');
  }

  function createLocalResponse(endpoint, body = {}) {
    const topic = topicFromBody(body);
    const text = cleanInput(body.text || body.context || body.message, topic);
    const commonClose = '\n\n下一步：你可以直接复制，或在 PPT 中选中一个文本框后点击“写回 Office”。';

    if (endpoint.includes('generate-outline')) {
      return [
        `《${topic}》PPT 大纲`,
        styleHeading(),
        '',
        'Slide 1：封面',
        `- 标题：${topic}`,
        '- 副标题：用一句话说明这次汇报要解决的问题',
        '',
        'Slide 2：为什么现在重要',
        makeBullets(topic, ['说明', '量化', '引出']),
        '',
        'Slide 3：现状与痛点',
        makeBullets(topic, ['梳理', '呈现', '定位']),
        '',
        'Slide 4：核心方案',
        makeBullets(topic, ['提出', '拆解', '强调']),
        '',
        'Slide 5：执行路径',
        '- 第一步：明确目标和负责人',
        '- 第二步：拆分里程碑和交付物',
        '- 第三步：用数据复盘并持续优化',
        '',
        'Slide 6：预期收益',
        '- 效率提升：减少重复沟通和返工',
        '- 质量提升：统一标准和判断依据',
        '- 决策提升：让风险、成本、收益更透明',
        '',
        'Slide 7：总结与行动',
        '- 重申核心结论',
        '- 给出下一步行动',
        '- 明确需要观众支持的事项',
        commonClose,
      ].join('\n');
    }

    if (endpoint.includes('generate-content')) {
      return [
        `《${topic}》完整 PPT 内容草稿`,
        styleHeading(),
        '',
        '1. 封面',
        `标题：${topic}`,
        '备注：今天我们聚焦一个实际问题：如何把想法变成可执行、可衡量的方案。',
        '',
        '2. 背景',
        makeBullets(topic, ['介绍', '解释', '指出']),
        '备注：先让听众理解问题的背景和紧迫性。',
        '',
        '3. 痛点',
        makeBullets(topic, ['拆解', '举例', '总结']),
        '备注：痛点越具体，后面的方案越容易被接受。',
        '',
        '4. 方案',
        makeBullets(topic, ['设计', '连接', '落地']),
        '备注：方案要少而清楚，避免堆概念。',
        '',
        '5. 路线图',
        '- 0-2 周：确认目标、资料和评估指标',
        '- 3-6 周：完成第一版试点',
        '- 7-12 周：根据反馈扩大应用范围',
        '备注：用时间线降低不确定性。',
        '',
        '6. 总结',
        '- 现在的问题值得解决',
        '- 方案可以从小范围开始',
        '- 下一步需要明确负责人和时间表',
        commonClose,
      ].join('\n');
    }

    if (endpoint.includes('generate-slide')) {
      return [
        `单页标题：${topic}`,
        '',
        '核心信息',
        makeBullets(topic, ['定义', '突出', '证明', '落到']),
        '',
        '推荐视觉',
        '- 左侧放一句核心结论，右侧放 3 个关键数据或图标',
        '- 使用高对比标题和浅色信息卡片',
        '- 每页只表达一个主结论',
        commonClose,
      ].join('\n');
    }

    if (endpoint.includes('polish')) {
      return [`精修版本：`, '', text.replace(/\s+/g, ' '), '', '优化建议：', '- 把长句拆成短句', '- 每个要点保留一个动词和一个结果', '- 结尾补一个明确行动'].join('\n');
    }

    if (endpoint.includes('expand')) {
      return [`扩展版本：${text}`, '', '- 背景：补充这个观点为什么重要', '- 例子：加入一个具体场景或数据', '- 影响：说明它对效率、成本或体验的改变', '- 行动：给出下一步可执行动作'].join('\n');
    }

    if (endpoint.includes('condense')) {
      const compact = text.split(/[。.!！?\n]/).map((item) => item.trim()).filter(Boolean).slice(0, 3);
      return ['精简版本：', ...(compact.length ? compact : [text]).map((item) => `- ${item}`)].join('\n');
    }

    if (endpoint.includes('translate')) {
      return ['翻译辅助：', '', '当前为免配置模板模式，无法做真正语义翻译。', '建议把要翻译文本复制到已配置 API Key 的 AI 模式，或使用服务端模式。', '', `原文：${text}`].join('\n');
    }

    if (endpoint.includes('speaker-notes')) {
      return ['演讲备注：', `开场：这一页主要讲 ${topic}。`, '展开：先说明背景，再解释关键点，最后落到行动建议。', '收束：请观众记住一个结论，并确认下一步。'].join('\n');
    }

    if (endpoint.includes('design-tips')) {
      return ['设计建议：', '- 一页只保留一个核心结论', '- 标题使用结论句，不要只写名词', '- 正文控制在 3-5 条要点', '- 用图标、流程图或数字卡片替代大段文字', '- 重要信息使用同一种强调色，避免页面失焦'].join('\n');
    }

    return [`回复：${topic}`, '', '当前处于免配置模板模式。我可以生成结构化草稿、润色建议、备注和设计建议；如果配置 API Key 或服务端地址，就能升级为真正 AI 生成。'].join('\n');
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

    if (provider.type === 'local') {
      return createLocalResponse('/api/ai/chat', { message: prompt });
    }

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
    const provider = PROVIDERS[settings.provider] || PROVIDERS.local;
    const shouldUseProxy = mode === 'proxy' || canUseLocalProxy(provider);
    const result = shouldUseProxy
      ? await callServer(endpoint, {
          ...body,
          provider: settings.provider,
          model: settings.model,
          apiKey: settings.apikey,
          style: settings.style,
        })
      : settings.provider === 'local'
        ? { text: createLocalResponse(endpoint, body) }
        : { text: await callDirectAI(prompt) };

    const text = result.text || '';
    if (text) {
      addChatMessage('assistant', text);
    }
    return text;
  }

  function checkConnection() {
    const provider = PROVIDERS[settings.provider] || PROVIDERS.local;

    if (mode === 'proxy' || canUseLocalProxy(provider)) {
      return callServer('/api/ai/check-connection', {
        provider: settings.provider,
        model: settings.model,
        apiKey: settings.apikey,
        style: settings.style,
      }).then((data) => Boolean(data.connected));
    }

    if (settings.provider === 'local') {
      return Promise.resolve(true);
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
