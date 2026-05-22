/**
 * ═══════════════════════════════════════════════════
 *  PPT 智能助手 - 功能引擎 v2.0
 *  ─────────────────────────────────────────────
 *  前端只需调用此模块暴露的函数，无需关心内部实现
 * ═══════════════════════════════════════════════════
 */

const PPTEngine = (() => {
  'use strict';

  // ═══ 内部状态 ═══
  let _settings = {
    provider: 'gemini',
    endpoint: '',
    model: '',
    apikey: '',
    style: 'business'
  };
  let _officeReady = false;
  let _host = 'Browser';
  let _chatHistory = [];

  // ═══ Provider 配置表 ═══
  const PROVIDERS = {
    gemini:    { endpoint: 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-preview-09-2025:generateContent', model: 'gemini-2.5-flash', type: 'gemini' },
    deepseek:  { endpoint: 'https://api.deepseek.com/v1/chat/completions', model: 'deepseek-chat', type: 'openai' },
    groq:      { endpoint: 'https://api.groq.com/openai/v1/chat/completions', model: 'llama-3.3-70b-versatile', type: 'openai' },
    openai:    { endpoint: 'https://api.openai.com/v1/chat/completions', model: 'gpt-4o', type: 'openai' },
    claude:    { endpoint: 'https://api.anthropic.com/v1/messages', model: 'claude-sonnet-4-20250514', type: 'claude' },
    ollama:    { endpoint: 'http://localhost:11434', model: 'qwen2.5:7b', type: 'ollama' },
    custom:    { endpoint: '', model: '', type: 'openai' }
  };

  // ═══ PPT 风格配置 ═══
  const STYLES = {
    business: '商务专业风格，语言精炼有力，适合企业汇报。',
    academic: '学术严谨风格，逻辑清晰，适合论文答辩。',
    creative: '创意活泼风格，生动有趣，适合路演展示。',
    minimal:  '极简现代风格，言简意赅，适合快节奏演讲。',
    tech:     '科技感风格，充满未来感，适合技术发布。'
  };

  // ═══ 系统提示词 ═══
  const SYSTEM_PROMPT = '你是 PowerPoint 演示文稿专家。请输出可直接用于PPT的高质量内容，结构清晰，重点突出，不要多余的开场白。';

  // ═══ 初始化（必须调用） ═══
  function init(options = {}) {
    return new Promise(resolve => {
      // 恢复设置
      _settings.provider  = localStorage.getItem('ppt_provider')  || options.provider  || 'gemini';
      _settings.endpoint  = localStorage.getItem('ppt_endpoint')  || options.endpoint  || '';
      _settings.model     = localStorage.getItem('ppt_model')     || options.model     || '';
      _settings.apikey    = localStorage.getItem('ppt_apikey')    || options.apikey    || '';
      _settings.style     = localStorage.getItem('ppt_style')     || options.style     || 'business';

      // 恢复聊天记录
      try { _chatHistory = JSON.parse(localStorage.getItem('ppt_chat') || '[]'); } catch(_) {}

      // Office 初始化
      if (typeof Office !== 'undefined' && Office.onReady) {
        Office.onReady(info => {
          _officeReady = true;
          _host = info.host ? info.host.toString() : 'Browser';
          if (options.onReady) options.onReady({ host: _host, settings: getSettings() });
          resolve({ host: _host, settings: getSettings() });
        });
      } else {
        if (options.onReady) options.onReady({ host: 'Browser', settings: getSettings() });
        resolve({ host: 'Browser', settings: getSettings() });
      }
    });
  }

  // ═══ 配置管理 ═══
  function getSettings() { return { ..._settings }; }

  function getProviders() { return Object.keys(PROVIDERS).map(k => ({ id: k, ...PROVIDERS[k] })); }

  function getStyles() { return Object.keys(STYLES).map(k => ({ id: k, label: STYLES[k] })); }

  function saveSettings(newSettings) {
    Object.assign(_settings, newSettings);
    localStorage.setItem('ppt_provider',  _settings.provider);
    localStorage.setItem('ppt_endpoint',  _settings.endpoint);
    localStorage.setItem('ppt_model',     _settings.model);
    localStorage.setItem('ppt_apikey',    _settings.apikey);
    localStorage.setItem('ppt_style',     _settings.style);
    return getSettings();
  }

  // ═══ 连接检测 ═══
  async function checkConnection() {
    const p = PROVIDERS[_settings.provider];
    try {
      const ctrl = new AbortController();
      setTimeout(() => ctrl.abort(), 5000);
      let url, headers = {};
      if (_settings.provider === 'ollama') {
        url = `${_settings.endpoint || p.endpoint}/api/tags`;
      } else if (_settings.provider === 'gemini') {
        url = `${p.endpoint}?key=${_settings.apikey}`;
      } else {
        const base = _settings.endpoint || p.endpoint;
        url = base.replace('/chat/completions','/models').replace('/v1/messages','');
        if (_settings.apikey) headers['Authorization'] = `Bearer ${_settings.apikey}`;
      }
      const res = await fetch(url, { headers, signal: ctrl.signal });
      return res.ok || res.status === 400;
    } catch(_) {
      return false;
    }
  }

  // ═══ Office 文档操作 ═══
  function isOfficeReady() { return _officeReady; }

  function getHost() { return _host; }

  function getSelection() {
    return new Promise(resolve => {
      if (!_officeReady) { resolve(''); return; }
      Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, r => {
        resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : '');
      });
    });
  }

  function writeToOffice(text) {
    return new Promise((resolve, reject) => {
      if (!_officeReady) { resolve(); return; }
      Office.context.document.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, r => {
        r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(new Error(r.error.message));
      });
    });
  }

  // ═══ 聊天历史 ═══
  function getChatHistory() { return [..._chatHistory]; }

  function addChatMessage(sender, text) {
    _chatHistory.push({ sender, text, time: Date.now() });
    if (_chatHistory.length > 200) _chatHistory.splice(0, _chatHistory.length - 200);
    localStorage.setItem('ppt_chat', JSON.stringify(_chatHistory));
  }

  function clearChatHistory() {
    _chatHistory = [];
    localStorage.removeItem('ppt_chat');
  }

  function exportChatHistory() {
    let md = '# PPT助手对话记录\n\n';
    _chatHistory.forEach(m => {
      md += `### ${m.sender === 'user' ? '🧑 用户' : '📊 AI'}\n${m.text}\n\n---\n\n`;
    });
    return md;
  }

  // ═══ 构建 API 请求 ═══
  function _buildRequest(prompt, context) {
    const cfg = PROVIDERS[_settings.provider];
    const styleHint = STYLES[_settings.style] || '';
    const system = `${SYSTEM_PROMPT}${styleHint ? ' ' + styleHint : ''}`;
    let fullPrompt = context ? `[PPT当前选中内容]:\n${context}\n\n[用户指令]:\n${prompt}` : prompt;

    let url, headers = {}, body = {};

    switch (_settings.provider) {
      case 'gemini':
        url = `${cfg.endpoint}?key=${_settings.apikey}`;
        headers = { 'Content-Type': 'application/json' };
        body = { contents: [{ parts: [{ text: fullPrompt }] }], systemInstruction: { parts: [{ text: system }] } };
        return { url, headers, body, type: 'gemini' };
      case 'claude':
        url = _settings.endpoint || cfg.endpoint;
        headers = { 'Content-Type': 'application/json', 'x-api-key': _settings.apikey, 'anthropic-version': '2023-06-01' };
        body = { model: _settings.model || cfg.model, max_tokens: 4096, system, messages: [{ role: 'user', content: fullPrompt }] };
        return { url, headers, body, type: 'claude' };
      case 'ollama':
        url = (_settings.endpoint || cfg.endpoint) + '/api/generate';
        headers = { 'Content-Type': 'application/json' };
        body = { model: _settings.model || cfg.model, prompt: fullPrompt, stream: false, options: { temperature: 0.7 } };
        return { url, headers, body, type: 'ollama' };
      default: // OpenAI-compatible
        url = _settings.endpoint || cfg.endpoint;
        headers = { 'Content-Type': 'application/json' };
        if (_settings.apikey) headers['Authorization'] = `Bearer ${_settings.apikey}`;
        body = { model: _settings.model || cfg.model, messages: [{ role: 'system', content: system }, { role: 'user', content: fullPrompt }], temperature: 0.7 };
        return { url, headers, body, type: 'openai' };
    }
  }

  function _parseResponse(data, type) {
    if (type === 'gemini') return data.candidates?.[0]?.content?.parts?.[0]?.text;
    if (type === 'claude')  return data.content?.[0]?.text;
    if (type === 'ollama')  return data.response;
    return data.choices?.[0]?.message?.content;
  }

  async function _callAI(prompt, context) {
    const { url, headers, body, type } = _buildRequest(prompt, context);
    const res = await fetch(url, { method: 'POST', headers, body: JSON.stringify(body) });
    if (!res.ok) {
      let msg = `HTTP ${res.status}`;
      if (res.status === 401) msg = 'API Key 无效';
      else if (res.status === 403) msg = '访问被拒绝，请检查权限';
      else if (res.status === 429) msg = '请求太频繁，请稍后重试';
      else if (res.status === 404) msg = '接口地址或模型名称错误';
      throw new Error(msg);
    }
    const data = await res.json();
    const text = _parseResponse(data, type);
    if (!text) throw new Error('AI 返回为空，请检查配置');
    return text;
  }

  // ═══ PPT 编写功能 ═══
  async function generateOutline(topic) {
    if (!topic) throw new Error('请提供PPT主题');
    const prompt = `请为以下主题生成一份完整的PPT演示文稿大纲。要求：1.封面页标题 2.目录/议程页（3-5项）3.每一页的标题+3-5个要点 4.总结/致谢页。请用清晰的格式输出，标注"第X页："和要点。主题：${topic}`;
    const result = await _callAI(prompt, topic);
    addChatMessage('user', `📋 生成大纲：${topic.substring(0,50)}`);
    addChatMessage('ai', result);
    return result;
  }

  async function generateFullContent(topic) {
    if (!topic) throw new Error('请提供PPT主题');
    const prompt = `请为以下PPT主题生成完整的、可直接使用的幻灯片内容。对每一页，请提供：•页面标题 •核心要点（3-5条，每条10-20字）•该页的演讲备注（50字左右）。格式清晰，总共6-10页。主题：${topic}`;
    const result = await _callAI(prompt, topic);
    addChatMessage('user', `📝 生成完整内容：${topic.substring(0,50)}`);
    addChatMessage('ai', result);
    return result;
  }

  async function generateSlideContent(title) {
    if (!title) throw new Error('请提供页面标题');
    const prompt = `这是一页PPT的标题/主题。请为该页生成完整的幻灯片内容：•该页标题（可优化措辞）•核心要点（4-6条，每条简洁有力）•演讲者备注（50-80字）•建议的配图/图表类型。选中内容：${title}`;
    const result = await _callAI(prompt, title);
    addChatMessage('user', `📄 生成单页：${title.substring(0,50)}`);
    addChatMessage('ai', result);
    return result;
  }

  // ═══ PPT 修改功能 ═══
  async function polishText(text) {
    if (!text) throw new Error('请先选中文字');
    const prompt = `请对以下PPT幻灯片文字进行精修润色，使其更专业、更有说服力。保留原意和关键数据，优化措辞和表达：`;
    const result = await _callAI(prompt, text);
    addChatMessage('user', `✨ 精修润色：${text.substring(0,50)}...`);
    addChatMessage('ai', result);
    return result;
  }

  async function expandContent(text) {
    if (!text) throw new Error('请先选中文字');
    const prompt = `请将以下PPT幻灯片要点进行扩展，增加更多细节、数据支撑或案例说明，让内容更丰满：`;
    const result = await _callAI(prompt, text);
    addChatMessage('user', `📈 内容扩展：${text.substring(0,50)}...`);
    addChatMessage('ai', result);
    return result;
  }

  async function condenseContent(text) {
    if (!text) throw new Error('请先选中文字');
    const prompt = `请将以下PPT幻灯片内容精简缩写，保留核心信息，删减冗余表述：`;
    const result = await _callAI(prompt, text);
    addChatMessage('user', `📉 精简缩写：${text.substring(0,50)}...`);
    addChatMessage('ai', result);
    return result;
  }

  async function translate(text) {
    if (!text) throw new Error('请先选中文字');
    const prompt = `请将以下文本进行中英互译（中文翻译成英文，英文翻译成中文），保持商务PPT的专业语气：`;
    const result = await _callAI(prompt, text);
    addChatMessage('user', `🌐 翻译`);
    addChatMessage('ai', result);
    return result;
  }

  async function generateSpeakerNotes(text) {
    if (!text) throw new Error('请先选中文字');
    const prompt = `请为以下PPT幻灯片内容撰写演讲者备注。包括：本页核心信息、各要点讲解方式、过渡语。适合演讲者现场使用：`;
    const result = await _callAI(prompt, text);
    addChatMessage('user', `🎤 演讲备注`);
    addChatMessage('ai', result);
    return result;
  }

  async function getDesignTips(text) {
    if (!text) throw new Error('请先选中文字');
    const prompt = `作为PPT设计专家，请分析以下幻灯片内容，给出排版美化建议。包括：配色方案、字体建议、图示类型、动画推荐、布局优化：`;
    const result = await _callAI(prompt, text);
    addChatMessage('user', `🎨 设计建议：${text.substring(0,50)}...`);
    addChatMessage('ai', result);
    return result;
  }

  // ═══ 自由对话 ═══
  async function chat(message, context = '') {
    if (!message) throw new Error('请输入内容');
    const result = await _callAI(message, context);
    addChatMessage('user', message);
    addChatMessage('ai', result);
    return result;
  }

  // ═══ 公开 API ═══
  return {
    // 初始化
    init,
    // 配置
    getSettings, getProviders, getStyles, saveSettings, checkConnection,
    // Office
    isOfficeReady, getHost, getSelection, writeToOffice,
    // 聊天历史
    getChatHistory, addChatMessage, clearChatHistory, exportChatHistory,
    // PPT 编写
    generateOutline, generateFullContent, generateSlideContent,
    // PPT 修改
    polishText, expandContent, condenseContent, translate, generateSpeakerNotes, getDesignTips,
    // 自由对话
    chat
  };
})();

// Node / 浏览器兼容
if (typeof module !== 'undefined' && module.exports) { module.exports = PPTEngine; }
