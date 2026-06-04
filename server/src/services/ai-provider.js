const config = require('../config');
const logger = require('../utils/logger');
const { AppError } = require('../middleware/error-handler');

const PROVIDERS = {
  gemini: {
    name: 'Gemini',
    buildRequest({ apiKey, model, prompt, systemPrompt }) {
      const selectedModel = model || 'gemini-2.5-flash';
      return {
        url: `https://generativelanguage.googleapis.com/v1beta/models/${selectedModel}:generateContent?key=${apiKey}`,
        headers: { 'Content-Type': 'application/json' },
        body: {
          contents: [{ parts: [{ text: prompt }] }],
          ...(systemPrompt ? { systemInstruction: { parts: [{ text: systemPrompt }] } } : {}),
        },
      };
    },
    parseResponse(data) {
      return data.candidates?.[0]?.content?.parts?.[0]?.text || '';
    },
  },
  deepseek: {
    name: 'DeepSeek',
    buildRequest({ apiKey, model, prompt, systemPrompt }) {
      return {
        url: 'https://api.deepseek.com/v1/chat/completions',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${apiKey}` },
        body: {
          model: model || 'deepseek-chat',
          messages: [
            ...(systemPrompt ? [{ role: 'system', content: systemPrompt }] : []),
            { role: 'user', content: prompt },
          ],
          temperature: 0.7,
        },
      };
    },
    parseResponse(data) {
      return data.choices?.[0]?.message?.content || '';
    },
  },
  claude: {
    name: 'Claude',
    buildRequest({ apiKey, model, prompt, systemPrompt }) {
      return {
        url: 'https://api.anthropic.com/v1/messages',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': apiKey,
          'anthropic-version': '2023-06-01',
        },
        body: {
          model: model || 'claude-sonnet-4-20250514',
          max_tokens: 4096,
          ...(systemPrompt ? { system: systemPrompt } : {}),
          messages: [{ role: 'user', content: prompt }],
        },
      };
    },
    parseResponse(data) {
      return data.content?.[0]?.text || '';
    },
  },
  openai: {
    name: 'OpenAI',
    buildRequest({ apiKey, model, prompt, systemPrompt }) {
      return {
        url: 'https://api.openai.com/v1/chat/completions',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${apiKey}` },
        body: {
          model: model || 'gpt-4o',
          messages: [
            ...(systemPrompt ? [{ role: 'system', content: systemPrompt }] : []),
            { role: 'user', content: prompt },
          ],
          temperature: 0.7,
        },
      };
    },
    parseResponse(data) {
      return data.choices?.[0]?.message?.content || '';
    },
  },
  groq: {
    name: 'Groq',
    buildRequest({ apiKey, model, prompt, systemPrompt }) {
      return {
        url: 'https://api.groq.com/openai/v1/chat/completions',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${apiKey}` },
        body: {
          model: model || 'llama-3.3-70b-versatile',
          messages: [
            ...(systemPrompt ? [{ role: 'system', content: systemPrompt }] : []),
            { role: 'user', content: prompt },
          ],
          temperature: 0.7,
        },
      };
    },
    parseResponse(data) {
      return data.choices?.[0]?.message?.content || '';
    },
  },
  ollama: {
    name: 'Ollama',
    noKey: true,
    buildRequest({ model, prompt, systemPrompt }) {
      return {
        url: `${config.ai.ollama.baseUrl || 'http://localhost:11434'}/api/generate`,
        headers: { 'Content-Type': 'application/json' },
        body: {
          model: model || 'qwen2.5:7b',
          prompt: systemPrompt ? `${systemPrompt}\n\n${prompt}` : prompt,
          stream: false,
          options: { temperature: 0.7 },
        },
      };
    },
    parseResponse(data) {
      return data.response || '';
    },
  },
};

const STYLES = {
  business: 'Use a polished, executive presentation tone.',
  academic: 'Use a structured, rigorous, evidence-driven tone.',
  creative: 'Use a lively and engaging presentation tone.',
  minimal: 'Use a concise, modern, minimalist presentation tone.',
  tech: 'Use a forward-looking, technical presentation tone.',
};

const SYSTEM_PROMPT_BASE = 'You are a PowerPoint content expert. Produce presentation-ready output with clear structure, concise points, and high information density.';

async function callAI(provider, prompt, options = {}) {
  const providerConfig = PROVIDERS[provider];
  if (!providerConfig) {
    throw new AppError(`Unsupported AI provider: ${provider}`, 400, 'INVALID_PROVIDER');
  }

  const apiKey = options.apiKey || config.ai[provider]?.apiKey || '';
  if (!providerConfig.noKey && !apiKey) {
    throw new AppError(`Provider ${provider} is missing an API key`, 400, 'MISSING_API_KEY');
  }

  const styleDescription = STYLES[options.style || 'business'] || '';
  const systemPrompt = options.systemPrompt || `${SYSTEM_PROMPT_BASE} ${styleDescription}`.trim();

  const { url, headers, body } = providerConfig.buildRequest({
    apiKey,
    model: options.model || '',
    prompt: typeof prompt === 'string' ? prompt : JSON.stringify(prompt),
    systemPrompt,
  });

  logger.info(`[AI] provider=${provider} model=${body.model || options.model || 'n/a'} promptLength=${String(prompt).length}`);

  const response = await fetch(url, {
    method: 'POST',
    headers,
    body: JSON.stringify(body),
    signal: AbortSignal.timeout(120000),
  });

  if (!response.ok) {
    const errorText = await response.text().catch(() => '');
    logger.error(`[AI] ${provider} failed with ${response.status}: ${errorText}`);

    const statusMap = {
      400: 'The AI request parameters were invalid',
      401: 'The API key is invalid or unauthorized',
      403: 'The AI provider rejected access',
      404: 'The model or endpoint could not be found',
      429: 'The AI provider rate limit was reached',
    };

    throw new AppError(
      statusMap[response.status] || `AI service error (HTTP ${response.status})`,
      response.status >= 500 ? 502 : 400,
      'AI_ERROR'
    );
  }

  const data = await response.json();
  const text = providerConfig.parseResponse(data);

  if (!text) {
    throw new AppError('The AI provider returned an empty response', 502, 'EMPTY_RESPONSE');
  }

  return text;
}

function getProviders() {
  return Object.entries(PROVIDERS).map(([id, provider]) => ({
    id,
    name: provider.name,
    noKey: Boolean(provider.noKey),
    hasKey: Boolean(config.ai[id]?.apiKey) || Boolean(provider.noKey),
  }));
}

function getStyles() {
  return Object.entries(STYLES).map(([id, label]) => ({ id, label }));
}

async function checkConnection(provider) {
  try {
    if (provider === 'ollama') {
      const response = await fetch(`${config.ai.ollama.baseUrl || 'http://localhost:11434'}/api/tags`, {
        signal: AbortSignal.timeout(5000),
      });
      return response.ok;
    }

    await callAI(provider, 'Reply with OK only.', {
      systemPrompt: 'Reply with OK only.',
    });
    return true;
  } catch (error) {
    logger.warn(`[AI] Connectivity check failed for ${provider}: ${error.message}`);
    return false;
  }
}

module.exports = { callAI, getProviders, getStyles, checkConnection, PROVIDERS };
