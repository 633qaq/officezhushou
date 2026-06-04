const { Router } = require('express');
const config = require('../config');
const { callAI } = require('../services/ai-provider');
const { AppError } = require('../middleware/error-handler');
const logger = require('../utils/logger');

const router = Router();

const MODEL_PROVIDER_MAP = {
  'gemini-2.5-flash': 'gemini',
  'gemini-2.0-flash': 'gemini',
  'deepseek-chat': 'deepseek',
  'deepseek-reasoner': 'deepseek',
  'deepseek-v3': 'deepseek',
  'deepseek-r1': 'deepseek',
  'gpt-4o': 'openai',
  'gpt-4': 'openai',
  'gpt-3.5-turbo': 'openai',
  'claude-sonnet-4-20250514': 'claude',
  'claude-3-opus': 'claude',
  'claude-3-sonnet': 'claude',
  'llama-3.3-70b-versatile': 'groq',
  'llama-3.1-8b': 'groq',
};

const PROVIDER_KEY_MAP = {
  gemini: () => config.ai.gemini.apiKey,
  deepseek: () => config.ai.deepseek.apiKey,
  claude: () => config.ai.claude.apiKey,
  openai: () => config.ai.openai.apiKey,
  groq: () => config.ai.groq.apiKey,
  ollama: () => null,
};

function resolveProvider(model, headers) {
  const headerProvider = headers['x-provider'];
  if (headerProvider && PROVIDER_KEY_MAP[headerProvider]) {
    return { provider: headerProvider, apiKey: PROVIDER_KEY_MAP[headerProvider]() };
  }

  const matchedModel = Object.keys(MODEL_PROVIDER_MAP).find((key) => String(model || '').startsWith(key));
  if (matchedModel) {
    const provider = MODEL_PROVIDER_MAP[matchedModel];
    return { provider, apiKey: PROVIDER_KEY_MAP[provider]() };
  }

  return { provider: 'deepseek', apiKey: config.ai.deepseek.apiKey };
}

router.post('/v1/chat/completions', async (req, res, next) => {
  try {
    const { model, messages, stream } = req.body;

    if (!Array.isArray(messages) || messages.length === 0) {
      throw new AppError('messages must be a non-empty array', 400, 'VALIDATION_ERROR');
    }
    if (stream) {
      throw new AppError('Streaming is not supported by this proxy yet', 400, 'UNSUPPORTED_FEATURE');
    }

    const systemPrompt = messages.filter((item) => item.role === 'system').map((item) => item.content).join('\n');
    const userMessage = messages.filter((item) => item.role !== 'system').map((item) => item.content).join('\n\n').trim();

    if (!userMessage) {
      throw new AppError('A user message is required', 400, 'VALIDATION_ERROR');
    }

    const { provider, apiKey } = resolveProvider(model, req.headers);
    const effectiveApiKey = apiKey || String(req.headers.authorization || '').replace(/^Bearer\s+/i, '');

    logger.info(`[Proxy] OpenAI-compatible request model=${model || 'n/a'} provider=${provider}`);

    const content = await callAI(provider, userMessage, {
      apiKey: effectiveApiKey,
      model: model || '',
      systemPrompt,
    });

    res.json({
      id: `chatcmpl-${Date.now()}`,
      object: 'chat.completion',
      created: Math.floor(Date.now() / 1000),
      model: model || 'unknown',
      choices: [
        {
          index: 0,
          message: { role: 'assistant', content },
          finish_reason: 'stop',
        },
      ],
      usage: {
        prompt_tokens: Math.ceil(userMessage.length / 4),
        completion_tokens: Math.ceil(content.length / 4),
        total_tokens: Math.ceil((userMessage.length + content.length) / 4),
      },
    });
  } catch (error) {
    next(error);
  }
});

router.post('/v1beta/models/:modelAction', async (req, res, next) => {
  try {
    const modelAction = req.params.modelAction || '';
    const model = modelAction.split(':')[0] || 'gemini-2.5-flash';
    const contents = Array.isArray(req.body.contents) ? req.body.contents : [];
    const prompt = contents.flatMap((item) => item.parts || []).map((part) => part.text).filter(Boolean).join('\n');
    const systemPrompt = (req.body.systemInstruction?.parts || []).map((part) => part.text).filter(Boolean).join('\n');

    if (!prompt) {
      throw new AppError('Prompt content is required', 400, 'VALIDATION_ERROR');
    }

    const { provider, apiKey } = resolveProvider(model, req.headers);
    const effectiveApiKey = apiKey || req.query.key || '';

    logger.info(`[Proxy] Gemini-compatible request model=${model} provider=${provider}`);

    const content = await callAI(provider, prompt, {
      apiKey: effectiveApiKey,
      model,
      systemPrompt,
    });

    res.json({
      candidates: [
        {
          content: {
            role: 'model',
            parts: [{ text: content }],
          },
          finishReason: 'STOP',
          index: 0,
        },
      ],
      usageMetadata: {
        promptTokenCount: Math.ceil(prompt.length / 4),
        candidatesTokenCount: Math.ceil(content.length / 4),
        totalTokenCount: Math.ceil((prompt.length + content.length) / 4),
      },
    });
  } catch (error) {
    next(error);
  }
});

router.get('/health', (_req, res) => {
  res.json({
    success: true,
    name: 'Office AI Assistant Proxy',
    version: '2.0.0',
    providers: Object.keys(PROVIDER_KEY_MAP),
    configuredKeys: Object.entries(PROVIDER_KEY_MAP)
      .filter(([provider, getter]) => provider === 'ollama' || Boolean(getter()))
      .map(([provider]) => provider),
  });
});

module.exports = router;
