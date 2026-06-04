const { Router } = require('express');
const { v4: uuidv4 } = require('uuid');
const { callAI, getProviders, getStyles, checkConnection } = require('../services/ai-provider');
const { optionalAuth } = require('../middleware/auth');
const { AppError } = require('../middleware/error-handler');
const { execute } = require('../database');
const logger = require('../utils/logger');

const router = Router();

const PROMPTS = {
  outline: (topic) => `Create a structured PPT outline for: ${topic}\nRequirements:\n1. Cover slide title\n2. Agenda section\n3. 6-10 logical slides with 3-5 bullets each\n4. Closing summary slide\nFormat each slide as Slide N: Title + bullets.`,
  content: (topic) => `Create complete PowerPoint slide content for: ${topic}\nFor each slide include a title, 3-5 concise bullet points, and short speaker notes. Aim for 6-10 slides.`,
  slide: (title) => `Create one complete PowerPoint slide for this topic: ${title}\nInclude an improved title, 4-6 bullets, short speaker notes, and suggested visual/chart ideas.`,
  polish: (text) => `Polish the following PPT text to make it more professional and persuasive while preserving the original meaning:\n${text}`,
  expand: (text) => `Expand the following PPT content with more detail, examples, and support points:\n${text}`,
  condense: (text) => `Condense the following PPT content while keeping the core message intact:\n${text}`,
  translate: (text) => `Translate the following text between Chinese and English while preserving business presentation tone:\n${text}`,
  speakerNotes: (text) => `Write speaker notes for the following PPT content. Keep them practical for live presentation delivery:\n${text}`,
  designTips: (text) => `Review the following PPT content and provide design suggestions for layout, typography, color, visuals, and animation:\n${text}`,
};

function saveChat(userId, role, content, provider, model) {
  if (!userId) {
    return;
  }

  try {
    execute(
      'INSERT INTO chat_history (id, user_id, role, content, provider, model) VALUES (?, ?, ?, ?, ?, ?)',
      [uuidv4(), userId, role, content, provider || '', model || '']
    );
  } catch (error) {
    logger.warn(`Failed to save chat history: ${error.message}`);
  }
}

async function runAiAction({ userId, provider, prompt, model, apiKey, systemPrompt, style }) {
  const text = await callAI(provider, prompt, { model, apiKey, systemPrompt, style });

  if (userId) {
    saveChat(userId, 'user', prompt, provider, model);
    saveChat(userId, 'assistant', text, provider, model);
  }

  return text;
}

router.post('/chat', optionalAuth, async (req, res) => {
  const message = String(req.body.message || '').trim();
  const context = String(req.body.context || '').trim();
  const provider = req.body.provider || 'gemini';
  const { model, apiKey, systemPrompt, style } = req.body;

  if (!message) {
    throw new AppError('Message is required', 400, 'VALIDATION_ERROR');
  }

  const prompt = context ? `Context:\n${context}\n\nUser request:\n${message}` : message;
  const text = await runAiAction({ userId: req.user?.id, provider, prompt, model, apiKey, systemPrompt, style });
  res.json({ success: true, data: { text } });
});

router.post('/generate-outline', optionalAuth, async (req, res) => {
  const topic = String(req.body.topic || '').trim();
  if (!topic) {
    throw new AppError('Topic is required', 400, 'VALIDATION_ERROR');
  }
  const text = await runAiAction({ userId: req.user?.id, provider: req.body.provider || 'gemini', prompt: PROMPTS.outline(topic), model: req.body.model, apiKey: req.body.apiKey, style: req.body.style });
  res.json({ success: true, data: { text } });
});

router.post('/generate-content', optionalAuth, async (req, res) => {
  const topic = String(req.body.topic || '').trim();
  if (!topic) {
    throw new AppError('Topic is required', 400, 'VALIDATION_ERROR');
  }
  const text = await runAiAction({ userId: req.user?.id, provider: req.body.provider || 'gemini', prompt: PROMPTS.content(topic), model: req.body.model, apiKey: req.body.apiKey, style: req.body.style });
  res.json({ success: true, data: { text } });
});

router.post('/generate-slide', optionalAuth, async (req, res) => {
  const title = String(req.body.title || '').trim();
  if (!title) {
    throw new AppError('Slide title is required', 400, 'VALIDATION_ERROR');
  }
  const text = await runAiAction({ userId: req.user?.id, provider: req.body.provider || 'gemini', prompt: PROMPTS.slide(title), model: req.body.model, apiKey: req.body.apiKey, style: req.body.style });
  res.json({ success: true, data: { text } });
});

router.post('/polish', optionalAuth, async (req, res) => {
  const sourceText = String(req.body.text || '').trim();
  if (!sourceText) {
    throw new AppError('Text is required', 400, 'VALIDATION_ERROR');
  }
  const text = await runAiAction({ userId: req.user?.id, provider: req.body.provider || 'gemini', prompt: PROMPTS.polish(sourceText), model: req.body.model, apiKey: req.body.apiKey, style: req.body.style });
  res.json({ success: true, data: { text } });
});

router.post('/expand', optionalAuth, async (req, res) => {
  const sourceText = String(req.body.text || '').trim();
  if (!sourceText) {
    throw new AppError('Text is required', 400, 'VALIDATION_ERROR');
  }
  const text = await runAiAction({ userId: req.user?.id, provider: req.body.provider || 'gemini', prompt: PROMPTS.expand(sourceText), model: req.body.model, apiKey: req.body.apiKey, style: req.body.style });
  res.json({ success: true, data: { text } });
});

router.post('/condense', optionalAuth, async (req, res) => {
  const sourceText = String(req.body.text || '').trim();
  if (!sourceText) {
    throw new AppError('Text is required', 400, 'VALIDATION_ERROR');
  }
  const text = await runAiAction({ userId: req.user?.id, provider: req.body.provider || 'gemini', prompt: PROMPTS.condense(sourceText), model: req.body.model, apiKey: req.body.apiKey, style: req.body.style });
  res.json({ success: true, data: { text } });
});

router.post('/translate', optionalAuth, async (req, res) => {
  const sourceText = String(req.body.text || '').trim();
  if (!sourceText) {
    throw new AppError('Text is required', 400, 'VALIDATION_ERROR');
  }
  const text = await runAiAction({ userId: req.user?.id, provider: req.body.provider || 'gemini', prompt: PROMPTS.translate(sourceText), model: req.body.model, apiKey: req.body.apiKey, style: req.body.style });
  res.json({ success: true, data: { text } });
});

router.post('/speaker-notes', optionalAuth, async (req, res) => {
  const sourceText = String(req.body.text || '').trim();
  if (!sourceText) {
    throw new AppError('Text is required', 400, 'VALIDATION_ERROR');
  }
  const text = await runAiAction({ userId: req.user?.id, provider: req.body.provider || 'gemini', prompt: PROMPTS.speakerNotes(sourceText), model: req.body.model, apiKey: req.body.apiKey, style: req.body.style });
  res.json({ success: true, data: { text } });
});

router.post('/design-tips', optionalAuth, async (req, res) => {
  const sourceText = String(req.body.text || '').trim();
  if (!sourceText) {
    throw new AppError('Text is required', 400, 'VALIDATION_ERROR');
  }
  const text = await runAiAction({ userId: req.user?.id, provider: req.body.provider || 'gemini', prompt: PROMPTS.designTips(sourceText), model: req.body.model, apiKey: req.body.apiKey, style: req.body.style });
  res.json({ success: true, data: { text } });
});

router.get('/providers', (_req, res) => {
  res.json({ success: true, data: getProviders() });
});

router.get('/styles', (_req, res) => {
  res.json({ success: true, data: getStyles() });
});

router.post('/check-connection', async (req, res) => {
  const provider = req.body.provider || 'gemini';
  const connected = await checkConnection(provider);
  res.json({ success: true, data: { connected } });
});

module.exports = router;
