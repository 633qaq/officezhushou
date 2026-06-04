const { Router } = require('express');
const { queryOne, execute } = require('../database');
const { authenticate } = require('../middleware/auth');
const { AppError } = require('../middleware/error-handler');

const router = Router();
const VALID_PROVIDERS = new Set(['gemini', 'deepseek', 'claude', 'openai', 'groq', 'ollama', 'custom']);
const VALID_STYLES = new Set(['business', 'academic', 'creative', 'minimal', 'tech']);

router.use(authenticate);

router.get('/', (req, res) => {
  let settings = queryOne('SELECT * FROM user_settings WHERE user_id = ?', [req.user.id]);

  if (!settings) {
    execute('INSERT INTO user_settings (user_id, provider, style) VALUES (?, ?, ?)', [req.user.id, 'gemini', 'business']);
    settings = { provider: 'gemini', endpoint: '', model: '', style: 'business', settings: '{}' };
  }

  let extra = {};
  try {
    extra = JSON.parse(settings.settings || '{}');
  } catch (_error) {
    extra = {};
  }

  res.json({
    success: true,
    data: {
      provider: settings.provider,
      endpoint: settings.endpoint || '',
      model: settings.model || '',
      style: settings.style || 'business',
      ...extra,
    },
  });
});

router.put('/', (req, res) => {
  const { provider, endpoint, model, style, ...rest } = req.body;

  if (provider && !VALID_PROVIDERS.has(provider)) {
    throw new AppError('Unsupported provider', 400, 'VALIDATION_ERROR');
  }
  if (style && !VALID_STYLES.has(style)) {
    throw new AppError('Unsupported style', 400, 'VALIDATION_ERROR');
  }

  const existing = queryOne('SELECT * FROM user_settings WHERE user_id = ?', [req.user.id]);
  const baseExtra = existing ? JSON.parse(existing.settings || '{}') : {};
  const nextExtra = { ...baseExtra };

  for (const key of ['apiKey', 'temperature', 'maxTokens']) {
    if (rest[key] !== undefined) {
      nextExtra[key] = rest[key];
    }
  }

  if (existing) {
    execute(
      `UPDATE user_settings
       SET provider = ?, endpoint = ?, model = ?, style = ?, settings = ?, updated_at = datetime('now')
       WHERE user_id = ?`,
      [provider || existing.provider, endpoint ?? existing.endpoint ?? '', model ?? existing.model ?? '', style || existing.style || 'business', JSON.stringify(nextExtra), req.user.id]
    );
  } else {
    execute(
      'INSERT INTO user_settings (user_id, provider, endpoint, model, style, settings) VALUES (?, ?, ?, ?, ?, ?)',
      [req.user.id, provider || 'gemini', endpoint || '', model || '', style || 'business', JSON.stringify(nextExtra)]
    );
  }

  res.json({ success: true, message: 'Settings saved' });
});

module.exports = router;
