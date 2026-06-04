const { Router } = require('express');
const { queryOne, queryAll, execute } = require('../database');
const { authenticate } = require('../middleware/auth');
const { AppError } = require('../middleware/error-handler');

const router = Router();

router.use(authenticate);

router.get('/', (req, res) => {
  const page = Math.max(1, Number.parseInt(req.query.page || '1', 10));
  const limit = Math.min(100, Math.max(1, Number.parseInt(req.query.limit || '50', 10)));
  const offset = (page - 1) * limit;

  const total = Number(queryOne('SELECT COUNT(*) as count FROM chat_history WHERE user_id = ?', [req.user.id])?.count || 0);
  const messages = queryAll(
    'SELECT id, role, content, provider, model, created_at FROM chat_history WHERE user_id = ? ORDER BY created_at DESC LIMIT ? OFFSET ?',
    [req.user.id, limit, offset]
  ).reverse();

  res.json({
    success: true,
    data: {
      messages,
      pagination: {
        page,
        limit,
        total,
        totalPages: total === 0 ? 0 : Math.ceil(total / limit),
      },
    },
  });
});

router.delete('/', (req, res) => {
  execute('DELETE FROM chat_history WHERE user_id = ?', [req.user.id]);
  res.json({ success: true, message: 'Chat history cleared' });
});

router.get('/export', (req, res) => {
  const format = String(req.query.format || 'markdown');
  const messages = queryAll(
    'SELECT role, content, provider, created_at FROM chat_history WHERE user_id = ? ORDER BY created_at ASC',
    [req.user.id]
  );

  if (format === 'json') {
    res.json({ success: true, data: messages });
    return;
  }

  let markdown = `# Office AI Assistant Chat Export\n\nExported at: ${new Date().toISOString()}\n\n---\n\n`;
  for (const message of messages) {
    const roleLabel = message.role === 'user' ? 'User' : 'Assistant';
    const provider = message.provider ? ` (${message.provider})` : '';
    markdown += `### ${roleLabel}${provider}\n${message.content}\n\n---\n\n`;
  }

  res.json({ success: true, data: { format: 'markdown', content: markdown } });
});

router.get('/:id', (req, res) => {
  const message = queryOne('SELECT * FROM chat_history WHERE id = ? AND user_id = ?', [req.params.id, req.user.id]);
  if (!message) {
    throw new AppError('Message was not found', 404, 'NOT_FOUND');
  }
  res.json({ success: true, data: message });
});

router.delete('/:id', (req, res) => {
  const result = execute('DELETE FROM chat_history WHERE id = ? AND user_id = ?', [req.params.id, req.user.id]);
  if (result.changes === 0) {
    throw new AppError('Message was not found', 404, 'NOT_FOUND');
  }
  res.json({ success: true, message: 'Message deleted' });
});

module.exports = router;
