const { Router } = require('express');
const { v4: uuidv4 } = require('uuid');
const { queryOne, queryAll, execute } = require('../database');
const { authenticate } = require('../middleware/auth');
const { AppError } = require('../middleware/error-handler');

const router = Router();
const VALID_DOCUMENT_TYPES = new Set(['text', 'outline', 'slides', 'docx', 'pptx']);

function safeParseTags(value) {
  try {
    if (Array.isArray(value)) {
      return value;
    }
    return JSON.parse(value || '[]');
  } catch (_error) {
    return [];
  }
}

router.use(authenticate);

router.get('/', (req, res) => {
  const page = Math.max(1, Number.parseInt(req.query.page || '1', 10));
  const limit = Math.min(50, Math.max(1, Number.parseInt(req.query.limit || '20', 10)));
  const offset = (page - 1) * limit;
  const type = String(req.query.type || '').trim();

  let whereClause = 'WHERE user_id = ?';
  const params = [req.user.id];

  if (type) {
    if (!VALID_DOCUMENT_TYPES.has(type)) {
      throw new AppError('Unsupported document type', 400, 'VALIDATION_ERROR');
    }
    whereClause += ' AND type = ?';
    params.push(type);
  }

  const total = Number(queryOne(`SELECT COUNT(*) as count FROM documents ${whereClause}`, params)?.count || 0);
  const documents = queryAll(
    `SELECT id, title, type, summary, tags, source, created_at, updated_at FROM documents ${whereClause} ORDER BY updated_at DESC LIMIT ? OFFSET ?`,
    [...params, limit, offset]
  ).map((item) => ({ ...item, tags: safeParseTags(item.tags) }));

  res.json({
    success: true,
    data: {
      documents,
      pagination: {
        page,
        limit,
        total,
        totalPages: total === 0 ? 0 : Math.ceil(total / limit),
      },
    },
  });
});

router.post('/', (req, res) => {
  const title = String(req.body.title || '').trim() || 'Untitled document';
  const type = String(req.body.type || 'text').trim();
  const content = typeof req.body.content === 'string' ? req.body.content : '';
  const summary = typeof req.body.summary === 'string' ? req.body.summary : '';
  const source = String(req.body.source || 'manual').trim() || 'manual';
  const tags = Array.isArray(req.body.tags) ? req.body.tags : [];

  if (!VALID_DOCUMENT_TYPES.has(type)) {
    throw new AppError('Unsupported document type', 400, 'VALIDATION_ERROR');
  }

  const id = uuidv4();
  execute(
    'INSERT INTO documents (id, user_id, title, type, content, summary, tags, source) VALUES (?, ?, ?, ?, ?, ?, ?, ?)',
    [id, req.user.id, title, type, content, summary, JSON.stringify(tags), source]
  );

  res.status(201).json({ success: true, data: { id, title, type } });
});

router.get('/:id', (req, res) => {
  const document = queryOne('SELECT * FROM documents WHERE id = ? AND user_id = ?', [req.params.id, req.user.id]);
  if (!document) {
    throw new AppError('Document was not found', 404, 'NOT_FOUND');
  }
  res.json({ success: true, data: { ...document, tags: safeParseTags(document.tags) } });
});

router.put('/:id', (req, res) => {
  const existing = queryOne('SELECT id FROM documents WHERE id = ? AND user_id = ?', [req.params.id, req.user.id]);
  if (!existing) {
    throw new AppError('Document was not found', 404, 'NOT_FOUND');
  }

  const updates = [];
  const params = [];

  if (req.body.title !== undefined) {
    updates.push('title = ?');
    params.push(String(req.body.title || '').trim() || 'Untitled document');
  }
  if (req.body.type !== undefined) {
    const type = String(req.body.type || '').trim();
    if (!VALID_DOCUMENT_TYPES.has(type)) {
      throw new AppError('Unsupported document type', 400, 'VALIDATION_ERROR');
    }
    updates.push('type = ?');
    params.push(type);
  }
  if (req.body.content !== undefined) {
    updates.push('content = ?');
    params.push(typeof req.body.content === 'string' ? req.body.content : '');
  }
  if (req.body.summary !== undefined) {
    updates.push('summary = ?');
    params.push(typeof req.body.summary === 'string' ? req.body.summary : '');
  }
  if (req.body.tags !== undefined) {
    updates.push('tags = ?');
    params.push(JSON.stringify(Array.isArray(req.body.tags) ? req.body.tags : []));
  }

  if (updates.length === 0) {
    throw new AppError('No fields were provided for update', 400, 'VALIDATION_ERROR');
  }

  updates.push("updated_at = datetime('now')");
  params.push(req.params.id, req.user.id);
  execute(`UPDATE documents SET ${updates.join(', ')} WHERE id = ? AND user_id = ?`, params);

  res.json({ success: true, message: 'Document updated' });
});

router.delete('/:id', (req, res) => {
  const result = execute('DELETE FROM documents WHERE id = ? AND user_id = ?', [req.params.id, req.user.id]);
  if (result.changes === 0) {
    throw new AppError('Document was not found', 404, 'NOT_FOUND');
  }
  res.json({ success: true, message: 'Document deleted' });
});

module.exports = router;
