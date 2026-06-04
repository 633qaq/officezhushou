const { Router } = require('express');
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');
const { v4: uuidv4 } = require('uuid');
const config = require('../config');
const { queryOne, execute } = require('../database');
const { authenticate } = require('../middleware/auth');
const { AppError } = require('../middleware/error-handler');
const logger = require('../utils/logger');

const router = Router();

router.post('/register', async (req, res) => {
  const username = String(req.body.username || '').trim();
  const password = String(req.body.password || '');
  const displayName = String(req.body.displayName || username).trim();

  if (!username || !password) {
    throw new AppError('Username and password are required', 400, 'VALIDATION_ERROR');
  }
  if (username.length < 2 || username.length > 32) {
    throw new AppError('Username must be between 2 and 32 characters', 400, 'VALIDATION_ERROR');
  }
  if (password.length < 6) {
    throw new AppError('Password must be at least 6 characters', 400, 'VALIDATION_ERROR');
  }

  const existing = queryOne('SELECT id FROM users WHERE username = ?', [username]);
  if (existing) {
    throw new AppError('Username is already registered', 409, 'DUPLICATE_USERNAME');
  }

  const id = uuidv4();
  const passwordHash = await bcrypt.hash(password, 10);

  execute(
    'INSERT INTO users (id, username, password_hash, display_name) VALUES (?, ?, ?, ?)',
    [id, username, passwordHash, displayName || username]
  );
  execute(
    'INSERT INTO user_settings (user_id, provider, style) VALUES (?, ?, ?)',
    [id, config.defaultAiProvider, 'business']
  );

  logger.info(`[Auth] Registered user ${username}`);

  const token = jwt.sign({ id, username, role: 'user' }, config.jwt.secret, {
    expiresIn: config.jwt.expiresIn,
  });

  res.status(201).json({
    success: true,
    data: {
      token,
      user: {
        id,
        username,
        displayName: displayName || username,
        role: 'user',
      },
    },
  });
});

router.post('/login', async (req, res) => {
  const username = String(req.body.username || '').trim();
  const password = String(req.body.password || '');

  if (!username || !password) {
    throw new AppError('Username and password are required', 400, 'VALIDATION_ERROR');
  }

  const user = queryOne('SELECT * FROM users WHERE username = ?', [username]);
  if (!user) {
    throw new AppError('Invalid username or password', 401, 'INVALID_CREDENTIALS');
  }

  const valid = await bcrypt.compare(password, user.password_hash);
  if (!valid) {
    throw new AppError('Invalid username or password', 401, 'INVALID_CREDENTIALS');
  }

  const token = jwt.sign(
    { id: user.id, username: user.username, role: user.role },
    config.jwt.secret,
    { expiresIn: config.jwt.expiresIn }
  );

  res.json({
    success: true,
    data: {
      token,
      user: {
        id: user.id,
        username: user.username,
        displayName: user.display_name || user.username,
        role: user.role,
      },
    },
  });
});

router.get('/me', authenticate, (req, res) => {
  const user = queryOne(
    'SELECT id, username, display_name, role, created_at FROM users WHERE id = ?',
    [req.user.id]
  );

  if (!user) {
    throw new AppError('User was not found', 404, 'USER_NOT_FOUND');
  }

  res.json({
    success: true,
    data: {
      id: user.id,
      username: user.username,
      displayName: user.display_name || user.username,
      role: user.role,
      createdAt: user.created_at,
    },
  });
});

module.exports = router;
