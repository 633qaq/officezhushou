const express = require('express');
const path = require('path');
const cors = require('cors');
const helmet = require('helmet');
const morgan = require('morgan');
const rateLimit = require('express-rate-limit');
const config = require('./config');
const { errorHandler } = require('./middleware/error-handler');
const logger = require('./utils/logger');

const authRoutes = require('./routes/auth');
const aiRoutes = require('./routes/ai');
const chatRoutes = require('./routes/chat');
const documentRoutes = require('./routes/documents');
const settingsRoutes = require('./routes/settings');
const proxyRoutes = require('./routes/proxy');

const app = express();

app.use(helmet({
  crossOriginResourcePolicy: { policy: 'cross-origin' },
  contentSecurityPolicy: false,
}));

app.use(cors({
  origin: config.cors.origin === '*' ? '*' : config.cors.origin.split(',').map((item) => item.trim()),
  credentials: true,
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization', 'X-Provider'],
}));

app.use(rateLimit({
  windowMs: config.rateLimit.windowMs,
  max: config.rateLimit.max,
  standardHeaders: true,
  legacyHeaders: false,
}));

app.use(morgan(config.nodeEnv === 'production' ? 'combined' : 'dev', {
  stream: { write: (message) => logger.info(message.trim()) },
}));

app.use(express.json({ limit: '5mb' }));
app.use(express.urlencoded({ extended: true, limit: '5mb' }));

app.use('/api/auth', authRoutes);
app.use('/api/ai', aiRoutes);
app.use('/api/chat', chatRoutes);
app.use('/api/documents', documentRoutes);
app.use('/api/settings', settingsRoutes);
app.use('/api/proxy', proxyRoutes);

app.get('/api/health', (_req, res) => {
  res.json({
    success: true,
    data: {
      status: 'running',
      version: '2.0.0',
      uptime: process.uptime(),
      timestamp: new Date().toISOString(),
    },
  });
});

app.use(express.static(config.publicDir, {
  extensions: ['html'],
  index: 'office.html',
}));

app.get('/', (_req, res) => {
  res.sendFile(path.join(config.publicDir, 'office.html'));
});

app.use((req, res) => {
  res.status(404).json({
    success: false,
    error: {
      code: 'NOT_FOUND',
      message: `Route ${req.method} ${req.path} was not found`,
    },
  });
});

app.use(errorHandler);

module.exports = { app };
