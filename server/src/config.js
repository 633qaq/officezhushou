const dotenv = require('dotenv');
const path = require('path');

dotenv.config({ path: path.resolve(__dirname, '..', '.env') });

const config = {
  port: Number.parseInt(process.env.PORT || '3456', 10),
  nodeEnv: process.env.NODE_ENV || 'development',
  jwt: {
    secret: process.env.JWT_SECRET || 'dev-secret-change-in-production',
    expiresIn: process.env.JWT_EXPIRES_IN || '7d',
  },
  db: {
    path: process.env.DB_PATH || path.resolve(__dirname, '..', 'data', 'app.db'),
  },
  cors: {
    origin: process.env.CORS_ORIGIN || '*',
  },
  publicDir: process.env.PUBLIC_DIR || path.resolve(__dirname, '..', '..'),
  rateLimit: {
    windowMs: Number.parseInt(process.env.RATE_LIMIT_WINDOW_MS || '60000', 10),
    max: Number.parseInt(process.env.RATE_LIMIT_MAX || '120', 10),
  },
  logLevel: process.env.LOG_LEVEL || 'info',
  maxUploadSize: Number.parseInt(process.env.MAX_UPLOAD_SIZE || '10', 10) * 1024 * 1024,
  ai: {
    gemini: { apiKey: process.env.GEMINI_API_KEY || '' },
    deepseek: { apiKey: process.env.DEEPSEEK_API_KEY || '' },
    claude: { apiKey: process.env.ANTHROPIC_API_KEY || '' },
    openai: { apiKey: process.env.OPENAI_API_KEY || '' },
    groq: { apiKey: process.env.GROQ_API_KEY || '' },
    ollama: { baseUrl: process.env.OLLAMA_BASE_URL || 'http://localhost:11434' },
  },
  defaultAiProvider: process.env.DEFAULT_AI_PROVIDER || 'gemini',
};

module.exports = config;
