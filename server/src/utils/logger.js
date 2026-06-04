/**
 * 日志工具
 */
const config = require('../config');

const LEVELS = { debug: 0, info: 1, warn: 2, error: 3 };
const currentLevel = LEVELS[config.logLevel] ?? LEVELS.info;

const ts = () => new Date().toISOString();

const logger = {
  debug(...args) { if (currentLevel <= LEVELS.debug) console.debug(`[${ts()}] [DEBUG]`, ...args); },
  info(...args)  { if (currentLevel <= LEVELS.info)  console.info(`[${ts()}] [INFO]`, ...args); },
  warn(...args)  { if (currentLevel <= LEVELS.warn)  console.warn(`[${ts()}] [WARN]`, ...args); },
  error(...args) { if (currentLevel <= LEVELS.error) console.error(`[${ts()}] [ERROR]`, ...args); },
};

module.exports = logger;
