const config = require('../config');
const logger = require('../utils/logger');

class AppError extends Error {
  constructor(message, statusCode = 400, code = 'BAD_REQUEST') {
    super(message);
    this.statusCode = statusCode;
    this.code = code;
    this.isOperational = true;
  }
}

function errorHandler(err, req, res, _next) {
  const statusCode = err.statusCode || 500;
  const code = err.code || 'INTERNAL_ERROR';

  if (statusCode >= 500 || !err.isOperational) {
    logger.error(`[${req.method}] ${req.path} -> ${err.message}`, err.stack);
  } else {
    logger.warn(`[${req.method}] ${req.path} -> ${err.message}`);
  }

  res.status(statusCode).json({
    success: false,
    error: {
      code,
      message: err.isOperational ? err.message : 'Internal server error',
      ...(config.nodeEnv === 'development' ? { stack: err.stack } : {}),
    },
  });
}

module.exports = { AppError, errorHandler };
