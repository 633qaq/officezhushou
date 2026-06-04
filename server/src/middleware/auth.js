const jwt = require('jsonwebtoken');
const config = require('../config');
const { AppError } = require('./error-handler');

function authenticate(req, _res, next) {
  const authHeader = req.headers.authorization;
  if (!authHeader || !authHeader.startsWith('Bearer ')) {
    throw new AppError('Authentication token is required', 401, 'UNAUTHORIZED');
  }

  const token = authHeader.slice('Bearer '.length);

  try {
    req.user = jwt.verify(token, config.jwt.secret);
    next();
  } catch (error) {
    if (error.name === 'TokenExpiredError') {
      throw new AppError('Authentication token has expired', 401, 'TOKEN_EXPIRED');
    }
    throw new AppError('Invalid authentication token', 401, 'INVALID_TOKEN');
  }
}

function optionalAuth(req, _res, next) {
  const authHeader = req.headers.authorization;
  if (authHeader && authHeader.startsWith('Bearer ')) {
    try {
      req.user = jwt.verify(authHeader.slice('Bearer '.length), config.jwt.secret);
    } catch (_error) {
      req.user = undefined;
    }
  }
  next();
}

function requireAdmin(req, _res, next) {
  if (!req.user || req.user.role !== 'admin') {
    throw new AppError('Administrator privileges are required', 403, 'FORBIDDEN');
  }
  next();
}

module.exports = { authenticate, optionalAuth, requireAdmin };
