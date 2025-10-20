const config = require('../config/sharepoint.config');
const logger = require('../utils/logger');

const apiKeyAuth = (req, res, next) => {
  const apiKey = req.header('X-API-Key') || req.header('Authorization')?.replace('Bearer ', '');

  if (!apiKey) {
    logger.warn('API request without API key', { ip: req.ip, path: req.path });
    return res.status(401).json({
      success: false,
      error: 'API key is required'
    });
  }

  if (apiKey !== config.server.apiKey) {
    logger.warn('Invalid API key attempt', { ip: req.ip, path: req.path });
    return res.status(403).json({
      success: false,
      error: 'Invalid API key'
    });
  }

  next();
};

module.exports = { apiKeyAuth };