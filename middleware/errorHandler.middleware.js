const logger = require('../utils/logger');

const errorHandler = (err, req, res, next) => {
  logger.error('Error occurred', {
    error: err.message,
    stack: err.stack,
    path: req.path,
    method: req.method,
    ip: req.ip
  });

  // SharePoint specific errors
  if (err.response?.status === 401) {
    return res.status(401).json({
      success: false,
      error: 'SharePoint authentication failed',
      message: 'Token may be expired or invalid',
      details: err.response.data
    });
  }

  if (err.response?.status === 403) {
    return res.status(403).json({
      success: false,
      error: 'Access denied',
      message: 'Insufficient permissions to access SharePoint resource',
      details: err.response.data
    });
  }

  if (err.response?.status === 404) {
    return res.status(404).json({
      success: false,
      error: 'Resource not found',
      message: 'The requested SharePoint resource does not exist',
      details: err.response.data
    });
  }

  // General error
  res.status(err.status || 500).json({
    success: false,
    error: err.message || 'Internal server error',
    ...(process.env.NODE_ENV === 'development' && { stack: err.stack })
  });
};

module.exports = errorHandler;