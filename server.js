const express = require('express');
const cors = require('cors');
const helmet = require('helmet');
const rateLimit = require('express-rate-limit');
const config = require('./config/sharepoint.config');
const logger = require('./utils/logger');
const errorHandler = require('./middleware/errorHandler.middleware');
const { apiKeyAuth } = require('./middleware/auth.middleware');

// Routes
const authRoutes = require('./routes/auth.routes');
const listsRoutes = require('./routes/lists.routes');
const itemsRoutes = require('./routes/items.routes');

const app = express();

// Start server
const PORT = process.env.PORT || 10000;
const server = app.listen(PORT, '0.0.0.0', () => {
  logger.info(`SharePoint API Service running on port ${PORT}`);
  logger.info(`Environment: ${config.server.env}`);
  logger.info(`SharePoint Site: ${config.sharepoint.siteUrl}`);
});



// Security middleware
app.use(helmet());
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Rate limiting
const limiter = rateLimit({
  windowMs: config.rateLimit.windowMs,
  max: config.rateLimit.max,
  message: 'Too many requests from this IP, please try again later.',
  standardHeaders: true,
  legacyHeaders: false,
});
app.use(limiter);

// Request logging
app.use((req, res, next) => {
  logger.info(`${req.method} ${req.path}`, {
    ip: req.ip,
    userAgent: req.get('user-agent')
  });
  next();
});

// Health check endpoint (no auth required)
app.get('/health', (req, res) => {
  res.json({
    status: 'ok',
    timestamp: new Date().toISOString(),
    uptime: process.uptime()
  });
});

// API documentation
app.get('/', (req, res) => {
  res.json({
    name: 'SharePoint API Service',
    version: '1.0.0',
    endpoints: {
      health: 'GET /health',
      auth: {
        getToken: 'GET /api/auth/token',
        validateToken: 'GET /api/auth/validate'
      },
      lists: {
        getAllLists: 'GET /api/lists',
        getListByTitle: 'GET /api/lists/:listTitle'
      },
      items: {
        getItems: 'GET /api/items/:listTitle/items',
        getItemById: 'GET /api/items/:listTitle/items/:itemId',
        createItem: 'POST /api/items/:listTitle/items',
        updateItem: 'PUT /api/items/:listTitle/items/:itemId',
        deleteItem: 'DELETE /api/items/:listTitle/items/:itemId',
        search: 'GET /api/items/search?query=...'
      }
    },
    authentication: 'Include X-API-Key header in all requests (except /health and /)'
  });
});

// Protected routes (require API key)
app.use('/api/auth', apiKeyAuth, authRoutes);
app.use('/api/lists', apiKeyAuth, listsRoutes);
app.use('/api/items', apiKeyAuth, itemsRoutes);

// 404 handler
app.use((req, res) => {
  res.status(404).json({
    success: false,
    error: 'Endpoint not found'
  });
});

// Error handling middleware
app.use(errorHandler);



// Graceful shutdown
process.on('SIGTERM', () => {
  logger.info('SIGTERM signal received: closing HTTP server');
  server.close(() => {
    logger.info('HTTP server closed');
  });
});

module.exports = app;
