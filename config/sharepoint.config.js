require('dotenv').config();

module.exports = {
 
        server: {
          port: process.env.PORT || 10000, // Changed from 3000
          env: process.env.NODE_ENV || 'development',
          apiKey: process.env.API_KEY
        },
  sharepoint: {
    tenantId: process.env.SHAREPOINT_TENANT_ID,
    clientId: process.env.SHAREPOINT_CLIENT_ID,
    clientSecret: process.env.SHAREPOINT_CLIENT_SECRET,
    siteUrl: process.env.SHAREPOINT_SITE_URL,
    tenantName: process.env.SHAREPOINT_TENANT_NAME,
    tokenEndpoint: process.env.TOKEN_ENDPOINT,
    resource: process.env.SHAREPOINT_SITE_URL
  },
  cache: {
    tokenTTL: parseInt(process.env.TOKEN_CACHE_TTL) || 3500,
    dataTTL: parseInt(process.env.DATA_CACHE_TTL) || 300
  },
  rateLimit: {
    windowMs: parseInt(process.env.RATE_LIMIT_WINDOW_MS) || 900000,
    max: parseInt(process.env.RATE_LIMIT_MAX_REQUESTS) || 100
  },
  logging: {
    level: process.env.LOG_LEVEL || 'info'
  }
};