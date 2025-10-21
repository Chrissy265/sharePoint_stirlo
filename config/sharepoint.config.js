require('dotenv').config();

// Validate required environment variables on startup
const requiredEnvVars = [
  'SHAREPOINT_TENANT_ID',
  'SHAREPOINT_CLIENT_ID',
  'SHAREPOINT_CLIENT_SECRET',
  'SHAREPOINT_SITE_URL',
  'API_KEY'
];

const missingVars = requiredEnvVars.filter(varName => !process.env[varName]);

if (missingVars.length > 0) {
  console.error('========================================');
  console.error('❌ CONFIGURATION ERROR - Missing Required Environment Variables');
  console.error('========================================');
  console.error('The following environment variables are required but not set:');
  missingVars.forEach(varName => console.error(`  - ${varName}`));
  console.error('\nPlease set these variables in your Render dashboard or .env file');
  console.error('See .env.example for reference');
  console.error('========================================');
  process.exit(1);
}

// Validate and parse numeric environment variables
const parseIntSafe = (value, defaultValue, varName) => {
  if (!value) return defaultValue;
  const parsed = parseInt(value);
  if (isNaN(parsed)) {
    console.error(`❌ CONFIGURATION ERROR: ${varName} must be a number, got "${value}"`);
    process.exit(1);
  }
  return parsed;
};

module.exports = {
  server: {
    port: process.env.PORT || 10000,
    env: process.env.NODE_ENV || 'development',
    apiKey: process.env.API_KEY
  },
  sharepoint: {
    tenantId: process.env.SHAREPOINT_TENANT_ID,
    clientId: process.env.SHAREPOINT_CLIENT_ID,
    clientSecret: process.env.SHAREPOINT_CLIENT_SECRET,
    siteUrl: process.env.SHAREPOINT_SITE_URL,
    tenantName: process.env.SHAREPOINT_TENANT_NAME,
    // CRITICAL: Use v1.0 endpoint
    tokenEndpoint: `https://login.microsoftonline.com/${process.env.SHAREPOINT_TENANT_ID}/oauth2/token`,
    resource: process.env.SHAREPOINT_SITE_URL
  },
  cache: {
    tokenTTL: parseIntSafe(process.env.TOKEN_CACHE_TTL, 3500, 'TOKEN_CACHE_TTL'),
    dataTTL: parseIntSafe(process.env.DATA_CACHE_TTL, 300, 'DATA_CACHE_TTL')
  },
  rateLimit: {
    windowMs: parseIntSafe(process.env.RATE_LIMIT_WINDOW_MS, 900000, 'RATE_LIMIT_WINDOW_MS'),
    max: parseIntSafe(process.env.RATE_LIMIT_MAX_REQUESTS, 100, 'RATE_LIMIT_MAX_REQUESTS')
  },
  logging: {
    level: process.env.LOG_LEVEL || 'info'
  }
};