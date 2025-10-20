const sharepointService = require('../services/sharepoint.service');
const logger = require('../utils/logger');

class AuthController {
  async getToken(req, res, next) {
    try {
      const token = await sharepointService.getAccessToken();
      
      res.json({
        success: true,
        data: {
          access_token: token,
          token_type: 'Bearer',
          expires_in: 3600
        }
      });
    } catch (error) {
      next(error);
    }
  }

  async validateToken(req, res, next) {
    try {
      const token = await sharepointService.getAccessToken();
      
      // Try a simple request to validate
      await sharepointService.makeRequest('GET', '/_api/web');
      
      res.json({
        success: true,
        message: 'Token is valid',
        data: {
          valid: true
        }
      });
    } catch (error) {
      res.status(401).json({
        success: false,
        message: 'Token is invalid or expired',
        data: {
          valid: false
        }
      });
    }
  }
}

module.exports = new AuthController();