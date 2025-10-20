const axios = require('axios');
const NodeCache = require('node-cache');
const config = require('../config/sharepoint.config');
const logger = require('../utils/logger');

// Token cache (TTL: 58 minutes - tokens expire in 60 mins)
const tokenCache = new NodeCache({ stdTTL: config.cache.tokenTTL });

// Data cache for frequently accessed data
const dataCache = new NodeCache({ stdTTL: config.cache.dataTTL });

class SharePointService {
  /**
   * Get SharePoint access token (cached)
   */
  async getAccessToken() {
    const cachedToken = tokenCache.get('access_token');
    if (cachedToken) {
      logger.debug('Using cached SharePoint token');
      return cachedToken;
    }

    logger.info('Requesting new SharePoint access token');

    try {
      const response = await axios.post(
        config.sharepoint.tokenEndpoint,
        new URLSearchParams({
          grant_type: 'client_credentials',
          client_id: config.sharepoint.clientId,
          client_secret: config.sharepoint.clientSecret,
          resource: config.sharepoint.resource
        }),
        {
          headers: {
            'Content-Type': 'application/x-www-form-urlencoded'
          }
        }
      );

      const token = response.data.access_token;
      tokenCache.set('access_token', token);
      
      logger.info('Successfully obtained SharePoint access token');
      return token;
    } catch (error) {
      logger.error('Failed to get SharePoint access token', {
        error: error.message,
        response: error.response?.data
      });
      throw new Error(`Token acquisition failed: ${error.response?.data?.error_description || error.message}`);
    }
  }

  /**
   * Make authenticated request to SharePoint
   */
  async makeRequest(method, endpoint, data = null, options = {}) {
    const token = await this.getAccessToken();
    const url = `${config.sharepoint.siteUrl}${endpoint}`;

    const requestConfig = {
      method,
      url,
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        ...options.headers
      },
      ...options
    };

    if (data) {
      requestConfig.data = data;
    }

    try {
      logger.debug(`SharePoint ${method} request`, { url, endpoint });
      const response = await axios(requestConfig);
      return response.data;
    } catch (error) {
      logger.error(`SharePoint ${method} request failed`, {
        url,
        error: error.message,
        status: error.response?.status,
        data: error.response?.data
      });
      throw error;
    }
  }

  /**
   * Get all lists
   */
  async getLists() {
    const cacheKey = 'all_lists';
    const cached = dataCache.get(cacheKey);
    if (cached) return cached;

    const data = await this.makeRequest('GET', '/_api/web/lists');
    dataCache.set(cacheKey, data);
    return data;
  }

  /**
   * Get specific list by title
   */
  async getListByTitle(listTitle) {
    const cacheKey = `list_${listTitle}`;
    const cached = dataCache.get(cacheKey);
    if (cached) return cached;

    const data = await this.makeRequest('GET', `/_api/web/lists/getbytitle('${listTitle}')`);
    dataCache.set(cacheKey, data);
    return data;
  }

  /**
   * Get list items
   */
  async getListItems(listTitle, options = {}) {
    const { filter, select, top, skip, orderby } = options;
    
    let endpoint = `/_api/web/lists/getbytitle('${listTitle}')/items`;
    const params = [];

    if (filter) params.push(`$filter=${filter}`);
    if (select) params.push(`$select=${select}`);
    if (top) params.push(`$top=${top}`);
    if (skip) params.push(`$skip=${skip}`);
    if (orderby) params.push(`$orderby=${orderby}`);

    if (params.length > 0) {
      endpoint += `?${params.join('&')}`;
    }

    const cacheKey = `items_${listTitle}_${endpoint}`;
    const cached = dataCache.get(cacheKey);
    if (cached) return cached;

    const data = await this.makeRequest('GET', endpoint);
    dataCache.set(cacheKey, data);
    return data;
  }

  /**
   * Get single item by ID
   */
  async getItemById(listTitle, itemId) {
    const cacheKey = `item_${listTitle}_${itemId}`;
    const cached = dataCache.get(cacheKey);
    if (cached) return cached;

    const data = await this.makeRequest('GET', `/_api/web/lists/getbytitle('${listTitle}')/items(${itemId})`);
    dataCache.set(cacheKey, data);
    return data;
  }

  /**
   * Create new list item
   */
  async createItem(listTitle, itemData) {
    // Get list metadata
    const listInfo = await this.getListByTitle(listTitle);
    const listItemEntityTypeFullName = listInfo.d.ListItemEntityTypeFullName;

    const data = {
      __metadata: {
        type: listItemEntityTypeFullName
      },
      ...itemData
    };

    // Get form digest for POST requests
    const digestData = await this.makeRequest('POST', '/_api/contextinfo');
    const formDigestValue = digestData.d.GetContextWebInformation.FormDigestValue;

    const result = await this.makeRequest(
      'POST',
      `/_api/web/lists/getbytitle('${listTitle}')/items`,
      data,
      {
        headers: {
          'X-RequestDigest': formDigestValue
        }
      }
    );

    // Invalidate cache
    dataCache.del(`items_${listTitle}_*`);
    
    return result;
  }

  /**
   * Update list item
   */
  async updateItem(listTitle, itemId, itemData) {
    // Get list metadata
    const listInfo = await this.getListByTitle(listTitle);
    const listItemEntityTypeFullName = listInfo.d.ListItemEntityTypeFullName;

    // Get current item to get etag
    const currentItem = await this.getItemById(listTitle, itemId);
    const etag = currentItem.d.__metadata.etag;

    const data = {
      __metadata: {
        type: listItemEntityTypeFullName
      },
      ...itemData
    };

    // Get form digest
    const digestData = await this.makeRequest('POST', '/_api/contextinfo');
    const formDigestValue = digestData.d.GetContextWebInformation.FormDigestValue;

    const result = await this.makeRequest(
      'POST',
      `/_api/web/lists/getbytitle('${listTitle}')/items(${itemId})`,
      data,
      {
        headers: {
          'X-RequestDigest': formDigestValue,
          'X-HTTP-Method': 'MERGE',
          'IF-MATCH': etag || '*'
        }
      }
    );

    // Invalidate cache
    dataCache.del(`item_${listTitle}_${itemId}`);
    dataCache.del(`items_${listTitle}_*`);

    return result;
  }

  /**
   * Delete list item
   */
  async deleteItem(listTitle, itemId) {
    // Get current item to get etag
    const currentItem = await this.getItemById(listTitle, itemId);
    const etag = currentItem.d.__metadata.etag;

    // Get form digest
    const digestData = await this.makeRequest('POST', '/_api/contextinfo');
    const formDigestValue = digestData.d.GetContextWebInformation.FormDigestValue;

    const result = await this.makeRequest(
      'POST',
      `/_api/web/lists/getbytitle('${listTitle}')/items(${itemId})`,
      null,
      {
        headers: {
          'X-RequestDigest': formDigestValue,
          'X-HTTP-Method': 'DELETE',
          'IF-MATCH': etag || '*'
        }
      }
    );

    // Invalidate cache
    dataCache.del(`item_${listTitle}_${itemId}`);
    dataCache.del(`items_${listTitle}_*`);

    return result;
  }

  /**
   * Search SharePoint
   */
  async search(query, options = {}) {
    const { selectproperties, rowlimit = 50, startrow = 0 } = options;
    
    let endpoint = `/_api/search/query?querytext='${encodeURIComponent(query)}'`;
    endpoint += `&rowlimit=${rowlimit}&startrow=${startrow}`;
    
    if (selectproperties) {
      endpoint += `&selectproperties='${selectproperties}'`;
    }

    return await this.makeRequest('GET', endpoint);
  }

  /**
   * Clear all caches
   */
  clearCache() {
    tokenCache.flushAll();
    dataCache.flushAll();
    logger.info('All caches cleared');
  }
}

module.exports = new SharePointService();