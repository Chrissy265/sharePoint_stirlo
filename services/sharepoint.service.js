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
   * Get all folders in a library
   */
  async getFolders(libraryName = 'Documents') {
    const cacheKey = `folders_${libraryName}`;
    const cached = dataCache.get(cacheKey);
    if (cached) return cached;

    const endpoint = `/_api/web/lists/getbytitle('${libraryName}')/folders`;
    const data = await this.makeRequest('GET', endpoint);
    dataCache.set(cacheKey, data);
    return data;
  }

  /**
   * Get folder contents by path
   */
  async getFolderContents(folderPath) {
    const cacheKey = `folder_contents_${folderPath}`;
    const cached = dataCache.get(cacheKey);
    if (cached) return cached;

    // Get folder by server relative URL
    const endpoint = `/_api/web/GetFolderByServerRelativeUrl('${folderPath}')?$expand=Folders,Files`;
    const data = await this.makeRequest('GET', endpoint);
    dataCache.set(cacheKey, data);
    return data;
  }

  /**
   * Search for folders by name (supports partial matching)
   */
  async searchFolders(searchTerm, libraryName = 'Documents') {
    const folders = await this.getFolders(libraryName);
    const results = folders.d.results.filter(folder => 
      folder.Name.toLowerCase().includes(searchTerm.toLowerCase())
    );
    return { d: { results } };
  }

  /**
   * Get all files in a library with metadata
   */
  async getAllFiles(libraryName = 'Documents', options = {}) {
    const { select, expand, filter, top, orderby } = options;
    
    let endpoint = `/_api/web/lists/getbytitle('${libraryName}')/items`;
    const params = [];

    // Default expansions for file metadata
    const defaultExpand = 'File,Folder,Author,Editor';
    params.push(`$expand=${expand || defaultExpand}`);

    // Default selections for useful metadata
    const defaultSelect = 'Id,Title,FileLeafRef,FileRef,FileDirRef,File/Name,File/ServerRelativeUrl,File/TimeLastModified,File/Length,Author/Title,Editor/Title,Modified,Created';
    params.push(`$select=${select || defaultSelect}`);

    if (filter) params.push(`$filter=${filter}`);
    if (top) params.push(`$top=${top}`);
    if (orderby) params.push(`$orderby=${orderby}`);

    endpoint += `?${params.join('&')}`;

    const cacheKey = `all_files_${libraryName}_${endpoint}`;
    const cached = dataCache.get(cacheKey);
    if (cached) return cached;

    const data = await this.makeRequest('GET', endpoint);
    dataCache.set(cacheKey, data);
    return data;
  }

  /**
   * Search files by name (Level 1: Basic file finding)
   */
  async searchFilesByName(fileName, libraryName = 'Documents') {
    const filter = `substringof('${fileName}', FileLeafRef)`;
    return await this.getAllFiles(libraryName, { filter, top: 100 });
  }

  /**
   * Search files by file type/extension (Level 2: Filter by type)
   */
  async searchFilesByType(fileExtension, libraryName = 'Documents') {
    // Remove leading dot if present
    const ext = fileExtension.replace(/^\./, '');
    const filter = `endswith(FileLeafRef, '.${ext}')`;
    return await this.getAllFiles(libraryName, { filter, top: 500 });
  }

  /**
   * Search files by multiple types (e.g., ["docx", "doc", "pdf"])
   */
  async searchFilesByTypes(fileExtensions, libraryName = 'Documents') {
    const filters = fileExtensions.map(ext => {
      const cleanExt = ext.replace(/^\./, '');
      return `endswith(FileLeafRef, '.${cleanExt}')`;
    });
    const filter = filters.join(' or ');
    return await this.getAllFiles(libraryName, { filter, top: 500 });
  }

  /**
   * Search files by date range (Level 2: Date filtering)
   */
  async searchFilesByDateRange(startDate, endDate, libraryName = 'Documents') {
    let filter = '';
    
    if (startDate && endDate) {
      filter = `Modified ge datetime'${startDate}' and Modified le datetime'${endDate}'`;
    } else if (startDate) {
      filter = `Modified ge datetime'${startDate}'`;
    } else if (endDate) {
      filter = `Modified le datetime'${endDate}'`;
    }

    return await this.getAllFiles(libraryName, { 
      filter, 
      top: 500,
      orderby: 'Modified desc'
    });
  }

  /**
   * Search files modified in specific month (Level 2)
   */
  async searchFilesByMonth(year, month, libraryName = 'Documents') {
    const startDate = new Date(year, month - 1, 1).toISOString();
    const endDate = new Date(year, month, 0).toISOString();
    return await this.searchFilesByDateRange(startDate, endDate, libraryName);
  }

  /**
   * Search files by author/creator (Level 2: Search by creator)
   */
  async searchFilesByAuthor(authorName, libraryName = 'Documents') {
    const filter = `Author/Title eq '${authorName}'`;
    return await this.getAllFiles(libraryName, { 
      filter,
      top: 500,
      orderby: 'Modified desc'
    });
  }

  /**
   * Search files by editor (last modified by)
   */
  async searchFilesByEditor(editorName, libraryName = 'Documents') {
    const filter = `Editor/Title eq '${editorName}'`;
    return await this.getAllFiles(libraryName, { 
      filter,
      top: 500,
      orderby: 'Modified desc'
    });
  }

  /**
   * Advanced keyword search (Level 2: Keyword search)
   * Searches in file names and titles
   */
  async searchByKeyword(keyword, libraryName = 'Documents') {
    const filter = `(substringof('${keyword}', FileLeafRef) or substringof('${keyword}', Title))`;
    return await this.getAllFiles(libraryName, { 
      filter,
      top: 200,
      orderby: 'Modified desc'
    });
  }

  /**
   * Multi-criteria search (Level 3: Combined filters)
   * Example: Word docs in Templates by Nicole Stirling
   */
  async searchMultiCriteria(criteria, libraryName = 'Documents') {
    const filters = [];

    if (criteria.fileName) {
      filters.push(`substringof('${criteria.fileName}', FileLeafRef)`);
    }

    if (criteria.fileType) {
      const ext = criteria.fileType.replace(/^\./, '');
      filters.push(`endswith(FileLeafRef, '.${ext}')`);
    }

    if (criteria.author) {
      filters.push(`Author/Title eq '${criteria.author}'`);
    }

    if (criteria.editor) {
      filters.push(`Editor/Title eq '${criteria.editor}'`);
    }

    if (criteria.folderPath) {
      filters.push(`substringof('${criteria.folderPath}', FileDirRef)`);
    }

    if (criteria.startDate) {
      filters.push(`Modified ge datetime'${criteria.startDate}'`);
    }

    if (criteria.endDate) {
      filters.push(`Modified le datetime'${criteria.endDate}'`);
    }

    if (criteria.keyword) {
      filters.push(`(substringof('${criteria.keyword}', FileLeafRef) or substringof('${criteria.keyword}', Title))`);
    }

    const filter = filters.join(' and ');

    return await this.getAllFiles(libraryName, { 
      filter,
      top: 500,
      orderby: criteria.orderby || 'Modified desc'
    });
  }

  /**
   * Get files in specific folder (Level 1 & 3: Nested queries)
   */
  async getFilesInFolder(folderPath, libraryName = 'Documents') {
    const filter = `FileDirRef eq '${folderPath}'`;
    return await this.getAllFiles(libraryName, { 
      filter,
      top: 500,
      orderby: 'FileLeafRef asc'
    });
  }

  /**
   * Get most recently modified files (Level 2 & 4)
   */
  async getRecentFiles(count = 10, libraryName = 'Documents') {
    return await this.getAllFiles(libraryName, { 
      top: count,
      orderby: 'Modified desc'
    });
  }

  /**
   * Get file statistics (Level 4: Analytical)
   */
  async getFileStatistics(libraryName = 'Documents') {
    const files = await this.getAllFiles(libraryName, { top: 5000 });
    const items = files.d.results;

    const stats = {
      totalFiles: items.length,
      filesByType: {},
      filesByAuthor: {},
      filesByEditor: {},
      filesByFolder: {},
      totalSize: 0,
      largestFile: null,
      newestFile: null,
      oldestFile: null
    };

    items.forEach(item => {
      // File type stats
      if (item.File && item.File.Name) {
        const ext = item.File.Name.split('.').pop().toLowerCase();
        stats.filesByType[ext] = (stats.filesByType[ext] || 0) + 1;

        // Size stats
        if (item.File.Length) {
          stats.totalSize += item.File.Length;
          if (!stats.largestFile || item.File.Length > stats.largestFile.size) {
            stats.largestFile = {
              name: item.File.Name,
              size: item.File.Length,
              url: item.File.ServerRelativeUrl
            };
          }
        }
      }

      // Author stats
      if (item.Author && item.Author.Title) {
        stats.filesByAuthor[item.Author.Title] = (stats.filesByAuthor[item.Author.Title] || 0) + 1;
      }

      // Editor stats
      if (item.Editor && item.Editor.Title) {
        stats.filesByEditor[item.Editor.Title] = (stats.filesByEditor[item.Editor.Title] || 0) + 1;
      }

      // Folder stats
      if (item.FileDirRef) {
        const folderName = item.FileDirRef.split('/').pop();
        stats.filesByFolder[folderName] = (stats.filesByFolder[folderName] || 0) + 1;
      }

      // Date stats
      if (item.Modified) {
        if (!stats.newestFile || new Date(item.Modified) > new Date(stats.newestFile.date)) {
          stats.newestFile = {
            name: item.FileLeafRef,
            date: item.Modified,
            url: item.FileRef
          };
        }
        if (!stats.oldestFile || new Date(item.Modified) < new Date(stats.oldestFile.date)) {
          stats.oldestFile = {
            name: item.FileLeafRef,
            date: item.Modified,
            url: item.FileRef
          };
        }
      }
    });

    return stats;
  }

  /**
   * Smart search with natural language understanding (Level 5: Conversational)
   * This interprets queries and routes to appropriate methods
   */
  async smartSearch(query, libraryName = 'Documents') {
    query = query.toLowerCase();
    
    // Detect query intent
    const intent = this.detectSearchIntent(query);
    
    logger.info('Smart search intent detected', { query, intent });

    switch (intent.type) {
      case 'file_by_name':
        return await this.searchFilesByName(intent.term, libraryName);
      
      case 'file_by_type':
        return await this.searchFilesByType(intent.fileType, libraryName);
      
      case 'file_by_author':
        return await this.searchFilesByAuthor(intent.author, libraryName);
      
      case 'file_by_date':
        if (intent.month && intent.year) {
          return await this.searchFilesByMonth(intent.year, intent.month, libraryName);
        }
        return await this.searchFilesByDateRange(intent.startDate, intent.endDate, libraryName);
      
      case 'folder_contents':
        return await this.getFilesInFolder(intent.folderPath, libraryName);
      
      case 'keyword_search':
        return await this.searchByKeyword(intent.keyword, libraryName);
      
      case 'recent_files':
        return await this.getRecentFiles(intent.count || 10, libraryName);
      
      case 'statistics':
        return await this.getFileStatistics(libraryName);
      
      case 'multi_criteria':
        return await this.searchMultiCriteria(intent.criteria, libraryName);
      
      default:
        // Fallback to keyword search
        return await this.searchByKeyword(query, libraryName);
    }
  }

  /**
   * Detect intent from natural language query
   */
  detectSearchIntent(query) {
    const intent = { type: 'keyword_search', term: query };

    // File type detection
    const fileTypePatterns = {
      'word|docx|doc|word documents?': 'docx',
      'powerpoint|pptx|ppt|presentations?|slides?': 'pptx',
      'excel|xlsx|xls|spreadsheets?': 'xlsx',
      'pdf|pdfs': 'pdf'
    };

    for (const [pattern, type] of Object.entries(fileTypePatterns)) {
      if (new RegExp(pattern, 'i').test(query)) {
        intent.type = 'file_by_type';
        intent.fileType = type;
        return intent;
      }
    }

    // Author/Creator detection
    const authorMatch = query.match(/(?:by|created by|from|authored by)\s+([a-z\s]+)/i);
    if (authorMatch || query.includes('nicole stirling') || query.includes('christine gooding') || query.includes('sajjad')) {
      intent.type = 'file_by_author';
      intent.author = authorMatch ? authorMatch[1].trim() : 
                      query.includes('nicole') ? 'Nicole Stirling' :
                      query.includes('christine') ? 'Christine Gooding' : 'Sajjad';
      return intent;
    }

    // Date detection
    const months = {
      january: 1, february: 2, march: 3, april: 4, may: 5, june: 6,
      july: 7, august: 8, september: 9, october: 10, november: 11, december: 12
    };

    for (const [month, num] of Object.entries(months)) {
      if (query.includes(month)) {
        intent.type = 'file_by_date';
        intent.month = num;
        intent.year = new Date().getFullYear(); // Default to current year
        
        // Check for year in query
        const yearMatch = query.match(/\b(20\d{2})\b/);
        if (yearMatch) {
          intent.year = parseInt(yearMatch[1]);
        }
        return intent;
      }
    }

    // Recent files detection
    if (query.includes('recent') || query.includes('latest') || query.includes('newest')) {
      intent.type = 'recent_files';
      const countMatch = query.match(/(\d+)/);
      intent.count = countMatch ? parseInt(countMatch[1]) : 10;
      return intent;
    }

    // Folder detection
    const folderKeywords = ['templates', 'marketing', 'branding', 'clients', 'internal assets'];
    for (const folder of folderKeywords) {
      if (query.includes(folder)) {
        intent.type = 'folder_contents';
        intent.folderPath = `/sites/YourSite/Documents/${folder}`;
        return intent;
      }
    }

    // Statistics detection
    if (query.includes('statistics') || query.includes('summary') || query.includes('how many') || query.includes('count')) {
      intent.type = 'statistics';
      return intent;
    }

    // Specific file name detection
    if (query.includes('venue research') || query.includes('conference checklist') || query.includes('event signage')) {
      intent.type = 'file_by_name';
      intent.term = query.includes('venue') ? 'Venue Research' :
                    query.includes('conference') ? 'conference checklist' :
                    query.includes('signage') ? 'Event signage' : query;
      return intent;
    }

    return intent;
  }

  /**
   * Get list items (existing method)
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