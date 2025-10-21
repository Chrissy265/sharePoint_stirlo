const sharepointService = require('../services/sharepoint.service');
const logger = require('../utils/logger');
const Joi = require('joi');

class ItemsController {
  /**
   * Get items from a list
   */
  async getItems(req, res, next) {
    try {
      const { listTitle } = req.params;
      const { filter, select, top, skip, orderby } = req.query;

      const data = await sharepointService.getListItems(listTitle, {
        filter,
        select,
        top: top ? parseInt(top) : undefined,
        skip: skip ? parseInt(skip) : undefined,
        orderby
      });
      
      res.json({
        success: true,
        data: data.d.results,
        count: data.d.results.length
      });
    } catch (error) {
      next(error);
    }
  }

  /**
   * Get item by ID
   */
  async getItemById(req, res, next) {
    try {
      const { listTitle, itemId } = req.params;
      const data = await sharepointService.getItemById(listTitle, parseInt(itemId));
      
      res.json({
        success: true,
        data: data.d
      });
    } catch (error) {
      next(error);
    }
  }

  /**
   * Create new item
   */
  async createItem(req, res, next) {
    try {
      const { listTitle } = req.params;
      const itemData = req.body;

      if (!itemData || Object.keys(itemData).length === 0) {
        return res.status(400).json({
          success: false,
          error: 'Request body is empty'
        });
      }

      const data = await sharepointService.createItem(listTitle, itemData);
      
      res.status(201).json({
        success: true,
        message: 'Item created successfully',
        data: data.d
      });
    } catch (error) {
      next(error);
    }
  }

  /**
   * Update item
   */
  async updateItem(req, res, next) {
    try {
      const { listTitle, itemId } = req.params;
      const itemData = req.body;

      if (!itemData || Object.keys(itemData).length === 0) {
        return res.status(400).json({
          success: false,
          error: 'Request body is empty'
        });
      }

      await sharepointService.updateItem(listTitle, parseInt(itemId), itemData);
      
      res.json({
        success: true,
        message: 'Item updated successfully'
      });
    } catch (error) {
      next(error);
    }
  }

  /**
   * Delete item
   */
  async deleteItem(req, res, next) {
    try {
      const { listTitle, itemId } = req.params;
      await sharepointService.deleteItem(listTitle, parseInt(itemId));
      
      res.json({
        success: true,
        message: 'Item deleted successfully'
      });
    } catch (error) {
      next(error);
    }
  }

  /**
   * Smart search endpoint (handles natural language queries)
   */
  async smartSearch(req, res, next) {
    try {
      const { query, library = 'Documents' } = req.query;

      if (!query) {
        return res.status(400).json({
          success: false,
          error: 'Query parameter is required'
        });
      }

      logger.info('Smart search request', { query, library });

      const data = await sharepointService.smartSearch(query, library);
      
      res.json({
        success: true,
        query: query,
        data: data.d?.results || data,
        count: data.d?.results?.length || 0
      });
    } catch (error) {
      next(error);
    }
  }

  /**
   * Search by file type
   */
  async searchByType(req, res, next) {
    try {
      const { type, library = 'Documents' } = req.query;

      if (!type) {
        return res.status(400).json({
          success: false,
          error: 'File type parameter is required'
        });
      }

      const data = await sharepointService.searchFilesByType(type, library);
      
      res.json({
        success: true,
        fileType: type,
        data: data.d.results,
        count: data.d.results.length
      });
    } catch (error) {
      next(error);
    }
  }

  /**
   * Search by author
   */
  async searchByAuthor(req, res, next) {
    try {
      const { author, library = 'Documents' } = req.query;

      if (!author) {
        return res.status(400).json({
          success: false,
          error: 'Author parameter is required'
        });
      }

      const data = await sharepointService.searchFilesByAuthor(author, library);
      
      res.json({
        success: true,
        author: author,
        data: data.d.results,
        count: data.d.results.length
      });
    } catch (error) {
      next(error);
    }
  }

  /**
   * Get file statistics
   */
  async getStatistics(req, res, next) {
    try {
      const { library = 'Documents' } = req.query;
      const stats = await sharepointService.getFileStatistics(library);
      
      res.json({
        success: true,
        library: library,
        statistics: stats
      });
    } catch (error) {
      next(error);
    }
  }

  /**
   * Get recent files
   */
  async getRecentFiles(req, res, next) {
    try {
      const { count = 10, library = 'Documents' } = req.query;
      const data = await sharepointService.getRecentFiles(parseInt(count), library);
      
      res.json({
        success: true,
        data: data.d.results,
        count: data.d.results.length
      });
    } catch (error) {
      next(error);
    }
  }

  /**
   * Get folder contents
   */
  async getFolderContents(req, res, next) {
    try {
      const { path } = req.params;
      const { library = 'Documents' } = req.query;

      if (!path) {
        return res.status(400).json({
          success: false,
          error: 'Folder path is required'
        });
      }

      const data = await sharepointService.getFilesInFolder(path, library);
      
      res.json({
        success: true,
        folderPath: path,
        data: data.d.results,
        count: data.d.results.length
      });
    } catch (error) {
      next(error);
    }
  }

  /**
   * Multi-criteria search
   */
  async advancedSearch(req, res, next) {
    try {
      const { library = 'Documents' } = req.query;
      const criteria = req.body;

      if (!criteria || Object.keys(criteria).length === 0) {
        return res.status(400).json({
          success: false,
          error: 'Search criteria are required in request body'
        });
      }

      const data = await sharepointService.searchMultiCriteria(criteria, library);

      res.json({
        success: true,
        criteria: criteria,
        data: data.d?.results || data,
        count: data.d?.results?.length || 0
      });
    } catch (error) {
      next(error);
    }
  }

  /**
   * General search items (keyword search)
   */
  async searchItems(req, res, next) {
    try {
      const { keyword, library = 'Documents' } = req.query;

      if (!keyword) {
        return res.status(400).json({
          success: false,
          error: 'Keyword parameter is required'
        });
      }

      const data = await sharepointService.searchByKeyword(keyword, library);

      res.json({
        success: true,
        keyword: keyword,
        data: data.d?.results || data,
        count: data.d?.results?.length || 0
      });
    } catch (error) {
      next(error);
    }
  }
}

module.exports = new ItemsController();