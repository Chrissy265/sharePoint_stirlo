const sharepointService = require('../services/sharepoint.service');
const logger = require('../utils/logger');
const Joi = require('joi');

class ItemsController {
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

  async createItem(req, res, next) {
    try {
      const { listTitle } = req.params;
      const itemData = req.body;

      // Validate request body
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

  async searchItems(req, res, next) {
    try {
      const { query } = req.query;

      if (!query) {
        return res.status(400).json({
          success: false,
          error: 'Query parameter is required'
        });
      }

      const data = await sharepointService.search(query, req.query);
      
      res.json({
        success: true,
        data: data.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results
      });
    } catch (error) {
      next(error);
    }
  }
}

module.exports = new ItemsController();