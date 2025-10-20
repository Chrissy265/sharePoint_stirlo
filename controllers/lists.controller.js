const sharepointService = require('../services/sharepoint.service');
const logger = require('../utils/logger');

class ListsController {
  async getAllLists(req, res, next) {
    try {
      const data = await sharepointService.getLists();
      
      res.json({
        success: true,
        data: data.d.results,
        count: data.d.results.length
      });
    } catch (error) {
      next(error);
    }
  }

  async getListByTitle(req, res, next) {
    try {
      const { listTitle } = req.params;
      const data = await sharepointService.getListByTitle(listTitle);
      
      res.json({
        success: true,
        data: data.d
      });
    } catch (error) {
      next(error);
    }
  }
}

module.exports = new ListsController();