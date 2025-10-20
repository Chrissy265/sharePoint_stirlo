const express = require('express');
const router = express.Router();
const itemsController = require('../controllers/items.controller');

// Smart search (natural language)
router.get('/smart-search', itemsController.smartSearch);

// Specific search endpoints
router.get('/search/type', itemsController.searchByType);
router.get('/search/author', itemsController.searchByAuthor);
router.get('/recent', itemsController.getRecentFiles);
router.get('/statistics', itemsController.getStatistics);

// Advanced search with multiple criteria
router.post('/search/advanced', itemsController.advancedSearch);

// Folder operations
router.get('/folder/*', itemsController.getFolderContents);

// Original endpoints
router.get('/:listTitle/items', itemsController.getItems);
router.get('/:listTitle/items/:itemId', itemsController.getItemById);
router.post('/:listTitle/items', itemsController.createItem);
router.put('/:listTitle/items/:itemId', itemsController.updateItem);
router.patch('/:listTitle/items/:itemId', itemsController.updateItem);
router.delete('/:listTitle/items/:itemId', itemsController.deleteItem);

module.exports = router;
