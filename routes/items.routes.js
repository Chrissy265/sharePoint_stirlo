const express = require('express');
const router = express.Router();
const itemsController = require('../controllers/items.controller');

// Smart search (natural language) - MUST be before /:listTitle routes
router.get('/smart-search', itemsController.smartSearch);

// Specific search endpoints
router.get('/search/type', itemsController.searchByType);
router.get('/search/author', itemsController.searchByAuthor);
router.post('/search/advanced', itemsController.advancedSearch);
router.get('/search', itemsController.searchItems); 

// Analytics endpoints
router.get('/recent', itemsController.getRecentFiles);
router.get('/statistics', itemsController.getStatistics);

// Folder operations
router.get('/folder/:path(*)', itemsController.getFolderContents);

// Basic CRUD routes (parameterized routes must come LAST)
router.get('/:listTitle/items', itemsController.getItems);
router.get('/:listTitle/items/:itemId', itemsController.getItemById);
router.post('/:listTitle/items', itemsController.createItem);
router.put('/:listTitle/items/:itemId', itemsController.updateItem);
router.delete('/:listTitle/items/:itemId', itemsController.deleteItem);





module.exports = router;
