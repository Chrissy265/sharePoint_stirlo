const express = require('express');
const router = express.Router();
const itemsController = require('../controllers/items.controller');

router.get('/:listTitle/items', itemsController.getItems);
router.get('/:listTitle/items/:itemId', itemsController.getItemById);
router.post('/:listTitle/items', itemsController.createItem);
router.put('/:listTitle/items/:itemId', itemsController.updateItem);
router.patch('/:listTitle/items/:itemId', itemsController.updateItem);
router.delete('/:listTitle/items/:itemId', itemsController.deleteItem);
router.get('/search', itemsController.searchItems);

module.exports = router;