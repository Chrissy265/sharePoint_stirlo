const express = require('express');
const router = express.Router();
const listsController = require('../controllers/lists.controller');

router.get('/', listsController.getAllLists);
router.get('/:listTitle', listsController.getListByTitle);

module.exports = router;