const express = require('express');
const router = express.Router();
const authController = require('../controllers/auth.controller');

router.get('/token', authController.getToken);
router.get('/validate', authController.validateToken);

module.exports = router;