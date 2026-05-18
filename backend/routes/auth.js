const express = require('express');
const User = require('../models/User');

const router = express.Router();

// Login Route
router.post('/login', async (req, res) => {
  try {
    const { username, password } = req.body;

    // Find the user by username
    const user = await User.findOne({ username });
    
    if (user && (await user.matchPassword(password))) {
      // Send back user data (excluding password)
      res.json({
        success: true,
        user: {
          _id: user._id,
          username: user.username,
          name: user.name,
          role: user.role
        }
      });
    } else {
      res.status(401).json({ success: false, message: 'Invalid username or password' });
    }
  } catch (error) {
    res.status(500).json({ success: false, message: 'Server error during login', error: error.message });
  }
});

module.exports = router;
