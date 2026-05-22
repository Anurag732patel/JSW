const dns = require('node:dns');
dns.setServers(['8.8.8.8', '8.8.4.4']);

const express = require('express');
const cors = require('cors');
const mongoose = require('mongoose');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 5000;

// Middleware
app.use(cors());
app.use(express.json());

// Routes
const authRoutes = require('./routes/auth');
const dprRoutes = require('./routes/dpr');

app.use('/api/auth', authRoutes);
app.use('/api/dpr', dprRoutes);

// Health check route to verify DB status
app.get('/api/health', (req, res) => {
  const dbState = mongoose.connection.readyState;
  const states = { 0: 'disconnected', 1: 'connected', 2: 'connecting', 3: 'disconnecting' };
  res.json({ 
    server: 'running', 
    database: states[dbState] || 'unknown',
    dbReady: dbState === 1
  });
});

// Database Connection Strategy (Atlas first, automatic fallback to Local if blocked)
const connectDB = async () => {
  const localURI = "mongodb://127.0.0.1:27017/jsw_db";
  
  try {
    console.log("📡 Connecting to MongoDB Atlas cloud database...");
    await mongoose.connect(process.env.MONGODB_URI, {
      serverSelectionTimeoutMS: 5000, // Faster failure detection for blocked networks
    });
    console.log("✅ MongoDB Atlas connected successfully!");
  } catch (err) {
    console.log("⚠️ MongoDB Atlas connection failed (likely port 27017 blocked by current network/ISP).");
    console.log("🔄 Automatically falling back to Local MongoDB...");
    try {
      await mongoose.connect(localURI, {
        serverSelectionTimeoutMS: 3000,
      });
      console.log("✅ Local MongoDB connected successfully!");
    } catch (localErr) {
      console.error("❌ Both MongoDB Atlas and Local MongoDB connections failed:", localErr.message);
    }
  }
};

connectDB();

app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
