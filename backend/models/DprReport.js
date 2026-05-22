const mongoose = require('mongoose');

const dprReportSchema = new mongoose.Schema({
  reportDate: {
    type: String, // String format YYYY-MM-DD prevents timezone issues and is highly indexed
    required: true,
    unique: true,
    trim: true,
    index: true,
  },
  filename: {
    type: String,
    required: true,
  },
  uploadedBy: {
    type: String,
    default: 'System Administrator',
  },
  data: {
    type: mongoose.Schema.Types.Mixed,
    required: true,
  }
}, { timestamps: true });

const DprReport = mongoose.model('DprReport', dprReportSchema);
module.exports = DprReport;
