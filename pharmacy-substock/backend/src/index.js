require('dotenv').config();
const express = require('express');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 3001;

// Middleware
app.use(cors({
  origin: process.env.FRONTEND_URL || '*',
  credentials: true,
}));
app.use(express.json({ limit: '10mb' }));

// Health check
app.get('/api/health', (_req, res) => {
  res.json({ success: true, data: { ok: true } });
});

// Routes
app.use('/api/transactions', require('./routes/transactions'));
app.use('/api/stock', require('./routes/stock'));
app.use('/api/import', require('./routes/import'));
app.use('/api/formulary', require('./routes/formulary'));
app.use('/api/dashboard', require('./routes/dashboard'));
app.use('/api/expiry', require('./routes/expiry'));
app.use('/api/audit', require('./routes/audit'));
app.use('/api/procurement', require('./routes/procurement'));
app.use('/api/settings', require('./routes/settings'));
app.use('/api/line', require('./routes/line'));
app.use('/api/email', require('./routes/email'));
app.use('/api/disp-no', require('./routes/disp-no'));
app.use('/api/drug-list', require('./routes/drug-list'));

// Error handling middleware
app.use((err, _req, res, _next) => {
  console.error('Unhandled error:', err);
  res.status(500).json({ success: false, error: err.message || 'Internal server error' });
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
