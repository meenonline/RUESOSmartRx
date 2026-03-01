const router = require('express').Router();
const db = require('../db');

// GET / — all settings
router.get('/', async (req, res, next) => {
  try {
    const { rows } = await db.query('SELECT * FROM settings');
    const obj = {};
    rows.forEach(r => { obj[r.key] = r.value; });
    res.json({ success: true, data: obj });
  } catch (err) { next(err); }
});

// POST / — upsert a setting
router.post('/', async (req, res, next) => {
  try {
    const { key, value } = req.body;
    await db.query(
      `INSERT INTO settings(key, value) VALUES($1, $2)
       ON CONFLICT(key) DO UPDATE SET value=EXCLUDED.value`,
      [key, value]
    );
    res.json({ success: true, data: { key, value } });
  } catch (err) { next(err); }
});

module.exports = router;
