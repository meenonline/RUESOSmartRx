const router = require('express').Router();
const db = require('../db');

// GET / — distinct drug list from transactions
router.get('/', async (req, res, next) => {
  try {
    const { rows } = await db.query(
      `SELECT DISTINCT ON (drug_name, drug_code, pack_size)
        drug_name, drug_code, pack_size, price_per_unit
       FROM transactions
       ORDER BY drug_name, drug_code, pack_size, created_at DESC`
    );
    res.json({ success: true, data: rows });
  } catch (err) { next(err); }
});

module.exports = router;
