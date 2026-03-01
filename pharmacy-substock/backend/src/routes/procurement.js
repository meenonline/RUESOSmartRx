const router = require('express').Router();
const db = require('../db');

// GET / — drugs below min stock
router.get('/', async (req, res, next) => {
  try {
    const { rows } = await db.query(
      `SELECT n.*, f.min_stock, f.cabinet
       FROM stock_summary_nolot n
       JOIN formulary f USING(drug_name)
       WHERE n.balance < f.min_stock
       ORDER BY drug_name`
    );
    res.json({ success: true, data: rows });
  } catch (err) { next(err); }
});

module.exports = router;
