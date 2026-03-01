const router = require('express').Router();
const db = require('../db');

// GET /summary — stock summary by lot
router.get('/summary', async (req, res, next) => {
  try {
    const { rows } = await db.query(
      'SELECT * FROM stock_summary_lot ORDER BY drug_name'
    );
    res.json({ success: true, data: rows });
  } catch (err) { next(err); }
});

// GET /summary-nolot — stock summary without lot
router.get('/summary-nolot', async (req, res, next) => {
  try {
    const { rows } = await db.query(
      'SELECT * FROM stock_summary_nolot ORDER BY drug_name'
    );
    res.json({ success: true, data: rows });
  } catch (err) { next(err); }
});

// POST /adjust — adjust single drug stock
router.post('/adjust', async (req, res, next) => {
  try {
    const { drug_name, drug_code, pack_size, price_per_unit, new_balance, recorder, lot_no, exp_date } = req.body;

    // Get current balance
    const { rows: current } = await db.query(
      'SELECT COALESCE(SUM(qty),0) AS balance FROM transactions WHERE drug_name=$1',
      [drug_name]
    );
    const currentBal = parseInt(current[0].balance, 10);
    const diff = new_balance - currentBal;

    if (diff === 0) {
      return res.json({ success: true, data: null, message: 'No adjustment needed' });
    }

    const { rows } = await db.query(
      `INSERT INTO transactions
        (drug_name, drug_code, pack_size, price_per_unit, qty, txn_type, recorder, lot_no, exp_date, remark)
       VALUES ($1,$2,$3,$4,$5,'adjust',$6,$7,$8,'Stock adjustment')
       RETURNING *`,
      [drug_name, drug_code, pack_size, price_per_unit, diff, recorder, lot_no, exp_date || null]
    );
    res.json({ success: true, data: rows[0] });
  } catch (err) { next(err); }
});

// POST /batch-adjust — adjust multiple drugs
router.post('/batch-adjust', async (req, res, next) => {
  try {
    const { adjustments, recorder } = req.body;
    if (!adjustments || !adjustments.length) {
      return res.status(400).json({ success: false, error: 'No adjustments provided' });
    }

    const results = [];
    for (const adj of adjustments) {
      const { drug_name, drug_code, pack_size, price_per_unit, new_balance } = adj;

      const { rows: current } = await db.query(
        'SELECT COALESCE(SUM(qty),0) AS balance FROM transactions WHERE drug_name=$1',
        [drug_name]
      );
      const currentBal = parseInt(current[0].balance, 10);
      const diff = new_balance - currentBal;

      if (diff !== 0) {
        const { rows } = await db.query(
          `INSERT INTO transactions
            (drug_name, drug_code, pack_size, price_per_unit, qty, txn_type, recorder, remark)
           VALUES ($1,$2,$3,$4,$5,'adjust',$6,'Batch adjustment')
           RETURNING *`,
          [drug_name, drug_code, pack_size, price_per_unit, diff, recorder]
        );
        results.push(rows[0]);
      }
    }
    res.json({ success: true, data: results });
  } catch (err) { next(err); }
});

module.exports = router;
