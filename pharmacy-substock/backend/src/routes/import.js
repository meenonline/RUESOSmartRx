const router = require('express').Router();
const db = require('../db');

// POST /preview — preview import rows against current stock
router.post('/preview', async (req, res, next) => {
  try {
    const { rows } = req.body;
    if (!rows || !rows.length) {
      return res.status(400).json({ success: false, error: 'No rows provided' });
    }

    // Get current stock summary
    const { rows: stock } = await db.query(
      'SELECT * FROM stock_summary_nolot'
    );
    const stockMap = {};
    stock.forEach(s => { stockMap[s.drug_name] = s; });

    const result = rows.map(row => {
      const existing = stockMap[row.drug_name];
      return {
        ...row,
        currentBal: existing ? parseInt(existing.balance, 10) : 0,
        match: !!existing,
      };
    });

    res.json({ success: true, data: result });
  } catch (err) { next(err); }
});

// POST /confirm — batch insert import rows
router.post('/confirm', async (req, res, next) => {
  try {
    const { rows, txn_type, recorder } = req.body;
    if (!rows || !rows.length) {
      return res.status(400).json({ success: false, error: 'No rows provided' });
    }

    const inserted = [];
    for (const row of rows) {
      const qty = txn_type === 'dispense' ? -Math.abs(row.qty) : Math.abs(row.qty);
      const { rows: result } = await db.query(
        `INSERT INTO transactions
          (disp_no, drug_code, drug_name, pack_size, price_per_unit,
           lot_no, exp_date, qty, txn_type, dept, recorder)
         VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,'substock',$10)
         RETURNING *`,
        [
          row.disp_no, row.drug_code, row.drug_name, row.pack_size,
          row.price_per_unit, row.lot_no, row.exp_date || null,
          qty, txn_type === 'dispense' ? 'dispense' : 'receive', recorder
        ]
      );
      inserted.push(result[0]);
    }

    res.json({ success: true, data: inserted });
  } catch (err) { next(err); }
});

module.exports = router;
