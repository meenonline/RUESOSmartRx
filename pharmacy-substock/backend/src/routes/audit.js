const router = require('express').Router();
const db = require('../db');

// GET /random — get random drugs for audit
router.get('/random', async (req, res, next) => {
  try {
    const { rows: allDrugs } = await db.query(
      `SELECT n.*, f.cabinet, f.shelf_position, f.min_stock
       FROM stock_summary_nolot n
       LEFT JOIN formulary f USING(drug_name)`
    );

    const critical = allDrugs.filter(d => {
      const min = d.min_stock || 0;
      return parseInt(d.balance, 10) < min && min > 0;
    });

    const normal = allDrugs
      .filter(d => !critical.includes(d))
      .sort(() => Math.random() - 0.5)
      .slice(0, 5);

    res.json({ success: true, data: { critical, normal } });
  } catch (err) { next(err); }
});

// POST /save — save audit results
router.post('/save', async (req, res, next) => {
  try {
    const { rows } = req.body;
    if (!rows || !rows.length) {
      return res.status(400).json({ success: false, error: 'No audit rows provided' });
    }

    const inserted = [];
    for (const row of rows) {
      const { rows: result } = await db.query(
        `INSERT INTO audit_trail
          (drug_name, system_qty, counted_qty, diff, cabinet, shelf_position, lot_no, recorder)
         VALUES ($1,$2,$3,$4,$5,$6,$7,$8)
         RETURNING *`,
        [
          row.drug_name, row.system_qty, row.counted_qty,
          row.diff, row.cabinet, row.shelf_position,
          row.lot_no, row.recorder
        ]
      );
      inserted.push(result[0]);
    }
    res.json({ success: true, data: inserted });
  } catch (err) { next(err); }
});

module.exports = router;
