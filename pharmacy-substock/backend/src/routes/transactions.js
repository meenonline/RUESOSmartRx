const router = require('express').Router();
const db = require('../db');

// GET / — list transactions
router.get('/', async (req, res, next) => {
  try {
    const { rows } = await db.query(
      'SELECT * FROM transactions ORDER BY created_at DESC LIMIT 500'
    );
    res.json({ success: true, data: rows });
  } catch (err) { next(err); }
});

// POST / — create transaction
router.post('/', async (req, res, next) => {
  try {
    const {
      disp_no, drug_code, drug_name, pack_size,
      price_per_unit, lot_no, exp_date, qty,
      txn_type, dept, recorder, remark
    } = req.body;

    const { rows } = await db.query(
      `INSERT INTO transactions
        (disp_no, drug_code, drug_name, pack_size, price_per_unit,
         lot_no, exp_date, qty, txn_type, dept, recorder, remark)
       VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12)
       RETURNING *`,
      [disp_no, drug_code, drug_name, pack_size,
       price_per_unit, lot_no, exp_date || null, qty,
       txn_type, dept || 'substock', recorder, remark]
    );
    res.json({ success: true, data: rows[0] });
  } catch (err) { next(err); }
});

// PATCH /:id — update transaction fields
router.patch('/:id', async (req, res, next) => {
  try {
    const { id } = req.params;
    const { drug_name, lot_no, exp_date } = req.body;
    const fields = [];
    const values = [];
    let idx = 1;

    if (drug_name !== undefined) { fields.push(`drug_name=$${idx++}`); values.push(drug_name); }
    if (lot_no !== undefined)    { fields.push(`lot_no=$${idx++}`);    values.push(lot_no); }
    if (exp_date !== undefined)  { fields.push(`exp_date=$${idx++}`);  values.push(exp_date || null); }

    if (fields.length === 0) {
      return res.status(400).json({ success: false, error: 'No fields to update' });
    }

    values.push(id);
    const { rows } = await db.query(
      `UPDATE transactions SET ${fields.join(', ')} WHERE id=$${idx} RETURNING *`,
      values
    );
    res.json({ success: true, data: rows[0] });
  } catch (err) { next(err); }
});

module.exports = router;
