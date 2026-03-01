const router = require('express').Router();
const db = require('../db');

// GET / — list formulary with current balance
router.get('/', async (req, res, next) => {
  try {
    const { rows } = await db.query(
      `SELECT f.*, COALESCE(s.balance, 0) AS balance
       FROM formulary f
       LEFT JOIN stock_summary_nolot s USING(drug_name)
       ORDER BY f.drug_name`
    );
    res.json({ success: true, data: rows });
  } catch (err) { next(err); }
});

// PATCH /:id — update formulary fields
router.patch('/:id', async (req, res, next) => {
  try {
    const { id } = req.params;
    const { min_stock, cabinet, reorder_point, shelf_position } = req.body;
    const fields = [];
    const values = [];
    let idx = 1;

    if (min_stock !== undefined)      { fields.push(`min_stock=$${idx++}`);      values.push(min_stock); }
    if (cabinet !== undefined)        { fields.push(`cabinet=$${idx++}`);        values.push(cabinet); }
    if (reorder_point !== undefined)  { fields.push(`reorder_point=$${idx++}`);  values.push(reorder_point); }
    if (shelf_position !== undefined) { fields.push(`shelf_position=$${idx++}`); values.push(shelf_position); }

    if (fields.length === 0) {
      return res.status(400).json({ success: false, error: 'No fields to update' });
    }

    values.push(id);
    const { rows } = await db.query(
      `UPDATE formulary SET ${fields.join(', ')} WHERE id=$${idx} RETURNING *`,
      values
    );
    res.json({ success: true, data: rows[0] });
  } catch (err) { next(err); }
});

// POST /toggle — toggle inactive status
router.post('/toggle', async (req, res, next) => {
  try {
    const { id } = req.body;
    const { rows } = await db.query(
      'UPDATE formulary SET is_inactive = NOT is_inactive WHERE id=$1 RETURNING *',
      [id]
    );
    res.json({ success: true, data: rows[0] });
  } catch (err) { next(err); }
});

module.exports = router;
