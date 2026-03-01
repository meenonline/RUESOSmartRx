const router = require('express').Router();
const db = require('../db');

// POST /ignore — add drug/lot to ignore list
router.post('/ignore', async (req, res, next) => {
  try {
    const { drug_name, lot_no } = req.body;
    await db.query(
      'INSERT INTO ignore_expiry(drug_name, lot_no) VALUES($1, $2) ON CONFLICT DO NOTHING',
      [drug_name, lot_no]
    );
    res.json({ success: true, data: { drug_name, lot_no } });
  } catch (err) { next(err); }
});

// DELETE /ignore — remove from ignore list
router.delete('/ignore', async (req, res, next) => {
  try {
    const { drug_name, lot_no } = req.body;
    await db.query(
      'DELETE FROM ignore_expiry WHERE drug_name=$1 AND lot_no=$2',
      [drug_name, lot_no]
    );
    res.json({ success: true, data: { drug_name, lot_no } });
  } catch (err) { next(err); }
});

module.exports = router;
