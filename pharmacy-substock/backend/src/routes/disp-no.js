const router = require('express').Router();
const db = require('../db');

// GET / — compute next disp_no
router.get('/', async (req, res, next) => {
  try {
    const thaiYear = new Date().getFullYear() + 543;
    const yearPrefix = String(thaiYear).slice(-2); // last 2 digits

    const { rows } = await db.query(
      `SELECT MAX(disp_no::BIGINT) AS max_disp
       FROM transactions
       WHERE disp_no LIKE $1`,
      [`${yearPrefix}%`]
    );

    let nextNo;
    if (rows[0].max_disp) {
      nextNo = String(parseInt(rows[0].max_disp, 10) + 1);
    } else {
      nextNo = `${yearPrefix}001`;
    }

    res.json({ success: true, data: { disp_no: nextNo } });
  } catch (err) { next(err); }
});

module.exports = router;
