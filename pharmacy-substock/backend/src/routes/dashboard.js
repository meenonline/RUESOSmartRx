const router = require('express').Router();
const db = require('../db');

// GET /stats — dashboard statistics
router.get('/stats', async (req, res, next) => {
  try {
    // Get stock summary by lot
    const { rows: stockLots } = await db.query('SELECT * FROM stock_summary_lot');
    // Get formulary
    const { rows: formularyRows } = await db.query('SELECT * FROM formulary WHERE is_inactive = false');
    // Get ignore_expiry list
    const { rows: ignoreRows } = await db.query('SELECT * FROM ignore_expiry');
    // Get stock summary no lot for low-stock checks
    const { rows: stockNolot } = await db.query('SELECT * FROM stock_summary_nolot');

    const ignoreMap = {};
    ignoreRows.forEach(r => { ignoreMap[`${r.drug_name}||${r.lot_no}`] = true; });

    const formularyMap = {};
    formularyRows.forEach(f => { formularyMap[f.drug_name] = f; });

    const now = new Date();
    const m3 = new Date(); m3.setMonth(m3.getMonth() + 3);
    const m6 = new Date(); m6.setMonth(m6.getMonth() + 6);

    let totalValue = 0;
    const expiredList = [];
    const exp3List = [];
    const exp6List = [];
    const expUnknownList = [];

    // Auto-remove from ignore_expiry if balance > 0
    const toRemoveFromIgnore = [];

    for (const lot of stockLots) {
      const bal = parseInt(lot.balance, 10);
      totalValue += parseFloat(lot.total_value) || 0;
      const key = `${lot.drug_name}||${lot.lot_no}`;
      const isIgnored = ignoreMap[key];

      // If ignored and balance=0, skip
      if (isIgnored && bal === 0) continue;
      // If ignored and balance>0, auto-remove from ignore
      if (isIgnored && bal > 0) {
        toRemoveFromIgnore.push({ drug_name: lot.drug_name, lot_no: lot.lot_no });
      }

      if (!lot.exp_date) {
        expUnknownList.push(lot);
        continue;
      }

      const expDate = new Date(lot.exp_date);
      if (expDate < now) {
        expiredList.push(lot);
      } else if (expDate <= m3) {
        exp3List.push(lot);
      } else if (expDate <= m6) {
        exp6List.push(lot);
      }
    }

    // Auto-delete from ignore_expiry where balance > 0
    for (const item of toRemoveFromIgnore) {
      await db.query(
        'DELETE FROM ignore_expiry WHERE drug_name=$1 AND lot_no=$2',
        [item.drug_name, item.lot_no]
      );
    }

    // Low stock & out of stock
    const lowStockList = [];
    const outOfStockList = [];
    for (const s of stockNolot) {
      const bal = parseInt(s.balance, 10);
      const f = formularyMap[s.drug_name];
      const minStock = f ? (f.min_stock || 0) : 0;

      if (bal <= 0) {
        outOfStockList.push({ ...s, min_stock: minStock });
      } else if (f && bal < minStock) {
        lowStockList.push({ ...s, min_stock: minStock });
      }
    }

    res.json({
      success: true,
      data: {
        totalValue,
        totalDrugs: stockNolot.length,
        expiredList,
        exp3List,
        exp6List,
        expUnknownList,
        lowStockList,
        outOfStockList,
      },
    });
  } catch (err) { next(err); }
});

module.exports = router;
