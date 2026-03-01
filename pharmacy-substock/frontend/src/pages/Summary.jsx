import { useState, useEffect } from 'react';
import { apiGet, apiPatch, apiPost } from '../api';
import toast from 'react-hot-toast';

const fmtNum = n => Number(n || 0).toLocaleString('th-TH', { minimumFractionDigits: 2 });

export default function Summary() {
  const [tab, setTab] = useState('lot'); // 'lot' | 'nolot'
  const [data, setData] = useState([]);
  const [formulary, setFormulary] = useState({});
  const [search, setSearch] = useState('');
  const [sortCol, setSortCol] = useState('drug_name');
  const [sortDir, setSortDir] = useState('asc');
  const [adjustments, setAdjustments] = useState({});
  const [editCell, setEditCell] = useState(null);
  const [editVal, setEditVal] = useState('');
  const [loading, setLoading] = useState(true);

  const load = async () => {
    setLoading(true);
    try {
      const endpoint = tab === 'lot' ? '/api/stock/summary' : '/api/stock/summary-nolot';
      const res = await apiGet(endpoint);
      setData(res.data);
      const fRes = await apiGet('/api/formulary');
      const fMap = {};
      fRes.data.forEach(f => { fMap[f.drug_name] = f; });
      setFormulary(fMap);
    } catch {
      toast.error('โหลดข้อมูลไม่สำเร็จ');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { load(); }, [tab]);

  const handleSort = col => {
    if (sortCol === col) setSortDir(d => d === 'asc' ? 'desc' : 'asc');
    else { setSortCol(col); setSortDir('asc'); }
  };

  const filtered = data
    .filter(r => r.drug_name?.toLowerCase().includes(search.toLowerCase()))
    .sort((a, b) => {
      const av = a[sortCol], bv = b[sortCol];
      const cmp = typeof av === 'number' ? av - bv : String(av || '').localeCompare(String(bv || ''));
      return sortDir === 'asc' ? cmp : -cmp;
    });

  const handleInlineEdit = async (id, field, value) => {
    try {
      await apiPatch(`/api/transactions/${id}`, { [field]: value });
      toast.success('อัปเดตแล้ว');
      load();
    } catch {
      toast.error('อัปเดตไม่สำเร็จ');
    }
    setEditCell(null);
  };

  const handleBatchAdjust = async () => {
    const entries = Object.entries(adjustments).filter(([, v]) => v !== '' && v !== undefined);
    if (entries.length === 0) return toast.error('กรุณาระบุยอดปรับ');

    const adj = entries.map(([drugName, newBal]) => {
      const item = data.find(d => d.drug_name === drugName);
      return {
        drug_name: drugName,
        drug_code: item?.drug_code,
        pack_size: item?.pack_size,
        price_per_unit: item?.price_per_unit,
        new_balance: Number(newBal),
      };
    });

    try {
      await apiPost('/api/stock/batch-adjust', { adjustments: adj, recorder: 'system' });
      toast.success(`ปรับยอด ${adj.length} รายการสำเร็จ`);
      setAdjustments({});
      load();
    } catch {
      toast.error('ปรับยอดไม่สำเร็จ');
    }
  };

  const balColor = (bal, drugName) => {
    const min = formulary[drugName]?.min_stock || 0;
    if (bal <= 0) return '#ef4444';
    if (min > 0 && bal < min) return '#f59e0b';
    return '#14b8a6';
  };

  const SortIcon = ({ col }) => (
    <span className="ml-1 text-xs">{sortCol === col ? (sortDir === 'asc' ? '▲' : '▼') : '⇅'}</span>
  );

  return (
    <div>
      <h2 className="text-2xl font-bold mb-4">สรุปยอดคงเหลือ</h2>

      {/* Tab toggle */}
      <div className="flex gap-2 mb-4">
        {[['lot', 'แยก Lot'], ['nolot', 'ไม่แยก Lot']].map(([key, label]) => (
          <button
            key={key}
            onClick={() => setTab(key)}
            className="px-4 py-2 rounded-lg text-sm font-medium"
            style={{
              background: tab === key ? 'var(--accent)' : 'var(--bg-secondary)',
              color: tab === key ? '#fff' : 'var(--text-primary)',
              border: `1px solid ${tab === key ? 'var(--accent)' : 'var(--border)'}`,
              cursor: 'pointer',
            }}
          >{label}</button>
        ))}
      </div>

      {/* Search */}
      <input
        type="text"
        placeholder="ค้นหาชื่อยา..."
        value={search}
        onChange={e => setSearch(e.target.value)}
        className="w-full max-w-md px-4 py-2 rounded-lg mb-4 text-sm"
        style={{ background: 'var(--bg-secondary)', border: '1px solid var(--border)', color: 'var(--text-primary)' }}
      />

      {loading ? (
        <div className="text-center py-10">กำลังโหลด...</div>
      ) : (
        <>
          <div className="overflow-x-auto rounded-xl shadow-sm" style={{ background: 'var(--bg-secondary)', border: '1px solid var(--border)' }}>
            <table className="w-full text-sm">
              <thead>
                <tr style={{ borderBottom: '2px solid var(--border)' }}>
                  {['drug_name', 'drug_code', 'pack_size', ...(tab === 'lot' ? ['lot_no', 'exp_date'] : []), 'balance', 'total_value'].map(col => (
                    <th
                      key={col}
                      className="text-left p-3 cursor-pointer select-none"
                      onClick={() => handleSort(col)}
                      style={{ color: 'var(--text-secondary)' }}
                    >
                      {{
                        drug_name: 'ชื่อยา',
                        drug_code: 'รหัสยา',
                        pack_size: 'ขนาดบรรจุ',
                        lot_no: 'Lot No',
                        exp_date: 'วันหมดอายุ',
                        balance: 'คงเหลือ',
                        total_value: 'มูลค่า',
                      }[col]}
                      <SortIcon col={col} />
                    </th>
                  ))}
                  {tab === 'nolot' && <th className="p-3 text-center" style={{ color: 'var(--text-secondary)' }}>ปรับยอด</th>}
                </tr>
              </thead>
              <tbody>
                {filtered.map((row, idx) => (
                  <tr key={idx} style={{ borderBottom: '1px solid var(--border)' }} className="hover:bg-black/5">
                    <td className="p-3">{row.drug_name}</td>
                    <td className="p-3">{row.drug_code}</td>
                    <td className="p-3">{row.pack_size}</td>
                    {tab === 'lot' && (
                      <>
                        <td className="p-3">{row.lot_no || '-'}</td>
                        <td className="p-3">{row.exp_date ? new Date(row.exp_date).toLocaleDateString('th-TH') : '-'}</td>
                      </>
                    )}
                    <td className="p-3 font-bold" style={{ color: balColor(parseInt(row.balance), row.drug_name) }}>
                      {row.balance}
                    </td>
                    <td className="p-3 text-right">{fmtNum(row.total_value)}</td>
                    {tab === 'nolot' && (
                      <td className="p-3 text-center">
                        <input
                          type="number"
                          placeholder="ยอดใหม่"
                          value={adjustments[row.drug_name] ?? ''}
                          onChange={e => setAdjustments({ ...adjustments, [row.drug_name]: e.target.value })}
                          className="w-20 px-2 py-1 rounded text-center text-sm"
                          style={{ background: 'var(--bg-primary)', border: '1px solid var(--border)', color: 'var(--text-primary)' }}
                        />
                      </td>
                    )}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {tab === 'nolot' && Object.keys(adjustments).some(k => adjustments[k] !== '') && (
            <div className="mt-4 flex justify-end">
              <button
                onClick={handleBatchAdjust}
                className="px-6 py-2 rounded-lg text-white font-medium"
                style={{ background: 'var(--accent)', border: 'none', cursor: 'pointer' }}
              >บันทึกปรับยอดทั้งหมด</button>
            </div>
          )}
        </>
      )}
    </div>
  );
}
