import { useState, useEffect } from 'react';
import { apiGet, apiPost } from '../api';
import toast from 'react-hot-toast';

const fmtNum = n => Number(n || 0).toLocaleString('th-TH', { minimumFractionDigits: 2 });
const fmtDate = d => d ? new Date(d).toLocaleDateString('th-TH') : '-';

export default function Dashboard() {
  const [stats, setStats] = useState(null);
  const [overlay, setOverlay] = useState(null); // { title, list }
  const [loading, setLoading] = useState(true);

  const load = async () => {
    try {
      setLoading(true);
      const res = await apiGet('/api/dashboard/stats');
      setStats(res.data);
    } catch (e) {
      toast.error('โหลดข้อมูลไม่สำเร็จ');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { load(); }, []);

  const handleHide = async (drug_name, lot_no) => {
    try {
      await apiPost('/api/expiry/ignore', { drug_name, lot_no });
      toast.success('ซ่อนแล้ว');
      load();
      setOverlay(null);
    } catch {
      toast.error('ซ่อนไม่สำเร็จ');
    }
  };

  if (loading) return <div className="text-center py-20 text-lg">กำลังโหลด...</div>;
  if (!stats) return <div className="text-center py-20">ไม่พบข้อมูล</div>;

  const cards = [
    { label: 'มูลค่ารวม', value: `฿${fmtNum(stats.totalValue)}`, color: '#0ea5e9', list: null },
    { label: 'รายการยาทั้งหมด', value: stats.totalDrugs, color: '#8b5cf6', list: null },
    { label: 'หมดอายุแล้ว', value: stats.expiredList.length, color: '#ef4444', list: stats.expiredList, title: 'รายการยาหมดอายุ' },
    { label: 'ใกล้หมดอายุ ≤3 เดือน', value: stats.exp3List.length, color: '#f59e0b', list: stats.exp3List, title: 'ใกล้หมดอายุ ≤3 เดือน' },
    { label: 'ใกล้หมดอายุ ≤6 เดือน', value: stats.exp6List.length, color: '#f97316', list: stats.exp6List, title: 'ใกล้หมดอายุ ≤6 เดือน' },
    { label: 'ยาใกล้หมด (Low Stock)', value: stats.lowStockList.length, color: '#eab308', list: stats.lowStockList, title: 'ยาใกล้หมด' },
    { label: 'ยาหมด (Out of Stock)', value: stats.outOfStockList.length, color: '#dc2626', list: stats.outOfStockList, title: 'ยาหมด' },
  ];

  return (
    <div>
      <h2 className="text-2xl font-bold mb-6">Dashboard</h2>

      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 mb-6">
        {cards.map((c, i) => (
          <div
            key={i}
            onClick={() => c.list && setOverlay({ title: c.title, list: c.list })}
            className="rounded-xl p-5 shadow-sm transition-transform hover:scale-[1.02]"
            style={{
              background: 'var(--bg-secondary)',
              border: '1px solid var(--border)',
              cursor: c.list ? 'pointer' : 'default',
              borderLeft: `4px solid ${c.color}`,
            }}
          >
            <div className="text-sm mb-1" style={{ color: 'var(--text-secondary)' }}>{c.label}</div>
            <div className="text-2xl font-bold" style={{ color: c.color }}>{c.value}</div>
          </div>
        ))}
      </div>

      {/* Overlay */}
      {overlay && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center"
          style={{ background: 'rgba(0,0,0,0.5)' }}
          onClick={() => setOverlay(null)}
        >
          <div
            className="rounded-xl p-6 max-w-3xl w-full max-h-[80vh] overflow-auto shadow-2xl"
            style={{ background: 'var(--bg-secondary)' }}
            onClick={e => e.stopPropagation()}
          >
            <div className="flex justify-between items-center mb-4">
              <h3 className="text-xl font-bold">{overlay.title}</h3>
              <button
                onClick={() => setOverlay(null)}
                style={{ background: 'none', border: 'none', fontSize: '1.5rem', cursor: 'pointer', color: 'var(--text-secondary)' }}
              >✕</button>
            </div>
            <table className="w-full text-sm">
              <thead>
                <tr style={{ borderBottom: '2px solid var(--border)' }}>
                  <th className="text-left p-2">ชื่อยา</th>
                  <th className="text-left p-2">Lot No</th>
                  <th className="text-left p-2">วันหมดอายุ</th>
                  <th className="text-right p-2">คงเหลือ</th>
                  <th className="text-right p-2">มูลค่า</th>
                  <th className="p-2"></th>
                </tr>
              </thead>
              <tbody>
                {overlay.list.map((item, idx) => (
                  <tr key={idx} style={{ borderBottom: '1px solid var(--border)' }}>
                    <td className="p-2">{item.drug_name}</td>
                    <td className="p-2">{item.lot_no || '-'}</td>
                    <td className="p-2">{fmtDate(item.exp_date)}</td>
                    <td className="p-2 text-right">{item.balance}</td>
                    <td className="p-2 text-right">{fmtNum(item.total_value)}</td>
                    <td className="p-2">
                      {item.lot_no && (
                        <button
                          onClick={() => handleHide(item.drug_name, item.lot_no)}
                          className="text-xs px-2 py-1 rounded"
                          style={{ background: 'var(--border)', border: 'none', cursor: 'pointer' }}
                        >ซ่อน</button>
                      )}
                    </td>
                  </tr>
                ))}
                {overlay.list.length === 0 && (
                  <tr><td colSpan={6} className="text-center p-4" style={{ color: 'var(--text-secondary)' }}>ไม่พบรายการ</td></tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}
