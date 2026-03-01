import { useState, useEffect } from 'react';
import { apiGet, apiPatch, apiPost } from '../api';
import toast from 'react-hot-toast';

export default function Formulary() {
  const [data, setData] = useState([]);
  const [search, setSearch] = useState('');
  const [loading, setLoading] = useState(true);

  const load = async () => {
    setLoading(true);
    try {
      const res = await apiGet('/api/formulary');
      setData(res.data);
    } catch {
      toast.error('โหลดข้อมูลไม่สำเร็จ');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { load(); }, []);

  const handleUpdate = async (id, field, value) => {
    try {
      await apiPatch(`/api/formulary/${id}`, { [field]: value });
      toast.success('อัปเดตแล้ว');
    } catch {
      toast.error('อัปเดตไม่สำเร็จ');
    }
  };

  const handleToggle = async (id) => {
    try {
      await apiPost('/api/formulary/toggle', { id });
      load();
    } catch {
      toast.error('ไม่สำเร็จ');
    }
  };

  const filtered = data.filter(r =>
    r.drug_name?.toLowerCase().includes(search.toLowerCase())
  );

  return (
    <div>
      <h2 className="text-2xl font-bold mb-4">รายการยา (Formulary)</h2>

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
        <div className="overflow-x-auto rounded-xl shadow-sm" style={{ background: 'var(--bg-secondary)', border: '1px solid var(--border)' }}>
          <table className="w-full text-sm">
            <thead>
              <tr style={{ borderBottom: '2px solid var(--border)' }}>
                {['ชื่อยา', 'รหัส', 'คงเหลือ', 'Min Stock', 'ตู้/ชั้น', 'Reorder Point', 'ตำแหน่ง', 'สถานะ', ''].map(h => (
                  <th key={h} className="text-left p-3 text-xs" style={{ color: 'var(--text-secondary)' }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.map(row => (
                <tr key={row.id} style={{
                  borderBottom: '1px solid var(--border)',
                  opacity: row.is_inactive ? 0.5 : 1,
                }} className="hover:bg-black/5">
                  <td className="p-3">{row.drug_name}</td>
                  <td className="p-3">{row.drug_code || '-'}</td>
                  <td className="p-3 font-bold" style={{
                    color: parseInt(row.balance) <= 0 ? '#ef4444'
                      : parseInt(row.balance) < (row.min_stock || 0) ? '#f59e0b'
                      : '#14b8a6'
                  }}>{row.balance}</td>
                  <td className="p-3">
                    <input
                      type="number"
                      defaultValue={row.min_stock}
                      onBlur={e => handleUpdate(row.id, 'min_stock', Number(e.target.value))}
                      className="w-16 px-2 py-1 rounded text-center text-sm"
                      style={{ background: 'var(--bg-primary)', border: '1px solid var(--border)', color: 'var(--text-primary)' }}
                    />
                  </td>
                  <td className="p-3">
                    <input
                      defaultValue={row.cabinet || ''}
                      onBlur={e => handleUpdate(row.id, 'cabinet', e.target.value)}
                      className="w-20 px-2 py-1 rounded text-sm"
                      style={{ background: 'var(--bg-primary)', border: '1px solid var(--border)', color: 'var(--text-primary)' }}
                    />
                  </td>
                  <td className="p-3">
                    <input
                      type="number"
                      defaultValue={row.reorder_point}
                      onBlur={e => handleUpdate(row.id, 'reorder_point', Number(e.target.value))}
                      className="w-16 px-2 py-1 rounded text-center text-sm"
                      style={{ background: 'var(--bg-primary)', border: '1px solid var(--border)', color: 'var(--text-primary)' }}
                    />
                  </td>
                  <td className="p-3">
                    <input
                      defaultValue={row.shelf_position || ''}
                      onBlur={e => handleUpdate(row.id, 'shelf_position', e.target.value)}
                      className="w-20 px-2 py-1 rounded text-sm"
                      style={{ background: 'var(--bg-primary)', border: '1px solid var(--border)', color: 'var(--text-primary)' }}
                    />
                  </td>
                  <td className="p-3">
                    <span className="px-2 py-0.5 rounded text-xs" style={{
                      background: row.is_inactive ? '#fee2e2' : '#dcfce7',
                      color: row.is_inactive ? '#991b1b' : '#166534',
                    }}>
                      {row.is_inactive ? 'ไม่ใช้งาน' : 'ใช้งาน'}
                    </span>
                  </td>
                  <td className="p-3">
                    <button
                      onClick={() => handleToggle(row.id)}
                      className="text-xs px-2 py-1 rounded"
                      style={{ background: 'var(--border)', border: 'none', cursor: 'pointer', color: 'var(--text-primary)' }}
                    >{row.is_inactive ? 'เปิดใช้' : 'ปิดใช้'}</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
