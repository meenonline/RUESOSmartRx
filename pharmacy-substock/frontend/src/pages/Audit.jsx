import { useState } from 'react';
import { apiGet, apiPost } from '../api';
import toast from 'react-hot-toast';

export default function Audit() {
  const [recorder, setRecorder] = useState('');
  const [critical, setCritical] = useState([]);
  const [normal, setNormal] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [saving, setSaving] = useState(false);

  const handleRandom = async () => {
    if (!recorder) return toast.error('กรุณาระบุชื่อผู้นับ');
    try {
      const res = await apiGet('/api/audit/random');
      const addCountField = list => list.map(d => ({ ...d, counted_qty: '', diff: null }));
      setCritical(addCountField(res.data.critical));
      setNormal(addCountField(res.data.normal));
      setLoaded(true);
    } catch {
      toast.error('โหลดข้อมูลไม่สำเร็จ');
    }
  };

  const updateCount = (type, idx, val) => {
    const setter = type === 'critical' ? setCritical : setNormal;
    setter(prev => {
      const arr = [...prev];
      const num = val === '' ? '' : Number(val);
      arr[idx] = {
        ...arr[idx],
        counted_qty: num,
        diff: num === '' ? null : num - parseInt(arr[idx].balance, 10),
      };
      return arr;
    });
  };

  const handleSave = async () => {
    const allRows = [...critical, ...normal].filter(r => r.counted_qty !== '');
    if (allRows.length === 0) return toast.error('กรุณานับยาอย่างน้อย 1 รายการ');

    setSaving(true);
    try {
      await apiPost('/api/audit/save', {
        rows: allRows.map(r => ({
          drug_name: r.drug_name,
          system_qty: parseInt(r.balance, 10),
          counted_qty: r.counted_qty,
          diff: r.diff,
          cabinet: r.cabinet || '',
          shelf_position: r.shelf_position || '',
          lot_no: r.lot_no || '',
          recorder,
        })),
      });
      toast.success('บันทึกผลการนับสำเร็จ');
      setLoaded(false);
      setCritical([]);
      setNormal([]);
    } catch {
      toast.error('บันทึกไม่สำเร็จ');
    } finally {
      setSaving(false);
    }
  };

  const renderTable = (title, list, type) => (
    <div className="mb-6">
      <h3 className="text-lg font-bold mb-3" style={{ color: type === 'critical' ? '#ef4444' : 'var(--text-primary)' }}>
        {title} ({list.length} รายการ)
      </h3>
      {list.length === 0 ? (
        <p className="text-sm" style={{ color: 'var(--text-secondary)' }}>ไม่พบรายการ</p>
      ) : (
        <div className="overflow-x-auto rounded-xl shadow-sm" style={{ background: 'var(--bg-secondary)', border: '1px solid var(--border)' }}>
          <table className="w-full text-sm">
            <thead>
              <tr style={{ borderBottom: '2px solid var(--border)' }}>
                {['ชื่อยา', 'คงเหลือ (ระบบ)', 'ตู้', 'ตำแหน่ง', 'นับได้', 'ผลต่าง'].map(h => (
                  <th key={h} className="text-left p-3 text-xs" style={{ color: 'var(--text-secondary)' }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {list.map((row, idx) => (
                <tr key={idx} style={{ borderBottom: '1px solid var(--border)' }}>
                  <td className="p-3">{row.drug_name}</td>
                  <td className="p-3 font-bold">{row.balance}</td>
                  <td className="p-3">{row.cabinet || '-'}</td>
                  <td className="p-3">{row.shelf_position || '-'}</td>
                  <td className="p-3">
                    <input
                      type="number"
                      value={row.counted_qty}
                      onChange={e => updateCount(type, idx, e.target.value)}
                      className="w-20 px-2 py-1 rounded text-center text-sm"
                      style={{ background: 'var(--bg-primary)', border: '1px solid var(--border)', color: 'var(--text-primary)' }}
                      placeholder="จำนวน"
                    />
                  </td>
                  <td className="p-3 font-bold" style={{
                    color: row.diff === null ? 'var(--text-secondary)' : row.diff === 0 ? '#10b981' : '#ef4444',
                  }}>
                    {row.diff === null ? '-' : row.diff > 0 ? `+${row.diff}` : row.diff}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );

  return (
    <div>
      <h2 className="text-2xl font-bold mb-4">สุ่มนับยา (Audit)</h2>

      <div className="flex items-end gap-3 mb-6">
        <div>
          <label className="block text-sm mb-1" style={{ color: 'var(--text-secondary)' }}>ผู้นับ</label>
          <input
            value={recorder}
            onChange={e => setRecorder(e.target.value)}
            className="px-3 py-2 rounded-lg text-sm"
            style={{ background: 'var(--bg-secondary)', border: '1px solid var(--border)', color: 'var(--text-primary)' }}
            placeholder="ชื่อผู้นับ"
          />
        </div>
        <button
          onClick={handleRandom}
          className="px-4 py-2 rounded-lg text-white text-sm font-medium"
          style={{ background: 'var(--accent)', border: 'none', cursor: 'pointer' }}
        >สุ่มนับยา</button>
      </div>

      {loaded && (
        <>
          {renderTable('รายการวิกฤต', critical, 'critical')}
          {renderTable('รายการปกติ', normal, 'normal')}

          <div className="flex justify-end mt-4">
            <button
              onClick={handleSave}
              disabled={saving}
              className="px-6 py-2 rounded-lg text-white font-medium"
              style={{ background: 'var(--success)', border: 'none', cursor: 'pointer' }}
            >{saving ? 'กำลังบันทึก...' : 'บันทึกผลการนับ'}</button>
          </div>
        </>
      )}
    </div>
  );
}
