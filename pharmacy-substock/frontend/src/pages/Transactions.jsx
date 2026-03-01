import { useState, useEffect } from 'react';
import { apiGet, apiPost } from '../api';
import toast from 'react-hot-toast';

export default function Transactions() {
  const [data, setData] = useState([]);
  const [drugList, setDrugList] = useState([]);
  const [loading, setLoading] = useState(true);
  const [showModal, setShowModal] = useState(false);
  const [form, setForm] = useState({
    drug_name: '', drug_code: '', pack_size: '', price_per_unit: '',
    disp_no: '', lot_no: '', exp_date: '', qty: '', txn_type: 'receive',
    dept: 'substock', recorder: '',
    date: new Date().toISOString().slice(0, 10),
  });

  const load = async () => {
    setLoading(true);
    try {
      const res = await apiGet('/api/transactions');
      setData(res.data);
    } catch {
      toast.error('โหลดข้อมูลไม่สำเร็จ');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { load(); }, []);

  const openModal = async () => {
    try {
      const [drugRes, dispRes] = await Promise.all([
        apiGet('/api/drug-list'),
        apiGet('/api/disp-no'),
      ]);
      setDrugList(drugRes.data);
      setForm(f => ({ ...f, disp_no: dispRes.data.disp_no }));
    } catch {}
    setShowModal(true);
  };

  const handleDrugSelect = e => {
    const name = e.target.value;
    const drug = drugList.find(d => d.drug_name === name);
    if (drug) {
      setForm(f => ({
        ...f,
        drug_name: drug.drug_name,
        drug_code: drug.drug_code || '',
        pack_size: drug.pack_size || '',
        price_per_unit: drug.price_per_unit || '',
      }));
    } else {
      setForm(f => ({ ...f, drug_name: name }));
    }
  };

  const handleSubmit = async e => {
    e.preventDefault();
    if (!form.drug_name || !form.qty) return toast.error('กรุณากรอกชื่อยาและจำนวน');
    const qty = form.txn_type === 'dispense' ? -Math.abs(Number(form.qty)) : Math.abs(Number(form.qty));
    try {
      await apiPost('/api/transactions', {
        ...form,
        qty,
        price_per_unit: Number(form.price_per_unit) || 0,
        exp_date: form.exp_date || null,
      });
      toast.success('บันทึกรายการสำเร็จ');
      setShowModal(false);
      setForm({
        drug_name: '', drug_code: '', pack_size: '', price_per_unit: '',
        disp_no: '', lot_no: '', exp_date: '', qty: '', txn_type: 'receive',
        dept: 'substock', recorder: '',
        date: new Date().toISOString().slice(0, 10),
      });
      load();
    } catch {
      toast.error('บันทึกไม่สำเร็จ');
    }
  };

  const fmtDate = d => d ? new Date(d).toLocaleDateString('th-TH') : '-';
  const fmtDateTime = d => d ? new Date(d).toLocaleString('th-TH') : '-';

  return (
    <div>
      <div className="flex justify-between items-center mb-4">
        <h2 className="text-2xl font-bold">รายการเคลื่อนไหว</h2>
        <button
          onClick={openModal}
          className="px-4 py-2 rounded-lg text-white font-medium text-sm"
          style={{ background: 'var(--accent)', border: 'none', cursor: 'pointer' }}
        >+ เพิ่มรายการ Manual</button>
      </div>

      {loading ? (
        <div className="text-center py-10">กำลังโหลด...</div>
      ) : (
        <div className="overflow-x-auto rounded-xl shadow-sm" style={{ background: 'var(--bg-secondary)', border: '1px solid var(--border)' }}>
          <table className="w-full text-sm">
            <thead>
              <tr style={{ borderBottom: '2px solid var(--border)' }}>
                {['วันที่', 'เลขที่', 'ชื่อยา', 'รหัส', 'Lot', 'วันหมดอายุ', 'จำนวน', 'ประเภท', 'ผู้บันทึก'].map(h => (
                  <th key={h} className="text-left p-3 text-xs" style={{ color: 'var(--text-secondary)' }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {data.map(row => (
                <tr key={row.id} style={{ borderBottom: '1px solid var(--border)' }} className="hover:bg-black/5">
                  <td className="p-3 text-xs">{fmtDateTime(row.created_at)}</td>
                  <td className="p-3">{row.disp_no || '-'}</td>
                  <td className="p-3">{row.drug_name}</td>
                  <td className="p-3">{row.drug_code || '-'}</td>
                  <td className="p-3">{row.lot_no || '-'}</td>
                  <td className="p-3">{fmtDate(row.exp_date)}</td>
                  <td className="p-3 font-bold" style={{ color: row.qty >= 0 ? '#10b981' : '#ef4444' }}>
                    {row.qty >= 0 ? `+${row.qty}` : row.qty}
                  </td>
                  <td className="p-3">
                    <span className="px-2 py-0.5 rounded text-xs" style={{
                      background: row.txn_type === 'receive' ? '#dcfce7' : row.txn_type === 'dispense' ? '#fee2e2' : '#e0e7ff',
                      color: row.txn_type === 'receive' ? '#166534' : row.txn_type === 'dispense' ? '#991b1b' : '#3730a3',
                    }}>
                      {row.txn_type === 'receive' ? 'รับเข้า' : row.txn_type === 'dispense' ? 'เบิกยา' : row.txn_type}
                    </span>
                  </td>
                  <td className="p-3">{row.recorder || '-'}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {/* Add Transaction Modal */}
      {showModal && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center"
          style={{ background: 'rgba(0,0,0,0.5)' }}
          onClick={() => setShowModal(false)}
        >
          <div
            className="rounded-xl p-6 w-full max-w-lg shadow-2xl"
            style={{ background: 'var(--bg-secondary)' }}
            onClick={e => e.stopPropagation()}
          >
            <h3 className="text-lg font-bold mb-4">เพิ่มรายการ Manual</h3>
            <form onSubmit={handleSubmit} className="space-y-3">
              <div>
                <label className="block text-sm mb-1" style={{ color: 'var(--text-secondary)' }}>ชื่อยา</label>
                <select
                  value={form.drug_name}
                  onChange={handleDrugSelect}
                  className="w-full px-3 py-2 rounded-lg text-sm"
                  style={{ background: 'var(--bg-primary)', border: '1px solid var(--border)', color: 'var(--text-primary)' }}
                >
                  <option value="">-- เลือกยา --</option>
                  {drugList.map((d, i) => (
                    <option key={i} value={d.drug_name}>{d.drug_name}</option>
                  ))}
                </select>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-sm mb-1" style={{ color: 'var(--text-secondary)' }}>รหัสยา</label>
                  <input value={form.drug_code} readOnly className="w-full px-3 py-2 rounded-lg text-sm"
                    style={{ background: 'var(--bg-primary)', border: '1px solid var(--border)', color: 'var(--text-secondary)' }} />
                </div>
                <div>
                  <label className="block text-sm mb-1" style={{ color: 'var(--text-secondary)' }}>ขนาดบรรจุ</label>
                  <input value={form.pack_size} readOnly className="w-full px-3 py-2 rounded-lg text-sm"
                    style={{ background: 'var(--bg-primary)', border: '1px solid var(--border)', color: 'var(--text-secondary)' }} />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-sm mb-1" style={{ color: 'var(--text-secondary)' }}>เลขที่ (Disp No)</label>
                  <input value={form.disp_no} readOnly className="w-full px-3 py-2 rounded-lg text-sm"
                    style={{ background: 'var(--bg-primary)', border: '1px solid var(--border)', color: 'var(--text-secondary)' }} />
                </div>
                <div>
                  <label className="block text-sm mb-1" style={{ color: 'var(--text-secondary)' }}>ราคาต่อหน่วย</label>
                  <input value={form.price_per_unit} readOnly className="w-full px-3 py-2 rounded-lg text-sm"
                    style={{ background: 'var(--bg-primary)', border: '1px solid var(--border)', color: 'var(--text-secondary)' }} />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-sm mb-1" style={{ color: 'var(--text-secondary)' }}>Lot No</label>
                  <input value={form.lot_no} onChange={e => setForm({ ...form, lot_no: e.target.value })}
                    className="w-full px-3 py-2 rounded-lg text-sm"
                    style={{ background: 'var(--bg-primary)', border: '1px solid var(--border)', color: 'var(--text-primary)' }} />
                </div>
                <div>
                  <label className="block text-sm mb-1" style={{ color: 'var(--text-secondary)' }}>วันหมดอายุ</label>
                  <input type="date" value={form.exp_date} onChange={e => setForm({ ...form, exp_date: e.target.value })}
                    className="w-full px-3 py-2 rounded-lg text-sm"
                    style={{ background: 'var(--bg-primary)', border: '1px solid var(--border)', color: 'var(--text-primary)' }} />
                </div>
              </div>

              <div className="grid grid-cols-3 gap-3">
                <div>
                  <label className="block text-sm mb-1" style={{ color: 'var(--text-secondary)' }}>จำนวน</label>
                  <input type="number" value={form.qty} onChange={e => setForm({ ...form, qty: e.target.value })}
                    className="w-full px-3 py-2 rounded-lg text-sm"
                    style={{ background: 'var(--bg-primary)', border: '1px solid var(--border)', color: 'var(--text-primary)' }} />
                </div>
                <div>
                  <label className="block text-sm mb-1" style={{ color: 'var(--text-secondary)' }}>ประเภท</label>
                  <select value={form.txn_type} onChange={e => setForm({ ...form, txn_type: e.target.value })}
                    className="w-full px-3 py-2 rounded-lg text-sm"
                    style={{ background: 'var(--bg-primary)', border: '1px solid var(--border)', color: 'var(--text-primary)' }}>
                    <option value="receive">รับเข้า</option>
                    <option value="dispense">เบิกยา</option>
                  </select>
                </div>
                <div>
                  <label className="block text-sm mb-1" style={{ color: 'var(--text-secondary)' }}>ผู้บันทึก</label>
                  <input value={form.recorder} onChange={e => setForm({ ...form, recorder: e.target.value })}
                    className="w-full px-3 py-2 rounded-lg text-sm"
                    style={{ background: 'var(--bg-primary)', border: '1px solid var(--border)', color: 'var(--text-primary)' }} />
                </div>
              </div>

              <div className="flex justify-end gap-2 mt-4">
                <button type="button" onClick={() => setShowModal(false)}
                  className="px-4 py-2 rounded-lg text-sm"
                  style={{ background: 'var(--bg-primary)', border: '1px solid var(--border)', cursor: 'pointer', color: 'var(--text-primary)' }}>
                  ยกเลิก
                </button>
                <button type="submit"
                  className="px-4 py-2 rounded-lg text-white text-sm font-medium"
                  style={{ background: 'var(--accent)', border: 'none', cursor: 'pointer' }}>
                  บันทึก
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}
