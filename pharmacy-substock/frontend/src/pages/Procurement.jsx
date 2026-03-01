import { useState, useEffect } from 'react';
import { apiGet } from '../api';
import toast from 'react-hot-toast';
import * as XLSX from 'xlsx';

export default function Procurement() {
  const [data, setData] = useState([]);
  const [selected, setSelected] = useState({});
  const [loading, setLoading] = useState(true);

  const load = async () => {
    setLoading(true);
    try {
      const res = await apiGet('/api/procurement');
      setData(res.data);
    } catch {
      toast.error('โหลดข้อมูลไม่สำเร็จ');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { load(); }, []);

  const toggleAll = () => {
    const allSel = data.every((_, i) => selected[i]);
    const sel = {};
    data.forEach((_, i) => { sel[i] = !allSel; });
    setSelected(sel);
  };

  const exportExcel = () => {
    const selRows = data.filter((_, i) => selected[i]);
    if (selRows.length === 0) return toast.error('กรุณาเลือกรายการ');

    // Group by cabinet
    const groups = {};
    selRows.forEach(r => {
      const cab = r.cabinet || 'ไม่ระบุตู้';
      if (!groups[cab]) groups[cab] = [];
      groups[cab].push(r);
    });

    const wb = XLSX.utils.book_new();
    const wsData = [];

    // Title
    wsData.push(['ใบเบิกยา Sub-stock โรงพยาบาลรือเสาะ']);
    wsData.push([`วันที่: ${new Date().toLocaleDateString('th-TH')}`]);
    wsData.push([]);

    let seq = 1;
    Object.entries(groups).forEach(([cabinet, items]) => {
      wsData.push([`ตู้: ${cabinet}`]);
      wsData.push(['ลำดับ', 'ชื่อยา', 'Min Stock', 'คงเหลือ', 'ที่ขอเบิก', 'ที่อนุมัติ', 'หมายเหตุ']);

      items.forEach(item => {
        const needed = Math.max(0, (item.min_stock || 0) - parseInt(item.balance, 10));
        wsData.push([seq++, item.drug_name, item.min_stock, parseInt(item.balance, 10), needed, '', '']);
      });

      // 5 blank rows for additional items
      wsData.push(['', 'รายการเพิ่มเติม', '', '', '', '', '']);
      for (let i = 0; i < 5; i++) {
        wsData.push([seq++, '', '', '', '', '', '']);
      }
      wsData.push([]);
    });

    // Signature rows
    wsData.push([]);
    wsData.push(['', 'ผู้เบิก ........................', '', '', 'ผู้อนุมัติ ........................']);
    wsData.push(['', 'วันที่ ........................', '', '', 'วันที่ ........................']);

    const ws = XLSX.utils.aoa_to_sheet(wsData);

    // Set column widths
    ws['!cols'] = [
      { wch: 6 },   // ลำดับ
      { wch: 40 },  // ชื่อยา
      { wch: 10 },  // Min Stock
      { wch: 10 },  // คงเหลือ
      { wch: 12 },  // ที่ขอเบิก
      { wch: 12 },  // ที่อนุมัติ
      { wch: 20 },  // หมายเหตุ
    ];

    // Set row heights
    ws['!rows'] = wsData.map(() => ({ hpt: 25 }));

    XLSX.utils.book_append_sheet(wb, ws, 'เบิกยา');
    XLSX.writeFile(wb, `procurement_${new Date().toISOString().slice(0, 10)}.xlsx`);
    toast.success('ส่งออก Excel สำเร็จ');
  };

  return (
    <div>
      <div className="flex justify-between items-center mb-4">
        <h2 className="text-2xl font-bold">เบิกยา (Procurement)</h2>
        <button
          onClick={exportExcel}
          className="px-4 py-2 rounded-lg text-white text-sm font-medium"
          style={{ background: '#10b981', border: 'none', cursor: 'pointer' }}
        >ส่งออก Excel</button>
      </div>

      {loading ? (
        <div className="text-center py-10">กำลังโหลด...</div>
      ) : data.length === 0 ? (
        <div className="text-center py-10 rounded-xl" style={{ background: 'var(--bg-secondary)', border: '1px solid var(--border)' }}>
          <p className="text-lg" style={{ color: 'var(--text-secondary)' }}>ไม่มีรายการยาที่ต้องเบิก</p>
          <p className="text-sm" style={{ color: 'var(--text-secondary)' }}>ยาทุกรายการมีจำนวนเพียงพอ</p>
        </div>
      ) : (
        <div className="overflow-x-auto rounded-xl shadow-sm" style={{ background: 'var(--bg-secondary)', border: '1px solid var(--border)' }}>
          <table className="w-full text-sm">
            <thead>
              <tr style={{ borderBottom: '2px solid var(--border)' }}>
                <th className="p-3"><input type="checkbox" onChange={toggleAll} checked={data.length > 0 && data.every((_, i) => selected[i])} /></th>
                {['ชื่อยา', 'ตู้', 'Min Stock', 'คงเหลือ', 'ต้องเบิก'].map(h => (
                  <th key={h} className="text-left p-3 text-xs" style={{ color: 'var(--text-secondary)' }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {data.map((row, i) => {
                const needed = Math.max(0, (row.min_stock || 0) - parseInt(row.balance, 10));
                return (
                  <tr key={i} style={{ borderBottom: '1px solid var(--border)' }} className="hover:bg-black/5">
                    <td className="p-3">
                      <input type="checkbox" checked={!!selected[i]} onChange={() => setSelected({ ...selected, [i]: !selected[i] })} />
                    </td>
                    <td className="p-3">{row.drug_name}</td>
                    <td className="p-3">{row.cabinet || '-'}</td>
                    <td className="p-3">{row.min_stock}</td>
                    <td className="p-3 font-bold" style={{ color: '#ef4444' }}>{row.balance}</td>
                    <td className="p-3 font-bold" style={{ color: '#f59e0b' }}>{needed}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
