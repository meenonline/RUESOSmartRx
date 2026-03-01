import { useState, useCallback } from 'react';
import { apiGet, apiPost } from '../api';
import toast from 'react-hot-toast';
import * as XLSX from 'xlsx';

export default function Import() {
  const [step, setStep] = useState(1);
  const [rows, setRows] = useState([]);
  const [selected, setSelected] = useState({});
  const [drugList, setDrugList] = useState([]);
  const [txnType, setTxnType] = useState('receive');
  const [recorder, setRecorder] = useState('');
  const [manualDrug, setManualDrug] = useState('');
  const [manualQty, setManualQty] = useState('');
  const [manualLot, setManualLot] = useState('');
  const [manualExp, setManualExp] = useState('');
  const [loading, setLoading] = useState(false);

  // Load drug list for manual add
  const loadDrugList = async () => {
    try {
      const res = await apiGet('/api/drug-list');
      setDrugList(res.data);
    } catch {}
  };

  const handleFile = useCallback(async (file) => {
    if (!file) return;
    setLoading(true);
    try {
      const ab = await file.arrayBuffer();
      const wb = XLSX.read(ab);
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(ws);

      // Map columns
      const mapped = json.map(r => ({
        drug_name: r['drug_name'] || r['ชื่อยา'] || r['DrugName'] || '',
        drug_code: r['drug_code'] || r['รหัสยา'] || '',
        pack_size: r['pack_size'] || r['ขนาดบรรจุ'] || '',
        price_per_unit: Number(r['price_per_unit'] || r['ราคา'] || 0),
        lot_no: r['lot_no'] || r['Lot'] || r['lot'] || '',
        exp_date: r['exp_date'] || r['วันหมดอายุ'] || '',
        qty: Number(r['qty'] || r['จำนวน'] || 0),
        disp_no: r['disp_no'] || r['เลขที่'] || '',
      })).filter(r => r.drug_name);

      // Preview against stock
      const res = await apiPost('/api/import/preview', { rows: mapped });
      setRows(res.data);
      const sel = {};
      res.data.forEach((_, i) => { sel[i] = true; });
      setSelected(sel);
      await loadDrugList();
    } catch (e) {
      toast.error('อ่านไฟล์ไม่สำเร็จ');
    } finally {
      setLoading(false);
    }
  }, []);

  const handleDrop = e => {
    e.preventDefault();
    const file = e.dataTransfer?.files[0];
    if (file) handleFile(file);
  };

  const handleAddManual = () => {
    if (!manualDrug || !manualQty) return toast.error('กรุณาเลือกยาและระบุจำนวน');
    const drug = drugList.find(d => d.drug_name === manualDrug);
    const newRow = {
      drug_name: manualDrug,
      drug_code: drug?.drug_code || '',
      pack_size: drug?.pack_size || '',
      price_per_unit: drug?.price_per_unit || 0,
      lot_no: manualLot,
      exp_date: manualExp,
      qty: Number(manualQty),
      currentBal: 0,
      match: !!drug,
    };
    setRows([...rows, newRow]);
    setSelected({ ...selected, [rows.length]: true });
    setManualDrug('');
    setManualQty('');
    setManualLot('');
    setManualExp('');
  };

  const toggleAll = () => {
    const allSelected = rows.every((_, i) => selected[i]);
    const sel = {};
    rows.forEach((_, i) => { sel[i] = !allSelected; });
    setSelected(sel);
  };

  const selectedRows = rows.filter((_, i) => selected[i]);

  const handleConfirm = async () => {
    if (!recorder) return toast.error('กรุณาระบุผู้บันทึก');
    if (selectedRows.length === 0) return toast.error('กรุณาเลือกรายการ');

    setLoading(true);
    try {
      await apiPost('/api/import/confirm', {
        rows: selectedRows,
        txn_type: txnType,
        recorder,
      });
      toast.success(`นำเข้า ${selectedRows.length} รายการสำเร็จ`);
      setStep(1);
      setRows([]);
      setSelected({});
    } catch {
      toast.error('นำเข้าไม่สำเร็จ');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div>
      <h2 className="text-2xl font-bold mb-4">นำเข้าข้อมูล</h2>

      {step === 1 && (
        <>
          {/* File drop zone */}
          <div
            onDrop={handleDrop}
            onDragOver={e => e.preventDefault()}
            className="border-2 border-dashed rounded-xl p-10 text-center mb-6 cursor-pointer"
            style={{ borderColor: 'var(--accent)', background: 'var(--bg-secondary)' }}
            onClick={() => document.getElementById('fileInput').click()}
          >
            <input
              id="fileInput"
              type="file"
              accept=".xlsx,.xls"
              className="hidden"
              onChange={e => handleFile(e.target.files[0])}
            />
            <div className="text-4xl mb-2">📂</div>
            <p className="text-lg font-medium">ลากไฟล์ .xlsx มาวางที่นี่</p>
            <p className="text-sm" style={{ color: 'var(--text-secondary)' }}>หรือคลิกเพื่อเลือกไฟล์</p>
          </div>

          {/* Manual add */}
          <div className="rounded-xl p-4 mb-4" style={{ background: 'var(--bg-secondary)', border: '1px solid var(--border)' }}>
            <h3 className="text-sm font-bold mb-3">เพิ่มรายการด้วยตนเอง</h3>
            <div className="flex flex-wrap gap-2 items-end">
              <div className="flex-1 min-w-[200px]">
                <label className="block text-xs mb-1" style={{ color: 'var(--text-secondary)' }}>ยา</label>
                <select
                  value={manualDrug}
                  onChange={e => setManualDrug(e.target.value)}
                  onFocus={loadDrugList}
                  className="w-full px-3 py-2 rounded text-sm"
                  style={{ background: 'var(--bg-primary)', border: '1px solid var(--border)', color: 'var(--text-primary)' }}
                >
                  <option value="">เลือกยา</option>
                  {drugList.map((d, i) => <option key={i} value={d.drug_name}>{d.drug_name}</option>)}
                </select>
              </div>
              <div className="w-20">
                <label className="block text-xs mb-1" style={{ color: 'var(--text-secondary)' }}>จำนวน</label>
                <input type="number" value={manualQty} onChange={e => setManualQty(e.target.value)}
                  className="w-full px-2 py-2 rounded text-sm"
                  style={{ background: 'var(--bg-primary)', border: '1px solid var(--border)', color: 'var(--text-primary)' }} />
              </div>
              <div className="w-28">
                <label className="block text-xs mb-1" style={{ color: 'var(--text-secondary)' }}>Lot No</label>
                <input value={manualLot} onChange={e => setManualLot(e.target.value)}
                  className="w-full px-2 py-2 rounded text-sm"
                  style={{ background: 'var(--bg-primary)', border: '1px solid var(--border)', color: 'var(--text-primary)' }} />
              </div>
              <div className="w-36">
                <label className="block text-xs mb-1" style={{ color: 'var(--text-secondary)' }}>วันหมดอายุ</label>
                <input type="date" value={manualExp} onChange={e => setManualExp(e.target.value)}
                  className="w-full px-2 py-2 rounded text-sm"
                  style={{ background: 'var(--bg-primary)', border: '1px solid var(--border)', color: 'var(--text-primary)' }} />
              </div>
              <button onClick={handleAddManual}
                className="px-4 py-2 rounded text-white text-sm"
                style={{ background: 'var(--accent)', border: 'none', cursor: 'pointer' }}>
                เพิ่ม
              </button>
            </div>
          </div>

          {loading && <div className="text-center py-4">กำลังประมวลผล...</div>}

          {rows.length > 0 && (
            <>
              <div className="overflow-x-auto rounded-xl shadow-sm mb-4" style={{ background: 'var(--bg-secondary)', border: '1px solid var(--border)' }}>
                <table className="w-full text-sm">
                  <thead>
                    <tr style={{ borderBottom: '2px solid var(--border)' }}>
                      <th className="p-3"><input type="checkbox" onChange={toggleAll} checked={rows.every((_, i) => selected[i])} /></th>
                      <th className="text-left p-3" style={{ color: 'var(--text-secondary)' }}>ชื่อยา</th>
                      <th className="text-left p-3" style={{ color: 'var(--text-secondary)' }}>Lot</th>
                      <th className="text-left p-3" style={{ color: 'var(--text-secondary)' }}>วันหมดอายุ</th>
                      <th className="text-right p-3" style={{ color: 'var(--text-secondary)' }}>จำนวน</th>
                      <th className="text-right p-3" style={{ color: 'var(--text-secondary)' }}>คงเหลือปัจจุบัน</th>
                      <th className="text-center p-3" style={{ color: 'var(--text-secondary)' }}>สถานะ</th>
                    </tr>
                  </thead>
                  <tbody>
                    {rows.map((row, i) => (
                      <tr key={i} style={{ borderBottom: '1px solid var(--border)' }}>
                        <td className="p-3"><input type="checkbox" checked={!!selected[i]} onChange={() => setSelected({ ...selected, [i]: !selected[i] })} /></td>
                        <td className="p-3">{row.drug_name}</td>
                        <td className="p-3">{row.lot_no || '-'}</td>
                        <td className="p-3">{row.exp_date || '-'}</td>
                        <td className="p-3 text-right font-bold">{row.qty}</td>
                        <td className="p-3 text-right">{row.currentBal}</td>
                        <td className="p-3 text-center">
                          <span className="px-2 py-0.5 rounded text-xs" style={{
                            background: row.match ? '#dcfce7' : '#fef3c7',
                            color: row.match ? '#166534' : '#92400e',
                          }}>
                            {row.match ? 'พบในระบบ' : 'ยาใหม่'}
                          </span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              <div className="flex justify-end">
                <button
                  onClick={() => setStep(2)}
                  disabled={selectedRows.length === 0}
                  className="px-6 py-2 rounded-lg text-white font-medium"
                  style={{ background: selectedRows.length > 0 ? 'var(--accent)' : '#94a3b8', border: 'none', cursor: 'pointer' }}
                >ถัดไป ({selectedRows.length} รายการ)</button>
              </div>
            </>
          )}
        </>
      )}

      {step === 2 && (
        <>
          <button onClick={() => setStep(1)} className="mb-4 text-sm" style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'var(--accent)' }}>
            ← กลับ
          </button>

          <div className="rounded-xl p-4 mb-4" style={{ background: 'var(--bg-secondary)', border: '1px solid var(--border)' }}>
            <div className="flex flex-wrap gap-4 items-end">
              <div>
                <label className="block text-sm mb-1" style={{ color: 'var(--text-secondary)' }}>ประเภทรายการ</label>
                <select value={txnType} onChange={e => setTxnType(e.target.value)}
                  className="px-3 py-2 rounded text-sm"
                  style={{ background: 'var(--bg-primary)', border: '1px solid var(--border)', color: 'var(--text-primary)' }}>
                  <option value="receive">รับเข้า</option>
                  <option value="dispense">เบิกยา</option>
                </select>
              </div>
              <div>
                <label className="block text-sm mb-1" style={{ color: 'var(--text-secondary)' }}>ผู้บันทึก</label>
                <input value={recorder} onChange={e => setRecorder(e.target.value)}
                  className="px-3 py-2 rounded text-sm"
                  style={{ background: 'var(--bg-primary)', border: '1px solid var(--border)', color: 'var(--text-primary)' }} />
              </div>
            </div>
          </div>

          <div className="overflow-x-auto rounded-xl shadow-sm mb-4" style={{ background: 'var(--bg-secondary)', border: '1px solid var(--border)' }}>
            <table className="w-full text-sm">
              <thead>
                <tr style={{ borderBottom: '2px solid var(--border)' }}>
                  <th className="text-left p-3" style={{ color: 'var(--text-secondary)' }}>ชื่อยา</th>
                  <th className="text-left p-3" style={{ color: 'var(--text-secondary)' }}>Lot</th>
                  <th className="text-right p-3" style={{ color: 'var(--text-secondary)' }}>จำนวน</th>
                  <th className="text-right p-3" style={{ color: 'var(--text-secondary)' }}>ราคา/หน่วย</th>
                </tr>
              </thead>
              <tbody>
                {selectedRows.map((row, i) => (
                  <tr key={i} style={{ borderBottom: '1px solid var(--border)' }}>
                    <td className="p-3">{row.drug_name}</td>
                    <td className="p-3">{row.lot_no || '-'}</td>
                    <td className="p-3 text-right font-bold">{row.qty}</td>
                    <td className="p-3 text-right">{row.price_per_unit}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div className="flex justify-end">
            <button
              onClick={handleConfirm}
              disabled={loading}
              className="px-6 py-2 rounded-lg text-white font-medium"
              style={{ background: 'var(--success)', border: 'none', cursor: 'pointer' }}
            >{loading ? 'กำลังบันทึก...' : `ยืนยันนำเข้า ${selectedRows.length} รายการ`}</button>
          </div>
        </>
      )}
    </div>
  );
}
