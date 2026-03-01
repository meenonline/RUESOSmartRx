import { useState, useEffect } from 'react';
import { apiGet, apiPost } from '../api';
import toast from 'react-hot-toast';

export default function Settings() {
  const [settings, setSettings] = useState({});
  const [lineTestResult, setLineTestResult] = useState(null);
  const [testing, setTesting] = useState(false);
  const [loading, setLoading] = useState(true);

  const load = async () => {
    try {
      const res = await apiGet('/api/settings');
      setSettings(res.data);
    } catch {
      toast.error('โหลดข้อมูลไม่สำเร็จ');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { load(); }, []);

  const saveSetting = async (key, value) => {
    try {
      await apiPost('/api/settings', { key, value });
      setSettings({ ...settings, [key]: value });
      toast.success(`บันทึก ${key} แล้ว`);
    } catch {
      toast.error('บันทึกไม่สำเร็จ');
    }
  };

  const testLine = async () => {
    setTesting(true);
    setLineTestResult(null);
    try {
      const res = await apiPost('/api/line/test', {
        token: settings.line_token || '',
        targetId: settings.line_target_id || '',
      });
      setLineTestResult(res);
    } catch (e) {
      setLineTestResult({ success: false, error: 'เชื่อมต่อไม่สำเร็จ' });
    } finally {
      setTesting(false);
    }
  };

  const fields = [
    { key: 'line_token', label: 'LINE Channel Access Token', type: 'password' },
    { key: 'line_target_id', label: 'LINE Target ID', type: 'text' },
    { key: 'email', label: 'Email สำหรับรายงาน', type: 'email' },
    { key: 'hospital_name', label: 'ชื่อโรงพยาบาล', type: 'text' },
  ];

  if (loading) return <div className="text-center py-20">กำลังโหลด...</div>;

  return (
    <div>
      <h2 className="text-2xl font-bold mb-6">ตั้งค่า</h2>

      <div className="max-w-xl space-y-4">
        {fields.map(f => (
          <div key={f.key} className="rounded-xl p-4" style={{ background: 'var(--bg-secondary)', border: '1px solid var(--border)' }}>
            <label className="block text-sm font-medium mb-2">{f.label}</label>
            <input
              type={f.type}
              value={settings[f.key] || ''}
              onChange={e => setSettings({ ...settings, [f.key]: e.target.value })}
              onBlur={e => saveSetting(f.key, e.target.value)}
              className="w-full px-3 py-2 rounded-lg text-sm"
              style={{ background: 'var(--bg-primary)', border: '1px solid var(--border)', color: 'var(--text-primary)' }}
            />
          </div>
        ))}

        {/* Test LINE button */}
        <div className="rounded-xl p-4" style={{ background: 'var(--bg-secondary)', border: '1px solid var(--border)' }}>
          <button
            onClick={testLine}
            disabled={testing}
            className="px-4 py-2 rounded-lg text-white text-sm font-medium"
            style={{ background: 'var(--accent)', border: 'none', cursor: 'pointer' }}
          >{testing ? 'กำลังทดสอบ...' : 'ทดสอบ LINE'}</button>

          {lineTestResult && (
            <div className="mt-3 p-3 rounded-lg text-sm" style={{
              background: lineTestResult.success ? '#dcfce7' : '#fee2e2',
              color: lineTestResult.success ? '#166534' : '#991b1b',
            }}>
              {lineTestResult.success ? (
                <p>{lineTestResult.data?.message || 'สำเร็จ!'}</p>
              ) : (
                <>
                  <p className="font-bold">{lineTestResult.error}</p>
                  {lineTestResult.hint && <p className="mt-1 text-xs">{lineTestResult.hint}</p>}
                </>
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
