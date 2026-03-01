import { useState } from 'react';
import { BrowserRouter, Routes, Route, NavLink } from 'react-router-dom';
import { Toaster } from 'react-hot-toast';
import Dashboard from './pages/Dashboard';
import Summary from './pages/Summary';
import Transactions from './pages/Transactions';
import Import from './pages/Import';
import Formulary from './pages/Formulary';
import Audit from './pages/Audit';
import Procurement from './pages/Procurement';
import Settings from './pages/Settings';

const NAV = [
  { to: '/',              label: 'Dashboard',    icon: '📊' },
  { to: '/summary',       label: 'สรุปยอดคงเหลือ', icon: '📦' },
  { to: '/transactions',  label: 'รายการเคลื่อนไหว', icon: '📋' },
  { to: '/import',        label: 'นำเข้าข้อมูล',    icon: '📥' },
  { to: '/formulary',     label: 'รายการยา',      icon: '💊' },
  { to: '/audit',         label: 'สุ่มนับยา',      icon: '🔍' },
  { to: '/procurement',   label: 'เบิกยา',        icon: '🛒' },
  { to: '/settings',      label: 'ตั้งค่า',        icon: '⚙️' },
];

export default function App() {
  const [dark, setDark] = useState(false);
  const [sideOpen, setSideOpen] = useState(true);

  return (
    <div className={dark ? 'dark' : ''}>
      <BrowserRouter>
        <div className="flex min-h-screen" style={{ background: 'var(--bg-primary)', color: 'var(--text-primary)' }}>
          {/* Sidebar */}
          <aside
            className={`${sideOpen ? 'w-60' : 'w-16'} transition-all duration-200 flex flex-col shrink-0`}
            style={{ background: 'var(--bg-sidebar)', minHeight: '100vh' }}
          >
            <div className="p-4 flex items-center justify-between">
              {sideOpen && (
                <h1 className="text-lg font-bold" style={{ color: 'var(--text-sidebar)' }}>
                  {import.meta.env.VITE_APP_NAME || 'Pharmacy'}
                </h1>
              )}
              <button
                onClick={() => setSideOpen(!sideOpen)}
                className="text-gray-400 hover:text-white p-1"
                style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: '1.2rem' }}
              >
                {sideOpen ? '◀' : '▶'}
              </button>
            </div>

            <nav className="flex-1 mt-2">
              {NAV.map(n => (
                <NavLink
                  key={n.to}
                  to={n.to}
                  end={n.to === '/'}
                  className={({ isActive }) =>
                    `flex items-center gap-3 px-4 py-3 text-sm transition-colors no-underline ${
                      isActive
                        ? 'bg-sky-600/20 text-sky-400 border-r-3 border-sky-400'
                        : 'hover:bg-white/5'
                    }`
                  }
                  style={{ color: 'var(--text-sidebar)', textDecoration: 'none' }}
                >
                  <span className="text-lg">{n.icon}</span>
                  {sideOpen && <span>{n.label}</span>}
                </NavLink>
              ))}
            </nav>

            <div className="p-4">
              <button
                onClick={() => setDark(!dark)}
                className="w-full py-2 px-3 rounded text-sm"
                style={{
                  background: 'rgba(255,255,255,0.1)',
                  color: 'var(--text-sidebar)',
                  border: 'none',
                  cursor: 'pointer',
                }}
              >
                {dark ? '☀️ Light' : '🌙 Dark'}
              </button>
            </div>
          </aside>

          {/* Main content */}
          <main className="flex-1 p-6 overflow-auto">
            <Routes>
              <Route path="/" element={<Dashboard />} />
              <Route path="/summary" element={<Summary />} />
              <Route path="/transactions" element={<Transactions />} />
              <Route path="/import" element={<Import />} />
              <Route path="/formulary" element={<Formulary />} />
              <Route path="/audit" element={<Audit />} />
              <Route path="/procurement" element={<Procurement />} />
              <Route path="/settings" element={<Settings />} />
            </Routes>
          </main>
        </div>

        <Toaster position="top-right" />
      </BrowserRouter>
    </div>
  );
}
