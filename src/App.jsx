import { useMemo, useState } from 'react';
import DashboardPage from './pages/DashboardPage';
import SpreadsheetPage from './pages/SpreadsheetPage';

export default function App() {
  const [activePage, setActivePage] = useState('dashboard');

  const isDashboard = activePage === 'dashboard';

  const heroTitle = useMemo(() => {
    if (isDashboard) return 'Operations dashboard';
    return 'XLSX upload';
  }, [isDashboard]);

  return (
    <div className="app-shell">
      <header className="hero">
        <div className="hero-copy">
          <p className="eyebrow">CRM / COO Workspace</p>
          <h1>{heroTitle}</h1>
        </div>

        <nav className="page-nav">
          <button
            type="button"
            className={`page-nav-button${isDashboard ? ' active' : ''}`}
            onClick={() => setActivePage('dashboard')}
          >
            Dashboard
          </button>
        </nav>
      </header>

      {activePage === 'dashboard' ? <DashboardPage /> : <SpreadsheetPage />}
    </div>
  );
}
