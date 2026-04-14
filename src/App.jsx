import { Suspense, lazy, useEffect, useState } from 'react';

const DashboardPage = lazy(() => import('./pages/DashboardPage'));
const SpreadsheetPage = lazy(() => import('./pages/SpreadsheetPage'));

const pages = {
  dashboard: {
    label: 'Dashboard',
    render: () => <DashboardPage />,
  },
  spreadsheet: {
    label: 'XLSX Upload',
    render: () => <SpreadsheetPage />,
  },
};

const getPageFromHash = () => {
  const hash = window.location.hash.replace('#', '');
  return pages[hash] ? hash : 'dashboard';
};

export default function App() {
  const [activePage, setActivePage] = useState(getPageFromHash);

  useEffect(() => {
    const syncPage = () => setActivePage(getPageFromHash());
    window.addEventListener('hashchange', syncPage);

    if (!window.location.hash) {
      window.location.hash = 'dashboard';
    }

    return () => window.removeEventListener('hashchange', syncPage);
  }, []);

  const navigate = (page) => {
    window.location.hash = page;
  };

  return (
    <div className="app-shell">
      <header className="hero">
        <div className="hero-copy">
          <p className="eyebrow">CRM / COO Workspace</p>
          <h1>Operations dashboard and spreadsheet viewer</h1>
          <p className="subtitle">
            Switch between the live dashboard and a drag-and-drop Excel page for quick inspection of
            uploaded `.xlsx` files.
          </p>
        </div>

        <nav className="page-nav" aria-label="Primary">
          {Object.entries(pages).map(([key, page]) => (
            <button
              key={key}
              type="button"
              className={`page-nav-button${activePage === key ? ' active' : ''}`}
              onClick={() => navigate(key)}
            >
              {page.label}
            </button>
          ))}
        </nav>
      </header>

      <Suspense fallback={<div className="card loading-card">Loading page...</div>}>
        {pages[activePage].render()}
      </Suspense>
    </div>
  );
}
