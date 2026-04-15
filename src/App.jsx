import DashboardPage from './pages/DashboardPage';

export default function App() {
  return (
    <div className="app-shell">
      <header className="hero">
        <div className="hero-copy">
          <p className="eyebrow">CRM / COO Workspace</p>
          <h1>Operations dashboard</h1>
        </div>
      </header>

      <DashboardPage />
    </div>
  );
}
